// =============================================================================
// xlsx_shifter — Excel单元格去重核心处理程序（Go实现）
//
// 功能：
//   模式1 "full" : 扫描去重+上移写入一体化（推荐，内存最低）
//   模式2 "shift": 仅上移写入（兼容旧版，Python已扫描完毕）
//
// 调用方式：通过stdin JSON传递任务，stdout返回JSON结果
// 设计目标：极低内存占用、作为Python GUI的子进程调用
//
// 作者：zyq (Python工具配套)
// 版本：4.0.0 — XML直接操作实现真正上移，内存<200MB
// =============================================================================

package main

import (
	"archive/zip"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"reflect"
	"runtime"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// =============================================================================
// 通过反射获取 excelize.Rows 内部的物理行号（<row r="N"> 中的 N）
// 这是统一阶段1(scanDuplicates)和阶段2(shiftUpSheetXML)行号体系的关键
// =============================================================================

func getPhysicalRowNum(rows *excelize.Rows) int {
	v := reflect.ValueOf(rows).Elem()
	field := v.FieldByName("curRow")
	if !field.IsValid() {
		return -1 // 反射失败时返回-1
	}
	return int(field.Int())
}

// =============================================================================
// 数据结构定义
// =============================================================================

type RuleDef struct {
	SheetName string `json:"sheet_name"`
	RedCols   []int  `json:"red_cols"`
	GreenCols []int  `json:"green_cols"`
}

type FullTaskInput struct {
	FilePath   string    `json:"file_path"`
	Rules      []RuleDef `json:"rules"`
	SkipHeader bool      `json:"skip_header"`
	RuleMode   string    `json:"rule_mode"`
}

type ShiftTaskInput struct {
	FilePath   string `json:"file_path"`
	SheetName  string `json:"sheet_name"`
	ColIndices []int  `json:"col_indices"`
	DupRows    []int  `json:"dup_rows"`
}

type TaskResult struct {
	Success      bool             `json:"success"`
	ErrorMessage string           `json:"error_message,omitempty"`
	TotalRows    int              `json:"total_rows,omitempty"`
	ModifiedRows int              `json:"modified_rows,omitempty"`
	CellsChanged int              `json:"cells_changed,omitempty"`
	DupBySheet   map[string][]int `json:"dup_by_sheet,omitempty"`
	TotalDups    int              `json:"total_dups,omitempty"`
	MemoryMB     float64          `json:"memory_mb,omitempty"`
	RuleResults  []PerRuleResult  `json:"rule_results,omitempty"`
}

type PerRuleResult struct {
	SheetName  string         `json:"sheet_name"`
	Success    bool           `json:"success"`
	TotalDups  int            `json:"total_dups"`
	DupBySheet map[string]int `json:"dup_by_sheet"`
	TimeMs     float64        `json:"time_ms"`
}

type SheetMod struct {
	SheetName string       `json:"sheet_name"`
	DupSet    map[int]bool `json:"-"`
	TotalRows int          `json:"total_rows"`
	GreenCols []int        `json:"green_cols"`
	DupCount  int          `json:"dup_count"`
}

// =============================================================================
// 主函数入口
// =============================================================================

func main() {
	// 顶层panic捕获：确保任何崩溃都能输出错误信息
	defer func() {
		if r := recover(); r != nil {
			fmt.Fprintf(os.Stderr, "[FATAL] 程序崩溃: %v\n", r)
			// 输出堆栈
			buf := make([]byte, 4096)
			n := runtime.Stack(buf, false)
			fmt.Fprintf(os.Stderr, "[FATAL] 堆栈:\n%s\n", buf[:n])
			outputJSON(TaskResult{Success: false, ErrorMessage: fmt.Sprintf("程序崩溃: %v", r)})
			os.Exit(1)
		}
	}()

	rawInput := make(map[string]interface{})
	decoder := json.NewDecoder(os.Stdin)
	if err := decoder.Decode(&rawInput); err != nil {
		outputJSON(TaskResult{Success: false, ErrorMessage: fmt.Sprintf("JSON解析失败: %v", err)})
		os.Exit(1)
	}

	mode, _ := rawInput["mode"].(string)
	if mode == "" && len(rawInput) > 0 {
		if _, hasRules := rawInput["rules"]; hasRules {
			mode = "full"
		} else {
			mode = "shift"
		}
	}

	var result TaskResult
	switch mode {
	case "full":
		result = executeFullPipeline(rawInput)
	default:
		result = executeShiftOnly(rawInput)
	}
	outputJSON(result)
	if !result.Success {
		os.Exit(1)
	}
}

// =============================================================================
// 阶段1：流式扫描（使用excelize Rows API — 低内存）
// 阶段2：XML直接操作上移（ZIP内字符串改写 — 极低内存+真正上移）
// =============================================================================

func executeFullPipeline(raw map[string]interface{}) TaskResult {
	startTime := currentTimeMs()

	inputJSON, _ := json.Marshal(raw)
	var input FullTaskInput
	if err := json.Unmarshal(inputJSON, &input); err != nil {
		return failResult(fmt.Sprintf("参数解析失败: %v", err))
	}
	if input.FilePath == "" {
		return failResult("file_path 不能为空")
	}
	if len(input.Rules) == 0 {
		return failResult("rules 不能为空")
	}

	result := TaskResult{Success: true}
	ruleResults := make([]PerRuleResult, 0, len(input.Rules))

	// 根据规则模式决定去重策略
	// multi: 所有规则共享同一个seen map（跨规则全局去重）
	// single: 每条规则独立seen map（规则间互不影响）
	var globalSeen map[string]struct{}
	if input.RuleMode == "multi" {
		globalSeen = make(map[string]struct{})
	}

	fmt.Fprintf(os.Stderr, "[INFO] === 阶段1: 流式扫描开始 ===\n")
	fmt.Fprintf(os.Stderr, "[INFO] 规则模式: %s\n", input.RuleMode)

	type RuleScanResult struct {
		Rule      RuleDef
		DupSet    map[int]bool
		TotalRows int
		DupCount  int
	}
	scanResults := make([]RuleScanResult, 0, len(input.Rules))

	f, err := excelize.OpenFile(input.FilePath)
	if err != nil {
		return failResult(fmt.Sprintf("无法打开文件 '%s': %v", input.FilePath, err))
	}

	for _, rule := range input.Rules {
		// 单规则模式：每次创建独立的seen；多规则模式：共享globalSeen
		var ruleSeen map[string]struct{}
		if globalSeen != nil {
			ruleSeen = globalSeen // 多规则：复用全局map
		} else {
			ruleSeen = make(map[string]struct{}) // 单规则：独立新map
		}

		dupSet, totalRows, dups := scanDuplicates(f, rule,
			func() map[string]struct{} { return ruleSeen }, input.SkipHeader)

		if totalRows > 0 {
			scanResults = append(scanResults, RuleScanResult{
				Rule: rule, DupSet: dupSet, TotalRows: totalRows, DupCount: dups,
			})
			ruleResults = append(ruleResults, PerRuleResult{
				SheetName: rule.SheetName, Success: true,
				TotalDups: dups, DupBySheet: map[string]int{rule.SheetName: dups},
			})
			result.TotalDups += dups
			fmt.Fprintf(os.Stderr, "[INFO] 扫描「%s」完成: %d行/%d重复\n", rule.SheetName, totalRows, dups)
		}
	}

	f.Close()
	f = nil

	fmt.Fprintf(os.Stderr, "[INFO] === 扫描阶段完成，共发现%d个重复项 ===\n", result.TotalDups)
	fmt.Fprintf(os.Stderr, "[INFO] 当前内存: %.0fMB\n", getMemoryMB())

	// === 阶段2: XML直接操作上移 ===
	fmt.Fprintf(os.Stderr, "[INFO] === 阶段2: XML精确上移开始 ===\n")

	var sheetMods []SheetMod
	for _, sr := range scanResults {
		if sr.DupCount > 0 {
			sheetMods = append(sheetMods, SheetMod{
				SheetName: sr.Rule.SheetName, DupSet: sr.DupSet,
				TotalRows: sr.TotalRows, GreenCols: sr.Rule.GreenCols, DupCount: sr.DupCount,
			})
		}
	}

	if len(sheetMods) > 0 {
		batchStart := currentTimeMs()
		err := batchShiftUpXML(input.FilePath, sheetMods)
		batchMs := currentTimeMs() - batchStart

		if err != nil {
			return failResult(fmt.Sprintf("批量上移失败: %v", err))
		}
		result.ModifiedRows = len(sheetMods)
		memMB := getMemoryMB()
		fmt.Fprintf(os.Stderr, "[INFO] 上移完成! %d个Sheet, 耗时%.0fms, 内存%.0fMB\n",
			len(sheetMods), batchMs, memMB)
	}

	result.RuleResults = ruleResults
	result.MemoryMB = getMemoryMB()
	result.Success = true
	totalMs := currentTimeMs() - startTime
	fmt.Fprintf(os.Stderr, "[INFO] 全部完成! 耗时%.0fms\n", totalMs)
	return result
}

func scanDuplicates(f *excelize.File, rule RuleDef,
	seenProvider func() map[string]struct{}, skipHeader bool) (map[int]bool, int, int) {

	dupSet := make(map[int]bool)
	totalRows := 0
	rowIdx := 0
	dupCount := 0

	// 轻量级采样：每个Sheet只记录少量key用于诊断
	const sampleSize = 3
	var firstKeys []string // 前N个首次出现的key
	var lastKeys []string  // 后N个首次出现的key
	var dupKeySamples []string // 前N个被判定重复的key（含所属行号）

	// 精准关键词追踪：从环境变量读取，逗号分隔，仅匹配时输出（零开销）
	traceKeywords := make(map[string]struct{})
	if envVal := os.Getenv("DEDUP_TRACE_KEYWORDS"); envVal != "" {
		for _, kw := range strings.Split(envVal, ",") {
			kw = strings.TrimSpace(kw)
			if kw != "" {
				traceKeywords[kw] = struct{}{}
			}
		}
	}
	hasTrace := len(traceKeywords) > 0
	traceLog := func(row int, rawVal, normVal, key string, isDup bool, reason string) {
		fmt.Fprintf(os.Stderr, "[TRACE] Sheet「%s」行%d: %s | dup=%v (%s)\n",
			rule.SheetName, row+1, reason, isDup, key)
		fmt.Fprintf(os.Stderr, "[TRACE]   原始值: [%s]\n", rawVal)
		fmt.Fprintf(os.Stderr, "[TRACE]   规范值: [%s]\n", normVal)
		fmt.Fprintf(os.Stderr, "[TRACE]   hex: %x\n", key)
	}

	idx, _ := f.GetSheetIndex(rule.SheetName)
	if idx == -1 {
		fmt.Fprintf(os.Stderr, "[WARN] Sheet '%s' 不存在\n", rule.SheetName)
		return dupSet, 0, 0
	}

	fmt.Fprintf(os.Stderr, "[INFO] 开始扫描Sheet「%s」...\n", rule.SheetName)

	rows, err := f.Rows(rule.SheetName)
	if err != nil {
		fmt.Fprintf(os.Stderr, "[ERR] 打开失败 '%s': %v\n", rule.SheetName, err)
		return dupSet, 0, 0
	}
	defer rows.Close()

	seen := seenProvider()

	for rows.Next() {
		// 获取物理行号（即 <row r="N"> 中的 N），用于统一阶段2的匹配
		physicalRow := getPhysicalRowNum(rows)

		if skipHeader && rowIdx == 0 {
			rowIdx++
			continue
		}

		rowData, err := rows.Columns()
		if err != nil || len(rowData) == 0 {
			rowIdx++
			if rowIdx > totalRows {
				totalRows = rowIdx
			}
			continue
		}
		totalRows = rowIdx + 1

		if len(rowData) <= maxOfInts(rule.RedCols...) {
			rowIdx++
			continue
		}

		key := buildDedupKey(rowData, rule.RedCols)
		if key == "" {
			rowIdx++
			continue
		}

		// 精准追踪：仅在匹配到关键词时记录详情
		if hasTrace {
			rawVal := rowData[rule.RedCols[0]]
			normVal := normalizeValue(rawVal)
			matched := false
			for kw := range traceKeywords {
				if strings.Contains(normVal, kw) || strings.Contains(rawVal, kw) {
					matched = true
					break
				}
			}
			if matched {
				if _, exists := seen[key]; exists {
					dupSet[physicalRow+1] = true // +1 转为1-based，与XML r="N" 对齐
					dupCount++
					traceLog(physicalRow, rawVal, normVal, key, true, "重复命中")
				} else {
					seen[key] = struct{}{}
					traceLog(physicalRow, rawVal, normVal, key, false, "首次出现")
				}
				rowIdx++
				// 跳过后面的通用采样逻辑，避免重复计数
				continue
			}
		}

		if _, exists := seen[key]; exists {
			dupSet[physicalRow+1] = true // +1 转为1-based，与XML r="N" 对齐
			dupCount++
			// 采集前N个重复key样本
			if len(dupKeySamples) < sampleSize {
				dupKeySamples = append(dupKeySamples,
					fmt.Sprintf("物理行r=%d:%s", physicalRow, key))
			}
		} else {
			seen[key] = struct{}{}
			// 采集首次出现的key（前N和后N）
			if len(firstKeys) < sampleSize {
				firstKeys = append(firstKeys, key)
			}
			// 始终维护后N个（覆盖方式）
			if len(lastKeys) < sampleSize {
				lastKeys = append(lastKeys, key)
			} else {
				lastKeys[len(lastKeys)-1] = key
				// 用滑动窗口方式保持最后N个
				if len(lastKeys) >= sampleSize {
					lastKeys = lastKeys[1:]
					lastKeys = append(lastKeys, key)
				}
			}
		}
		rowIdx++

		if rowIdx%10000 == 0 {
			fmt.Fprintf(os.Stderr, "[PROGRESS] 「%s」已扫描 %d 行...\n", rule.SheetName, rowIdx)
		}
	}

	// ====== 轻量级诊断输出（每个Sheet仅此一次） ======
	fmt.Fprintf(os.Stderr, "[SAMPLE] === Sheet「%s」扫描摘要 ===\n", rule.SheetName)
	fmt.Fprintf(os.Stderr, "[SAMPLE]   总行数:%d | 发现重复:%d | seen map大小:%d\n",
		totalRows, dupCount, len(seen))
	fmt.Fprintf(os.Stderr, "[SAMPLE]   前%d个新key: %v\n", len(firstKeys), firstKeys)
	fmt.Fprintf(os.Stderr, "[SAMPLE]   后%d个新key: %v\n", len(lastKeys), lastKeys)
	if len(dupKeySamples) > 0 {
		fmt.Fprintf(os.Stderr, "[SAMPLE]   前%d个重复: %v\n", len(dupKeySamples), dupKeySamples)
	} else {
		fmt.Fprintf(os.Stderr, "[SAMPLE]   重复样本: 无\n")
	}
	// 显示key中是否包含可疑不可见字符（hex dump第一个key）
	if len(firstKeys) > 0 {
		fmt.Fprintf(os.Stderr, "[SAMPLE]   首key hex: %x\n", firstKeys[0])
	}
	fmt.Fprintf(os.Stderr, "[SAMPLE] === 摘要结束 ===\n")

	return dupSet, totalRows, dupCount
}

// =============================================================================
// 模式2：仅上移（兼容旧版）
// =============================================================================

func executeShiftOnly(raw map[string]interface{}) TaskResult {
	inputJSON, _ := json.Marshal(raw)
	var input ShiftTaskInput
	if err := json.Unmarshal(inputJSON, &input); err != nil {
		return failResult(fmt.Sprintf("参数解析失败: %v", err))
	}
	if input.FilePath == "" || input.SheetName == "" || len(input.ColIndices) == 0 {
		return failResult("缺少必要参数")
	}

	f, err := excelize.OpenFile(input.FilePath)
	if err != nil {
		return failResult(fmt.Sprintf("无法打开文件: %v", err))
	}
	defer f.Close()

	idx, _ := f.GetSheetIndex(input.SheetName)
	if idx == -1 {
		return failResult(fmt.Sprintf("工作表 '%s' 不存在", input.SheetName))
	}

	rows, err := f.GetRows(input.SheetName)
	if err != nil {
		return failResult(fmt.Sprintf("读取失败: %v", err))
	}
	totalRows := len(rows)

	dupSet := make(map[int]bool)
	for _, r := range input.DupRows {
		dupSet[r] = true
	}

	cellsChanged := 0
	for _, colIdx := range input.ColIndices {
		colLetter := colNumToLetter(colIdx + 1)
		newValues := make([]string, 0, totalRows-len(dupSet))
		for rowIdx := 0; rowIdx < totalRows; rowIdx++ {
			if dupSet[rowIdx] {
				continue
			}
			var val string
			if colIdx < len(rows[rowIdx]) {
				val = rows[rowIdx][colIdx]
			}
			newValues = append(newValues, val)
		}
		for rowIdx := 0; rowIdx < totalRows; rowIdx++ {
			cellRef := fmt.Sprintf("%s%d", colLetter, rowIdx+1)
			if rowIdx < len(newValues) {
				f.SetCellValue(input.SheetName, cellRef, newValues[rowIdx])
			} else {
				f.SetCellValue(input.SheetName, cellRef, "")
			}
			cellsChanged++
		}
	}
	if err := f.Save(); err != nil {
		return failResult(fmt.Sprintf("保存失败: %v", err))
	}

	return TaskResult{Success: true, TotalRows: totalRows, ModifiedRows: len(input.DupRows), CellsChanged: cellsChanged}
}

// =============================================================================
// 辅助函数
// =============================================================================

// normalizeValue 清洗单元格值，去除各种不可见/空白字符，确保跨Sheet匹配一致
func normalizeValue(s string) string {
	// 去除BOM
	s = strings.ReplaceAll(s, "\ufeff", "")
	// 去除零宽字符
	s = strings.ReplaceAll(s, "\u200b", "")
	s = strings.ReplaceAll(s, "\u200c", "")
	s = strings.ReplaceAll(s, "\u200d", "")
	// 将不间断空格(NBSP)、全角空格等替换为普通空格
	s = strings.ReplaceAll(s, "\xa0", " ")
	s = strings.ReplaceAll(s, "\u3000", " ")
	// 去除首尾所有Unicode空白(包括普通空格、制表符、换行等)
	s = strings.TrimSpace(s)
	// 压缩中间连续空白为单个空格
	s = strings.Join(strings.Fields(s), " ")
	return s
}

func buildDedupKey(rowData []string, cols []int) string {
	parts := make([]string, 0, len(cols))
	for _, colIdx := range cols {
		if colIdx >= len(rowData) {
			return ""
		}
		val := rowData[colIdx]
		if val == "" {
			return ""
		}
		s := normalizeValue(val)
		if s == "" {
			return ""
		}
		parts = append(parts, s)
	}
	if len(parts) == 0 {
		return ""
	}
	return strings.Join(parts, "\x00")
}

func maxOfInts(nums ...int) int {
	maxV := -1
	for _, n := range nums {
		if n > maxV {
			maxV = n
		}
	}
	return maxV
}

func colNumToLetter(n int) string {
	var result []byte
	for n > 0 {
		n--
		result = append([]byte{byte('A' + n%26)}, result...)
		n /= 26
	}
	return string(result)
}

func outputJSON(result TaskResult) {
	encoder := json.NewEncoder(os.Stdout)
	encoder.SetIndent("", "  ")
	if err := encoder.Encode(result); err != nil {
		fmt.Fprintln(os.Stderr, "JSON输出失败:", err)
	}
}

func failResult(msg string) TaskResult { return TaskResult{Success: false, ErrorMessage: msg} }

func getMemoryMB() float64 {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	return float64(m.Alloc) / 1024 / 1024
}

func currentTimeMs() float64 { return float64(time.Now().UnixNano()) / 1e6 }

// =============================================================================
// 阶段2核心：ZIP/XML 直接操作 — 真正上移 + 极低内存
//
// 算法（对每个 green 列）：
//   维护 outputRow = 1
//   遍历 XML 中该列的所有 cell：
//     - 非 → 改写 r="D12" 为 r="D{outputRow}"，outputRow++  （数据上移！）
//     - 是 → 清空值（outputRow 不变，产生空缺被上方填充）
//     - 超 → 清空值
//
// 内存：单次 strings.Builder 遍历 O(n)，不加载 DOM，80MB文件约50~150MB
// 准确度：只修改行号引用和清空值，保留所有原始属性/格式/SST引用不变
// =============================================================================

// batchShiftUpXML 批量处理多个 Sheet 的真正上移
func batchShiftUpXML(xlsxPath string, mods []SheetMod) error {
	modMap := make(map[string]SheetMod)
	for _, m := range mods {
		modMap[m.SheetName] = m
	}

	reader, err := zip.OpenReader(xlsxPath)
	if err != nil {
		return fmt.Errorf("无法打开ZIP: %v", err)
	}
	sheetFileMap := buildSheetFileMap(&reader.Reader)
	reader.Close()
	if sheetFileMap == nil {
		return fmt.Errorf("解析sheet映射失败")
	}

	fmt.Fprintf(os.Stderr, "[INFO] 发现%d个sheet, 需修改%d个\n", len(sheetFileMap), len(mods))

	tmpPath := xlsxPath + ".tmp"
	destFile, err := os.Create(tmpPath)
	if err != nil {
		return fmt.Errorf("创建临时文件: %v", err)
	}
	zipWriter := zip.NewWriter(destFile)

	reader, err = zip.OpenReader(xlsxPath)
	if err != nil {
		return fmt.Errorf("重新打开ZIP: %v", err)
	}
	defer reader.Close()

	for _, srcFile := range reader.File {
		var targetMod *SheetMod
		for sName, zPath := range sheetFileMap {
			if srcFile.Name == zPath {
				if mod, ok := modMap[sName]; ok && mod.DupCount > 0 {
					targetMod = &mod
				}
				break
			}
		}

		writer, err := zipWriter.Create(srcFile.Name)
		if err != nil {
			rc, _ := srcFile.Open()
			if rc != nil {
				io.Copy(writer, rc)
				rc.Close()
			}
			continue
		}

		if targetMod != nil {
			rc, _ := srcFile.Open()
			xmlData, _ := io.ReadAll(rc)
			rc.Close()

			fmt.Fprintf(os.Stderr, "[INFO] 正在处理「%s」(%s): %d重复×%d列...\n",
				targetMod.SheetName, srcFile.Name, targetMod.DupCount, len(targetMod.GreenCols))

			modifiedXML := shiftUpSheetXML(xmlData, *targetMod)
			writer.Write(modifiedXML)

			memMB := getMemoryMB()
			fmt.Fprintf(os.Stderr, "[INFO]   「%s」完成, 内存%.0fMB\n", targetMod.SheetName, memMB)
			if memMB > 1500 {
				return fmt.Errorf("内存过高(%.0fMB)，建议拆分文件", memMB)
			}
		} else {
			rc, _ := srcFile.Open()
			if rc != nil {
				io.Copy(writer, rc)
				rc.Close()
			}
		}
	}

	zipWriter.Close()
	destFile.Close()
	reader.Close()
	os.Remove(xlsxPath)
	if err := os.Rename(tmpPath, xlsxPath); err != nil {
		return fmt.Errorf("替换文件: %v", err)
	}
	return nil
}

// buildSheetFileMap 从 workbook.xml 提取 sheet名→XML路径映射
func buildSheetFileMap(reader *zip.Reader) map[string]string {
	result := make(map[string]string)
	for _, f := range reader.File {
		if f.Name == "xl/workbook.xml" || f.Name == "xl\\workbook.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			extractSheets(string(data), result)
			break
		}
	}
	return result
}

// extractSheets 解析 workbook.xml 填充映射
func extractSheets(wb string, result map[string]string) {
	pat := `<sheet name="`
	i := 0
	for {
		s := strings.Index(wb[i:], pat)
		if s < 0 {
			break
		}
		s += i + len(pat)
		e := strings.Index(wb[s:], `"`)
		if e < 0 {
			break
		}
		name := wb[s : s+e]
		ip := strings.Index(wb[s+e:], `sheetId="`)
		if ip < 0 {
			break
		}
		ip += s + e + len(`sheetId="`)
		ie := strings.Index(wb[ip:], `"`)
		if ie < 0 {
			break
		}
		id := 0
		fmt.Sscanf(wb[ip:ip+ie], "%d", &id)
		result[name] = fmt.Sprintf("xl/worksheets/sheet%d.xml", id)
		i = s + e
	}
}

// shiftUpSheetXML 对单个 sheet XML 执行真正的整行上移
// 算法：以 <row> 为单位切分，同步改 <row r="N"> + 内部所有 <c r="XN">
//
//	重复行的整个 <row> 块删除 → 和 Python ElementTree 移动效果一致
func shiftUpSheetXML(xmlData []byte, mod SheetMod) []byte {
	content := string(xmlData)
	
	fmt.Fprintf(os.Stderr, "[START] 开始处理SheetXML，数据大小=%d字节, 重复数=%d\n", len(content), mod.DupCount)

	// 定位 <sheetData> ... </sheetData> 区域
	sdStart := strings.Index(content, `<sheetData`)
	if sdStart < 0 {
		return xmlData
	}
	sdEnd := strings.Index(content, `</sheetData>`)
	if sdEnd < 0 {
		return xmlData
	}
	sdClose := strings.Index(content[sdStart:], `>`) + sdStart

	header := content[:sdClose+1]      // <sheetData ...>
	footer := content[sdEnd:]          // </sheetData>
	body := content[sdClose+1 : sdEnd] // 中间的 rows

	outputRow := 1
	var builder strings.Builder
	builder.Grow(len(content))
	builder.WriteString(header)

	rowCount := 0
	skippedCount := 0
	keptCount := 0
	consecRecoveries := 0    // 连续恢复次数
	lastRecoverPos := 0       // 上次恢复位置（检测微小步进）
	const maxConsecRecoveries = 15 // 连续恢复上限
	const minRecoverStep = 200     // 恢复最小步进(字节)

	// 按行遍历
	for offset := 0; offset < len(body); {
		rs := strings.Index(body[offset:], `<row`)
		if rs < 0 {
			builder.WriteString(body[offset:])
			break
		}
		rs += offset

		re := findTagEnd(body, rs)
		if re < 0 || re >= len(body) {
			fmt.Fprintf(os.Stderr, "[WARN] findTagEnd失败(pos=%d,re=%d)，放弃剩余%d字节\n",
				rs, re, len(body)-rs)
			builder.WriteString(body[rs:])
			break
		}

		rCloseIdx := re // findTagEnd 返回的就是闭合 > 的位置

		endRow := strings.Index(body[rCloseIdx:], `</row>`)
		if endRow < 0 {
			// 检测是否为自闭合 <row ... />
			// findTagEnd返回的rCloseIdx指向'>'字符
			// 对于自闭合标签 <row ... />，'>'的前一个字符应该是'/'
			isSelfClosing := false
			if rCloseIdx > rs && body[rCloseIdx-1] == '/' {
				// 标准：<row ... />
				isSelfClosing = true
			} else if rCloseIdx+1 < len(body) && body[rCloseIdx+1] == '/' {
				// 非标准但可能的格式：<row ... ></（不太可能但防御性处理）
				isSelfClosing = true
				rCloseIdx++ // 跳过'/'
			}
			
			if isSelfClosing {
				scEnd := rCloseIdx + 1  // 包含'>'
				if scEnd > len(body) { scEnd = len(body) }
				builder.WriteString(body[rs:scEnd])
				offset = scEnd
				consecRecoveries = 0 // 正常路径，重置
				continue
			}

			// 确实不是自闭合 → 进入恢复逻辑
			if consecRecoveries <= 2 {
				fmt.Fprintf(os.Stderr, "[DIAG] 非自闭合格式: rs=%d rCloseIdx=%d 前后=[%s]\n",
					rs, rCloseIdx,
					truncateStr(strings.ReplaceAll(strings.ReplaceAll(
						body[max(rs,rCloseIdx-5):min(rCloseIdx+15,len(body))],
						"\n", "\\n"), "\t", "\\t"), 60))
			}

			// ===== 恢复逻辑（带防死循环保护）=====
			consecRecoveries++

			// 日志节流：只在前3次和每第10次输出详情
			if consecRecoveries <= 3 || consecRecoveries%10 == 0 {
				dumpLen := rCloseIdx - rs + 300
				maxDump := len(body) - rs
				if dumpLen > maxDump || dumpLen < 0 { dumpLen = maxDump }
				fmt.Fprintf(os.Stderr, "[WARN] 行%d附近找不到</row>(第%d次连续,rCloseIdx=%d)，<row>: [%s]\n",
					rowCount+1, consecRecoveries, rCloseIdx,
					truncateStr(strings.ReplaceAll(strings.ReplaceAll(body[rs:rs+dumpLen], "\n", "\\n"), "\t", "\\t"), 200))
			}

			// 连续恢复超限 → 放弃剩余
			if consecRecoveries > maxConsecRecoveries {
				fmt.Fprintf(os.Stderr, "[WARN] 连续恢复超限(%d次)，原样保留剩余%d字节\n",
					consecRecoveries, len(body)-rCloseIdx)
				// 输出残留（<row标签之前的内容）
				if rs > offset {
					builder.WriteString(body[offset:rs])
				}
				builder.WriteString(body[rCloseIdx:])
				break
			}

			// 输出残留
			if rs > offset {
				builder.WriteString(body[offset:rs])
			}

			// 查找下一个 <row 或 </sheetData>
			nextRow := strings.Index(body[rCloseIdx:], `<row`)
			nextSD := strings.Index(body[rCloseIdx:], `</sheetData>`)
			var skipTo int
			if nextRow >= 0 && (nextSD < 0 || nextRow < nextSD) {
				skipTo = rCloseIdx + nextRow
				// ★ 最小步进保护：防止微小移动死循环
				stepSize := skipTo - lastRecoverPos
				if lastRecoverPos > 0 && stepSize < minRecoverStep && nextRow > 0 {
					// 步进太小 → 强制用更大范围搜索（跳过可能的假匹配）
					widerSearch := body[rCloseIdx+minRecoverStep:]
					widerHit := strings.Index(widerSearch, `<row`)
					if widerHit >= 0 {
						skipTo = rCloseIdx + minRecoverStep + widerHit
						if consecRecoveries <= 3 || consecRecoveries%10 == 0 {
							fmt.Fprintf(os.Stderr, "[WARN]   → 步进过小(%d)，强制跳到pos=%d\n", stepSize, skipTo)
						}
					}
				}
				if consecRecoveries <= 3 || consecRecoveries%10 == 0 {
					fmt.Fprintf(os.Stderr, "[WARN]   → 跳到下一个<row>(pos=%d)\n", skipTo)
				}
			} else if nextSD >= 0 {
				skipTo = rCloseIdx + nextSD
				if skipTo > len(body) { skipTo = len(body) }
				fmt.Fprintf(os.Stderr, "[WARN]   → 到达</sheetData>(pos=%d)，结束\n", skipTo)
				builder.WriteString(body[rCloseIdx:skipTo])
				break
			} else {
				fmt.Fprintf(os.Stderr, "[WARN]   → 无法恢复，保留剩余%d字节\n", len(body)-rCloseIdx)
				builder.WriteString(body[rCloseIdx:])
				break
			}

			lastRecoverPos = skipTo
			offset = skipTo
			continue
		}
		// 正常找到 </row> → 重置连续恢复计数
		consecRecoveries = 0
		lastRecoverPos = 0
		endRow += rCloseIdx + len(`</row>`)

		// 安全clamp：确保endRow不超出body范围
		if endRow > len(body) {
			fmt.Fprintf(os.Stderr, "[WARN] 行%d endRow=%d超出body长度%d，截断\n",
				rowCount+1, endRow, len(body))
			endRow = len(body)
		}

		fullRowBlock := body[rs:endRow]

		rowNum := extractRowNum(body, rs, re)
		nxt := endRow

		// 直接用物理行号查 DupSet（阶段1已改为用物理行号作为key）
		if mod.DupSet[rowNum] {
			// 重复行：整行删除（不输出任何内容）
			skippedCount++
			rowCount++ // 也计入rowCount以正确触发progress日志
		} else {
			// 非重复行：重编号 <row r="N"> + 内部所有 <c r="?N">
			builder.WriteString(renumberRow(fullRowBlock, rowNum, outputRow))
			outputRow++
			keptCount++
			rowCount++
		}

		offset = nxt
		if rowCount%5000 == 0 && rowCount > 0 {
			fmt.Fprintf(os.Stderr, "[PROGRESS]   已处理%d行(保留%d/跳过%d重复), 内存%.0fMB\n",
				rowCount, keptCount, skippedCount, getMemoryMB())
		}
	}

	builder.WriteString(footer)

	resultSize := builder.Len()
	originalSize := len(content)
	resultStr := builder.String()

	// ====== 诊断：全量XML结构完整性检查 ======
	totalRows := strings.Count(resultStr, `<row`)
	totalCells := strings.Count(resultStr, `<c `)
	totalValues := strings.Count(resultStr, `<v>`)
	hasSheetDataOpen := strings.Contains(resultStr, `<sheetData`)
	hasSheetDataClose := strings.Contains(resultStr, `</sheetData>`)
	hasCellR := strings.Count(resultStr, `<c r="`)    // 正确的cell引用格式
	brokenRefs := strings.Count(resultStr, `"<c r="`) // 错误格式（前缀丢失）
	sheetDataPos := strings.Index(resultStr, `<sheetData`)

	fmt.Fprintf(os.Stderr, "[DIAG] === XML结构诊断 ===\n")
	fmt.Fprintf(os.Stderr, "[DIAG] 总大小: %.0fKB(原始%.0fKB), <row总数: %d, <c总数: %d, <v>值: %d\n",
		float64(resultSize)/1024, float64(originalSize)/1024, totalRows, totalCells, totalValues)
	fmt.Fprintf(os.Stderr, "[DIAG] <sheetData>: %v, </sheetData>: %v\n", hasSheetDataOpen, hasSheetDataClose)
	fmt.Fprintf(os.Stderr, "[DIAG] <c r=\" 格式正确: %d, 异常(缺失<): %d\n", hasCellR, brokenRefs)

	// 提取前3个完整行块用于对比
	if sheetDataPos >= 0 {
		fmt.Fprintf(os.Stderr, "[DIAG] --- 前3个完整行块 ---\n")
		bodyStart := sheetDataPos + strings.Index(resultStr[sheetDataPos:], `>`) + 1
		offset := bodyStart
		for rowIdx := 0; rowIdx < 3 && offset < len(resultStr); rowIdx++ {
			rs := strings.Index(resultStr[offset:], `<row`)
			if rs < 0 {
				break
			}
			rs += offset
			reEnd := findTagEnd(resultStr, rs)
			if reEnd < 0 {
				break
			}
			rowClose := strings.Index(resultStr[reEnd:], `</row>`)
			if rowClose < 0 {
				break
			}
			rowEnd := reEnd + rowClose + len(`</row>`)
			rowBlock := resultStr[rs:rowEnd]
			// 截断显示（每行最多500字符）
			display := rowBlock
			if len(display) > 500 {
				display = display[:497] + "..."
			}
			fmt.Fprintf(os.Stderr, "[DIAG] 行%d(%d字符): %s\n", rowIdx+1, len(rowBlock), display)
			offset = rowEnd
		}

		// 也显示原始XML的前3个行块作为对照
		fmt.Fprintf(os.Stderr, "[DIAG] --- 原始前3个行块(对照用) ---\n")
		origBodyStart := sdClose + 1
		origOffset := origBodyStart
		for rowIdx := 0; rowIdx < 3 && origOffset < len(content); rowIdx++ {
			rs := strings.Index(content[origOffset:], `<row`)
			if rs < 0 {
				break
			}
			rs += origOffset
			reEnd := findTagEnd(content, rs)
			if reEnd < 0 {
				break
			}
			rowClose := strings.Index(content[reEnd:], `</row>`)
			if rowClose < 0 {
				break
			}
			rowEnd := reEnd + rowClose + len(`</row>`)
			rowBlock := content[rs:rowEnd]
			display := rowBlock
			if len(display) > 500 {
				display = display[:497] + "..."
			}
			fmt.Fprintf(os.Stderr, "[DIAG] 原始行%d(原r=%d, %d字符): %s\n",
				rowIdx+1, extractRowNum(content, rs, reEnd), len(rowBlock), display)
			origOffset = rowEnd
		}

		fmt.Fprintf(os.Stderr, "[DIAG] === 结束 ===\n")
	} else {
		fmt.Fprintf(os.Stderr, "[DIAG] ⚠ 未找到<sheetData>标签！\n")
	}

	if originalSize > 1000 && resultSize < originalSize/10 {
		fmt.Fprintf(os.Stderr, "[WARN]   ⚠ 输出仅为原始的 %.1f%%！可能异常\n",
			float64(resultSize)*100/float64(originalSize))
	}

	return []byte(builder.String())
}

// extractRowNum 从 <row r="12" ...> 提取行号
func extractRowNum(body string, tagStart, tagAttrEnd int) int {
	pat := `r="`
	pi := strings.Index(body[tagStart:tagAttrEnd], pat)
	if pi < 0 {
		return 0
	}
	pi += tagStart + len(pat)
	pe := strings.Index(body[pi:], `"`)
	if pe < 0 {
		return 0
	}
	n := 0
	fmt.Sscanf(body[pi:pi+pe], "%d", &n)
	return n
}

// truncateStr 截断字符串到指定最大长度，超出部分用"..."代替
func truncateStr(s string, maxLen int) string {
	if len(s) <= maxLen {
		return s
	}
	if maxLen > 3 {
		return s[:maxLen-3] + "..."
	}
	return s[:maxLen]
}

// findTagEnd 找到 <row ... > 的闭合 >（考虑属性中的引号）
func findTagEnd(s string, start int) int {
	i := start + 4 // 跳过 "<row"
	inQuote := false
	for i < len(s) {
		c := s[i]
		if c == '"' {
			inQuote = !inQuote
		} else if c == '>' && !inQuote {
			return i
		}
		i++
	}
	return -1
}

// renumberRow 对一个完整的 <row ...>块</row> 重编号：
//
//	<row r="12" ...> → <row r="5" ...>
//	内部所有 <c r="D12"> → <c r="D5">
func renumberRow(rowBlock string, oldRow, newRow int) string {
	oldStr := strconv.Itoa(oldRow)
	newStr := strconv.Itoa(newRow)

	var b strings.Builder
	b.Grow(len(rowBlock))

	// 第一处替换：<row r="old">
	ri := strings.Index(rowBlock, `r="`)
	if ri >= 0 {
		b.WriteString(rowBlock[:ri+3])
		b.WriteString(newStr)
		rest := rowBlock[ri+3+len(oldStr):]

		// 后续所有 cell 引用中出现的 oldRow 数字
		for off := 0; off < len(rest); {
			ci := strings.Index(rest[off:], `<c r="`)
			if ci < 0 {
				b.WriteString(rest[off:])
				break
			}
			ci += off
			b.WriteString(rest[off:ci]) // <c r=" 前面的内容
			b.WriteString(`<c r="`)     // ← 补回 cell 标签前缀（之前漏掉了！）
			refStart := ci + 6          // " 之后

			qi := strings.Index(rest[refStart:], `"`)
			if qi < 0 {
				b.WriteString(rest[ci:])
				break
			}
			refStr := rest[refStart : refStart+qi] // D12 这种引用

			colLetter := extract(refStr)
			newRef := colLetter + newStr
			b.WriteString(newRef) // 写入新引用
			b.WriteByte('"')

			off = refStart + qi + 1 // 跳过旧引用的 "
		}
	} else {
		b.WriteString(rowBlock)
	}
	return b.String()
}

// extract 从 cell 引用提取列字母
func extract(ref string) string {
	i := 0
	for i < len(ref) && ref[i] >= 'A' && ref[i] <= 'Z' {
		i++
	}
	return ref[:i]
}

// clearValue 安全清空 cell 的值部分，保留结构
func clearValue(cell string) string {
	// 富文本 <is> 内多 <t>
	if strings.Contains(cell, "<is>") {
		out := cell
		pos := 0
		for {
			ti := strings.Index(out[pos:], "<t")
			if ti < 0 {
				break
			}
			ti += pos
			gi := strings.Index(out[ti:], ">")
			if gi < 0 {
				break
			}
			cs := ti + gi + 1
			te := strings.Index(out[cs:], "</t>")
			if te < 0 {
				break
			}
			out = out[:cs] + out[cs+te:]
			pos = ti + 3
		}
		return out
	}
	// 普通 <v>
	vi := strings.Index(cell, "<v")
	if vi >= 0 {
		ve := strings.Index(cell[vi:], "</v>")
		if ve > 0 {
			gi := strings.Index(cell[vi:vi+ve], ">")
			if gi >= 0 {
				return cell[:vi+gi+1] + cell[vi+ve:]
			}
		}
	}
	// 普通 <t>
	ti := strings.Index(cell, "<t")
	if ti >= 0 {
		te := strings.Index(cell[ti:], "</t>")
		if te > 0 {
			gi := strings.Index(cell[ti:ti+te], ">")
			if gi >= 0 {
				return cell[:ti+gi+1] + cell[ti+te:]
			}
		}
	}
	return cell
}
