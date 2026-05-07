# Excel 数据清洗工具 v1.0

一个用于批量处理 Excel 文件的跨Sheet数据去重工具，支持多Sheet联合查重、高性能去重引擎。

## 功能特性

- 📊 **跨Sheet联合查重**：多个Sheet共享同一去重上下文，确保整个工作簿中数据唯一性
- ⚡ **高性能Go引擎**：Go 语言编写底层引擎，直接操作 xlsx ZIP 内部 XML，内存占用极低
- 🎯 **灵活的规则配置**：支持按指定列去重，可配置多规则联查或单规则独立查重
- 📝 **详细的处理日志**：生成完整处理日志，包含每阶段行数变化、重复删除详情
- 🖥️ **精美GUI界面**：基于 customtkinter 的现代化图形界面

## 项目结构

```
xlsx表格数据清洗工具/
├── excel_cleaner_1.0.py     # Python主程序（GUI界面）
├── xlsx_shifter.exe          # Go语言去重引擎（编译版）
├── xlsx_shifter/             # Go语言源码目录
│   ├── main.go              # 核心去重逻辑 + XML上移引擎
│   ├── go.mod               # Go模块配置
│   └── go.sum               # 依赖校验
├── 数据清洗.ico               # 程序图标
└── README.md                # 项目说明文档
```

## 技术栈

- **前端界面**：Python 3.x + customtkinter
- **去重引擎**：Go + excelize v2
- **Excel处理**：底层XML字符串操作（高性能，无DOM解析）
- **打包分发**：PyInstaller 单文件打包

## 使用方法

### 方式一：便携版（推荐）

下载 `Excel数据清洗工具_v1.0.exe`，直接双击运行，**无需安装任何环境**。

### 方式二：源码运行

```bash
# 安装依赖
pip install customtkinter openpyxl

# 运行程序
python excel_cleaner_1.0.py
```

### 方式三：自行编译Go引擎

```bash
cd xlsx_shifter
go build -o ../xlsx_shifter.exe main.go
```

## 版本历史

### v1.0 (2026-05-07)

**修复：跨Sheet去重核心BUG**

**根因**：`scanDuplicates` 中 `dupSet[physicalRow+1]` 应改为 `dupSet[physicalRow]`。

excelize 的 `curRow`（通过 `getPhysicalRowNum` 反射获取）已经是 **1-based**。多加的 `+1` 导致所有 `dupSet` 的key比实际XML行号大1，结果：

- **应删除行 r=8 和 r=9，实际删除行 r=9 和 r=10**
- 重复的行8被保留，非重复的行9被错删
- 所有Sheet的去重均存在1行偏移

**修复内容**：
- `go.Shifter/main.go`: `dupSet[physicalRow+1]` → `dupSet[physicalRow]`
- `excel_cleaner_1.0.py`: 升级主程序，优化Go全托管模式
- 新增阶段2后验证：3次一致性扫描 + ZIP行数直读验证
- 新增行号对齐诊断（DUPKEY-DEBUG）
- Python源码支持 PyInstaller 打包

### v0.7 (2026-04-30)

- 初始版本，实现基本Excel去重功能
- Go引擎初步实现ZIP级XML行操作
- 存在跨Sheet去重失效BUG

## 已知问题

- **杀毒软件误报**：PyInstaller 打包的单文件 exe 可能被部分杀毒软件报毒，添加信任即可
- **大文件内存**：处理超大型文件（>500MB）时建议使用源码方式运行

## 注意事项

⚠️ **SQL Server 关键字保留**：SQL Server中 `timestamp` 是保留字，必须使用 `update_time` 作为列名。

## Git仓库

- **仓库地址**：https://github.com/shangqing2005/codebuddy-excal-cleaner
- **提交用户**：shangqing2005 (2892911463@qq.com)

---

**最后更新**：2026-05-07
**维护者**：shangqing2005
