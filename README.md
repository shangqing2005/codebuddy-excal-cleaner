# Excel 表格数据清洗工具

一个用于批量处理 Excel 文件的数据清洗工具，支持多Sheet联合查重、数据去重、格式清理等功能。

## 功能特性

- 📊 **多Sheet联合查重**：支持跨Sheet的数据去重，确保整个工作簿中数据唯一性
- 🔧 **高性能去重引擎**：使用 Go 语言编写底层去重引擎，处理大文件速度快
- 🎯 **灵活的规则配置**：支持按指定列进行去重，可配置多个去重规则
- 📝 **详细的处理日志**：生成详细的处理日志，记录每一步操作
- 🖥️ **图形化界面**：基于 Python Tkinter 的用户友好界面

## 项目结构

```
xlsx表格数据清洗工具/
├── excel_cleaner_0.7.py    # Python主程序（GUI界面）
├── xlsx_shifter/           # Go语言去重引擎
│   ├── main.go            # 核心去重逻辑
│   ├── go.mod             # Go模块配置
│   └── go.sum             # 依赖校验
├── 数据清洗.ico            # 程序图标
├── .gitignore             # Git忽略配置
└── README.md              # 项目说明文档
```

## 技术栈

- **前端界面**：Python 3.x + Tkinter
- **去重引擎**：Go 1.x
- **Excel处理**：Excelize (Go)
- **数据库支持**：pyodbc + SQLAlchemy (可选)

## 使用方法

### 环境要求

- Python 3.7+
- Go 1.16+ (仅修改Go引擎时需要)
- Windows 操作系统

### 运行方式

1. 直接运行打包好的 `.exe` 文件（如果有）
2. 或运行 Python 脚本：
   ```bash
   python excel_cleaner_0.7.py
   ```

### 去重规则说明

- 工具支持配置多个去重规则
- 每个规则指定按哪些列进行去重
- 支持跨Sheet联合查重（多规则联查）
- 重复数据会被标记并删除，只保留第一条

## 已知问题

### BUG: 阶段1/阶段2 行号对齐问题 (off-by-one)

**现象**：部分Sheet去重未生效

**原因分析**：
- 阶段1扫描使用 `getPhysicalRowNum()` 获取0-based行号
- 阶段2 XML删除使用 `extractRowNum()` 获取1-based行号（Excel标准）
- 导致 `dupSet` 中的行号与XML中的行号不匹配

**修复状态**：已尝试修复（main.go 第351行和365行），但问题可能未完全解决

**后续排查方向**：
1. 验证 `getPhysicalRowNum()` 反射获取的行号是否准确
2. 检查 `extractRowNum()` XML解析是否正确
3. 添加更详细的行号调试日志

## 开发笔记

### Git仓库

项目已上传至 GitHub：
- 仓库地址：<https://github.com/shangqing2005/codebuddy-excal-cleaner>
- 提交用户：shangqing2005 (2892911463@qq.com)

### 版本历史

- **v0.7** (2026-04-30)：初始版本提交至GitHub
  - 实现基本的Excel去重功能
  - 修复行号对齐问题（待验证）

## 注意事项

⚠️ **SQL Server 关键字保留**：SQL Server中 `timestamp` 是保留字，必须使用 `update_time` 代替作为列名。

## 待办事项

- [ ] 完全修复行号对齐BUG
- [ ] 添加单元测试
- [ ] 优化大文件处理性能
- [ ] 添加更多数据清洗功能（去空格、格式化等）
- [ ] 支持更多Excel格式（.xls等）

## 许可证

本项目为内部工具，暂不公开许可证。

---

**最后更新**：2026-04-30
**维护者**：shangqing2005