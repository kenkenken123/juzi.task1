# 项目详情记录

## 项目基本信息

**项目名称**: juzi.task1  
**项目类型**: .NET 8.0 命令行应用程序  
**创建时间**: 2024-12-XX  
**当前状态**: 开发中

## 项目描述

根据 Excel 模板数据（12月办事处日常费用预算财务.xlsx）自动生成 Word 文档（类似 1天河.doc 格式），用于日常费用预算财务报表生成。

## 技术栈

- **框架**: .NET 8.0
- **语言**: C#
- **项目类型**: Console Application
- **开发工具**: Visual Studio / Rider / VS Code
- **Excel 处理**: EPPlus 7.0.0
- **Word 处理**: DocumentFormat.OpenXml 3.0.0

## 项目结构

```
juzi.task1/
├── juzi.task1.csproj    # 项目文件
├── Program.cs            # 程序入口（菜单驱动）
├── AGENTS.MD            # 项目详情记录（本文件）
├── .gitignore           # Git 忽略规则
├── Features/            # 功能模块
│   └── DocumentGenerationFeature.cs  # 文档生成功能
├── Services/            # 服务类
│   ├── ExcelReader.cs   # Excel 读取服务
│   ├── WordGenerator.cs # Word 生成服务（旧版，已废弃）
│   └── WordTemplateProcessor.cs  # Word 模板处理服务
├── data/                # 数据目录
│   ├── 12月办事处日常费用预算财务.xlsx
│   └── 1天河.doc
└── output/              # 输出目录（生成的文件）
```

## 核心功能

- [x] 读取 Excel 文件的所有工作表（Sheet）
- [x] 根据 Sheet 名称生成对应的 Word 文档（文件名 = Sheet 名称）
- [x] 解析 Excel 多级表头结构（合并单元格）
- [x] 生成格式化的 Word 表格文档（包含多级表头、数据行、合计行）
- [x] 正确处理数字格式和空值

## 重要决策记录

### 决策 #001 - 选择 Excel/Word 处理库
**日期**: 2024-12-XX  
**背景**: 需要处理 Excel 和 Word 文件，需要选择合适的库  
**方案**: 
- Excel: EPPlus 7.0.0（免费版，.NET 8 支持良好）
- Word: DocumentFormat.OpenXml 3.0.0（Microsoft 官方，免费）  
**原因**: 
- EPPlus 免费版功能足够，API 简洁
- DocumentFormat.OpenXml 是官方库，稳定可靠，无需商业授权
- 两者都是 NuGet 包，易于集成  
**影响**: 项目依赖两个外部库，但都是成熟稳定的选择

---

## 开发日志

| 日期 | 内容 | 备注 |
|------|------|------|
| 2024-12-XX | 项目初始化，创建 .NET 命令行项目结构 | - |
| 2024-12-XX | 实现任务 #001：日常费用预算财务文档生成 | 添加 ExcelReader 和 WordGenerator 服务 |
| 2024-12-XX | 优化实现：根据 Sheet 名称生成文档 | 支持多 Sheet，每个 Sheet 生成一个 Word 文档 |
| 2024-12-XX | 重构：使用模板填充方式生成文档 | 添加 WordTemplateProcessor，基于模板文件填充数据 |
| 2024-12-XX | 重构：功能模块化，菜单驱动 | Program.cs 改为菜单驱动，功能实现封装到 Features 目录 |

---

## 待办事项

- [ ] 测试 Excel 文件读取功能（多 Sheet 支持）
- [ ] 测试 Word 文档生成功能（模板填充）
- [ ] 验证生成的 Word 文档格式是否符合要求
- [ ] 处理 .doc 格式模板文件（当前仅支持 .docx）
- [ ] 添加更多功能模块（如数据统计、报表导出等）
- [ ] 优化错误处理和日志记录
- [ ] 处理特殊字符和文件名非法字符

---

## 参考资料

- [相关文档链接]
- [技术文档]
