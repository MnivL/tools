# 项目目录树

项目名称：any2excel  
分析时间：2026-10-06  
项目类型：Python 代码项目

## 目录结构

```text
any2excel/
+-- src/any2excel/
|   +-- __init__.py          包版本入口
|   +-- excel_writer.py      Excel 工作簿构建与公共序列化逻辑
|   +-- json2excel.py        JSON 数组转换器及 CLI
|   +-- jsonl2excel.py       JSONL 转换器及 CLI
+-- tests/
|   +-- test_any2excel.py    unittest 自动化测试
+-- examples/                JSON/JSONL 输入与示例输出
+-- data/                    项目样例数据
+-- docs/project-structure/  结构、架构和变更说明
|   +-- changelog.md         本次流式与分片改造记录
+-- pyproject.toml           包、依赖和 CLI 入口配置
+-- uv.lock                  uv 锁定依赖
+-- README.md                使用说明和 API 示例
```

## 关键文件说明

| 文件/目录 | 作用 |
|---|---|
| `src/any2excel/excel_writer.py` | 统一处理表头推断、字典拍平、值序列化、列宽和保存 |
| `src/any2excel/json2excel.py` | 读取 JSON 数组，提供 `json2excel` 命令和 `convert` API |
| `src/any2excel/jsonl2excel.py` | 逐行解析 JSONL，提供 `jsonl2excel` 命令和 `convert` API |
| `tests/test_any2excel.py` | 当前自动化回归测试 |
| `pyproject.toml` | Python 版本、依赖、构建后端和命令行入口 |

## 文件统计

- 核心 Python 文件：4
- 测试文件：1
- 配置/锁定文件：2
- 示例与数据文件：5
