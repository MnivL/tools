# 项目结构说明

## 项目概览

`any2excel` 是一个轻量的 JSON/JSONL 到 Excel（`.xlsx`）转换工具。它适合一次性数据导出、日志/接口结果转表格，以及被其他 Python 脚本作为转换库调用。

当前实现不包含服务端、数据库或复杂配置，核心链路短，部署成本低。输入主要要求是对象列表：JSON 是 `List[Dict]`，JSONL 是每行一个 JSON 对象。

## 目录结构详解

### `src/any2excel`

- `excel_writer.py`：通用写入层。负责递归拍平嵌套 Mapping、按首次出现顺序收集列、序列化日期/Decimal/其他对象、调整列宽并保存文件。
- `json2excel.py`：标准 JSON 数组入口，包含 Pydantic 请求校验、文件/字符串输入、CLI 和程序化 `convert`。
- `jsonl2excel.py`：JSONL 入口，支持空行与 `#` 注释，错误消息带行号。
- `__init__.py`：包说明与版本号。

### `tests`

使用标准库 `unittest`，不额外依赖 pytest。测试通过 `openpyxl` 回读生成的工作簿，验证结果而不只验证“不报错”。

### `examples` / `data`

提供 JSON、JSONL 和对应 Excel 样例，适合作为手工试运行与文档示例输入。

## 快速导航

| 需求 | 位置 |
|---|---|
| 修改字段拍平/序列化/列宽 | `src/any2excel/excel_writer.py` |
| 修改 JSON 数组 CLI | `src/any2excel/json2excel.py` |
| 修改 JSONL 解析规则 | `src/any2excel/jsonl2excel.py` |
| 增加回归测试 | `tests/test_any2excel.py` |
| 修改依赖或命令入口 | `pyproject.toml` |
