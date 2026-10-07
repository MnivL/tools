# 项目结构变更日志

## 2026-10-07

### 变更概览

- 新增流式 JSON 数组解析和 JSONL 文件迭代读取。
- 新增 Excel 行列上限分片与多工作簿输出。
- 新增临时磁盘缓存，降低超大输入的内存占用。
- CLI 增加非法输入的简洁错误处理。
- 测试从 4 个增加到 6 个，覆盖小块读取和工作簿分片。
- 更新 README、结构说明和架构说明。

### 影响范围

- `src/any2excel/excel_writer.py`：新增 `stream_rows_to_files`，并增强 `rows_to_workbook`。
- `src/any2excel/json2excel.py`：新增顶层数组增量解析。
- `src/any2excel/jsonl2excel.py`：新增逐行文件读取。
- `tests/test_any2excel.py`：新增流式与分片回归测试。

### 兼容性

原有 `convert`、`load_rows_from_file`、`iter_jsonl` API 保留。CLI 输出单文件时仍使用用户给定的 `.xlsx` 路径；只有发生工作簿分片时才生成 `.partNNN.xlsx`。
