# any2excel

将 JSON 数据（标准 JSON 列表与 JSONL）转换为 Excel 工作簿的轻量工具集。

本项目采用 [MIT License](LICENSE) 开源。

## 目录

- `src/any2excel/`：核心源码
  - `excel_writer.py`：通用写入工具（表头推断、字典拍平、值序列化等）
  - `json2excel.py`：JSON 列表（List[Dict]）转 Excel
  - `jsonl2excel.py`：JSONL（每行一个 JSON 对象）转 Excel

大文件输入会使用流式解析和临时磁盘缓存，不会把整个输入文件一次性放入内存。超过 Excel 单工作表限制时会自动拆分工作表；当工作簿工作表数量达到分片阈值时，会继续生成 `output.part002.xlsx` 等文件。

## 环境

- Python >= 3.10
- 依赖：`openpyxl >= 3.1.5`、`pydantic >= 2.7`

## 快速开始

使用 `uv` 创建虚拟环境并安装依赖：

```bash
uv sync
```

### JSON 列表 -> Excel

```bash
python -m any2excel.json2excel -i examples/list_sample.json -o examples/list_output.xlsx
```

示例输入文件：`examples/list_sample.json`

文件内容示例：

```json
[
  {"id": 1, "name": "张三", "score": 98.5},
  {"id": 2, "name": "李四", "score": 88.0}
]
```

也可以直接传入 JSON 字符串：

```bash
python -m any2excel.json2excel --data '[{"id":1,"name":"张三"}]' -o out.xlsx
```

### JSONL -> Excel

```bash
python -m any2excel.jsonl2excel -i examples/jsonl_sample.jsonl -o examples/jsonl_output.xlsx
```

示例输入文件：`examples/jsonl_sample.jsonl`

文件内容示例：

```text
{"id": 1, "name": "张三", "score": 98.5}
{"id": 2, "name": "李四", "score": 88.0}
```

直接传入 JSONL 字符串：

```bash
python -m any2excel.jsonl2excel --data '{"id":1,"name":"张三"}
{"id":2,"name":"李四"}' -o out.xlsx
```

### 输出文件名

如果输出路径没有 `.xlsx` 后缀，工具会自动补上。例如：

```bash
python -m any2excel.json2excel -i examples/list_sample.json -o examples/list_output
```

输出文件为 `examples/list_output.xlsx`。示例输出目录也可以换成任意已有写权限的目录。

## 行为说明

- 嵌套对象会被自动拍平为 `点.分隔` 的表头，例如 `{"a": {"b": 1}}` -> 列名 `a.b`。
- 字段以首次出现的顺序作为 Excel 表头列序。
- `datetime` / `date` / `Decimal` 等特殊类型会被转换为字符串后写入。
- 无法识别的对象统一通过 `json.dumps(default=str)` 序列化。
- 输出文件名若不以 `.xlsx` 结尾，将自动追加。

## 排错

| 现象 | 原因 | 处理 |
| ---- | ---- | ---- |
| `JSON 数据必须是数组` | 传入的不是 JSON 数组 | 检查 `--data` 或文件内容是否为 `[...]` |
| `第 N 行 JSON 解析失败` | JSONL 单行 JSON 非法 | 检查对应行是否漏写引号或多逗号 |
| 输出文件未生成 | 输出目录无写权限 | 确认 `-o` 路径或其父目录可写 |

## API（程序化调用）

```python
from any2excel.json2excel import convert as json_convert
from any2excel.jsonl2excel import convert as jsonl_convert

json_convert([{"a": 1}, {"a": 2}], "examples/api_json_output.xlsx")
jsonl_convert([{"a": 1}, {"a": 2}], "examples/api_jsonl_output.xlsx")
```

### 大文件与 Excel 限制

- Excel 单工作表最多 1,048,576 行、16,384 列；工具会按行和列自动分片。
- JSON 文件采用顶层数组增量解析，JSONL 采用逐行读取。
- 普通 `convert` API 适合已经在内存中的数据；CLI 适合大文件。
- 如需在程序中处理生成器，可使用：

```python
from any2excel.excel_writer import stream_rows_to_files

files = stream_rows_to_files(({"id": i} for i in range(10_000_000)), "out.xlsx")
```

返回值始终是文件路径列表；没有发生工作簿分片时列表只包含 `out.xlsx`。

### 已覆盖的异常场景

- 空数组、空行和 `#` 开头的 JSONL 注释
- JSON 顶层不是数组、数组元素不是对象
- JSONL 单行 JSON 损坏，并提示行号
- 字段在不同记录中增减、嵌套对象、列表、日期、Decimal 和其他对象
- 输出路径无 `.xlsx` 后缀、输出目录不存在
- 超过工作表行列上限时的自动分片
