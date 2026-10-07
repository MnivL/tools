# 项目架构说明

## 架构概览

项目采用“输入适配器 + 共享写入层”的小型分层结构：

```text
JSON 文件/字符串  ─┐
                  ├─> json2excel / jsonl2excel ─> rows_to_workbook
JSONL 文件/字符串 ─┘                              │
                                                 ├─> flatten_dict
                                                 ├─> serialize_value
                                                 ├─> auto_adjust_width
                                                 └─> save_workbook -> .xlsx
```

## 核心模块

### 输入适配层

- `json2excel`：解析完整 JSON 文档，要求顶层为数组，数组元素必须是对象。
- `jsonl2excel`：按行解析，跳过空行和 `#` 注释，遇到错误报告行号。
- 两者都暴露 CLI 和 `convert(rows, output, sheet_name)` API。

### 共享写入层

- `flatten_dict`：把嵌套对象转换成 `a.b` 形式的列名；列表保持整体值。
- `collect_columns`：遍历所有行，按首次出现顺序合并列。
- `serialize_value`：保留 Excel 原生基础类型，日期转 ISO 字符串，复杂对象转 JSON 字符串。
- `rows_to_workbook` / `save_workbook`：创建工作簿、写行、调整列宽并确保输出目录存在。

## 可行性评估

结论：作为“中小规模到超大规模、结构相对规整的 JSON 导出工具”可行，当前实现具备 CLI、库 API、流式解析、自动分片、示例和回归测试。

适用场景：接口响应导出、日志抽样导出、配置/清单转表格、一次性批处理。  
不宜直接承诺的场景：复杂 Excel 样式/公式、强 schema 约束、需要保留原始数值精度的财务数据，以及需要在内存中长期保留全部数据的二次处理流程。

## 风险与建议

1. `stream_rows_to_files` 会把行先写入临时磁盘以确定完整表头，磁盘空间需要纳入部署容量评估。
2. 工作表按 Excel 行列上限自动切分；工作簿达到默认 50 个工作表时自动生成 `.partNNN.xlsx` 文件。
3. 工作表名称会清理 Excel 非法字符并截断到 31 个字符，但业务上最好仍提供简短的 `--sheet-name`。
4. 当前没有 pytest/CI 配置，后续可增加性能基准、损坏文件恢复和真实超大文件测试。
5. 依赖声明允许 Python 3.10+，但本次机器上的 uv 缓存 Python 3.13 无法加载 `pyexpat`，导致 openpyxl 导入失败。这是运行时环境问题，不是项目代码异常；建议在发布/部署环境固定经过验证的 Python 版本并执行安装后冒烟测试。

## 扩展指南

- 新增输入格式：在 `src/any2excel` 增加输入适配模块，最终转换为 `Iterable[Mapping]`，复用 `stream_rows_to_files`。
- 增加 Excel 能力：优先扩展 `excel_writer.py`，避免 JSON 和 JSONL 两套逻辑分叉。
- 增加性能能力：继续优化临时文件格式、压缩策略和写入批次；不要回退到一次性 `json.load`。
