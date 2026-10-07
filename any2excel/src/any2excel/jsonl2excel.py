"""将 JSONL（每行一个 JSON 对象）转换为 Excel。

文件格式示例::

    {"a": 1, "b": "x"}
    {"a": 2, "b": "y"}

使用示例::

    python -m any2excel.jsonl2excel -i data.jsonl -o out.xlsx
    python -m any2excel.jsonl2excel --data '{"a": 1}
    {"a": 2}' -o out.xlsx
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Iterable, Iterator, Mapping, Sequence, TextIO

from pydantic import BaseModel, ConfigDict, Field, field_validator

from .excel_writer import rows_to_workbook, save_workbook, stream_rows_to_files


class ConvertRequest(BaseModel):
    """JSONL 转 Excel 的请求模型。"""

    model_config = ConfigDict(extra="forbid")

    rows: Sequence[Mapping[str, Any]] = Field(default_factory=list)
    output: Path
    sheet_name: str = Field(default="data")

    @field_validator("output")
    @classmethod
    def _ensure_xlsx(cls, value: Path) -> Path:
        if value.suffix.lower() != ".xlsx":
            return value.with_suffix(".xlsx")
        return value


def iter_jsonl(text: str) -> Iterator[Mapping[str, Any]]:
    """逐行解析 JSONL 文本，跳过空行和以 ``#`` 开头的注释行。"""

    for lineno, raw in enumerate(text.splitlines(), start=1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        try:
            item = json.loads(line)
        except json.JSONDecodeError as exc:
            raise ValueError(f"第 {lineno} 行 JSON 解析失败：{exc}") from exc
        if not isinstance(item, Mapping):
            raise ValueError(f"第 {lineno} 行不是 JSON 对象")
        yield item


def load_rows_from_file(path: str | Path) -> list[Mapping[str, Any]]:
    """从 JSONL 文件读取所有行。"""

    text = Path(path).read_text(encoding="utf-8")
    return list(iter_jsonl(text))


def iter_jsonl_file(path: str | Path) -> Iterator[Mapping[str, Any]]:
    """逐行读取 JSONL 文件，不把整个文件载入内存。"""
    with Path(path).open("r", encoding="utf-8") as stream:
        for lineno, raw in enumerate(stream, start=1):
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            try:
                item = json.loads(line)
            except json.JSONDecodeError as exc:
                raise ValueError(f"第 {lineno} 行 JSON 解析失败：{exc}") from exc
            if not isinstance(item, Mapping):
                raise ValueError(f"第 {lineno} 行不是 JSON 对象")
            yield item


def convert(rows: Sequence[Mapping[str, Any]], output: str | Path, sheet_name: str = "data") -> Path:
    """执行 JSONL -> Excel 转换。"""

    request = ConvertRequest(rows=list(rows), output=Path(output), sheet_name=sheet_name)
    workbook = rows_to_workbook(request.rows, sheet_name=request.sheet_name)
    return save_workbook(workbook, request.output)


def build_parser() -> argparse.ArgumentParser:
    """构建命令行参数解析器。"""

    parser = argparse.ArgumentParser(description="将 JSONL 文件转换为 Excel")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("-i", "--input", help="输入 JSONL 文件路径")
    group.add_argument("--data", help="直接传入 JSONL 字符串")
    parser.add_argument("-o", "--output", required=True, help="输出 Excel 文件路径")
    parser.add_argument("--sheet-name", default="data", help="工作表名称")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """CLI 入口函数。"""

    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        rows: Iterable[Mapping[str, Any]] = iter_jsonl_file(args.input) if args.input else iter_jsonl(args.data)
        output_paths = stream_rows_to_files(rows, args.output, sheet_name=args.sheet_name)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        parser.error(str(exc))
    for output_path in output_paths:
        print(f"已生成 Excel 文件：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
