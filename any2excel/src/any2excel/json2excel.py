"""将 JSON 列表（List[Dict]）转换为 Excel。

支持从以下来源读取数据：
- 标准 JSON 文件
- 通过 ``--data`` 直接传入 JSON 字符串

使用示例::

    python -m any2excel.json2excel -i data.json -o out.xlsx
    python -m any2excel.json2excel --data '[{"a": 1}]' -d out.xlsx
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Iterable, Iterator, Mapping, Sequence, TextIO

from pydantic import BaseModel, ConfigDict, Field, field_validator

from .excel_writer import rows_to_workbook, save_workbook, stream_rows_to_files


class ConvertRequest(BaseModel):
    """JSON 列表转 Excel 的请求模型。"""

    model_config = ConfigDict(extra="forbid")

    rows: Sequence[Mapping[str, Any]] = Field(default_factory=list, description="待写入的数据行")
    output: Path = Field(..., description="输出 Excel 路径")
    sheet_name: str = Field(default="data", description="工作表名称")

    @field_validator("rows")
    @classmethod
    def _ensure_list_of_dict(cls, value: Sequence[Mapping[str, Any]]) -> Sequence[Mapping[str, Any]]:
        if not isinstance(value, list):
            raise ValueError("JSON 数据必须是数组")
        for index, item in enumerate(value):
            if not isinstance(item, Mapping):
                raise ValueError(f"第 {index} 项不是对象")
        return value

    @field_validator("output")
    @classmethod
    def _ensure_xlsx(cls, value: Path) -> Path:
        if value.suffix.lower() != ".xlsx":
            return value.with_suffix(".xlsx")
        return value


def load_rows_from_file(path: str | Path) -> list[Mapping[str, Any]]:
    """从 JSON 文件加载列表数据。"""

    text = Path(path).read_text(encoding="utf-8")
    data = json.loads(text)
    if not isinstance(data, list):
        raise ValueError(f"{path} 内容不是 JSON 数组")
    return data


def iter_json_array_stream(stream: TextIO, chunk_size: int = 1024 * 1024) -> Iterator[Mapping[str, Any]]:
    """增量解析顶层 JSON 数组，避免 ``json.load`` 一次性占满内存。"""
    decoder = json.JSONDecoder()
    buffer = ""
    eof = False

    def fill() -> None:
        nonlocal buffer, eof
        if not eof:
            part = stream.read(chunk_size)
            if part:
                buffer += part
            else:
                eof = True

    def skip_ws() -> None:
        nonlocal buffer
        buffer = buffer.lstrip()

    fill()
    while True:
        skip_ws()
        if buffer or eof:
            break
        fill()
    if not buffer.startswith("["):
        raise ValueError("JSON 数据必须是数组")
    buffer = buffer[1:]
    while True:
        while True:
            skip_ws()
            if buffer or eof:
                break
            fill()
        if buffer.startswith("]"):
            return
        while True:
            try:
                value, end = decoder.raw_decode(buffer)
                break
            except json.JSONDecodeError:
                if eof:
                    raise ValueError("JSON 数组解析失败")
                fill()
        if not isinstance(value, Mapping):
            raise ValueError("JSON 数组中的每一项必须是对象")
        yield value
        buffer = buffer[end:]
        while True:
            skip_ws()
            if not buffer and not eof:
                fill()
                continue
            if buffer.startswith("]"):
                return
            if buffer.startswith(","):
                buffer = buffer[1:]
                break
            raise ValueError("JSON 数组缺少逗号或结束符 ]")


def iter_rows_from_file(path: str | Path) -> Iterator[Mapping[str, Any]]:
    with Path(path).open("r", encoding="utf-8") as stream:
        yield from iter_json_array_stream(stream)


def convert(rows: Sequence[Mapping[str, Any]], output: str | Path, sheet_name: str = "data") -> Path:
    """执行转换并保存到 Excel 文件。"""

    request = ConvertRequest(rows=list(rows), output=Path(output), sheet_name=sheet_name)
    workbook = rows_to_workbook(request.rows, sheet_name=request.sheet_name)
    return save_workbook(workbook, request.output)


def build_parser() -> argparse.ArgumentParser:
    """构建命令行参数解析器。"""

    parser = argparse.ArgumentParser(description="将 JSON 列表转换为 Excel")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("-i", "--input", help="输入 JSON 文件路径")
    group.add_argument("--data", help="直接传入 JSON 字符串")
    parser.add_argument("-o", "--output", required=True, help="输出 Excel 文件路径")
    parser.add_argument("--sheet-name", default="data", help="工作表名称")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """CLI 入口函数。"""

    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if args.input:
            rows: Iterable[Mapping[str, Any]] = iter_rows_from_file(args.input)
        else:
            rows = iter_json_array_stream(__import__("io").StringIO(args.data))
        output_paths = stream_rows_to_files(rows, args.output, sheet_name=args.sheet_name)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        parser.error(str(exc))
    for output_path in output_paths:
        print(f"已生成 Excel 文件：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
