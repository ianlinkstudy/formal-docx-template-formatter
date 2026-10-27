#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import sys
import tempfile
from pathlib import Path

from core.common import (
    DEFAULT_LIBRARY_DIR,
    FormatterError,
    atomic_write_json,
    prepare_docx_copy,
)
from core.formatter_core import format_with_template, verify_file
from core.template_store import get_template


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="选择 GB 参考版式或已确认保存的自定义模板排版 DOCX。"
    )
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument(
        "--template",
        required=True,
        help="gb9704 或已确认并命名保存的自定义模板名称",
    )
    parser.add_argument("--library-dir", type=Path, default=DEFAULT_LIBRARY_DIR)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--verify-report", type=Path)
    parser.add_argument("--overwrite", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report_path = args.report or args.output.with_suffix(
        args.output.suffix + ".format-report.json"
    )
    verify_path = args.verify_report or args.output.with_suffix(
        args.output.suffix + ".verify-report.json"
    )
    if args.input.resolve() == args.output.resolve():
        print("失败：输出路径必须与输入路径不同。", file=sys.stderr)
        return 2
    if not args.overwrite and (
        args.output.exists() or report_path.exists() or verify_path.exists()
    ):
        print("失败：输出或报告已存在；如需覆盖请启用 --overwrite。", file=sys.stderr)
        return 2
    try:
        template = get_template(args.template, args.library_dir)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            prepared_input = prepare_docx_copy(args.input, temp / "input.docx")
            staged_output = temp / "formatted.docx"
            result = format_with_template(
                prepared_input,
                staged_output,
                template,
                overwrite=True,
            )
            verification = verify_file(staged_output, prepared_input, template)
            if verification["status"] != "passed":
                raise FormatterError("排版结果未通过核验，未生成正式输出。")
            os.replace(staged_output, args.output)
        result["output"] = str(args.output)
        format_report = {
            "schema_version": 1,
            "status": "success",
            "input": str(args.input),
            "output": str(args.output),
            "template": template["name"],
            "template_id": template["id"],
            "template_mode": template["mode"],
            "semantic_integrity": {
                "passed": True,
                "semantic_hash": result["semantic_hash"],
            },
            "changes": {
                "paragraphs_formatted": result["formatted_paragraphs"],
                "tables_formatted": result["formatted_tables"],
                "heading2_split": result["split_heading2"],
                "inline_images_resized": result["resized_inline_images"],
            },
            "rich_objects": {
                "formula_count": result["formula_count"],
                "ole_object_count": result["ole_object_count"],
                "nested_table_count": result["nested_table_count"],
                "floating_object_count": result["floating_object_count"],
            },
            "warnings": result["warnings"],
        }
        atomic_write_json(report_path, format_report)
        atomic_write_json(verify_path, verification)
        print(f"已生成：{args.output}")
        print(f"模板：{template['name']}（{template['mode']}）")
        print("正文与富对象完整性：通过")
        print(f"核验状态：{verification['status']}")
        for warning in result["warnings"]:
            print(f"提醒：{warning}")
        return 0
    except (FormatterError, OSError, ValueError) as exc:
        failure = {
            "schema_version": 1,
            "status": "failed",
            "input": str(args.input),
            "output": str(args.output),
            "error": str(exc),
        }
        atomic_write_json(report_path, failure)
        print(f"失败：{exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
