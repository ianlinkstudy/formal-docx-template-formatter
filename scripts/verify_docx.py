#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

from core.common import (
    DEFAULT_LIBRARY_DIR,
    FormatterError,
    atomic_write_json,
    prepare_docx_copy,
)
from core.formatter_core import verify_file
from core.template_store import get_template


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="核验模板排版结果。")
    parser.add_argument("formatted", type=Path)
    parser.add_argument("--reference", required=True, type=Path)
    parser.add_argument("--template", required=True)
    parser.add_argument("--library-dir", type=Path, default=DEFAULT_LIBRARY_DIR)
    parser.add_argument("--report", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report_path = args.report or args.formatted.with_suffix(
        args.formatted.suffix + ".verify-report.json"
    )
    try:
        template = get_template(args.template, args.library_dir)
        with tempfile.TemporaryDirectory() as temp_dir:
            reference = prepare_docx_copy(
                args.reference, Path(temp_dir) / "reference.docx"
            )
            report = verify_file(args.formatted, reference, template)
        atomic_write_json(report_path, report)
        print(f"核验状态：{report['status']}")
        print(f"报告：{report_path}")
        for check in report["checks"]:
            print(
                f"- {check['name']}：{'通过' if check['passed'] else '失败'}"
            )
        return 0 if report["status"] == "passed" else 2
    except (FormatterError, OSError, ValueError) as exc:
        print(f"失败：{exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
