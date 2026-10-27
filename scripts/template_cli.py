#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from core.common import DEFAULT_LIBRARY_DIR, FormatterError, safe_json
from core.template_store import (
    amend_candidate,
    archive_template,
    confirm_candidate,
    inspect_template,
    rename_template,
    restore_template,
    save_candidate,
    show_template,
    template_choices,
)


def _common(name: str, help_text: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--library-dir", type=Path, default=DEFAULT_LIBRARY_DIR)
    return parser


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="管理 formal-docx-template-formatter 的本地模板生命周期。"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    list_parser = subparsers.add_parser(
        "list", parents=[_common("list", "list")], help="列出 GB、自定义模板和槽位"
    )
    list_parser.set_defaults(action="list")

    inspect_parser = subparsers.add_parser(
        "inspect",
        parents=[_common("inspect", "inspect")],
        help="解析 DOCX/DOTX 并创建未保存候选",
    )
    inspect_parser.add_argument("source", type=Path)
    inspect_parser.add_argument(
        "--mode", choices=["auto", "style", "shell"], default="auto"
    )
    inspect_parser.add_argument("--body-start", type=int)
    inspect_parser.add_argument("--body-end", type=int)
    inspect_parser.set_defaults(action="inspect")

    amend_parser = subparsers.add_parser(
        "amend",
        parents=[_common("amend", "amend")],
        help="修正候选并使其确认失效",
    )
    amend_parser.add_argument("candidate_id")
    amend_parser.add_argument("--changes-json")
    amend_parser.add_argument("--changes-file", type=Path)
    amend_parser.add_argument("--mode", choices=["style", "shell"])
    amend_parser.add_argument("--body-start", type=int)
    amend_parser.add_argument("--body-end", type=int)
    amend_parser.set_defaults(action="amend")

    confirm_parser = subparsers.add_parser(
        "confirm",
        parents=[_common("confirm", "confirm")],
        help="绑定候选哈希并记录用户明确确认",
    )
    confirm_parser.add_argument("candidate_id")
    confirm_parser.add_argument("--candidate-sha256", required=True)
    confirm_parser.add_argument("--user-confirmation-text", required=True)
    confirm_parser.set_defaults(action="confirm")

    save_parser = subparsers.add_parser(
        "save",
        parents=[_common("save", "save")],
        help="为已确认候选命名并保存",
    )
    save_parser.add_argument("candidate_id")
    save_parser.add_argument("--name", required=True)
    save_parser.set_defaults(action="save")

    for command, help_text in (
        ("archive", "归档有效模板并释放槽位"),
        ("restore", "恢复归档模板"),
        ("show", "查看模板详情"),
    ):
        child = subparsers.add_parser(
            command, parents=[_common(command, command)], help=help_text
        )
        child.add_argument("name")
        child.set_defaults(action=command)

    rename_parser = subparsers.add_parser(
        "rename",
        parents=[_common("rename", "rename")],
        help="重命名有效模板",
    )
    rename_parser.add_argument("current_name")
    rename_parser.add_argument("new_name")
    rename_parser.set_defaults(action="rename")
    return parser


def _changes(args: argparse.Namespace) -> dict | None:
    if args.changes_json and args.changes_file:
        raise FormatterError("--changes-json 与 --changes-file 只能使用一个。")
    if args.changes_json:
        try:
            payload = json.loads(args.changes_json)
        except json.JSONDecodeError as exc:
            raise FormatterError("--changes-json 不是有效 JSON。") from exc
    elif args.changes_file:
        try:
            payload = json.loads(args.changes_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise FormatterError("无法读取 --changes-file。") from exc
    else:
        return None
    if not isinstance(payload, dict):
        raise FormatterError("修正内容必须是 JSON 对象。")
    return payload


def main() -> int:
    args = build_parser().parse_args()
    try:
        if args.action == "list":
            result = template_choices(args.library_dir)
        elif args.action == "inspect":
            result = inspect_template(
                args.source,
                mode=args.mode,
                body_start=args.body_start,
                body_end=args.body_end,
                library_dir=args.library_dir,
            )
        elif args.action == "amend":
            result = amend_candidate(
                args.candidate_id,
                changes=_changes(args),
                mode=args.mode,
                body_start=args.body_start,
                body_end=args.body_end,
                library_dir=args.library_dir,
            )
        elif args.action == "confirm":
            result = confirm_candidate(
                args.candidate_id,
                args.candidate_sha256,
                args.user_confirmation_text,
                library_dir=args.library_dir,
            )
        elif args.action == "save":
            result = save_candidate(
                args.candidate_id, args.name, library_dir=args.library_dir
            )
        elif args.action == "archive":
            result = archive_template(args.name, library_dir=args.library_dir)
        elif args.action == "restore":
            result = restore_template(args.name, library_dir=args.library_dir)
        elif args.action == "rename":
            result = rename_template(
                args.current_name, args.new_name, library_dir=args.library_dir
            )
        elif args.action == "show":
            result = show_template(args.name, library_dir=args.library_dir)
        else:  # pragma: no cover
            raise FormatterError("未知命令。")
        print(json.dumps(safe_json(result), ensure_ascii=False, indent=2))
        return 0
    except (FormatterError, OSError, ValueError) as exc:
        print(
            json.dumps(
                {"status": "failed", "error": str(exc)},
                ensure_ascii=False,
                indent=2,
            ),
            file=sys.stderr,
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
