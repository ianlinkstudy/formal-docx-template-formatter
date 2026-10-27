#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from xml.sax.saxutils import escape

from core.common import FormatterError


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成 DOCX 的 PDF/PNG 预览。")
    parser.add_argument("docx", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--soffice", type=Path)
    parser.add_argument("--pdftoppm", type=Path)
    parser.add_argument("--dpi", type=int, default=144)
    return parser.parse_args()


def _binary(explicit: Path | None, name: str) -> str | None:
    if explicit:
        return str(explicit) if explicit.is_file() else None
    return shutil.which(name)


def _font_directories() -> list[Path]:
    home = Path.home()
    candidates = [
        home / "Library" / "Fonts",
        Path("/Library/Fonts"),
        Path("/System/Library/Fonts"),
        Path("/System/Library/Fonts/Supplemental"),
        home / ".fonts",
        home / ".local" / "share" / "fonts",
        Path("/opt/homebrew/share/fonts"),
        Path("/usr/local/share/fonts"),
        Path("/usr/share/fonts"),
    ]
    windows = os.environ.get("WINDIR")
    if windows:
        candidates.append(Path(windows) / "Fonts")
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        candidates.append(Path(local_app_data) / "Microsoft" / "Windows" / "Fonts")
    return [path for path in candidates if path.is_dir()]


def _write_fontconfig(path: Path, cache_dir: Path) -> None:
    directories = "\n".join(
        f"  <dir>{escape(str(directory))}</dir>"
        for directory in _font_directories()
    )
    path.write_text(
        "<?xml version=\"1.0\"?>\n"
        "<!DOCTYPE fontconfig SYSTEM \"urn:fontconfig:fonts.dtd\">\n"
        "<fontconfig>\n"
        f"{directories}\n"
        f"  <cachedir>{escape(str(cache_dir))}</cachedir>\n"
        "</fontconfig>\n",
        encoding="utf-8",
    )


def main() -> int:
    args = parse_args()
    if not args.docx.is_file():
        print(f"失败：文件不存在：{args.docx}", file=sys.stderr)
        return 2
    args.output.mkdir(parents=True, exist_ok=True)
    soffice = _binary(args.soffice, "soffice") or _binary(None, "libreoffice")
    pdftoppm = _binary(args.pdftoppm, "pdftoppm")
    try:
        if soffice and pdftoppm:
            with tempfile.TemporaryDirectory() as temp_dir:
                profile = Path(temp_dir) / "lo-profile"
                cache = Path(temp_dir) / "font-cache"
                profile.mkdir()
                cache.mkdir()
                fontconfig = Path(temp_dir) / "fonts.conf"
                _write_fontconfig(fontconfig, cache)
                env = os.environ.copy()
                env["FONTCONFIG_FILE"] = str(fontconfig)
                env["FONTCONFIG_PATH"] = temp_dir
                env["XDG_CACHE_HOME"] = str(Path(temp_dir) / "xdg-cache")
                env["XDG_CONFIG_HOME"] = str(Path(temp_dir) / "xdg-config")
                command = [
                    soffice,
                    f"-env:UserInstallation={profile.as_uri()}",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    temp_dir,
                    str(args.docx),
                ]
                subprocess.run(
                    command, check=True, capture_output=True, env=env
                )
                pdfs = list(Path(temp_dir).glob("*.pdf"))
                if not pdfs:
                    raise FormatterError("LibreOffice 未生成 PDF。")
                pdf = args.output / f"{args.docx.stem}.pdf"
                shutil.copyfile(pdfs[0], pdf)
                subprocess.run(
                    [
                        pdftoppm,
                        "-png",
                        "-r",
                        str(args.dpi),
                        str(pdf),
                        str(args.output / "page"),
                    ],
                    check=True,
                    capture_output=True,
                    env=env,
                )
                print(f"已生成 PDF：{pdf}")
                print(f"已生成逐页 PNG：{args.output}")
                return 0

        qlmanage = shutil.which("qlmanage")
        if qlmanage and sys.platform == "darwin":
            subprocess.run(
                [qlmanage, "-t", "-s", "2000", "-o", str(args.output), str(args.docx)],
                check=True,
                capture_output=True,
            )
            generated = list(args.output.glob("*.png"))
            if generated:
                target = args.output / "page-1.png"
                if generated[0] != target:
                    os.replace(generated[0], target)
                print(f"已生成首页预览：{target}")
                print("提醒：Quick Look 只生成首页，不能替代逐页视觉检查。")
                return 0
        raise FormatterError(
            "未找到 LibreOffice/poppler 或 Quick Look；请使用 Word/WPS 人工检查。"
        )
    except (FormatterError, subprocess.SubprocessError, OSError) as exc:
        print(f"失败：{exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
