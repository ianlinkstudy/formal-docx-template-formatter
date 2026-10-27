#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import shutil
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIST = ROOT / "dist"
VERSION = (ROOT / "VERSION").read_text(encoding="utf-8").strip()
NAME = "formal-docx-template-formatter"
ARCHIVE_STEM = f"{NAME}-v{VERSION}"
ZIP_PATH = DIST / f"{ARCHIVE_STEM}.zip"
SKILL_PATH = DIST / f"{ARCHIVE_STEM}.skill"
PDF_PATH = DIST / f"{ARCHIVE_STEM}-usage-guide.zh-CN.pdf"
SUMS_PATH = DIST / "SHA256SUMS"

EXCLUDED_DIRS = {".git", "__pycache__", ".venv", "dist"}
EXCLUDED_SUFFIXES = {".pyc", ".pyo", ".zip", ".skill"}
EXCLUDED_NAMES = {".DS_Store"}
REQUIRED_FILES = {
    "SKILL.md",
    "README.md",
    "VERSION",
    "agents/openai.yaml",
    "profiles/gb9704.json",
    "requirements.txt",
    "scripts/format_docx.py",
    "scripts/template_cli.py",
    "scripts/verify_docx.py",
    "docs/usage-guide.zh-CN.pdf",
}


def archive_files() -> list[Path]:
    files: list[Path] = []
    for path in ROOT.rglob("*"):
        if not path.is_file():
            continue
        relative = path.relative_to(ROOT)
        if any(part in EXCLUDED_DIRS for part in relative.parts):
            continue
        if path.name in EXCLUDED_NAMES or path.suffix in EXCLUDED_SUFFIXES:
            continue
        files.append(path)
    return sorted(files, key=lambda item: item.relative_to(ROOT).as_posix())


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    DIST.mkdir(parents=True, exist_ok=True)
    pdf_source = ROOT / "docs" / "usage-guide.zh-CN.pdf"
    if not pdf_source.is_file():
        raise SystemExit("缺少 docs/usage-guide.zh-CN.pdf")

    files = archive_files()
    relative_files = {path.relative_to(ROOT).as_posix() for path in files}
    missing = sorted(REQUIRED_FILES - relative_files)
    if missing:
        raise SystemExit(f"发布包缺少必要文件：{', '.join(missing)}")

    with zipfile.ZipFile(
        ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
    ) as package:
        for path in files:
            relative = path.relative_to(ROOT)
            package.write(path, f"{NAME}/{relative.as_posix()}")

    shutil.copyfile(ZIP_PATH, SKILL_PATH)
    shutil.copyfile(pdf_source, PDF_PATH)

    with zipfile.ZipFile(ZIP_PATH) as package:
        names = package.namelist()
        if any(not name.startswith(f"{NAME}/") for name in names):
            raise SystemExit("压缩包包含错误的顶层目录。")
        for required in REQUIRED_FILES:
            if f"{NAME}/{required}" not in names:
                raise SystemExit(f"压缩包缺少：{required}")

    sums = "".join(
        f"{sha256(path)}  {path.name}\n"
        for path in (ZIP_PATH, SKILL_PATH, PDF_PATH)
    )
    SUMS_PATH.write_text(sums, encoding="utf-8")
    print(f"已生成：{DIST}")
    print(sums, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
