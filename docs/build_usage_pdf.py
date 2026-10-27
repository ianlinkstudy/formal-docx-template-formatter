#!/usr/bin/env python3
from __future__ import annotations

import html
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    Preformatted,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "docs" / "usage-guide.zh-CN.md"
OUTPUT = ROOT / "docs" / "usage-guide.zh-CN.pdf"

NAVY = colors.HexColor("#163C43")
TEAL = colors.HexColor("#2F6F75")
MINT = colors.HexColor("#EAF2F1")
PALE = colors.HexColor("#F6F8F8")
INK = colors.HexColor("#23383B")
MUTED = colors.HexColor("#637A7D")
LINE = colors.HexColor("#D4DFDF")
CODE = colors.HexColor("#EEF3F3")
CJK_FONT = "SourceHanSansCN"


def register_cjk_font() -> None:
    explicit = os.environ.get("FORMAL_DOCX_CJK_FONT")
    candidates = [
        Path(explicit) if explicit else None,
        ROOT.parent / "fonts" / "SourceHanSansCN-Normal.ttf",
        Path("/Library/Fonts/SourceHanSansCN-Normal.otf"),
        Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
    ]
    for candidate in candidates:
        if candidate is None or not candidate.is_file():
            continue
        if candidate.suffix.lower() == ".ttf":
            pdfmetrics.registerFont(TTFont(CJK_FONT, str(candidate)))
            return
        if candidate.suffix.lower() == ".otf":
            with tempfile.TemporaryDirectory() as temp_dir:
                converted = Path(temp_dir) / "SourceHanSansCN-Normal.ttf"
                converter = Path(sys.executable).parent / "otf2ttf"
                subprocess.run(
                    [str(converter), str(candidate), "-o", str(converted)],
                    check=True,
                    capture_output=True,
                )
                pdfmetrics.registerFont(TTFont(CJK_FONT, str(converted)))
                return
    raise RuntimeError(
        "未找到可嵌入的中文 TTF。请通过 FORMAL_DOCX_CJK_FONT 指定开源字体。"
    )


register_cjk_font()


def styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "body": ParagraphStyle(
            "BodyCN",
            parent=base["BodyText"],
            fontName=CJK_FONT,
            fontSize=10.5,
            leading=16,
            textColor=INK,
            spaceAfter=7,
        ),
        "h2": ParagraphStyle(
            "Heading2CN",
            parent=base["Heading2"],
            fontName=CJK_FONT,
            fontSize=18,
            leading=24,
            textColor=NAVY,
            spaceBefore=10,
            spaceAfter=10,
            keepWithNext=True,
        ),
        "h3": ParagraphStyle(
            "Heading3CN",
            parent=base["Heading3"],
            fontName=CJK_FONT,
            fontSize=13,
            leading=19,
            textColor=TEAL,
            spaceBefore=7,
            spaceAfter=6,
            keepWithNext=True,
        ),
        "bullet": ParagraphStyle(
            "BulletCN",
            parent=base["BodyText"],
            fontName=CJK_FONT,
            fontSize=10.5,
            leading=16,
            textColor=INK,
            leftIndent=14,
            firstLineIndent=-10,
            bulletIndent=2,
            spaceAfter=4,
        ),
        "code": ParagraphStyle(
            "CodeCN",
            parent=base["Code"],
            fontName=CJK_FONT,
            fontSize=8,
            leading=11.5,
            textColor=colors.HexColor("#26464A"),
            backColor=CODE,
            borderColor=LINE,
            borderWidth=0.6,
            borderPadding=7,
            leftIndent=4,
            rightIndent=4,
            spaceBefore=3,
            spaceAfter=9,
        ),
        "quote": ParagraphStyle(
            "QuoteCN",
            parent=base["BodyText"],
            fontName=CJK_FONT,
            fontSize=10,
            leading=15,
            textColor=MUTED,
            leftIndent=12,
            borderColor=TEAL,
            borderWidth=0,
            borderPadding=0,
            spaceBefore=3,
            spaceAfter=8,
        ),
        "table": ParagraphStyle(
            "TableCN",
            parent=base["BodyText"],
            fontName=CJK_FONT,
            fontSize=8.3,
            leading=12,
            textColor=INK,
        ),
    }


def inline_markup(text: str) -> str:
    escaped = html.escape(text, quote=False)
    escaped = re.sub(
        r"`([^`]+)`",
        rf'<font name="{CJK_FONT}" color="#245A60">\1</font>',
        escaped,
    )
    escaped = re.sub(
        r"\*\*([^*]+)\*\*",
        r'<font color="#163C43">\1</font>',
        escaped,
    )
    escaped = re.sub(
        r"\[([^\]]+)\]\(([^)]+)\)",
        r'<link href="\2" color="#2F6F75">\1</link>',
        escaped,
    )
    return escaped


def wrap_code(lines: list[str], width: int = 88) -> str:
    wrapped: list[str] = []
    for line in lines:
        if len(line) <= width:
            wrapped.append(line)
            continue
        remaining = line
        while remaining:
            wrapped.append(remaining[:width])
            remaining = remaining[width:]
    return "\n".join(wrapped)


def parse_markdown(lines: list[str], text_styles: dict[str, ParagraphStyle]) -> list:
    story: list = []
    index = 0
    after_cover_rule = False
    while index < len(lines):
        line = lines[index].rstrip()
        stripped = line.strip()
        if not stripped:
            index += 1
            continue
        if stripped.startswith("# "):
            index += 1
            continue
        if stripped.startswith("> 鼓翼成长星球"):
            index += 1
            continue
        if stripped == "---" and not after_cover_rule:
            after_cover_rule = True
            index += 1
            continue
        if stripped.startswith("```"):
            index += 1
            block: list[str] = []
            while index < len(lines) and not lines[index].strip().startswith("```"):
                block.append(lines[index].rstrip())
                index += 1
            story.append(Preformatted(wrap_code(block), text_styles["code"]))
            index += 1
            continue
        if stripped.startswith("|") and stripped.endswith("|"):
            rows: list[list[str]] = []
            while (
                index < len(lines)
                and lines[index].strip().startswith("|")
                and lines[index].strip().endswith("|")
            ):
                row = [
                    cell.strip()
                    for cell in lines[index].strip().strip("|").split("|")
                ]
                if not all(re.fullmatch(r":?-{3,}:?", cell) for cell in row):
                    rows.append(row)
                index += 1
            if rows:
                data = [
                    [
                        Paragraph(inline_markup(cell), text_styles["table"])
                        for cell in row
                    ]
                    for row in rows
                ]
                width = 166 * mm
                columns = len(rows[0])
                table = Table(
                    data,
                    colWidths=[width / columns] * columns,
                    repeatRows=1,
                    hAlign="LEFT",
                )
                table.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), MINT),
                            ("TEXTCOLOR", (0, 0), (-1, 0), NAVY),
                            ("GRID", (0, 0), (-1, -1), 0.5, LINE),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 6),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                            ("TOPPADDING", (0, 0), (-1, -1), 6),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, PALE]),
                        ]
                    )
                )
                story.extend([table, Spacer(1, 7)])
            continue
        heading_match = re.match(r"^(#{2,3})\s+(.+)$", stripped)
        if heading_match:
            level = len(heading_match.group(1))
            style = text_styles["h2" if level == 2 else "h3"]
            story.append(
                Paragraph(inline_markup(heading_match.group(2)), style)
            )
            index += 1
            continue
        if stripped.startswith(">"):
            story.append(
                Paragraph(
                    inline_markup(stripped.lstrip("> ").strip()),
                    text_styles["quote"],
                )
            )
            index += 1
            continue
        if re.match(r"^[-*]\s+", stripped):
            content = re.sub(r"^[-*]\s+", "", stripped)
            story.append(
                Paragraph(
                    inline_markup(content),
                    text_styles["bullet"],
                    bulletText="•",
                )
            )
            index += 1
            continue
        paragraph_lines = [stripped]
        index += 1
        while index < len(lines):
            next_line = lines[index].strip()
            if (
                not next_line
                or next_line.startswith(("#", "```", ">", "|", "- ", "* "))
                or next_line == "---"
            ):
                break
            paragraph_lines.append(next_line)
            index += 1
        story.append(
            Paragraph(
                inline_markup(" ".join(paragraph_lines)),
                text_styles["body"],
            )
        )
    return story


def cover(text_styles: dict[str, ParagraphStyle]) -> list:
    cover_title = ParagraphStyle(
        "CoverTitle",
        parent=text_styles["h2"],
        fontSize=34,
        leading=44,
        alignment=TA_LEFT,
        textColor=NAVY,
        spaceAfter=12,
    )
    cover_subtitle = ParagraphStyle(
        "CoverSubtitle",
        parent=text_styles["body"],
        fontSize=14,
        leading=23,
        textColor=MUTED,
        spaceAfter=8,
    )
    cover_brand = ParagraphStyle(
        "CoverBrand",
        parent=text_styles["body"],
        fontSize=10,
        leading=14,
        alignment=TA_LEFT,
        textColor=TEAL,
        spaceAfter=42,
    )
    cover_meta = ParagraphStyle(
        "CoverMeta",
        parent=text_styles["body"],
        fontSize=9,
        leading=15,
        alignment=TA_LEFT,
        textColor=MUTED,
    )
    flow = Table(
        [
            [
                Paragraph("解析", text_styles["table"]),
                Paragraph("确认", text_styles["table"]),
                Paragraph("命名", text_styles["table"]),
                Paragraph("保存", text_styles["table"]),
            ]
        ],
        colWidths=[30 * mm] * 4,
        hAlign="LEFT",
    )
    flow.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), MINT),
                ("TEXTCOLOR", (0, 0), (-1, -1), NAVY),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTNAME", (0, 0), (-1, -1), CJK_FONT),
                ("FONTSIZE", (0, 0), (-1, -1), 11),
                ("BOX", (0, 0), (-1, -1), 0.8, TEAL),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.white),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    return [
        Spacer(1, 54 * mm),
        Paragraph("鼓翼成长星球 · WORD 模板排版 SKILL", cover_brand),
        Paragraph("Word 模板排版", cover_title),
        Paragraph("先选择模板，再改变格式。", cover_subtitle),
        Paragraph(
            "导入 DOCX/DOTX 模板，经过解析、明确确认、命名和保存；"
            "每次排版再选择模板，输出新文档与可复核报告。",
            cover_subtitle,
        ),
        Spacer(1, 15 * mm),
        flow,
        Spacer(1, 18 * mm),
        Paragraph("1.0.0 · 2026-10-27", cover_meta),
        Paragraph("Python 3.10+ · Codex / Claude Code / Cursor", cover_meta),
        PageBreak(),
    ]


def draw_cover(canvas, doc) -> None:
    width, height = A4
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.rect(0, height - 16 * mm, width, 16 * mm, fill=1, stroke=0)
    canvas.setFillColor(TEAL)
    canvas.rect(0, 0, 9 * mm, height, fill=1, stroke=0)
    canvas.restoreState()


def draw_later(canvas, doc) -> None:
    width, height = A4
    canvas.saveState()
    canvas.setStrokeColor(LINE)
    canvas.setLineWidth(0.5)
    canvas.line(22 * mm, height - 16 * mm, width - 22 * mm, height - 16 * mm)
    canvas.setFont(CJK_FONT, 8)
    canvas.setFillColor(MUTED)
    canvas.drawString(22 * mm, height - 12 * mm, "WORD 模板排版 · 使用手册")
    canvas.drawRightString(
        width - 22 * mm,
        height - 12 * mm,
        "鼓翼成长星球",
    )
    canvas.line(22 * mm, 14 * mm, width - 22 * mm, 14 * mm)
    canvas.drawString(22 * mm, 9 * mm, "formal-docx-template-formatter")
    canvas.drawRightString(
        width - 22 * mm,
        9 * mm,
        f"{canvas.getPageNumber():02d}",
    )
    canvas.restoreState()


def main() -> int:
    text_styles = styles()
    lines = SOURCE.read_text(encoding="utf-8").splitlines()
    doc = SimpleDocTemplate(
        str(OUTPUT),
        pagesize=A4,
        leftMargin=22 * mm,
        rightMargin=22 * mm,
        topMargin=23 * mm,
        bottomMargin=20 * mm,
        title="Word 模板排版使用手册",
        author="鼓翼成长星球",
        subject="formal-docx-template-formatter",
    )
    story = cover(text_styles)
    story.extend(parse_markdown(lines, text_styles))
    doc.build(story, onFirstPage=draw_cover, onLaterPages=draw_later)
    print(f"已生成：{OUTPUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
