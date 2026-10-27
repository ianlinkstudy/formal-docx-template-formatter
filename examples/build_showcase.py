#!/usr/bin/env python3
from __future__ import annotations

import os
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
BEFORE = (
    ROOT
    / "examples"
    / "previews"
    / "before"
    / "01-project-report"
    / "page-1.png"
)
AFTER = (
    ROOT
    / "examples"
    / "previews"
    / "after"
    / "02-company-shell-project-report"
    / "page-1.png"
)
OUTPUT = ROOT / "docs" / "formal-docx-template-before-after.png"
FONT_CANDIDATES = (
    Path(os.environ["FORMAL_DOCX_CJK_FONT"])
    if os.environ.get("FORMAL_DOCX_CJK_FONT")
    else None,
    ROOT.parent / "fonts" / "SourceHanSansCN-Normal.ttf",
    Path("/Library/Fonts/SourceHanSansCN-Normal.otf"),
    Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
)


def font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for path in FONT_CANDIDATES:
        if path is not None and path.is_file():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def page_image(path: Path, height: int) -> Image.Image:
    image = Image.open(path).convert("RGB")
    width = round(image.width * height / image.height)
    return image.resize((width, height), Image.Resampling.LANCZOS)


def main() -> int:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    page_height = 1120
    before = page_image(BEFORE, page_height)
    after = page_image(AFTER, page_height)
    margin = 78
    gap = 78
    label_height = 70
    footer_height = 76
    canvas_width = margin * 2 + before.width + gap + after.width
    canvas_height = margin * 2 + label_height + page_height + footer_height

    canvas = Image.new("RGB", (canvas_width, canvas_height), "#F2F5F5")
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle(
        (0, 0, canvas_width, 18),
        radius=0,
        fill="#163C43",
    )

    labels = (
        (margin, "排版前", "普通 DOCX"),
        (margin + before.width + gap, "排版后", "已确认公司母版"),
    )
    for x, title, subtitle in labels:
        draw.text((x, margin), title, font=font(34), fill="#163C43")
        draw.text(
            (x + 146, margin + 10),
            subtitle,
            font=font(22),
            fill="#5E7376",
        )

    page_top = margin + label_height
    for x, image in (
        (margin, before),
        (margin + before.width + gap, after),
    ):
        shadow = Image.new("RGBA", (image.width + 30, image.height + 30), (0, 0, 0, 0))
        shadow_draw = ImageDraw.Draw(shadow)
        shadow_draw.rounded_rectangle(
            (12, 16, image.width + 18, image.height + 22),
            radius=8,
            fill=(23, 52, 56, 45),
        )
        canvas.paste(shadow, (x - 12, page_top - 12), shadow)
        canvas.paste(image, (x, page_top))
        draw.rectangle(
            (x, page_top, x + image.width, page_top + image.height),
            outline="#CBD7D7",
            width=2,
        )

    footer_y = page_top + page_height + 24
    draw.line(
        (margin, footer_y, canvas_width - margin, footer_y),
        fill="#C8D5D5",
        width=2,
    )
    draw.text(
        (margin, footer_y + 18),
        "先选择模板，再改变格式。",
        font=font(24),
        fill="#163C43",
    )
    footer = "内容、表格、图片、公式与页眉页脚完整性均通过核验"
    footer_width = draw.textlength(footer, font=font(20))
    draw.text(
        (canvas_width - margin - footer_width, footer_y + 22),
        footer,
        font=font(20),
        fill="#5E7376",
    )

    canvas.save(OUTPUT, optimize=True)
    print(f"已生成：{OUTPUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
