from __future__ import annotations

import argparse
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_dir", type=Path)
    parser.add_argument("--output_dir", required=True, type=Path)
    parser.add_argument("--columns", type=int, default=4)
    parser.add_argument("--rows", type=int, default=4)
    args = parser.parse_args()

    pages = sorted(args.input_dir.glob("page-*.png"), key=page_number)
    if not pages:
        raise SystemExit(f"No page PNGs found in {args.input_dir}")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    font = ImageFont.load_default()
    cell_width, cell_height, label_height = 330, 460, 24
    page_limit = args.columns * args.rows
    for sheet_index in range(math.ceil(len(pages) / page_limit)):
        batch = pages[sheet_index * page_limit : (sheet_index + 1) * page_limit]
        sheet = Image.new("RGB", (args.columns * cell_width, args.rows * cell_height), "#d8d8d8")
        draw = ImageDraw.Draw(sheet)
        for index, path in enumerate(batch):
            row, column = divmod(index, args.columns)
            x, y = column * cell_width, row * cell_height
            with Image.open(path) as page:
                page = page.convert("RGB")
                page.thumbnail((cell_width - 16, cell_height - label_height - 16), Image.Resampling.LANCZOS)
                px = x + (cell_width - page.width) // 2
                py = y + label_height + (cell_height - label_height - page.height) // 2
                sheet.paste(page, (px, py))
            label = f"Trang {page_number(path)}"
            draw.text((x + 8, y + 6), label, fill="black", font=font)
        output = args.output_dir / f"contact-{sheet_index + 1:03d}.png"
        sheet.save(output, "PNG", optimize=True)
        print(output.name)


def page_number(path: Path) -> int:
    return int(path.stem.rsplit("-", 1)[-1])


if __name__ == "__main__":
    main()
