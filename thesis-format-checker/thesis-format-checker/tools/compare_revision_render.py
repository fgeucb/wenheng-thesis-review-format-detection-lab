"""Pixel comparison of Word-exported content PDFs, using bundled pypdfium2/Pillow."""
import json
import argparse
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageChops, ImageDraw


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", type=Path, default=Path("samples/_revision_20261012"))
    folder = parser.parse_args().folder
    qa = folder / "word_render"
    before, after = [pdfium.PdfDocument(str(qa / (name + ".pdf"))) for name in ("original", "revised")]
    assert len(before) == len(after)
    comparisons = []
    thumbs = []
    for i in range(len(after)):
        images = [pdf[i].render(scale=2).to_pil().convert("RGB") for pdf in (before, after)]
        bbox = ImageChops.difference(*images).getbbox() if images[0].size == images[1].size else "size"
        texts = [pdf[i].get_textpage() for pdf in (before, after)]
        same_text = texts[0].get_text_range() == texts[1].get_text_range()
        delta = None
        delta_x = delta_y = None
        if same_text and texts[0].count_chars() == texts[1].count_chars():
            boxes = [(texts[0].get_charbox(j), texts[1].get_charbox(j))
                     for j in range(texts[0].count_chars()) if texts[0].get_text_range(j, 1).strip()]
            delta_x = max((abs(a[k] - b[k]) for a, b in boxes for k in (0, 2)), default=0)
            delta_y = max((abs(a[k] - b[k]) for a, b in boxes for k in (1, 3)), default=0)
            delta = max(delta_x, delta_y)
        comparisons.append({"page": i + 1, "identical": bbox is None, "difference_bbox": bbox,
                            "text_identical": same_text, "max_character_box_delta_pt": delta,
                            "max_x_delta_pt": delta_x, "max_y_delta_pt": delta_y})
        images[1].save(qa / f"page-{i + 1}.png")
        if bbox is not None:
            images[0].save(qa / f"original-page-{i + 1}.png")
            ImageChops.difference(*images).save(qa / f"diff-page-{i + 1}.png")
        thumb = images[1].copy()
        thumb.thumbnail((210, 300))
        thumbs.append(thumb)
    for first in range(0, len(thumbs), 12):
        sheet = Image.new("RGB", (4 * 230, 3 * 335), "white")
        draw = ImageDraw.Draw(sheet)
        for index, thumb in enumerate(thumbs[first:first + 12]):
            x, y = (index % 4) * 230, (index // 4) * 335
            sheet.paste(thumb, (x, y + 22))
            draw.text((x + 5, y + 3), f"Page {first + index + 1}", fill="black")
        sheet.save(qa / f"contact-{first + 1}.png")
    payload = {"pages": len(after), "dpi": 144, "identical_pages": sum(x["identical"] for x in comparisons),
               "comparisons": comparisons}
    (folder / "render_verification.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in payload.items() if k != "comparisons"}))


if __name__ == "__main__":
    main()
