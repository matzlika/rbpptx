#!/usr/bin/env python3
"""Generate a synthetic .pptx for benchmarking.

Run via uv so dependencies don't pollute site-packages:

    uv run --with python-pptx benchmark/fixture_gen.py [SLIDES] [OUT_PATH]

Produces a deck with the requested slide count, each slide carrying a
title placeholder + a body placeholder containing several runs split by
soft line breaks. The runs are tagged with ``{{slot_N}}`` markers so the
benchmark can target individual runs for surgical-edit measurements.
"""

import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt


def build(slides: int, out_path: Path) -> None:
    pres = Presentation()
    blank_layout = pres.slide_layouts[5]
    title_only_layout = pres.slide_layouts[1]

    for i in range(slides):
        slide = pres.slides.add_slide(title_only_layout if i % 2 else blank_layout)
        if slide.shapes.title is not None:
            slide.shapes.title.text = f"Slide {i + 1} title"

        textbox = slide.shapes.add_textbox(
            Inches(1), Inches(2), Inches(8), Inches(4)
        )
        tf = textbox.text_frame
        tf.text = f"slide_{i + 1}_intro"
        for j in range(4):
            p = tf.add_paragraph()
            run_a = p.add_run()
            run_a.text = f"prefix_{i + 1}_{j} "
            run_a.font.size = Pt(14)
            run_b = p.add_run()
            run_b.text = f"{{{{slot_{i + 1}_{j}}}}}"
            run_b.font.size = Pt(14)
            run_b.font.bold = True
            run_c = p.add_run()
            run_c.text = " suffix"
            run_c.font.size = Pt(14)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    pres.save(out_path)


if __name__ == "__main__":
    slide_count = int(sys.argv[1]) if len(sys.argv) > 1 else 50
    target = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("samples/bench.pptx")
    build(slide_count, target)
    print(f"wrote {target} ({slide_count} slides)")
