#!/usr/bin/env python3
"""python-pptx comparison driver invoked by ``benchmark/compare.rb``.

Run directly via uv (the Ruby driver does this for you):

    uv run --with python-pptx benchmark/python_pptx_compare.py SRC SCRATCH_DIR
"""

import json
import os
import resource
import sys
import time
from pathlib import Path

from pptx import Presentation


WARMUP = int(os.environ.get("RBPPTX_BENCH_WARMUP", "1"))
ITERATIONS = int(os.environ.get("RBPPTX_BENCH_ITERATIONS", "5"))


def rss_kb() -> int:
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss


def benchmark(label: str, fn) -> dict:
    for _ in range(WARMUP):
        fn()

    samples = []
    rss_deltas = []
    for _ in range(ITERATIONS):
        before = rss_kb()
        t0 = time.perf_counter()
        fn()
        samples.append(time.perf_counter() - t0)
        rss_deltas.append(rss_kb() - before)

    mean = sum(samples) / len(samples)
    variance = sum((s - mean) ** 2 for s in samples) / len(samples)
    return {
        "label": label,
        "real": mean,
        "real_min": min(samples),
        "real_stddev": variance**0.5,
        "rss_delta_kb": max(rss_deltas),
        "iterations": ITERATIONS,
    }


def read_text(src: Path) -> int:
    pres = Presentation(src)
    total = 0
    for slide in pres.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                total += len(shape.text_frame.text)
    return total


def fill_one(src: Path, dst: Path) -> None:
    pres = Presentation(src)
    slide = pres.slides[0]
    shape = next(s for s in slide.shapes if s.has_text_frame)
    shape.text_frame.text = "Bench Title"
    pres.save(dst)


def fill_many(src: Path, dst: Path, count: int = 10) -> None:
    pres = Presentation(src)
    edited = 0
    for slide in pres.slides:
        if edited >= count:
            break
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            tf = shape.text_frame
            if not tf.paragraphs or not tf.paragraphs[0].runs:
                continue
            tf.paragraphs[0].runs[0].text = f"fill_{edited}"
            edited += 1
            break
    pres.save(dst)


def main() -> None:
    src = Path(sys.argv[1])
    scratch = Path(sys.argv[2])
    scratch.mkdir(parents=True, exist_ok=True)

    results = [
        benchmark("python-pptx read", lambda: read_text(src)),
        benchmark("python-pptx fill_one", lambda: fill_one(src, scratch / "py_one.pptx")),
        benchmark("python-pptx fill_many", lambda: fill_many(src, scratch / "py_many.pptx")),
    ]
    print(json.dumps(results))


if __name__ == "__main__":
    main()
