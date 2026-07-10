#!/usr/bin/env python3
"""Render a .pptx file to PNG via macOS Quick Look (`qlmanage`) so Claude can
read it with the Read tool and visually verify what PowerPoint will display.

Quick Look's renderer is the same engine that produces .pptx previews in
Finder — it lets us close the feedback loop without depending on Microsoft
PowerPoint's flaky AppleScript export. The PNG it produces shows the FIRST
slide of the deck at the requested resolution.

This is the single most important guardrail when adjusting title polish:
fontTools width math drifts from PowerPoint's actual render, so the only
trustworthy verification is "render it and look at it."

Usage:
    python3 scripts/render_for_review.py <pptx_path> [output_dir]

Returns the path to the generated PNG on stdout.
"""
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def _require_qlmanage():
    """Visual verification depends on macOS Quick Look (`qlmanage`). Fail with a
    clear, actionable message on systems where it isn't available, rather than a
    raw FileNotFoundError."""
    if sys.platform != "darwin" or shutil.which("qlmanage") is None:
        raise RuntimeError(
            "Visual verification uses macOS Quick Look (`qlmanage`), which is "
            "not available on this system. The deck was still built correctly — "
            "open the .pptx in PowerPoint to verify it visually."
        )


def render_cover_to_png(pptx_path, output_dir=None, size=1600):
    """Render the cover (slide 1) of `pptx_path` to a PNG via qlmanage and
    return the absolute path. `output_dir` defaults to `<pptx_dir>/_review`.

    Quick Look produces one thumbnail per file — slide 1 — which is exactly
    the slide most likely to need verification (cover title sizing). To check
    other slides, the workflow is: temporarily reorder the deck or extract
    individual slides into their own .pptx and re-run."""
    _require_qlmanage()
    pptx_abs = os.path.abspath(pptx_path)
    if not os.path.exists(pptx_abs):
        raise FileNotFoundError(pptx_abs)

    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(pptx_abs), "_review")
    out_abs = os.path.abspath(output_dir)
    os.makedirs(out_abs, exist_ok=True)

    # qlmanage names the output `<filename>.png` — clear any stale copy so
    # we never read a previous render by accident.
    expected_png = os.path.join(out_abs, os.path.basename(pptx_abs) + ".png")
    if os.path.exists(expected_png):
        os.unlink(expected_png)

    result = subprocess.run(
        ["qlmanage", "-t", "-s", str(size), "-o", out_abs, pptx_abs],
        capture_output=True, text=True, timeout=60,
    )
    if not os.path.exists(expected_png):
        raise RuntimeError(
            f"qlmanage produced no PNG at {expected_png}\n"
            f"stdout: {result.stdout.strip()}\n"
            f"stderr: {result.stderr.strip()}"
        )
    return expected_png


def render_slide_to_png(pptx_path, slide_index, output_dir=None, size=1600):
    """Render slide #`slide_index` (0-based) of `pptx_path` to a PNG via qlmanage.

    Since qlmanage only renders slide 1, we make a temporary copy of the deck
    and delete every slide except the target. The result is a single-slide
    .pptx whose qlmanage thumbnail IS the target slide. Returns the absolute
    PNG path (renamed to `slide-<N>.png` in `output_dir`)."""
    pptx_abs = os.path.abspath(pptx_path)
    if not os.path.exists(pptx_abs):
        raise FileNotFoundError(pptx_abs)

    if slide_index == 0:
        return render_cover_to_png(pptx_path, output_dir, size)

    _require_qlmanage()

    # Lazy imports so the cover-only path stays light
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from pptx import Presentation
    from pptx_helpers import delete_slide

    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(pptx_abs), "_review")
    out_abs = os.path.abspath(output_dir)
    os.makedirs(out_abs, exist_ok=True)

    tmp_dir = tempfile.mkdtemp(prefix="slide_render_")
    try:
        tmp_pptx = os.path.join(tmp_dir, f"only-slide-{slide_index + 1}.pptx")
        shutil.copy(pptx_abs, tmp_pptx)

        prs = Presentation(tmp_pptx)
        n = len(prs.slides)
        if slide_index >= n:
            raise IndexError(f"slide_index {slide_index} out of range (deck has {n} slides)")
        # Delete every other slide, reverse order to keep indices stable
        for i in range(n - 1, -1, -1):
            if i != slide_index:
                delete_slide(prs, i)
        prs.save(tmp_pptx)

        # qlmanage names its output `<pptx_basename>.png`
        ql_png = os.path.join(out_abs, os.path.basename(tmp_pptx) + ".png")
        target_png = os.path.join(out_abs, f"slide-{slide_index + 1}.png")
        for p in (ql_png, target_png):
            if os.path.exists(p):
                os.unlink(p)

        result = subprocess.run(
            ["qlmanage", "-t", "-s", str(size), "-o", out_abs, tmp_pptx],
            capture_output=True, text=True, timeout=60,
        )
        if not os.path.exists(ql_png):
            raise RuntimeError(
                f"qlmanage produced no PNG at {ql_png}\n"
                f"stdout: {result.stdout.strip()}\n"
                f"stderr: {result.stderr.strip()}"
            )
        os.rename(ql_png, target_png)
        return target_png
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    pptx_path = sys.argv[1]
    if len(sys.argv) >= 3 and sys.argv[2].isdigit():
        slide_index = int(sys.argv[2])
        out_dir = sys.argv[3] if len(sys.argv) >= 4 else None
        png = render_slide_to_png(pptx_path, slide_index, out_dir)
    else:
        out_dir = sys.argv[2] if len(sys.argv) >= 3 else None
        png = render_cover_to_png(pptx_path, out_dir)
    print(png)
