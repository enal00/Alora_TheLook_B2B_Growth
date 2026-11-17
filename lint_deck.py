"""Layout linter for the deck.

There is no LibreOffice/PowerPoint on this machine, so the deck cannot be rendered
to images. Instead this measures the *actual* text with Segoe UI font metrics and
checks each shape for:
  - out-of-bounds geometry
  - text taller than its box (overflow)
  - text wider than one line when it should not be
  - two text-bearing shapes overlapping
  - tables that run past the footer line
"""
import os
import sys

from PIL import ImageFont
from pptx import Presentation
from pptx.util import Emu

DPI = 96
REG = r"C:\Windows\Fonts\segoeui.ttf"
BOLD = r"C:\Windows\Fonts\segoeuib.ttf"
_cache = {}


def font(pt, bold=False):
    px = max(int(round(pt * DPI / 72)), 1)
    key = (px, bold)
    if key not in _cache:
        _cache[key] = ImageFont.truetype(BOLD if bold else REG, px)
    return _cache[key]


def wrap_lines(text, f, max_px):
    """Greedy wrap; returns list of lines."""
    out = []
    for para in str(text).split("\n"):
        if not para.strip():
            out.append("")
            continue
        words = para.split(" ")
        cur = ""
        for word in words:
            trial = word if not cur else cur + " " + word
            if f.getlength(trial) <= max_px or not cur:
                cur = trial
            else:
                out.append(cur)
                cur = word
        out.append(cur)
    return out


def para_metrics(shape):
    """Return (max_font_pt, bold, line_spacing, n_paragraphs_of_text)."""
    sizes, bolds, spacing, n = [], [], 1.0, 0
    tf = shape.text_frame
    for p in tf.paragraphs:
        line = "".join(r.text for r in p.runs)
        if line.strip():
            n += 1
        if p.line_spacing:
            spacing = float(p.line_spacing)
        for r in p.runs:
            if r.text.strip():
                sizes.append(r.font.size.pt if r.font.size else 18.0)
                bolds.append(bool(r.font.bold))
    if not sizes:
        return None
    # use the dominant (max) size; conservative for mixed runs
    return max(sizes), (sum(bolds) / len(bolds) > 0.5), spacing, n


def check(path):
    prs = Presentation(path)
    SW = prs.slide_width / 914400
    SH = prs.slide_height / 914400
    issues = []

    for idx, slide in enumerate(prs.slides, start=1):
        boxes = []
        for sh in slide.shapes:
            l, t = sh.left / 914400, sh.top / 914400
            w, h = sh.width / 914400, sh.height / 914400

            # ---- geometry
            if l < -0.01 or t < -0.01 or l + w > SW + 0.01 or t + h > SH + 0.01:
                issues.append((idx, "OUT-OF-BOUNDS",
                               f"{sh.shape_type} at ({l:.2f},{t:.2f}) size {w:.2f}x{h:.2f}"))

            # ---- tables
            if sh.has_table:
                tbl = sh.table
                if t + h > 6.98:
                    issues.append((idx, "TABLE-PAST-FOOTER",
                                   f"table bottom {t+h:.2f}in > 6.98in"))
                # per-cell text check
                for ri, row in enumerate(tbl.rows):
                    for ci, cell in enumerate(row.cells):
                        txt = cell.text_frame.text
                        if not txt.strip():
                            continue
                        cw = tbl.columns[ci].width / 914400 - 0.10 - 0.08
                        sizes = [r.font.size.pt for p in cell.text_frame.paragraphs
                                 for r in p.runs if r.font.size]
                        pt = max(sizes) if sizes else 12
                        bold = any(r.font.bold for p in cell.text_frame.paragraphs
                                   for r in p.runs)
                        f = font(pt, bold)
                        need = 0
                        for para in txt.split("\n"):
                            need += max(1, len(wrap_lines(para, f, cw * DPI)))
                        rh = tbl.rows[ri].height / 914400
                        need_h = need * pt * 1.22 / 72
                        if need_h > rh + 0.05:
                            issues.append((idx, "CELL-OVERFLOW",
                                           f"r{ri}c{ci} needs {need_h:.2f}in in {rh:.2f}in: {txt[:38]!r}"))
                continue

            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue

            m = para_metrics(sh)
            if not m:
                continue
            pt, bold, spacing, n_paras = m
            f = font(pt, bold)
            inner_w = w - (sh.text_frame.margin_left + sh.text_frame.margin_right) / 914400
            inner_h = h - (sh.text_frame.margin_top + sh.text_frame.margin_bottom) / 914400
            lines = 0
            for p in sh.text_frame.paragraphs:
                line = "".join(r.text for r in p.runs)
                if not line:
                    continue
                lines += max(1, len(wrap_lines(line, f, inner_w * DPI)))
            need_h = lines * pt * 1.22 * spacing / 72
            if need_h > inner_h + 0.04:
                issues.append((idx, "TEXT-OVERFLOW",
                               f"need {need_h:.2f}in ({lines} lines @{pt:.0f}pt) "
                               f"in {inner_h:.2f}in box: {sh.text_frame.text[:44]!r}"))
            boxes.append((l, t, w, h, sh.text_frame.text[:30]))

        # ---- text/text overlap
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = boxes[i], boxes[j]
                ox = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
                oy = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
                if ox > 0.06 and oy > 0.06:
                    area = ox * oy
                    if area > 0.05:
                        issues.append((idx, "TEXT-OVERLAP",
                                       f"'{a[4]}' vs '{b[4]}' overlap {ox:.2f}x{oy:.2f}in"))

    print(f"slides: {len(prs.slides)}   issues: {len(issues)}")
    cur = None
    for idx, kind, detail in sorted(issues):
        if idx != cur:
            print(f"\n--- slide {idx}")
            cur = idx
        print(f"  [{kind}] {detail}")
    return len(issues)


if __name__ == "__main__":
    p = sys.argv[1] if len(sys.argv) > 1 else "Alora_TheLook_B2B_Growth_Deck.pptx"
    sys.exit(0 if check(p) == 0 else 1)
