"""Approximate PPTX -> PNG preview renderer.

No LibreOffice/PowerPoint on this host, so this re-draws the deck from the OOXML
primitives the builder used: rect/rounded-rect fills, connectors, pictures, tables,
and rich text (runs with their own colour/size/bold).
It is a design preview, not a pixel-exact renderer - good enough to judge
hierarchy, balance, spacing and collisions.
"""
import os
import sys

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

DPI = 96
EMU = 914400.0
REG = r"C:\Windows\Fonts\segoeui.ttf"
BOLD = r"C:\Windows\Fonts\segoeuib.ttf"
_fc = {}


def fnt(pt, bold=False):
    px = max(int(round(pt * DPI / 72)), 1)
    k = (px, bool(bold))
    if k not in _fc:
        _fc[k] = ImageFont.truetype(BOLD if bold else REG, px)
    return _fc[k]


def rgb(c):
    return (c[0], c[1], c[2]) if isinstance(c, tuple) else (int(c[0]), int(c[1]), int(c[2]))


def tokens_of(shape):
    """Flatten paragraphs into token lists: [[(word, font, color), ...], ...]"""
    paras = []
    for p in shape.text_frame.paragraphs:
        toks = []
        for r in p.runs:
            if r.text == "":
                continue
            pt = r.font.size.pt if r.font.size else 18.0
            bold = bool(r.font.bold)
            col = (0, 0, 0)
            try:
                if r.font.color and r.font.color.rgb is not None:
                    col = rgb(r.font.color.rgb)
            except Exception:
                pass
            f = fnt(pt, bold)
            for word in r.text.replace("\n", " \n ").split(" "):
                if word == "\n":
                    toks.append(("\n", f, col))
                elif word:
                    toks.append((word, f, col))
        paras.append((p, toks))
    return paras


def draw_text_shape(d, shape, ox=0, oy=0):
    tf = shape.text_frame
    l = shape.left / EMU + ox
    t = shape.top / EMU + oy
    w = shape.width / EMU
    h = shape.height / EMU
    ml = tf.margin_left / EMU if tf.margin_left is not None else 0.1
    mr = tf.margin_right / EMU if tf.margin_right is not None else 0.1
    mt = tf.margin_top / EMU if tf.margin_top is not None else 0.05
    mb = tf.margin_bottom / EMU if tf.margin_bottom is not None else 0.05
    inner_w = (w - ml - mr) * DPI
    paras = tokens_of(shape)
    if not any(toks for _, toks in paras):
        return

    # lay out
    lines = []          # (segments, height_px, spacing)
    for p, toks in paras:
        if not toks:
            lines.append(([], 0, 1.0, p.alignment))
            continue
        spacing = float(p.line_spacing) if p.line_spacing else 1.0
        cur, curw = [], 0.0
        for tok in toks:
            if tok[0] == "\n":
                lines.append((cur, max(f.size for _, f, _ in cur) if cur else 12, spacing, p.alignment))
                cur, curw = [], 0.0
                continue
            wd, f, col = tok
            space_w = f.getlength(" ") if cur else 0
            tw = f.getlength(wd)
            if cur and curw + space_w + tw > inner_w:
                lines.append((cur, max(x[1].size for x in cur), spacing, p.alignment))
                cur, curw = [(wd, f, col)], tw
            else:
                cur.append((wd, f, col))
                curw += space_w + tw
        if cur:
            lines.append((cur, max(x[1].size for x in cur), spacing, p.alignment))

    heights = []
    for segs, sz, spacing, _ in lines:
        if not segs:
            heights.append(sz * 1.22 * spacing)
        else:
            heights.append(max(f.size for _, f, _ in segs) * 1.22 * spacing)
    total = sum(heights)

    y = (t + mt) * DPI
    anchor = tf.vertical_anchor
    if anchor == MSO_ANCHOR.MIDDLE:
        y = t * DPI + (h * DPI - total) / 2
    elif anchor == MSO_ANCHOR.BOTTOM:
        y = t * DPI + h * DPI - mb * DPI - total

    for (segs, sz, spacing, align), lh in zip(lines, heights):
        if not segs:
            y += lh
            continue
        line_w = sum(f.getlength(wd) for wd, f, _ in segs) + \
            sum(f.getlength(" ") for _, f, _ in segs[1:])
        x = (l + ml) * DPI
        if align == PP_ALIGN.CENTER:
            x = l * DPI + (w * DPI - line_w) / 2
        elif align == PP_ALIGN.RIGHT:
            x = (l + w - mr) * DPI - line_w
        for i, (wd, f, col) in enumerate(segs):
            if i:
                x += f.getlength(" ")
            d.text((x, y + (lh - f.size * 1.22) / 2), wd, font=f, fill=col)
            x += f.getlength(wd)
        y += lh


def render(path, outdir):
    prs = Presentation(path)
    SW = int(prs.slide_width / EMU * DPI)
    SH = int(prs.slide_height / EMU * DPI)
    os.makedirs(outdir, exist_ok=True)
    made = []

    for idx, slide in enumerate(prs.slides, start=1):
        # background
        bg = (255, 255, 255)
        try:
            f = slide.background.fill
            if f.type is not None and f.fore_color and f.fore_color.rgb is not None:
                bg = rgb(f.fore_color.rgb)
        except Exception:
            pass
        img = Image.new("RGB", (SW, SH), bg)
        d = ImageDraw.Draw(img)

        for sh in slide.shapes:
            l = sh.left / EMU * DPI
            t = sh.top / EMU * DPI
            w = sh.width / EMU * DPI
            h = sh.height / EMU * DPI

            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    pic = Image.open(sh.image.blob and __import__("io").BytesIO(sh.image.blob))
                    pic = pic.convert("RGBA").resize((max(int(w), 1), max(int(h), 1)), Image.LANCZOS)
                    img.paste(pic, (int(l), int(t)), pic)
                except Exception as e:
                    d.rectangle([l, t, l + w, t + h], outline=(255, 0, 0))
                    print("pic fail", e)
                continue

            if sh.has_table:
                tbl = sh.table
                cy = t
                for ri, row in enumerate(tbl.rows):
                    cx = l
                    rh = row.height / EMU * DPI
                    for ci, cell in enumerate(row.cells):
                        cw = tbl.columns[ci].width / EMU * DPI
                        fill = (255, 255, 255)
                        try:
                            if cell.fill.type is not None:
                                fill = rgb(cell.fill.fore_color.rgb)
                        except Exception:
                            pass
                        d.rectangle([cx, cy, cx + cw, cy + rh], fill=fill)
                        if cell.text_frame.text.strip():
                            draw_text_shape(d, cell_shape_wrap(cell, cx, cy, cw, rh))
                        cx += cw
                    cy += rh
                continue

            is_conn = sh.shape_type == MSO_SHAPE_TYPE.LINE or sh.__class__.__name__.endswith("Connector")
            if is_conn:
                try:
                    col = rgb(sh.line.color.rgb)
                    d.line([l, t, l + w, t + h], fill=col,
                           width=max(int(sh.line.width.pt * DPI / 72), 1))
                except Exception:
                    pass
                continue

            # autoshape / textbox: fill + outline
            fill = None
            try:
                if sh.fill.type is not None and sh.fill.type == 1:
                    fill = rgb(sh.fill.fore_color.rgb)
            except Exception:
                fill = None
            line = None
            try:
                if sh.line.color and sh.line.color.rgb is not None:
                    line = rgb(sh.line.color.rgb)
            except Exception:
                line = None

            rounded = "ROUNDED" in str(sh.shape_type).upper() or \
                (sh.__class__.__name__ == "Shape" and getattr(sh, "adjustments", None)
                 and len(sh.adjustments) and sh.adjustments[0] > 0)
            if rounded:
                d.rounded_rectangle([l, t, l + w, t + h], radius=min(w, h) * 0.18,
                                    fill=fill, outline=line, width=1 if line else 0)
            else:
                d.rectangle([l, t, l + w, t + h], fill=fill, outline=line,
                            width=1 if line else 0)

            if sh.has_text_frame and sh.text_frame.text.strip():
                draw_text_shape(d, sh)

        p = os.path.join(outdir, f"slide_{idx:02d}.png")
        img.save(p)
        made.append(p)
    return made


class cell_shape_wrap:
    """Adapter so a table cell can reuse the text renderer."""
    def __init__(self, cell, x, y, w, h):
        self._c = cell
        self.text_frame = cell.text_frame
        # x/y/w/h arrive in PIXELS; convert px -> EMU (1 px = 9525 EMU at 96 dpi)
        self.left = int(x * 9525)
        self.top = int(y * 9525)
        self.width = int(w * 9525)
        self.height = int(h * 9525)
        self.shape_type = None
        self.adjustments = []


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "Alora_TheLook_B2B_Growth_Deck.pptx"
    out = sys.argv[2] if len(sys.argv) > 2 else "preview"
    files = render(src, out)
    print("\n".join(files))
    print(len(files), "slides rendered")
