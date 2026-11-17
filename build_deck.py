"""Build the Alora / TheLook B2B growth deck.

Design: 16:9, restrained consulting layout.
  eyebrow -> headline -> subhead -> [chart + reading panel] -> takeaway -> footer
Every number here is computed from rfm_segmentation_CLEAN.csv / B2B_LEAD_SCORING.csv
or is a published figure of the BigQuery public dataset (marked "approx").
"""
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt
from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
VIZ = os.path.join(BASE, "Visualizations")
OUTFILE = os.path.join(BASE, "Alora_TheLook_B2B_Growth_Deck.pptx")

# ------------------------------------------------------------------ tokens
INK = RGBColor(0x10, 0x14, 0x18)
MUTED = RGBColor(0x57, 0x60, 0x6B)
TEAL = RGBColor(0x0E, 0x7C, 0x86)
TEAL_D = RGBColor(0x0A, 0x5A, 0x62)
AMBER = RGBColor(0xC8, 0x91, 0x2A)
RED = RGBColor(0xB4, 0x43, 0x3C)
DEEP = RGBColor(0x8E, 0x2F, 0x2A)
PANEL = RGBColor(0xF4, 0xF6, 0xF8)
PANEL2 = RGBColor(0xEA, 0xEF, 0xF2)
RULE = RGBColor(0xDD, 0xE2, 0xE7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x0E, 0x12, 0x16)
FONT = "Segoe UI"

W, H = 13.333, 7.5
ML = 0.78                      # left margin
CW = W - 2 * ML                # content width = 11.773
RIGHT_COL = ML + 8.25 + 0.28   # reading panel x
RIGHT_W = W - ML - RIGHT_COL

prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)
BLANK = prs.slide_layouts[6]

SLIDES = []          # (slide, kind) for the footer pass


# ------------------------------------------------------------------ helpers
def new_slide(dark=False):
    s = prs.slides.add_slide(BLANK)
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = DARK if dark else WHITE
    return s


def rect(slide, l, t, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE,
         line_w=0.75, radius=None):
    sh = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_w)
    sh.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = radius
    if sh.has_text_frame:
        sh.text_frame.word_wrap = True
    return sh


def tx(slide, l, t, w, h, text, size=12, bold=False, color=INK,
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False,
       spacing=1.0, space_after=0, font=FONT):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = str(text).split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_after = Pt(space_after)
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.name = font
        r.font.color.rgb = color
    return box


def bullets(slide, l, t, w, h, items, size=11.5, color=INK, marker=TEAL,
            gap=8, spacing=1.12, bold_lead=False):
    """items: list of (lead, rest) or plain strings."""
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = spacing
        p.space_after = Pt(gap)
        rm = p.add_run()
        rm.text = "\u25aa  "
        rm.font.size = Pt(size)
        rm.font.name = FONT
        rm.font.color.rgb = marker
        rm.font.bold = True
        if isinstance(it, tuple):
            lead, rest = it
            if lead:
                r1 = p.add_run()
                r1.text = lead
                r1.font.size = Pt(size)
                r1.font.bold = True
                r1.font.name = FONT
                r1.font.color.rgb = color
            r2 = p.add_run()
            r2.text = rest
            r2.font.size = Pt(size)
            r2.font.name = FONT
            r2.font.color.rgb = color
        else:
            r = p.add_run()
            r.text = it
            r.font.size = Pt(size)
            r.font.name = FONT
            r.font.color.rgb = color
    return box


def eyebrow(slide, text, color=TEAL):
    tx(slide, ML, 0.42, CW, 0.24, text.upper(), size=9.5, bold=True, color=color)


def headline(slide, text, size=25, color=INK, top=0.70, h=0.95):
    return tx(slide, ML, top, CW, h, text, size=size, bold=True, color=color,
              spacing=1.06)


def subhead(slide, text, color=MUTED, top=1.60, size=12.5, w=None):
    return tx(slide, ML, top, w or CW, 0.6, text, size=size, color=color,
              spacing=1.2)


def rule(slide, y, color=RULE, l=None, w=None, weight=0.9):
    sh = slide.shapes.add_connector(1, Inches(l if l is not None else ML),
                                    Inches(y), Inches((l if l is not None else ML) + (w or CW)),
                                    Inches(y))
    sh.line.color.rgb = color
    sh.line.width = Pt(weight)
    return sh


def takeaway(slide, text, accent=TEAL):
    rect(slide, ML, 6.20, CW, 0.58, PANEL)
    rect(slide, ML, 6.20, 0.055, 0.58, accent)
    tx(slide, ML + 0.20, 6.24, CW - 0.35, 0.50,
       text, size=10.5, color=INK, anchor=MSO_ANCHOR.MIDDLE, spacing=1.15)


def footer(slide, n, label="TheLook B2B Growth Diagnostic  \u00b7  M. Zainal Salim"):
    rule(slide, 7.02, RULE, weight=0.75)
    tx(slide, ML, 7.08, 7.5, 0.24, label, size=8, color=MUTED)
    tx(slide, W - ML - 1.2, 7.08, 1.2, 0.24, f"{n:02d}", size=9, bold=True,
       color=INK, align=PP_ALIGN.RIGHT)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def fit_image(slide, path, l, t, w, h, align="center"):
    """Place image inside the box preserving aspect ratio."""
    with Image.open(path) as im:
        iw, ih = im.size
    ar = iw / ih
    box_ar = w / h
    if ar >= box_ar:
        nw = w
        nh = w / ar
    else:
        nh = h
        nw = h * ar
    nx = l + (w - nw) / 2 if align == "center" else l
    ny = t + (h - nh) / 2
    return slide.shapes.add_picture(path, Inches(nx), Inches(ny),
                                    Inches(nw), Inches(nh))


def kpi(slide, l, t, w, h, value, label, sub=None, accent=TEAL):
    rect(slide, l, t, w, h, PANEL)
    rect(slide, l, t, w, 0.05, accent)
    tx(slide, l + 0.20, t + 0.26, w - 0.36, 0.50, value, size=24, bold=True,
       color=accent)
    tx(slide, l + 0.20, t + 0.80, w - 0.36, 0.30, label, size=9.5, bold=True,
       color=INK, spacing=1.1)
    if sub:
        tx(slide, l + 0.20, t + 1.10, w - 0.36, 0.46, sub, size=9, color=MUTED,
           spacing=1.15)


def table(slide, l, t, w, cols, rows, widths, header_size=9.5, body_size=9.5,
          row_h=0.42, head_h=0.36, accent=TEAL, first_bold=True):
    n_rows = len(rows) + 1
    shape = slide.shapes.add_table(n_rows, len(cols), Inches(l), Inches(t),
                                   Inches(w), Inches(head_h + row_h * len(rows)))
    tbl = shape.table
    tbl.first_row = False
    tbl.horz_banding = False
    total = sum(widths)
    for i, wd in enumerate(widths):
        tbl.columns[i].width = Emu(int(Inches(w) * (wd / total)))
    tbl.rows[0].height = Inches(head_h)
    for i in range(1, n_rows):
        tbl.rows[i].height = Inches(row_h)
    for j, c in enumerate(cols):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = INK
        cell.margin_left = Inches(0.10)
        cell.margin_right = Inches(0.08)
        cell.margin_top = cell.margin_bottom = Inches(0.03)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.text = ""
        r = p.add_run()
        r.text = c
        r.font.size = Pt(header_size)
        r.font.bold = True
        r.font.name = FONT
        r.font.color.rgb = WHITE
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if i % 2 else PANEL
            cell.margin_left = Inches(0.10)
            cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = ""
            r = p.add_run()
            r.text = str(val)
            r.font.size = Pt(body_size)
            r.font.name = FONT
            r.font.bold = bool(first_bold and j == 0)
            r.font.color.rgb = INK if (j == 0 or j < len(cols)) else MUTED
            if j == 0:
                r.font.color.rgb = TEAL_D
    return tbl


# ============================================================ 01 · COVER
s = new_slide(dark=True)
rect(s, 0, 0, 0.10, H, TEAL)
tx(s, 1.05, 1.02, 9.0, 0.26, "DATA ANALYTICS PORTFOLIO  \u00b7  END-TO-END CASE STUDY",
   size=10, bold=True, color=RGBColor(0x5E, 0xC7, 0xCF))
tx(s, 1.05, 1.48, 10.6, 1.05, "The Retention Leak", size=54, bold=True, color=WHITE)
tx(s, 1.05, 2.62, 9.6, 0.70,
   "Where $1.52M of historical revenue went quiet \u2014 and what to do in 90 days",
   size=19, color=RGBColor(0xC7, 0xCF, 0xD6), spacing=1.2)
rect(s, 1.05, 3.30, 3.0, 0.035, TEAL)
tx(s, 1.05, 3.86, 6.2, 0.95,
   "A B2B growth diagnostic on TheLook E-commerce (BigQuery public data):\n"
   "SQL modelling, Python segmentation, and a prioritised commercial action plan.",
   size=11.5, color=RGBColor(0x9A, 0xA5, 0xB0), spacing=1.35)
s.shapes.add_picture(os.path.join(VIZ, "cover_ghost.png"),
                     Inches(7.900), Inches(3.550), Inches(5.433), Inches(1.716))
rect(s, 1.05, 5.34, 11.23, 0.012, RGBColor(0x2A, 0x32, 0x3A))
tx(s, 1.05, 5.60, 4.4, 0.30, "M. ZAINAL SALIM", size=11, bold=True, color=WHITE)
tx(s, 1.05, 5.90, 5.2, 0.30, "Data Analyst  \u00b7  Commercial & Growth Analytics",
   size=10, color=RGBColor(0x9A, 0xA5, 0xB0))
tx(s, 7.6, 5.60, 4.8, 0.60,
   "27,655 accounts   \u00b7   $2.74M revenue\n8 analyses   \u00b7   17,655-line SQL / Python pipeline",
   size=10, color=RGBColor(0xB6, 0xC0, 0xC9), align=PP_ALIGN.RIGHT, spacing=1.4)
tx(s, 1.05, 6.62, 11.2, 0.26,
   "BigQuery  \u00b7  SQL  \u00b7  Python (pandas)  \u00b7  matplotlib  \u00b7  Looker Studio  \u00b7  A/B test design",
   size=9.5, color=RGBColor(0x8A, 0x94, 0x9E))
notes(s, "Opening line: 'I did not want to show you a dashboard. I wanted to show you a "
         "decision.' This deck is one commercial question answered end-to-end: where does "
         "revenue leak after the first order, and what is it worth to stop it.")

# ================================================== 02 · EXEC SUMMARY
s = new_slide()
eyebrow(s, "Executive summary")
headline(s, "This business does not have an acquisition problem.\nIt has a repeat-purchase problem.")
kw, kg = (CW - 3 * 0.24) / 4, 0.24
for i, (v, lab, sub, ac) in enumerate([
    ("27,655", "ACCOUNTS ANALYSED", "$2.74M historical revenue", TEAL),
    ("27.1%", "STILL ACTIVE", "ordered within 180 days", AMBER),
    ("87.4%", "BOUGHT ONCE", "median gap: 479 days", RED),
    ("$661K", "REVENUE GONE QUIET", "in top-decile A-tier accounts", DEEP),
]):
    kpi(s, ML + i * (kw + kg), 1.98, kw, 1.72, v, lab, sub, ac)

COL2W = (CW - 0.60) / 2
COL2X = ML + COL2W + 0.60
tx(s, ML, 3.98, COL2W, 0.26, "WHAT THE DATA SAYS", size=9, bold=True, color=MUTED)
bullets(s, ML, 4.30, COL2W, 1.80, [
    ("Revenue is concentrated. ", "Top 10% of accounts = 34.5% of revenue ($945K); top 20% = 52.7%."),
    ("Retention is the leak. ", "87.4% of buyers never placed a second order; 55.5% of revenue sits in accounts silent 365+ days."),
    ("The best cohort is leaving. ", "The 176 'At Risk - High Value' accounts are worth $696 each \u2014 7x the average ($99)."),
], size=10.5, gap=9)

tx(s, COL2X, 3.98, COL2W, 0.26, "WHAT I WOULD DO FIRST", size=9, bold=True, color=TEAL_D)
bullets(s, COL2X, 4.30, COL2W, 1.80, [
    ("Score accounts on value x health. ", "Route 2,773 A-tier accounts to a named-account play, not to a newsletter."),
    ("Run a 30-day win-back sprint. ", "1,933 A-tier quiet accounts = $661K addressable, at near-zero media cost."),
    ("Enter Brazil on margin, not volume. ", "AOV is 49% above benchmark there \u2014 lead with the premium outerwear bundle."),
], size=10.5, gap=9, marker=AMBER)
takeaway(s, "Net: one win-back quarter on the existing base is worth more than a 25% increase in new-customer acquisition \u2014 and it costs a fraction of the media budget.")
footer(s, 2)
notes(s, "Lead with the reframe, not the numbers: acquisition is not the constraint. If 87% of "
         "buyers never return, paying more for new ones compounds the leak.")

# ============================================ 03 · BUSINESS UNDERSTANDING
s = new_slide()
eyebrow(s, "01  \u00b7  Business understanding")
headline(s, "One question drives the whole analysis: where does revenue leak\nafter the first order?")
tx(s, ML, 1.95, CW * 0.44, 0.26, "THE PROBLEM", size=9, bold=True, color=MUTED)
tx(s, ML, 2.25, CW * 0.44, 1.55,
   "The platform spends marketing budget to drive sessions and acquires buyers \u2014 then loses "
   "most of them after a single order. Growth targets are met by buying more traffic rather "
   "than earning more revenue per account, so every quarter starts from zero.",
   size=11, color=INK, spacing=1.35)
tx(s, ML, 3.92, CW * 0.44, 0.26, "WHY IT MATTERS COMMERCIALLY", size=9, bold=True, color=MUTED)
bullets(s, ML, 4.22, CW * 0.44, 1.75, [
    ("CAC is paid once, ", "revenue must be earned repeatedly."),
    ("Repeat buyers are cheaper ", "to serve \u2014 no new acquisition cost."),
    ("A quiet account is not lost revenue; ", "it is deferred revenue with an expiry date."),
], size=10.5, gap=8)

tx(s, ML + CW * 0.48, 1.95, CW * 0.52, 0.26,
   "HOW I SPLIT THE WORK ACROSS THE ROLE'S THREE PILLARS", size=9, bold=True, color=MUTED)
table(s, ML + CW * 0.48, 2.25, CW * 0.52,
      ["Pillar", "Business question", "Analysis"],
      [["Customer", "Who is worth keeping, and who is quietly leaving?",
        "RFM, value x health matrix"],
       ["Operational", "Where does the order pipeline break?",
        "Fulfilment lag, cancellations"],
       ["Market", "Where is revenue per order highest?",
        "Country x category AOV"],
       ["Commercial", "What do we do on Monday?",
        "Playbooks, A/B test, roadmap"]],
      widths=[0.9, 2.0, 1.35], row_h=0.66, head_h=0.34, body_size=9.5)
tx(s, ML + CW * 0.48, 5.42, CW * 0.52, 0.60,
   "Every pillar ends in a number a manager can act on and a metric a dashboard can track.",
   size=9.5, color=MUTED, italic=True, spacing=1.25)
takeaway(s, "Scope discipline: I only kept analyses that change a decision. Anything that was "
            "interesting but not actionable went to the appendix.")
footer(s, 3)
notes(s, "This maps directly to the job description's three pillars. Say it explicitly: "
         "operational, customer, market - and each one ends in a commercial decision.")

# ================================================= 04 · DATA SOURCE
s = new_slide()
eyebrow(s, "02  \u00b7  Data source & scope")
headline(s, "Public retail data, treated like a commercial account base")
subhead(s, "bigquery-public-data.thelook_ecommerce \u2014 a real, live Google dataset, queried "
           "directly rather than downloaded as a static CSV.")
table(s, ML, 2.20, CW * 0.60,
      ["Table", "Volume (approx.)", "What I use it for"],
      [["users", "100K rows", "Account master, geography"],
       ["orders", "125K rows", "Status, timestamps, fulfilment lag"],
       ["order_items", "181K rows", "Revenue, product mix, category"],
       ["products", "29K rows", "Category, cost, margin"],
       ["events", "2.4M rows", "Session funnel \u2014 acquisition stage"],
       ["distribution_centers", "10 rows", "Operational bottleneck analysis"]],
      widths=[1.35, 1.05, 2.1], row_h=0.44, head_h=0.34, body_size=9.5)

rect(s, ML + CW * 0.635, 2.20, CW * 0.365, 3.05, PANEL)
tx(s, ML + CW * 0.635 + 0.24, 2.42, CW * 0.31, 0.26,
   "SCOPE DECISIONS", size=9, bold=True, color=TEAL_D)
bullets(s, ML + CW * 0.635 + 0.24, 2.74, CW * 0.315, 2.35, [
    ("status = 'Complete' only. ", "Cancelled and Returned orders never became cashflow, so they cannot count as revenue."),
    ("27,655 buying accounts. ", "Users with at least one completed order \u2014 the only population where value can be measured."),
    ("Snapshot, not a time series. ", "One row per account: recency, frequency, monetary. Stated as a limitation later."),
], size=9.5, gap=8, spacing=1.18)
takeaway(s, "The filtering choice is the analysis: including cancelled orders would inflate revenue "
            "by orders of magnitude and make every segment ranking wrong.")
footer(s, 4)
notes(s, "If asked about the dataset being synthetic: yes, it is a public Google demo dataset. The "
         "value is that the method transfers - a real client account base has exactly these columns.")

# ============================================ 05 · PIPELINE / TECH STACK
s = new_slide()
eyebrow(s, "03  \u00b7  Cloud environment & tech stack")
headline(s, "Six stages, each one reproducible by someone else")
stages = [
    ("QUERY", "BigQuery\nSQL window functions"),
    ("CLEAN", "BigQuery\nassertion queries"),
    ("MODEL", "Python / pandas\nRFM + tiers"),
    ("SEGMENT", "Python\nvalue x health"),
    ("VISUALISE", "matplotlib\n+ Looker Studio"),
    ("DECIDE", "Memo, deck\n+ experiment plan"),
]
bw = (CW - 5 * 0.06) / 6
for i, (name, tool) in enumerate(stages):
    x = ML + i * (bw + 0.06)
    sh = rect(s, x, 2.35, bw, 1.62, PANEL if i < 5 else TEAL,
              shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    tx(s, x + 0.14, 2.52, bw - 0.28, 0.24, f"0{i+1}", size=9, bold=True,
       color=TEAL if i < 5 else WHITE)
    tx(s, x + 0.14, 2.78, bw - 0.28, 0.26, name, size=10.5, bold=True,
       color=INK if i < 5 else WHITE)
    tx(s, x + 0.14, 3.10, bw - 0.28, 0.70, tool, size=8.5,
       color=MUTED if i < 5 else RGBColor(0xE6, 0xF4, 0xF5), spacing=1.25)

tx(s, ML, 4.35, CW * 0.47, 0.26, "WHY THIS STACK", size=9, bold=True, color=MUTED)
bullets(s, ML, 4.65, CW * 0.465, 1.45, [
    ("BigQuery, not a local CSV. ", "The data never leaves the warehouse; only aggregates are exported."),
    ("Python for the modelling. ", "Percentile tiers and the value x health matrix are version-controlled code, not manual Excel."),
    ("One visual language. ", "Same palette and typography in charts, dashboard and deck."),
], size=10, gap=7)
tx(s, RIGHT_COL, 4.35, RIGHT_W, 0.26, "WHAT THIS PROVES TO A HIRING MANAGER",
   size=9, bold=True, color=TEAL_D)
bullets(s, RIGHT_COL, 4.65, RIGHT_W, 1.45, [
    "SQL strong enough for window functions and CTEs.",
    "Python that does real business logic, not just plotting.",
    "Deliverables a non-technical manager will actually open.",
], size=10, gap=7, marker=AMBER)
takeaway(s, "A portfolio is not proof that I can run a query. It is proof that someone else can "
            "re-run my work and get the same number.")
footer(s, 5)
notes(s, "Expect the question 'why not Power BI?'. Answer: for this deliverable Looker Studio was "
         "free and link-shareable; the modelling logic lives in Python either way.")

# ============================================== 06 · DATA CLEANING
s = new_slide()
eyebrow(s, "04  \u00b7  Data cleaning")
headline(s, "Five defects found before a single insight was claimed")
table(s, ML, 2.18, CW,
      ["#", "Defect", "How I caught it", "Rule applied"],
      [["1", "Future-dated order timestamps",
        "recency came out negative in the RFM extract",
        "Negative recency capped at 0; cause documented"],
       ["2", "Cancelled / Returned orders",
        "status distribution before aggregation",
        "Excluded \u2014 never became cashflow"],
       ["3", "Item-level fan-out",
        "order count inflated after joining order_items",
        "COUNT(DISTINCT order_id)"],
       ["4", "Orphan orders",
        "orders with no matching user record",
        "Inner join; lost volume measured and disclosed"],
       ["5", "Accounts with no completed order",
        "users table vs orders table mismatch",
        "Excluded from RFM, kept for funnel analysis"]],
      widths=[0.35, 2.2, 2.9, 2.9], row_h=0.62, head_h=0.34, body_size=9.5)
tx(s, ML, 5.72, CW, 0.34,
   "Every rule is written down so the numbers can be defended in a review \u2014 and reversed if "
   "the business disagrees with the rule.",
   size=9.5, color=MUTED, italic=True)
takeaway(s, "The negative-recency bug was the most dangerous one: it silently inverted the ranking "
            "and would have put the best customers at the bottom of the list.")
footer(s, 6)
notes(s, "This is the slide that wins credibility. Anyone can run SUM(sale_price). Almost nobody "
         "shows you the assertion queries that caught the broken recency column.")

# ========================================= 07 · FEATURE ENGINEERING
s = new_slide()
eyebrow(s, "05  \u00b7  Data extraction & feature engineering")
headline(s, "Two axes a sales team can act on separately: value, and health")
rect(s, ML, 1.98, CW * 0.33, 3.62, PANEL)
tx(s, ML + 0.24, 2.18, CW * 0.28, 0.24, "THE FEATURES", size=9, bold=True, color=TEAL_D)
groups = [
    ("recency", "DATE_DIFF(CURRENT_DATE(),\nMAX(order_date), DAY)"),
    ("frequency", "COUNT(DISTINCT order_id)"),
    ("monetary", "SUM(sale_price)"),
]
gy = 2.52
for name, formula in groups:
    tx(s, ML + 0.24, gy, CW * 0.28, 0.22, name, size=10, bold=True, color=INK)
    nl = formula.count("\n") + 1
    tx(s, ML + 0.24, gy + 0.24, CW * 0.28, 0.19 * nl, formula,
       size=8.5, color=MUTED, spacing=1.1)
    gy += 0.28 + 0.18 * nl + 0.14
rule(s, gy, RULE, l=ML + 0.24, w=CW * 0.275)
tx(s, ML + 0.24, gy + 0.14, CW * 0.28, 0.22, "TIER THRESHOLDS (PERCENTILE)",
   size=9, bold=True, color=TEAL_D)
bullets(s, ML + 0.24, gy + 0.42, CW * 0.275, 0.95, [
    ("A \u2014 top 10% (P90). ", "monetary \u2265 $220"),
    ("B \u2014 next 20% (P70). ", "monetary \u2265 $111"),
    ("C \u2014 bottom 70%. ", "monetary < $111"),
    ("Health. ", "\u2264180d active \u00b7 \u2264365d dormant \u00b7 >365d churned"),
], size=8.5, gap=5, spacing=1.12)

tx(s, ML + CW * 0.375, 2.16, CW * 0.62, 0.26,
   "THE 3\u00d73 ACCOUNT MATRIX \u2014 EVERY CELL IS A PLAYBOOK, NOT A LABEL",
   size=9, bold=True, color=MUTED)
mrows = [
    ["A \u2014 top 10%", "Expand: named-account\ndirect sales", "Win back: senior AE\n+ commercial offer", "Reactivate:\ntargeted campaign"],
    ["B \u2014 next 20%", "Grow: bundle upsell\n+ case-study nurture", "Nurture: automated\nre-engagement", "Low-touch\nwin-back email"],
    ["C \u2014 bottom 70%", "Retain: self-serve,\nno human cost", "Automate: email only", "Suppress: stop\npaid spend"],
]
tbl = s.shapes.add_table(4, 4, Inches(ML + CW * 0.375), Inches(2.50),
                         Inches(CW * 0.625), Inches(2.36)).table
tbl.first_row = False
tbl.horz_banding = False
wx = [1.15, 1.62, 1.62, 1.45]
for i, wd in enumerate(wx):
    tbl.columns[i].width = Emu(int(Inches(CW * 0.625) * (wd / sum(wx))))
for i in range(4):
    tbl.rows[i].height = Inches(0.62 if i else 0.44)
hdrs = ["VALUE \u2193 / HEALTH \u2192", "ACTIVE (\u2264180d)", "DORMANT (181\u2013365d)", "CHURNED (>365d)"]
for j, h in enumerate(hdrs):
    c = tbl.cell(0, j)
    c.fill.solid()
    c.fill.fore_color.rgb = INK
    c.margin_left = c.margin_right = Inches(0.07)
    c.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = c.text_frame.paragraphs[0].add_run()
    r.text = h
    r.font.size = Pt(8)
    r.font.bold = True
    r.font.name = FONT
    r.font.color.rgb = WHITE
for i, row in enumerate(mrows, start=1):
    for j, v in enumerate(row):
        c = tbl.cell(i, j)
        c.fill.solid()
        if j == 0:
            c.fill.fore_color.rgb = PANEL2
        elif i == 1 and j == 3:
            c.fill.fore_color.rgb = RGBColor(0xFB, 0xEC, 0xEA)
        elif i == 1:
            c.fill.fore_color.rgb = RGBColor(0xE6, 0xF3, 0xF4)
        else:
            c.fill.fore_color.rgb = WHITE
        c.margin_left = c.margin_right = Inches(0.07)
        c.margin_top = c.margin_bottom = Inches(0.02)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = c.text_frame.paragraphs[0]
        p.line_spacing = 1.1
        r = p.add_run()
        r.text = v
        r.font.size = Pt(8)
        r.font.name = FONT
        r.font.bold = (j == 0)
        r.font.color.rgb = TEAL_D if j == 0 else INK
tx(s, ML + CW * 0.375, 5.04, CW * 0.625, 0.42,
   "Read row A, column 'Churned': 1,501 accounts worth $509K that were once top-decile and "
   "have not ordered in a year.",
   size=9.5, color=MUTED, italic=True, spacing=1.2)
takeaway(s, "Separating value from health is the whole point: a $500 account that is active needs "
            "an upsell call, a $500 account that is silent needs a win-back call. One list cannot hold both.")
footer(s, 7)
notes(s, "If pushed on thresholds: they are percentile-based, so they survive a change of currency "
         "or a different dataset without recoding. The 180/365 boundaries are the commercial "
         "definition of 'active this half' and 'lost for a year'.")

# ============================================== 08 · FINDING 1 PARETO
s = new_slide()
eyebrow(s, "06  \u00b7  Key finding 1 of 5  \u2014  customer")
headline(s, "Revenue is concentrated \u2014 but not as extremely as folklore says")
fit_image(s, os.path.join(VIZ, "c4_pareto.png"), ML, 2.00, 8.25, 4.05)
tx(s, RIGHT_COL, 2.00, RIGHT_W, 0.26, "WHAT THIS MEANS", size=9, bold=True, color=TEAL_D)
bullets(s, RIGHT_COL, 2.32, RIGHT_W, 2.35, [
    ("Top 10% \u2192 34.5%. ", "2,773 accounts carry $945K of $2.74M."),
    ("Top 20% \u2192 52.7%. ", "Half the revenue sits in a fifth of the base."),
    ("Not 80/20. ", "The bottom 80% still generates 47% of revenue \u2014 it is not dead weight."),
], size=10, gap=9)
rect(s, RIGHT_COL, 4.86, RIGHT_W, 1.05, PANEL)
tx(s, RIGHT_COL + 0.18, 4.98, RIGHT_W - 0.36, 0.85,
   "Implication: account farming alone cannot hit a growth target. The long tail has to be "
   "monetised by automation, not by sales time.",
   size=9.5, color=INK, spacing=1.25)
takeaway(s, "I checked the 80/20 myth rather than assuming it. The honest answer \u2014 34/10, not "
            "80/20 \u2014 changes the strategy: automate the tail, protect the head.")
footer(s, 8)
notes(s, "This is where a weaker analyst says 'Pareto confirmed'. Say instead: I tested it, it did "
         "not hold, and here is what the real curve means.")

# ========================================== 09 · FINDING 2 ACTIVITY
s = new_slide()
eyebrow(s, "07  \u00b7  Key finding 2 of 5  \u2014  customer")
headline(s, "Only 27% of the account base is still buying")
fit_image(s, os.path.join(VIZ, "c6b_health_split.png"), ML, 2.00, 8.25, 4.05)
tx(s, RIGHT_COL, 2.00, RIGHT_W, 0.26, "THE DIAGNOSIS", size=9, bold=True, color=TEAL_D)
bullets(s, RIGHT_COL, 2.32, RIGHT_W, 2.55, [
    ("Median account: 479 days quiet. ", "Longer than a full annual buying cycle."),
    ("55.5% of revenue is 365+ days old. ", "$1.52M earned by accounts that have stopped ordering."),
    ("87.4% bought exactly once. ", "The first order converts; the second order does not."),
    ("Average frequency: 1.13 orders. ", "This is a trial product, not a habit."),
], size=10, gap=8)
rect(s, RIGHT_COL, 5.02, RIGHT_W, 0.95, PANEL)
tx(s, RIGHT_COL + 0.18, 5.13, RIGHT_W - 0.36, 0.75,
   "The commercial constraint is second-order conversion \u2014 not traffic, not assortment.",
   size=9.5, color=INK, spacing=1.25)
takeaway(s, "Reframe for a manager: the marketing spend is buying first orders. Nobody owns the "
            "second order, and that is where the compounding is.")
footer(s, 9)
notes(s, "Careful framing: this dataset is a snapshot, so 'churned' means 'no order in 365 days at "
         "extract time'. Say that plainly - it is the honest limitation of a snapshot RFM.")

# ======================================== 10 · FINDING 3 MATRIX
s = new_slide()
eyebrow(s, "08  \u00b7  Key finding 3 of 5  \u2014  customer")
headline(s, "$661K is sitting in top-decile accounts that went quiet")
fit_image(s, os.path.join(VIZ, "c6_account_matrix.png"), ML, 2.00, 8.25, 4.05)
tx(s, RIGHT_COL, 2.00, RIGHT_W, 0.26, "THE ADDRESSABLE POOL", size=9, bold=True, color=TEAL_D)
tx(s, RIGHT_COL, 2.30, RIGHT_W, 0.60, "$660,556", size=25, bold=True, color=RED)
tx(s, RIGHT_COL, 2.94, RIGHT_W, 0.30, "1,933 A-tier accounts, dormant or churned",
   size=9.5, color=MUTED, spacing=1.15)
rule(s, 3.36, RULE, l=RIGHT_COL, w=RIGHT_W)
bullets(s, RIGHT_COL, 3.50, RIGHT_W, 2.30, [
    ("Best single cell: ", "A-tier churned \u2014 1,501 accounts, $509K."),
    ("Smaller but richer: ", "A-tier dormant \u2014 432 accounts, $151K, still inside a year."),
    ("Win-back economics: ", "recovering 10% of that pool = $66K, with no media spend."),
], size=9.5, gap=8)
takeaway(s, "This matrix is the deliverable a sales manager can run tomorrow: nine lists, nine "
            "playbooks, one owner per cell.")
footer(s, 10)
notes(s, "The 10% recovery assumption is deliberately conservative and labelled as such. In the "
         "interview, say: I would rather commit to a small number I can defend.")

# ======================================= 11 · FINDING 4 SEGMENTS
s = new_slide()
eyebrow(s, "09  \u00b7  Key finding 4 of 5  \u2014  customer")
headline(s, "The most valuable segment per account is the one already leaving")
fit_image(s, os.path.join(VIZ, "c2_value_map.png"), ML, 2.00, 8.25, 4.05)
tx(s, RIGHT_COL, 2.00, RIGHT_W, 0.26, "SEGMENT READ", size=9, bold=True, color=TEAL_D)
bullets(s, RIGHT_COL, 2.32, RIGHT_W, 2.70, [
    ("At Risk \u2014 High Value: ", "176 accounts, $696 average lifetime value. Highest of any segment."),
    ("Champions: 57 accounts only. ", "$238 each, 3x purchase frequency \u2014 the head is thin."),
    ("Hibernating: 19,996 accounts. ", "66.9% of revenue, $92 each \u2014 a volume story, not a value story."),
], size=10, gap=9)
rect(s, RIGHT_COL, 5.06, RIGHT_W, 0.92, PANEL)
tx(s, RIGHT_COL + 0.18, 5.16, RIGHT_W - 0.36, 0.72,
   "Protect the 176 first. They are the smallest list and the largest value per head.",
   size=9.5, color=INK, spacing=1.25)
takeaway(s, "A small segment can out-earn a huge one per account. Ranking segments by total revenue "
            "alone would have buried the most profitable cohort.")
footer(s, 11)
notes(s, "Watch for the challenge: At Risk has higher value partly because high spenders churn "
         "harder. Acknowledge it - that is exactly why they need a playbook.")

# ========================================= 12 · FINDING 5 MARKET
s = new_slide()
eyebrow(s, "10  \u00b7  Key finding 5 of 5  \u2014  market")
headline(s, "Brazil converts 49% more revenue per order than the global benchmark")
fit_image(s, os.path.join(VIZ, "c7_market_aov.png"), ML, 2.00, 8.25, 4.05)
tx(s, RIGHT_COL, 2.00, RIGHT_W, 0.26, "MARKET READ", size=9, bold=True, color=TEAL_D)
bullets(s, RIGHT_COL, 2.32, RIGHT_W, 2.45, [
    ("Brazil \u00b7 Outerwear & Coats: ", "AOV $153.22 on 940 buyers, $147.5K revenue."),
    ("China: ", "largest revenue pool ($339.2K) but a lower AOV \u2014 a volume market."),
    ("Different plays. ", "China = defend share. Brazil = premium entry, not discount entry."),
], size=10, gap=9)
rect(s, RIGHT_COL, 4.92, RIGHT_W, 1.02, PANEL)
tx(s, RIGHT_COL + 0.18, 5.03, RIGHT_W - 0.36, 0.82,
   "Entering Brazil with the same discount playbook used in China would destroy the one "
   "advantage the market has: price per order.",
   size=9.5, color=INK, spacing=1.25)
takeaway(s, "Market choice is a pricing decision, not a geography decision. Brazil is worth more "
            "per order, and that is the number an expansion plan should be built on.")
footer(s, 12)
notes(s, "These two figures came from the country x category aggregation I ran in BigQuery. If "
         "asked for the query, offer to walk through the window function and the AOV join.")

# ============================================== 13 · A/B TEST
s = new_slide()
eyebrow(s, "11  \u00b7  Experiment design")
headline(s, "The next question is causal: does a premium bundle lift AOV,\nor just kill conversion?")
rect(s, ML, 1.95, CW * 0.315, 2.30, PANEL)
tx(s, ML + 0.22, 2.10, CW * 0.27, 0.26, "HYPOTHESIS", size=9, bold=True, color=TEAL_D)
tx(s, ML + 0.22, 2.40, CW * 0.275, 1.56,
   "Adding a curated 'Outerwear + Accessories' bundle module to the product page raises "
   "average order value without reducing conversion rate.",
   size=10.5, color=INK, spacing=1.3)
tx(s, ML + 0.22, 4.10, CW * 0.275, 0.20, "Control: current PDP  \u00b7  Treatment: bundle module + badge",
   size=8.5, color=MUTED, spacing=1.2)

table(s, ML + CW * 0.345, 1.95, CW * 0.655,
      ["Design element", "Decision"],
      [["Primary metric", "Average order value per session"],
       ["Guardrail metrics", "Conversion rate, refund rate (must not degrade)"],
       ["Baseline conversion", "3.0% \u2014 to be validated on 30 days of event data"],
       ["Minimum detectable effect", "+20% relative (3.0% \u2192 3.6%)"],
       ["Sample size", "13,914 per arm \u00b7 27,828 total (95% CI, 80% power)"],
       ["Duration", "\u2248 10 weeks at 3,000 sessions/week"],
       ["Ship rule", "AOV lift \u2265 MDE and CVR not down > 0.3pp"],
       ["Kill rule", "Guardrail breach at 95% CI, or no lift after full sample"]],
      widths=[1.55, 2.85], row_h=0.40, head_h=0.32, body_size=9)
takeaway(s, "The point of the test is not the bundle. It is proving the team can settle a commercial "
            "argument with a number instead of a meeting.")
footer(s, 13)
notes(s, "Have the sample-size maths ready: two-proportion test, alpha 0.05, power 0.80. If they "
         "ask why 20% MDE: anything smaller is not worth the engineering cost.")

# ============================================== 14 · DASHBOARD
s = new_slide()
eyebrow(s, "12  \u00b7  Dashboard & measurement")
headline(s, "Six tiles, one owner, one weekly review")
tiles = [
    ("Revenue by value tier", "A / B / C share of revenue", "Commercial lead", TEAL),
    ("Active account rate", "% ordered in last 180 days", "CRM owner", TEAL),
    ("Revenue at risk ($)", "A-tier dormant + churned pool", "Sales manager", RED),
    ("AOV by market", "Brazil vs China vs benchmark", "Market analyst", AMBER),
    ("Win-back conversion", "% of contacted accounts reactivated", "Campaign owner", AMBER),
    ("Experiment lift", "AOV and CVR delta with CI", "Analytics", TEAL),
]
tw, th, gx = (CW - 2 * 0.26) / 3, 1.42, 0.26
for i, (title, metric, owner, ac) in enumerate(tiles):
    col, row = i % 3, i // 3
    x = ML + col * (tw + gx)
    y = 2.00 + row * (th + 0.24)
    rect(s, x, y, tw, th, PANEL)
    rect(s, x, y, 0.05, th, ac)
    tx(s, x + 0.22, y + 0.16, tw - 0.40, 0.28, title, size=10.5, bold=True, color=INK)
    tx(s, x + 0.22, y + 0.50, tw - 0.40, 0.34, metric, size=9, color=MUTED, spacing=1.15)
    tx(s, x + 0.22, y + th - 0.34, tw - 0.40, 0.24, f"OWNER  {owner.upper()}", size=7.5,
       bold=True, color=ac)
tx(s, ML, 5.42, CW, 0.60,
   "Cadence: Monday review of the six tiles \u2192 one decision per tile \u2192 logged in a decision "
   "journal. A tile with no owner does not go on the dashboard.",
   size=9.5, color=MUTED, italic=True, spacing=1.3)
takeaway(s, "A dashboard is not a report. Every tile names the person who owns it and the action "
            "it is allowed to trigger.")
footer(s, 14)
notes(s, "If they ask what tool: Looker Studio for shareability, BigQuery scheduled queries for "
         "freshness. Tool choice is the least interesting part - ownership is the interesting part.")

# ============================================== 15 · ROADMAP
s = new_slide()
eyebrow(s, "13  \u00b7  Recommendations")
headline(s, "90 days, sequenced by cost of delay")
phases = [
    ("DAYS 0\u201330", "Stop the leak", [
        ("Rebuild the account matrix ", "in BigQuery as a scheduled table \u2014 one source of truth."),
        ("Protect the 176 At-Risk accounts. ", "Named AE outreach, executive-level offer."),
        ("Instrument the funnel. ", "Log second-order conversion so it can be tracked weekly."),
    ], TEAL),
    ("DAYS 31\u201360", "Recover the base", [
        ("Win-back sprint on 1,933 A-tier accounts ", "\u2014 $661K addressable pool."),
        ("Bundle pilot for Brazil entry ", "at premium price point, not discounted."),
        ("Automate the C tier. ", "Suppress paid spend on churned bottom-70% accounts."),
    ], AMBER),
    ("DAYS 61\u201390", "Prove the lift", [
        ("Run the A/B test ", "to full sample \u2014 27,828 sessions across both arms."),
        ("Report the causal result ", "to management with CI, not a single point estimate."),
        ("Rebuild targets ", "on repeat revenue, not new-customer count."),
    ], TEAL_D),
]
pw = (CW - 2 * 0.28) / 3
for i, (phase, theme, items, ac) in enumerate(phases):
    x = ML + i * (pw + 0.28)
    rect(s, x, 1.98, pw, 0.62, ac)
    tx(s, x + 0.20, 2.06, pw - 0.40, 0.22, phase, size=8.5, bold=True,
       color=RGBColor(0xE8, 0xF2, 0xF3))
    tx(s, x + 0.20, 2.28, pw - 0.40, 0.26, theme, size=12, bold=True, color=WHITE)
    rect(s, x, 2.60, pw, 2.98, PANEL)
    bullets(s, x + 0.20, 2.78, pw - 0.40, 2.70, items, size=9.5, gap=9, spacing=1.18)
takeaway(s, "Sequencing logic: recover money already earned before spending money on new demand. "
            "Day 30 pays for day 60.")
footer(s, 15)
notes(s, "The sequence is the argument: win-back has the shortest payback, so it funds the "
         "experiment. Lead with payback, not with ambition.")

# ============================================== 16 · LIMITATIONS
s = new_slide()
eyebrow(s, "14  \u00b7  Limitations & next steps")
headline(s, "What this analysis cannot tell you \u2014 stated up front")
table(s, ML, 2.06, CW,
      ["Limitation", "Why it matters", "How I would close it"],
      [["Snapshot RFM, not a time series",
        "Cannot separate true churn from seasonality",
        "Monthly cohort table in BigQuery; retention curves by cohort"],
       ["No contract or invoice data",
        "B2B account mapping is a proxy built on retail orders",
        "Add CRM/contract table; re-derive tiers on contracted value"],
       ["No margin or cost-to-serve",
        "Value ranking is revenue-based, not profit-based",
        "Join product cost; rank on gross margin contribution"],
       ["A/B test designed, not executed",
        "The causal claim is a hypothesis, not a result",
        "Run to full sample; report effect size with confidence interval"],
       ["Event funnel not yet modelled",
        "Acquisition leakage is inferred, not measured",
        "Funnel query on events: session \u2192 add-to-cart \u2192 purchase"]],
      widths=[2.3, 2.7, 3.1], row_h=0.66, head_h=0.34, body_size=9.5)
tx(s, ML, 6.02, CW, 0.24,
   "Stating limits is not a weakness in a portfolio. It is the difference between an analyst and "
   "a chart generator.", size=9.5, color=MUTED, italic=True)
footer(s, 16)
notes(s, "Volunteer the limitations before they are found. It signals that I know the difference "
         "between what the data supports and what I want it to say.")

# ============================================== 17 · CLOSING
s = new_slide(dark=True)
rect(s, 0, 0, 0.10, H, TEAL)
tx(s, 1.05, 1.10, 10.0, 0.26, "CLOSING", size=10, bold=True,
   color=RGBColor(0x5E, 0xC7, 0xCF))
tx(s, 1.05, 1.55, 10.8, 1.00, "Three things I would bring on day one",
   size=34, bold=True, color=WHITE, spacing=1.1)
items = [
    ("01", "A decision habit, not a report",
     "Every analysis ends in a named owner, a metric and a date."),
    ("02", "Commercial framing of technical work",
     "Percentiles and window functions only matter when they change who gets called."),
    ("03", "Honest numbers under pressure",
     "I test the 80/20 myth instead of quoting it, and I show the leaked revenue before the wins."),
]
for i, (num, title, body) in enumerate(items):
    y = 2.92 + i * 1.20
    tx(s, 1.05, y, 0.75, 0.40, num, size=15, bold=True, color=TEAL)
    tx(s, 1.85, y - 0.02, 8.6, 0.32, title, size=14.5, bold=True, color=WHITE)
    tx(s, 1.85, y + 0.32, 8.9, 0.52, body, size=10.5,
       color=RGBColor(0x9A, 0xA5, 0xB0), spacing=1.3)
rect(s, 1.05, 6.42, 3.0, 0.035, TEAL)
tx(s, 1.05, 6.62, 6.0, 0.30, "M. ZAINAL SALIM  \u00b7  Data Analyst",
   size=10, bold=True, color=WHITE)
tx(s, 7.4, 6.62, 4.9, 0.30, "github.com/enal00  \u00b7  portfolio & full SQL",
   size=9.5, color=RGBColor(0x9A, 0xA5, 0xB0), align=PP_ALIGN.RIGHT)
notes(s, "Close on the habit, not the tooling. The last line to say out loud: 'I would rather "
         "bring you a defended $66K than an optimistic $6M.'")

prs.save(OUTFILE)
print("saved:", OUTFILE)
print("slides:", len(prs.slides))
