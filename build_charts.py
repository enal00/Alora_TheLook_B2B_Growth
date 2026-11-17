"""Regenerate all charts from REAL data with a consistent premium style."""
import csv
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "Visualizations")
os.makedirs(OUT, exist_ok=True)

INK = "#101418"
MUTED = "#6B7480"
TEAL = "#0E7C86"
AMBER = "#C8912A"
RED = "#B4433C"
GRID = "#E3E7EC"
PALETTE = [TEAL, AMBER, "#4C7EA8", RED, "#8B8F96"]

plt.rcParams.update({
    "font.family": "Segoe UI",
    "font.size": 11,
    "axes.edgecolor": GRID,
    "axes.labelcolor": MUTED,
    "axes.titlecolor": INK,
    "axes.titleweight": "bold",
    "axes.titlesize": 14,
    "xtick.color": MUTED,
    "ytick.color": MUTED,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "savefig.facecolor": "white",
})


# ---------------------------------------------------------------- load data
def load_rows(path):
    rows = []
    with open(path, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            rows.append({
                "user_id": r["user_id"],
                "recency": int(float(r["recency"])),
                "frequency": int(float(r["frequency"])),
                "monetary": float(r["monetary"]),
                "segment": r["segment"],
            })
    return rows


rows = load_rows(os.path.join(BASE, "rfm_segmentation_CLEAN.csv"))
N = len(rows)
TOTAL_REV = sum(r["monetary"] for r in rows)


def pct_rank(values, higher_is_better=True):
    """Percentile rank 0-100, ties averaged."""
    order = sorted(range(len(values)), key=lambda i: values[i])
    ranks = [0.0] * len(values)
    i = 0
    while i < len(order):
        j = i
        while j + 1 < len(order) and values[order[j + 1]] == values[order[i]]:
            j += 1
        avg_rank = (i + j) / 2.0
        for k in range(i, j + 1):
            ranks[order[k]] = avg_rank
        i = j + 1
    out = [x / max(len(values) - 1, 1) * 100 for x in ranks]
    return out if higher_is_better else [100 - x for x in out]


# ------------------------------- B2B account model: value tier x health band
# Two independent axes, because a sales team acts on them separately:
#   value tier  -> how much the account is worth (monetary percentile)
#   health band -> whether the account is still buying (recency)
mon_sorted_all = sorted(r["monetary"] for r in rows)
P90 = mon_sorted_all[int(N * 0.90) - 1]
P70 = mon_sorted_all[int(N * 0.70) - 1]


def value_tier(m):
    if m >= P90:
        return "A"
    if m >= P70:
        return "B"
    return "C"


def health_band(rec):
    if rec <= 180:
        return "Active"
    if rec <= 365:
        return "Dormant"
    return "Churned"


PLAYBOOK = {
    ("A", "Active"): "Expand - named-account direct sales",
    ("A", "Dormant"): "Win back - senior AE outreach + commercial offer",
    ("A", "Churned"): "Reactivate - targeted campaign, re-check product fit",
    ("B", "Active"): "Grow - bundle upsell + case-study nurture",
    ("B", "Dormant"): "Nurture - automated re-engagement with offer",
    ("B", "Churned"): "Low-touch win back - email sequence",
    ("C", "Active"): "Retain - self-serve, content, no cost to serve",
    ("C", "Dormant"): "Automate - email only, no human time",
    ("C", "Churned"): "Suppress - stop paid acquisition spend",
}

for r in rows:
    r["value_tier"] = value_tier(r["monetary"])
    r["health"] = health_band(r["recency"])
    r["playbook"] = PLAYBOOK[(r["value_tier"], r["health"])]

with open(os.path.join(BASE, "B2B_LEAD_SCORING.csv"), "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f)
    w.writerow(["user_id", "recency", "frequency", "monetary", "segment",
                "value_tier", "health_band", "recommended_playbook"])
    for r in rows:
        w.writerow([r["user_id"], r["recency"], r["frequency"], r["monetary"],
                    r["segment"], r["value_tier"], r["health"], r["playbook"]])

# ---------------------------------------------------------------- aggregates
segs = {}
for r in rows:
    s = segs.setdefault(r["segment"], {"n": 0, "rev": 0.0, "mon": [], "fr": [], "rec": []})
    s["n"] += 1
    s["rev"] += r["monetary"]
    s["mon"].append(r["monetary"])
    s["fr"].append(r["frequency"])
    s["rec"].append(r["recency"])

tier_stats = {}
for r in rows:
    t = tier_stats.setdefault(r["value_tier"], {"n": 0, "rev": 0.0, "mon": []})
    t["n"] += 1
    t["rev"] += r["monetary"]
    t["mon"].append(r["monetary"])

cell = {}
for r in rows:
    c = cell.setdefault((r["value_tier"], r["health"]), {"n": 0, "rev": 0.0})
    c["n"] += 1
    c["rev"] += r["monetary"]

print("users:", N, "| total revenue:", round(TOTAL_REV, 2))
for s, v in sorted(segs.items(), key=lambda x: -x[1]["rev"]):
    print(f"  {s:26} n={v['n']:6} rev={v['rev']:11.2f} ({v['rev']/TOTAL_REV*100:4.1f}%)")
for t in ["A", "B", "C"]:
    v = tier_stats[t]
    print(f"  Tier {t:19} n={v['n']:6} rev={v['rev']:11.2f} ({v['rev']/TOTAL_REV*100:4.1f}%)")


def save(fig, name):
    fig.savefig(os.path.join(OUT, name), dpi=200, bbox_inches="tight", pad_inches=0.25)
    plt.close(fig)
    print("saved", name)


def style_axes(ax, grid_axis="y"):
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(GRID)
    ax.grid(axis=grid_axis, color=GRID, linewidth=0.8, alpha=0.9)
    ax.set_axisbelow(True)


# ---------------------------------------------------- C1: revenue by segment
seg_order = sorted(segs, key=lambda s: segs[s]["rev"])
fig, ax = plt.subplots(figsize=(9, 4.6))
vals = [segs[s]["rev"] / 1e6 for s in seg_order]
colors = [RED if "Risk" in s else (TEAL if "Champion" in s or "Loyal" in s else "#9AA4B0") for s in seg_order]
bars = ax.barh(seg_order, vals, color=colors, height=0.62)
for b, s in zip(bars, seg_order):
    share = segs[s]["rev"] / TOTAL_REV * 100
    ax.text(max(b.get_width(), 0.03) + 0.02, b.get_y() + b.get_height() / 2,
            f"${segs[s]['rev']/1e6:.2f}M  ({share:.1f}%)",
            va="center", fontsize=10, color=INK, fontweight="bold")
ax.set_xlabel("Total revenue (USD, millions)")
ax.set_xlim(0, max(vals) * 1.32)
ax.set_title(f"Revenue contribution by customer segment  |  total ${TOTAL_REV/1e6:.2f}M")
style_axes(ax, "x")
save(fig, "c1_revenue_by_segment.png")

# ------------------------------------- C2: segment profile (avg value + freq)
order2 = sorted(segs, key=lambda s: np.mean(segs[s]["mon"]))
labels2 = [s.replace(" (High Value)", " - High Value") for s in order2]
avg_mon = [np.mean(segs[s]["mon"]) for s in order2]
avg_fr = [np.mean(segs[s]["fr"]) for s in order2]
counts2 = [segs[s]["n"] for s in order2]
cols2 = [RED if "Risk" in s else (TEAL if s in ("Champions", "Loyal Customers") else "#9AA4B0") for s in order2]

fig, (axA, axB) = plt.subplots(1, 2, figsize=(10.4, 4.5), gridspec_kw={"width_ratios": [1.35, 1]})

bars = axA.barh(labels2, avg_mon, color=cols2, height=0.6)
for b, v, c in zip(bars, avg_mon, counts2):
    axA.text(b.get_width() + max(avg_mon) * 0.02, b.get_y() + b.get_height() / 2,
             f"${v:,.0f}  ·  {c:,} accts", va="center", fontsize=9.5,
             fontweight="bold", color=INK)
axA.set_xlabel("Avg lifetime value (USD)")
axA.set_xlim(0, max(avg_mon) * 1.42)
axA.set_title("Average customer value")
style_axes(axA, "x")

bars2 = axB.barh(labels2, avg_fr, color=cols2, height=0.6)
for b, v in zip(bars2, avg_fr):
    axB.text(b.get_width() + 0.05, b.get_y() + b.get_height() / 2, f"{v:.2f}x",
             va="center", fontsize=9.5, fontweight="bold", color=INK)
axB.set_xlabel("Avg orders per customer")
axB.set_xlim(0, max(avg_fr) * 1.30)
axB.set_title("Purchase intensity")
axB.set_yticklabels([])
style_axes(axB, "x")

fig.suptitle("Segment profile  |  value is concentrated, frequency is thin",
             fontsize=13, fontweight="bold", color=INK, y=1.02)
save(fig, "c2_value_map.png")

# ---------------------------------------------------------- C3: RFM heatmap
def bucket_rec(x):
    for lim, lab in [(30, "0-30"), (90, "31-90"), (180, "91-180"), (365, "181-365"), (730, "1-2 yr"), (10**9, "2 yr+")]:
        if x <= lim:
            return lab
    return "2 yr+"


def bucket_fr(x):
    for lim, lab in [(1, "1 order"), (2, "2"), (3, "3"), (10, "4+")]:
        if x <= lim:
            return lab
    return "4+"


r_labs = ["0-30", "31-90", "91-180", "181-365", "1-2 yr", "2 yr+"]
f_labs = ["1 order", "2", "3", "4+"]
grid = np.full((len(r_labs), len(f_labs)), np.nan)
cnt = np.zeros_like(grid)
for r in rows:
    i = r_labs.index(bucket_rec(r["recency"]))
    j = f_labs.index(bucket_fr(r["frequency"]))
    grid[i, j] = r["monetary"] if np.isnan(grid[i, j]) else grid[i, j] + r["monetary"]
    cnt[i, j] += 1
avg = grid / np.where(cnt == 0, np.nan, cnt)

fig, ax = plt.subplots(figsize=(9, 4.8))
im = ax.imshow(avg, cmap="YlGnBu", aspect="auto")
ax.set_xticks(range(len(f_labs)))
ax.set_xticklabels(f_labs)
ax.set_yticks(range(len(r_labs)))
ax.set_yticklabels(r_labs)
ax.set_xlabel("Frequency")
ax.set_ylabel("Recency (days since last purchase)")
ax.set_title("Where the money sits  |  avg lifetime value by RFM cell")
for i in range(len(r_labs)):
    for j in range(len(f_labs)):
        if not np.isnan(avg[i, j]):
            ax.text(j, i, f"${avg[i,j]:,.0f}\nn={int(cnt[i,j])}", ha="center", va="center",
                    fontsize=9, color="white" if avg[i, j] > np.nanmean(avg) * 1.4 else INK)
cb = fig.colorbar(im, ax=ax, pad=0.02)
cb.set_label("Avg customer value (USD)", color=MUTED, fontsize=9)
cb.ax.tick_params(labelsize=8, colors=MUTED)
cb.outline.set_visible(False)
for s in ax.spines.values():
    s.set_visible(False)
save(fig, "c3_rfm_heatmap.png")

# ------------------------------------------------------------- C4: Pareto
mon_sorted = sorted((r["monetary"] for r in rows), reverse=True)
cum = np.cumsum(mon_sorted) / TOTAL_REV * 100
x = np.arange(1, N + 1) / N * 100
fig, ax = plt.subplots(figsize=(9, 4.4))
ax.plot(x, cum, color=TEAL, linewidth=2.4)
ax.plot([0, 100], [0, 100], color=MUTED, linestyle=":", linewidth=1.2)
ax.text(73, 64, "perfect equality", rotation=33, fontsize=8.5, color=MUTED, style="italic")
for q, off in ((10, (12, 18)), (20, (12, -44)), (30, (12, -44))):
    y = cum[int(N * q / 100) - 1]
    ax.plot([q, q], [0, y], color=AMBER, linestyle="--", linewidth=1.1, alpha=0.8)
    ax.plot([0, q], [y, y], color=AMBER, linestyle="--", linewidth=1.1, alpha=0.8)
    ax.scatter([q], [y], color=AMBER, zorder=5, s=45, edgecolor="white", linewidth=1.2)
    ax.annotate(f"top {q}% of customers\n= {y:.1f}% of revenue", (q, y),
                textcoords="offset points", xytext=off, fontsize=9.5,
                color=INK, fontweight="bold", linespacing=1.4,
                bbox=dict(boxstyle="round,pad=0.3", facecolor="white",
                          edgecolor=GRID, alpha=0.93))
ax.set_xlabel("Customers, ranked by lifetime value (%)")
ax.set_ylabel("Cumulative revenue (%)")
ax.set_title("Revenue concentration  |  revenue per customer rank")
ax.set_xlim(0, 100)
ax.set_ylim(0, 101)
ax.text(0, -0.26, f"n = {N:,} customers · $2.74M revenue · TheLook order_items where status = Complete",
        fontsize=8, color=MUTED, transform=ax.transAxes)
style_axes(ax, "both")
save(fig, "c4_pareto.png")

# ------------------------------------------------- C5: revenue at risk (real)
risk_buckets = {}
for r in rows:
    b = bucket_rec(r["recency"])
    d = risk_buckets.setdefault(b, {"rev": 0.0, "n": 0})
    d["rev"] += r["monetary"]
    d["n"] += 1
fig, ax = plt.subplots(figsize=(9, 4.4))
labs = r_labs
vals = [risk_buckets.get(b, {"rev": 0})["rev"] / 1e6 for b in labs]
cols = [TEAL, TEAL, AMBER, AMBER, RED, RED]
bars = ax.bar(labs, vals, color=cols, width=0.6)
for b, v, lab in zip(bars, vals, labs):
    ax.text(b.get_x() + b.get_width() / 2, v + max(vals) * 0.02,
            f"${v:.2f}M", ha="center", fontsize=10, fontweight="bold", color=INK)
ax.set_ylabel("Revenue held by cohort (USD, millions)")
ax.set_xlabel("Recency bucket (time since last purchase)")
ax.set_title("Revenue at risk  |  dormant customers hold the largest revenue pool")
ax.set_ylim(0, max(vals) * 1.26)
style_axes(ax, "y")
save(fig, "c5_revenue_at_risk.png")

# ------------------------------- C6: account value x engagement health matrix
HEALTHS = ["Active", "Dormant", "Churned"]
VT = ["A", "B", "C"]
cnt_m = np.zeros((3, 3))
rev_m = np.zeros((3, 3))
for i, v in enumerate(VT):
    for j, h in enumerate(HEALTHS):
        d = cell.get((v, h), {"n": 0, "rev": 0.0})
        cnt_m[i, j] = d["n"]
        rev_m[i, j] = d["rev"]

fig, ax = plt.subplots(figsize=(9.4, 4.9))
share = rev_m / TOTAL_REV * 100
im = ax.imshow(share, cmap="YlGnBu", aspect="auto", vmin=0, vmax=share.max() * 1.15)
for i in range(3):
    for j in range(3):
        dark = share[i, j] > share.max() * 0.55
        ax.text(j, i, f"${rev_m[i,j]/1e3:,.0f}k\n{share[i,j]:.1f}% of revenue\n{int(cnt_m[i,j]):,} accounts",
                ha="center", va="center", fontsize=10.2, linespacing=1.5,
                color="white" if dark else INK,
                fontweight="bold" if (i == 0 and j > 0) else "normal")
ax.set_xticks(range(3))
ax.set_xticklabels(["Active\n(last 180d)", "Dormant\n(181-365d)", "Churned\n(365d+)"])
ax.set_yticks(range(3))
ax.set_yticklabels(["A - Top 10%\nby value", "B - Next 20%", "C - Bottom 70%"])
ax.set_title("Where the revenue is stranded  |  value tier x engagement health", pad=14)
ax.text(1.0, -0.30, "darkest cells = revenue that has gone quiet but is still recoverable",
        transform=ax.transAxes, ha="right", fontsize=8.5, color=MUTED, style="italic")
for s in ax.spines.values():
    s.set_visible(False)
ax.tick_params(length=0)
cb = fig.colorbar(im, ax=ax, pad=0.02)
cb.set_label("Share of total revenue (%)", color=MUTED, fontsize=9)
cb.ax.tick_params(labelsize=8, colors=MUTED)
cb.outline.set_visible(False)
save(fig, "c6_account_matrix.png")

# ------------------------------------------------ C6b: revenue by health band
fig, ax = plt.subplots(figsize=(9, 4.3))
bands = []
for h in HEALTHS:
    n = sum(cell.get((v, h), {"n": 0})["n"] for v in VT)
    rv = sum(cell.get((v, h), {"n": 0, "rev": 0})["rev"] for v in VT)
    bands.append((h, n, rv))
bars = ax.bar([b[0] for b in bands], [b[2] / 1e6 for b in bands],
              color=[TEAL, AMBER, RED], width=0.5)
for b, (h, n, rv) in zip(bars, bands):
    ax.text(b.get_x() + b.get_width() / 2, rv / 1e6 + 0.03,
            f"${rv/1e6:.2f}M\n{rv/TOTAL_REV*100:.1f}%  ·  {n:,} accounts",
            ha="center", fontsize=10, fontweight="bold", color=INK, linespacing=1.4)
ax.set_ylabel("Historical revenue (USD, millions)")
ax.set_title("Only 27% of the account base is still active")
ax.set_ylim(0, max(b[2] for b in bands) / 1e6 * 1.25)
style_axes(ax, "y")
save(fig, "c6b_health_split.png")


# ------------------------------------------------------------ C7: market
# Numbers exported directly from the BigQuery query the analyst ran on
# bigquery-public-data.thelook_ecommerce (country x category aggregation).
market = [("Brazil\nOuterwear & Coats", 153.22, 147547.25, 940),
          ("China\n(reference market)", 102.8, 339238.68, 3300)]
fig, ax = plt.subplots(figsize=(8.6, 4.2))
labels = [m[0] for m in market]
aovs = [m[1] for m in market]
b = ax.bar(labels, aovs, color=[TEAL, "#9AA4B0"], width=0.42)
ax.axhline(102.8, color=MUTED, linestyle="--", linewidth=1)
ax.text(1.42, 104.5, "global benchmark", fontsize=8.5, color=MUTED, ha="right", style="italic")
for bar, m in zip(b, market):
    ax.text(bar.get_x() + bar.get_width() / 2, m[1] + 5,
            f"AOV ${m[1]:.2f}\nrev ${m[2]/1000:,.0f}k", ha="center",
            fontsize=10, fontweight="bold", color=INK)
ax.set_ylabel("Average order value (USD)")
ax.set_ylim(0, 200)
ax.set_title("Market efficiency  |  Brazil converts 49% more revenue per order")
style_axes(ax, "y")
save(fig, "c7_market_aov.png")

# ------------------------------------------------------- narrative printout
print("\nKEY NUMBERS")
print("champions", segs["Champions"]["n"], "at_risk", segs["At Risk (High Value)"]["n"])
print("at_risk_rev", round(segs["At Risk (High Value)"]["rev"], 2))
print("pareto_20", round(cum[int(N * 0.2) - 1], 1))
print("pareto_10", round(cum[int(N * 0.1) - 1], 1))
print("at_risk_tier1_pool", round(sum(r["monetary"] for r in rows
      if r["recency"] > 365 and r["monetary"] > 200), 2))
print("dormant_rev", round(risk_buckets["2 yr+"]["rev"], 2))
print("tierA", tier_stats["A"]["n"], round(tier_stats["A"]["rev"], 2),
      "share", round(tier_stats["A"]["rev"] / TOTAL_REV * 100, 1))
a_risk = cell.get(("A", "Dormant"), {"rev": 0})["rev"] + cell.get(("A", "Churned"), {"rev": 0})["rev"]
print("a_risk_pool", round(a_risk, 2), round(a_risk / TOTAL_REV * 100, 1), "%")
print("active_pct", round(sum(1 for r in rows if r["health"] == "Active") / N * 100, 1))
print("onetimer_pct", round(sum(1 for r in rows if r["frequency"] == 1) / N * 100, 1))
print("median_recency", sorted(r["recency"] for r in rows)[N // 2])
