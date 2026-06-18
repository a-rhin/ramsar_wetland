import os
os.environ['PYTHONUTF8'] = '1'
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from matplotlib.patches import Rectangle, FancyBboxPatch
from pathlib import Path

BASE = Path('D:/ramsar_wetlands')
OUT  = BASE / 'outputs'

# ── Colours ───────────────────────────────────────────────────────────────────
CH  = '#1B5E20'   # dark forest green  – header
CS  = '#2E7D32'   # medium green       – section bars
CD  = '#37474F'   # dark slate blue    – alt section bars
CB  = '#FFFFFF'   # white background
CP  = '#F1F8E9'   # pale green panel   – text panels
CX  = '#E8F5E9'   # stat box bg
CK  = '#F8F8F8'   # near-white card
CT  = '#1A1A1A'   # near-black text
CC  = '#546E7A'   # caption grey
CW  = '#FFFFFF'   # white text
CV  = '#2d7d2d'   # vegetation green
CWB = '#4a90d9'   # water blue
CBL = '#e07b39'   # built orange

# ── Load images ───────────────────────────────────────────────────────────────
def load(fn):
    fp = OUT / fn
    return plt.imread(str(fp)) if fp.exists() else None

img_study  = load('study_area_map.png')
img_keta_m = load('keta_multipanel.png')
img_muni_m = load('muni_multipanel.png')
img_keta_p = load('keta_2035_vs_2025.png')
img_muni_p = load('muni_2035_vs_2025.png')
img_keta_l = load('keta_landscape_metrics.png')
img_muni_l = load('muni_landscape_metrics.png')

# ── Figure (A0 portrait = 33.1 × 46.8 in) ────────────────────────────────────
fig = plt.figure(figsize=(33.1, 46.8), facecolor=CB)

gs = gridspec.GridSpec(6, 3, figure=fig,
    height_ratios=[4, 9, 14, 12, 10, 3.5],
    width_ratios=[1, 1, 1],
    hspace=0.03, wspace=0.025,
    left=0.012, right=0.988, top=0.993, bottom=0.007)

# ── Panel helpers ─────────────────────────────────────────────────────────────
def panel_base(ax, bg=CK, border='#CCCCCC'):
    ax.set_xticks([]); ax.set_yticks([])
    ax.set_facecolor(bg)
    for sp in ax.spines.values():
        sp.set_linewidth(0.8); sp.set_color(border)

def section_bar(ax, title, hfrac=0.07, bg=CS, fsz=21):
    ax.add_patch(Rectangle((0, 1 - hfrac), 1, hfrac,
        transform=ax.transAxes, facecolor=bg, zorder=5, clip_on=False))
    ax.text(0.5, 1 - hfrac / 2, title,
        ha='center', va='center', transform=ax.transAxes,
        fontsize=fsz, fontweight='bold', color=CW, zorder=6)

def embed_img(ax, img, x0=0, y0=0, x1=1, y1=1):
    if img is None:
        return
    sub = ax.inset_axes([x0, y0, x1 - x0, y1 - y0])
    sub.imshow(img, aspect='auto', interpolation='antialiased')
    sub.axis('off')

# ── ROW 0 – HEADER ────────────────────────────────────────────────────────────
ax_h = fig.add_subplot(gs[0, :])
ax_h.set_facecolor(CH)
for sp in ax_h.spines.values(): sp.set_visible(False)
ax_h.set_xticks([]); ax_h.set_yticks([])

ax_h.text(0.5, 0.72,
    'Wetland Degradation at Keta Lagoon Complex and Muni-Pomadze Ramsar Sites, Ghana:\n'
    'Multi-temporal LULC Change Detection and CA-Markov Future Projection (1991–2035)',
    ha='center', va='center', transform=ax_h.transAxes,
    fontsize=41, fontweight='bold', color=CW, linespacing=1.30)

ax_h.text(0.5, 0.21,
    'Remote Sensing & GIS Laboratory  ·  June 2026  ·  github.com/a-rhin/ramsar_wetland',
    ha='center', va='center', transform=ax_h.transAxes,
    fontsize=22, color='#C8E6C9', fontstyle='italic')

# ── ROW 1 COL 0 – Introduction ────────────────────────────────────────────────
ax_in = fig.add_subplot(gs[1, 0])
panel_base(ax_in, bg=CP)
section_bar(ax_in, 'INTRODUCTION & OBJECTIVES', hfrac=0.10)

INTRO = (
    'Ramsar wetlands deliver essential ecosystem services including\n'
    'flood regulation, carbon storage, fisheries habitat, and coastal\n'
    'protection. Yet they face mounting pressure from urban expansion,\n'
    'agriculture, and climate variability across sub-Saharan Africa.\n\n'
    'This study delivers the first comprehensive multi-epoch LULC\n'
    'analysis for two Ghanaian Ramsar Sites using 34 years of Landsat\n'
    'imagery, with CA-Markov projections to 2035.\n\n'
    'OBJECTIVES\n'
    '▸  Classify LULC for 1991, 2001, 2015 and 2025 at both sites\n'
    '▸  Quantify net area change: Vegetation, Water body,\n'
    '    Built up/Bareland\n'
    '▸  Project 2035 LULC via CA-Markov transition probability\n'
    '    modelling\n'
    '▸  Evaluate landscape structural change: Shannon Diversity\n'
    '    Index (SDI), patch count, Largest Patch Index (LPI)'
)
ax_in.text(0.05, 0.88, INTRO,
    ha='left', va='top', transform=ax_in.transAxes,
    fontsize=17, color=CT, linespacing=1.50)

# ── ROW 1 COL 1 – Study area ──────────────────────────────────────────────────
ax_sa = fig.add_subplot(gs[1, 1])
panel_base(ax_sa, bg=CK)
section_bar(ax_sa, 'STUDY AREA — GHANA', hfrac=0.08)
embed_img(ax_sa, img_study, x0=0.04, y0=0.01, x1=0.96, y1=0.91)

ax_sa.text(0.5, 0.935,
    'Keta Lagoon Complex (Volta Region, SE Ghana) ·  Muni-Pomadze (Central Region, S Ghana)',
    ha='center', va='center', transform=ax_sa.transAxes,
    fontsize=13.5, color=CC, fontstyle='italic')

# ── ROW 1 COL 2 – Methods ─────────────────────────────────────────────────────
ax_me = fig.add_subplot(gs[1, 2])
panel_base(ax_me, bg=CP)
section_bar(ax_me, 'METHODOLOGY', hfrac=0.10)

METH = (
    'SATELLITE DATA\n'
    '  1991 — Landsat 4/5 TM    (30 m, Bands 1–5, 7)\n'
    '  2001 — Landsat 7 ETM+   (30 m, Bands 1–5, 7)\n'
    '  2015, 2025 — Landsat 8/9 OLI (30 m, Bands 2–7)\n\n'
    'CLASSIFICATION\n'
    '  Classifier: Support Vector Machine (SVM), RBF kernel\n'
    '  Classes: Vegetation | Water body | Built up/Bareland\n'
    '  Validation: Stratified random sampling\n'
    '  Accuracy — Keta: OA 81–87%, κ 0.70–0.79\n'
    '             Muni: OA 83–90%, κ 0.75–0.85\n\n'
    'CHANGE DETECTION\n'
    '  Class-area statistics (km²) across 4 epochs\n'
    '  Net change and rate per inter-epoch interval\n\n'
    'CA-MARKOV PROJECTION\n'
    '  Transition Probability Matrix (TPM): 2015 → 2025\n'
    '  Projection horizon: 2035 (10 years forward)\n\n'
    'LANDSCAPE METRICS\n'
    '  SDI = −Σ(pᵢ × ln pᵢ)   Patch count per class\n'
    '  Largest Patch Index (LPI, %)  across 5 epochs'
)
ax_me.text(0.05, 0.88, METH,
    ha='left', va='top', transform=ax_me.transAxes,
    fontsize=16.5, color=CT, linespacing=1.48)

# ── ROW 2 COLS 0–1 – Keta multipanel maps ────────────────────────────────────
ax_km = fig.add_subplot(gs[2, :2])
panel_base(ax_km, bg=CK)
section_bar(ax_km, 'KETA LAGOON COMPLEX — LULC CLASSIFICATION MAPS (1991–2025)', hfrac=0.055)
embed_img(ax_km, img_keta_m, x0=0.005, y0=0.01, x1=0.995, y1=0.940)

# ── ROW 2 COL 2 – Keta stats ─────────────────────────────────────────────────
ax_ks = fig.add_subplot(gs[2, 2])
panel_base(ax_ks, bg=CX)
section_bar(ax_ks, 'KETA — KEY STATISTICS', hfrac=0.06, bg='#1B5E20')

ax_ks.text(0.5, 0.895, 'LULC Area by Class and Year (km²)',
    ha='center', va='top', transform=ax_ks.transAxes,
    fontsize=17, fontweight='bold', color=CT)

# Table
hdrs = ['Class', '1991', '2001', '2015', '2025']
cxs  = [0.12, 0.38, 0.54, 0.68, 0.85]
ax_ks.add_patch(Rectangle((0.03, 0.775), 0.94, 0.075,
    transform=ax_ks.transAxes, facecolor=CS, zorder=2))
for h, x in zip(hdrs, cxs):
    ax_ks.text(x, 0.812, h, ha='center', va='center', transform=ax_ks.transAxes,
        fontsize=14.5, fontweight='bold', color=CW, zorder=3)

rows_k = [
    ('Vegetation', CV,   [845.4, 719.9, 787.4, 499.0]),
    ('Water body', CWB,  [340.6, 268.8, 297.6, 312.9]),
    ('Built up',   CBL,  [206.0, 408.5, 295.3, 568.0]),
]
for i, (cls, clr, vals) in enumerate(rows_k):
    y = 0.70 - i * 0.08
    bg_c = '#F9FBE7' if i % 2 == 0 else '#FFFFFF'
    ax_ks.add_patch(Rectangle((0.03, y - 0.025), 0.94, 0.055,
        transform=ax_ks.transAxes, facecolor=bg_c, zorder=1))
    ax_ks.text(cxs[0], y, cls, ha='center', va='center',
        transform=ax_ks.transAxes, fontsize=14, color=clr, fontweight='bold')
    for j, v in enumerate(vals):
        ax_ks.text(cxs[j+1], y, f'{v:.1f}', ha='center', va='center',
            transform=ax_ks.transAxes, fontsize=14.5, color=CT)

# Big callout boxes
ax_ks.add_patch(FancyBboxPatch((0.04, 0.365), 0.92, 0.155,
    boxstyle='round,pad=0.01', facecolor='#FFEBEE', edgecolor='#E53935', lw=2,
    transform=ax_ks.transAxes, zorder=2))
ax_ks.text(0.5, 0.495, 'Vegetation loss  1991 → 2025',
    ha='center', va='center', transform=ax_ks.transAxes,
    fontsize=16, fontweight='bold', color='#B71C1C', zorder=3)
ax_ks.text(0.5, 0.405, '−346.4 km²  (−41.0 %)',
    ha='center', va='center', transform=ax_ks.transAxes,
    fontsize=22, fontweight='bold', color='#C62828', zorder=3)

ax_ks.add_patch(FancyBboxPatch((0.04, 0.190), 0.92, 0.155,
    boxstyle='round,pad=0.01', facecolor='#FFF3E0', edgecolor='#FB8C00', lw=2,
    transform=ax_ks.transAxes, zorder=2))
ax_ks.text(0.5, 0.320, 'Built up/Bareland gain  1991 → 2025',
    ha='center', va='center', transform=ax_ks.transAxes,
    fontsize=16, fontweight='bold', color='#E65100', zorder=3)
ax_ks.text(0.5, 0.230, '+361.9 km²  (+175.7 %)',
    ha='center', va='center', transform=ax_ks.transAxes,
    fontsize=22, fontweight='bold', color='#E65100', zorder=3)

ax_ks.text(0.5, 0.135,
    'OA 81.4–86.9%    κ 0.697–0.792',
    ha='center', va='center', transform=ax_ks.transAxes,
    fontsize=15, color=CC, fontweight='bold')
ax_ks.text(0.5, 0.060,
    'Note: 2001–2015 partial vegetation recovery reversed by 2025',
    ha='center', va='center', transform=ax_ks.transAxes,
    fontsize=13.5, color=CC, fontstyle='italic')

# ── ROW 3 COLS 0–1 – Muni multipanel maps ────────────────────────────────────
ax_mm = fig.add_subplot(gs[3, :2])
panel_base(ax_mm, bg=CK)
section_bar(ax_mm, 'MUNI-POMADZE — LULC CLASSIFICATION MAPS (1991–2025)', hfrac=0.065)
embed_img(ax_mm, img_muni_m, x0=0.005, y0=0.01, x1=0.995, y1=0.930)

# ── ROW 3 COL 2 – Muni stats ─────────────────────────────────────────────────
ax_ms = fig.add_subplot(gs[3, 2])
panel_base(ax_ms, bg=CX)
section_bar(ax_ms, 'MUNI — KEY STATISTICS', hfrac=0.07, bg='#1B5E20')

ax_ms.text(0.5, 0.89, 'LULC Area by Class and Year (km²)',
    ha='center', va='top', transform=ax_ms.transAxes,
    fontsize=17, fontweight='bold', color=CT)

ax_ms.add_patch(Rectangle((0.03, 0.770), 0.94, 0.075,
    transform=ax_ms.transAxes, facecolor=CS, zorder=2))
for h, x in zip(hdrs, cxs):
    ax_ms.text(x, 0.807, h, ha='center', va='center', transform=ax_ms.transAxes,
        fontsize=14.5, fontweight='bold', color=CW, zorder=3)

rows_m = [
    ('Vegetation', CV,   [60.77, 54.41, 39.74, 53.34]),
    ('Water body', CWB,  [ 1.56,  0.81,  1.01,  0.99]),
    ('Built up',   CBL,  [55.59, 61.92, 76.38, 62.80]),
]
for i, (cls, clr, vals) in enumerate(rows_m):
    y = 0.695 - i * 0.08
    bg_c = '#F9FBE7' if i % 2 == 0 else '#FFFFFF'
    ax_ms.add_patch(Rectangle((0.03, y - 0.025), 0.94, 0.055,
        transform=ax_ms.transAxes, facecolor=bg_c, zorder=1))
    ax_ms.text(cxs[0], y, cls, ha='center', va='center',
        transform=ax_ms.transAxes, fontsize=14, color=clr, fontweight='bold')
    for j, v in enumerate(vals):
        ax_ms.text(cxs[j+1], y, f'{v:.2f}', ha='center', va='center',
            transform=ax_ms.transAxes, fontsize=14.5, color=CT)

ax_ms.add_patch(FancyBboxPatch((0.04, 0.355), 0.92, 0.155,
    boxstyle='round,pad=0.01', facecolor='#FFEBEE', edgecolor='#E53935', lw=2,
    transform=ax_ms.transAxes, zorder=2))
ax_ms.text(0.5, 0.483, 'Vegetation loss  1991 → 2025',
    ha='center', va='center', transform=ax_ms.transAxes,
    fontsize=16, fontweight='bold', color='#B71C1C', zorder=3)
ax_ms.text(0.5, 0.395, '−7.43 km²  (−12.2 %)',
    ha='center', va='center', transform=ax_ms.transAxes,
    fontsize=22, fontweight='bold', color='#C62828', zorder=3)

ax_ms.add_patch(FancyBboxPatch((0.04, 0.190), 0.92, 0.145,
    boxstyle='round,pad=0.01', facecolor='#FFF3E0', edgecolor='#FB8C00', lw=2,
    transform=ax_ms.transAxes, zorder=2))
ax_ms.text(0.5, 0.310, 'Built up/Bareland gain  1991 → 2025',
    ha='center', va='center', transform=ax_ms.transAxes,
    fontsize=16, fontweight='bold', color='#E65100', zorder=3)
ax_ms.text(0.5, 0.225, '+7.21 km²  (+13.0 %)',
    ha='center', va='center', transform=ax_ms.transAxes,
    fontsize=22, fontweight='bold', color='#E65100', zorder=3)

ax_ms.text(0.5, 0.130,
    'OA 83.3–90.0%    κ 0.750–0.850',
    ha='center', va='center', transform=ax_ms.transAxes,
    fontsize=15, color=CC, fontweight='bold')
ax_ms.text(0.5, 0.055,
    'Vegetation minimum in 2015; partial recovery by 2025',
    ha='center', va='center', transform=ax_ms.transAxes,
    fontsize=13.5, color=CC, fontstyle='italic')

# ── ROW 4 COL 0 – Keta landscape metrics ─────────────────────────────────────
ax_kl = fig.add_subplot(gs[4, 0])
panel_base(ax_kl, bg=CK)
section_bar(ax_kl, 'LANDSCAPE METRICS — KETA (1991–2035)', hfrac=0.085, bg=CD)
embed_img(ax_kl, img_keta_l, x0=0.01, y0=0.02, x1=0.99, y1=0.910)

# ── ROW 4 COL 1 – Muni landscape metrics ─────────────────────────────────────
ax_ml = fig.add_subplot(gs[4, 1])
panel_base(ax_ml, bg=CK)
section_bar(ax_ml, 'LANDSCAPE METRICS — MUNI (1991–2035)', hfrac=0.085, bg=CD)
embed_img(ax_ml, img_muni_l, x0=0.01, y0=0.02, x1=0.99, y1=0.910)

# ── ROW 4 COL 2 – 2035 Predictions ───────────────────────────────────────────
ax_pr = fig.add_subplot(gs[4, 2])
panel_base(ax_pr, bg=CK)
section_bar(ax_pr, 'CA-MARKOV 2035 PROJECTIONS', hfrac=0.085, bg=CD)
embed_img(ax_pr, img_keta_p, x0=0.01, y0=0.50, x1=0.99, y1=0.910)
ax_pr.text(0.5, 0.485, 'Keta Lagoon Complex — 2025 vs 2035',
    ha='center', va='center', transform=ax_pr.transAxes,
    fontsize=13.5, color=CC, fontstyle='italic')
embed_img(ax_pr, img_muni_p, x0=0.01, y0=0.06, x1=0.99, y1=0.465)
ax_pr.text(0.5, 0.045, 'Muni-Pomadze — 2025 vs 2035',
    ha='center', va='center', transform=ax_pr.transAxes,
    fontsize=13.5, color=CC, fontstyle='italic')

# ── ROW 5 – FOOTER ────────────────────────────────────────────────────────────
ax_ft = fig.add_subplot(gs[5, :])
ax_ft.set_facecolor(CH)
for sp in ax_ft.spines.values(): sp.set_visible(False)
ax_ft.set_xticks([]); ax_ft.set_yticks([])

ax_ft.text(0.012, 0.78, 'KEY CONCLUSIONS',
    ha='left', va='center', transform=ax_ft.transAxes,
    fontsize=18, fontweight='bold', color='#A5D6A7')

CONCL = (
    '▸  Keta: vegetation loss −41 % and built-up expansion +176 % over 34 years (1991–2025)\n'
    '▸  A 2001–2015 partial vegetation recovery at Keta was reversed sharply by 2025\n'
    '▸  CA-Markov 2035 projections indicate continued degradation at Keta; Muni remains relatively stable\n'
    '▸  Declining SDI at Keta (1.018 → 0.675) signals progressive landscape homogenisation\n'
    '▸  Muni shows a −12 % vegetation loss with partial 2025 recovery, consistent with conservation efforts'
)
ax_ft.text(0.012, 0.48, CONCL,
    ha='left', va='center', transform=ax_ft.transAxes,
    fontsize=15, color=CW, linespacing=1.60)

REFS = (
    'Mensah et al. (2020) Ocean & Coast. Mgmt.  ·  '
    'Vapnik (1995) Statistical Learning Theory. Springer.  ·  '
    'Pontius & Millones (2011) Int. J. Remote Sens.  ·  '
    'Ramsar Convention Secretariat (2016)  ·  '
    'Schiavina et al. (2022) GHS-POP R2022A'
)
ax_ft.text(0.012, 0.12, REFS,
    ha='left', va='center', transform=ax_ft.transAxes,
    fontsize=12, color='#81C784')

ax_ft.text(0.870, 0.58, 'Code & data:',
    ha='center', va='center', transform=ax_ft.transAxes,
    fontsize=15, color='#C8E6C9', fontweight='bold')
ax_ft.text(0.870, 0.32, 'github.com/a-rhin/\nramsar_wetland',
    ha='center', va='center', transform=ax_ft.transAxes,
    fontsize=14, color='#A5D6A7', fontstyle='italic', linespacing=1.4)

# ── Save ──────────────────────────────────────────────────────────────────────
out_png = OUT / 'wetland_poster_A0.png'
fig.savefig(str(out_png), dpi=150, bbox_inches='tight', facecolor=CB)
plt.close(fig)
sz = out_png.stat().st_size
print(f'Saved: {out_png.name}  ({sz/1024/1024:.1f} MB)')
print('Done.')