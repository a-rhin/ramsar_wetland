#!/usr/bin/env python3
"""
Targeted fix for three figures:
  1. keta_multipanel.png      — 1991 blank panel; background bleed
  2. keta_2035_vs_2025.png    — background bleed outside boundary
  3. muni_2035_vs_2025.png    — background bleed outside boundary
  4. keta_landscape_metrics.png — 2035 bar wrong colour
"""
import warnings
warnings.filterwarnings('ignore')

import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.colors import ListedColormap, BoundaryNorm
import rasterio
from rasterio.features import geometry_mask
import geopandas as gpd
from scipy.ndimage import label as ndlabel
from pathlib import Path

plt.rcParams.update({
    'font.family': 'DejaVu Sans',
    'axes.titlesize': 11, 'axes.labelsize': 10,
    'xtick.labelsize': 9,  'ytick.labelsize': 9,
    'figure.dpi': 150,
})

BASE = Path("D:/ramsar_wetlands")
OUT  = BASE / "outputs"

CNAMES = {1: 'Vegetation', 2: 'Water body', 3: 'Built up/Bareland'}
COLORS = {'Vegetation': '#2e7d32', 'Water body': '#1565c0', 'Built up/Bareland': '#e65100'}
NORM   = BoundaryNorm([0.5, 1.5, 2.5, 3.5], 3)


def make_cmap():
    c = ListedColormap(['#2e7d32', '#1565c0', '#e65100'])
    c.set_bad(color='white', alpha=1)
    return c


def to_masked(arr):
    """Mask NaN and sub-class values so they render as white."""
    return np.ma.masked_where(np.isnan(arr) | (arr < 0.5), arr)


# ── raster loader (mirrors main.py logic) ──────────────────────────────────
def load_raster(site, year, bnd_gdf=None):
    p = BASE / f"data/{site}/rasters/{site}_classified_{year}.tif"
    with rasterio.open(p) as src:
        arr  = src.read(1).astype(np.float32)
        meta = src.meta.copy()
        tf   = src.transform
        crs  = src.crs
    meta['crs'] = crs
    if int(arr.max()) <= 2:          # 0-indexed muni-style
        if bnd_gdf is not None:
            bnd = bnd_gdf.to_crs(crs)
            geoms   = [g for g in bnd.geometry if g is not None]
            in_bnd  = geometry_mask(geoms, transform=tf,
                                    out_shape=arr.shape, invert=True)
            result  = np.full(arr.shape, np.nan, dtype=np.float32)
            result[in_bnd] = arr[in_bnd] + 1   # 0→1, 1→2, 2→3
            return result, meta, tf
        return np.where(arr > 0, arr + 1, np.nan).astype(np.float32), meta, tf
    arr[arr == 0] = np.nan           # 1-indexed keta-style
    return arr, meta, tf


def load_boundary(site):
    return gpd.read_file(BASE / f"data/{site}/area/{site}_boundary.shp")


def load_pred_2035(site):
    """Load the saved CA-Markov prediction GeoTIFF (values 1–3, 0=NoData)."""
    with rasterio.open(OUT / f"{site}_predicted_2035.tif") as src:
        arr = src.read(1).astype(np.float32)
    arr[arr == 0] = np.nan
    return arr


# ── decorators ─────────────────────────────────────────────────────────────
def add_north_arrow(ax):
    ax.annotate('N', xy=(0.07, 0.97), xytext=(0.07, 0.90),
                xycoords='axes fraction', textcoords='axes fraction',
                ha='center', va='center', fontsize=9, fontweight='bold',
                arrowprops=dict(arrowstyle='->', color='black', lw=1.5),
                annotation_clip=False)


def add_scale_bar(ax, pixel_m):
    xlim, ylim = ax.get_xlim(), ax.get_ylim()
    w_px = xlim[1] - xlim[0]
    h_px = abs(ylim[1] - ylim[0])
    extent_km = w_px * pixel_m / 1000
    bar_km = 1
    for v in [1, 2, 5, 10, 20, 50, 100]:
        if v < extent_km * 0.25:
            bar_km = v
    bar_px = bar_km * 1000 / pixel_m
    x0 = xlim[0] + w_px * 0.05
    x1 = x0 + bar_px
    y0 = min(ylim) + h_px * 0.05
    tick = h_px * 0.01
    ax.plot([x0, x1], [y0, y0], 'k-', lw=2, solid_capstyle='butt')
    for xv in (x0, x1):
        ax.plot([xv, xv], [y0 - tick, y0 + tick], 'k-', lw=1.5)
    ax.text((x0 + x1) / 2, y0 + h_px * 0.025,
            f'{bar_km} km', ha='center', va='bottom', fontsize=7)


def overlay_boundary(ax, bnd_gdf, crs, tf):
    try:
        bnd = bnd_gdf.to_crs(crs)
        for geom in bnd.geometry:
            if geom is None:
                continue
            parts = list(geom.geoms) if hasattr(geom, 'geoms') else [geom]
            for part in parts:
                xy   = np.array(part.exterior.coords)
                cols = (xy[:, 0] - tf.c) / tf.a
                rows = (xy[:, 1] - tf.f) / tf.e
                ax.plot(cols, rows, 'k-', lw=0.9, alpha=0.85)
    except Exception as e:
        print(f"    [WARN] boundary overlay: {e}")


def pin_limits(ax, arr):
    """Force axes to exactly the raster's pixel extent (y inverted for imshow)."""
    ax.set_xlim(0, arr.shape[1])
    ax.set_ylim(arr.shape[0], 0)


# ═══════════════════════════════════════════════════════════════════════════
# FIX 1 — keta_multipanel.png
# ═══════════════════════════════════════════════════════════════════════════
print("=" * 60)
print("FIX 1: keta_multipanel.png")
print("=" * 60)

site  = 'keta'
bnd   = load_boundary(site)
cmap  = make_cmap()
YEARS = [1991, 2001, 2015, 2025]

rasters, metas = {}, {}
for year in YEARS:
    arr, meta, tf = load_raster(site, year, bnd)
    rasters[year] = arr
    metas[year]   = (meta, tf)
    print(f"  {year}: shape={arr.shape}, valid={int(np.sum(~np.isnan(arr))):,}")

pred = load_pred_2035(site)
rasters[2035] = pred
metas[2035]   = metas[2025]   # same spatial reference as 2025

map_data = [(y, rasters[y]) for y in [1991, 2001, 2015, 2025, 2035]]

fig, axes = plt.subplots(1, 5, figsize=(25, 6))

for ax, (year, arr) in zip(axes, map_data):
    meta_yr, tf_yr = metas[year]
    crs_yr = meta_yr.get('crs') or meta_yr['crs']
    pixel_m = abs(tf_yr.a)

    ax.imshow(to_masked(arr), cmap=cmap, norm=NORM, interpolation='nearest')
    pin_limits(ax, arr)                            # ← pin BEFORE any overlay

    title = f"{year}" + ("\n(Predicted)" if year == 2035 else "")
    ax.set_title(title, fontweight='bold', fontsize=10, pad=3)
    ax.axis('off')

    overlay_boundary(ax, bnd, crs_yr, tf_yr)
    pin_limits(ax, arr)                            # ← re-pin AFTER boundary

    add_north_arrow(ax)
    add_scale_bar(ax, pixel_m)

patches = [mpatches.Patch(color=c, label=n) for n, c in COLORS.items()]
fig.legend(handles=patches, loc='lower center', ncol=3, fontsize=10,
           frameon=True, bbox_to_anchor=(0.5, -0.04), framealpha=0.9)
fig.suptitle('Keta Lagoon Complex — LULC Maps (1991–2035)',
             fontsize=13, fontweight='bold', y=1.01)
plt.tight_layout(rect=[0, 0.07, 1, 1])

fig.savefig(OUT / 'keta_multipanel.png', dpi=300, bbox_inches='tight')
fig.savefig(OUT / 'keta_multipanel.pdf', bbox_inches='tight')
plt.close(fig)
print("  Saved: keta_multipanel.png + .pdf\n")


# ═══════════════════════════════════════════════════════════════════════════
# FIX 2 — 2035 vs 2025 comparison (keta & muni)
# ═══════════════════════════════════════════════════════════════════════════
print("=" * 60)
print("FIX 2: 2035 vs 2025 comparison maps")
print("=" * 60)

for site in ['keta', 'muni']:
    site_label = 'Keta Lagoon Complex' if site == 'keta' else 'Muni-Pomadze'
    bnd = load_boundary(site)
    arr_2025, meta_2025, tf_2025 = load_raster(site, 2025, bnd)
    pred                          = load_pred_2035(site)
    crs_2025                      = meta_2025.get('crs') or meta_2025['crs']
    pixel_m_2025                  = abs(tf_2025.a)
    cmap = make_cmap()

    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    for ax, data, title in zip(
        axes,
        [arr_2025, pred],
        ['2025 (Classified)', '2035 (CA-Markov Predicted)']
    ):
        ax.imshow(to_masked(data), cmap=cmap, norm=NORM, interpolation='nearest')
        pin_limits(ax, data)

        ax.set_title(f'{site_label} — {title}', fontweight='bold', fontsize=12)
        ax.axis('off')

        overlay_boundary(ax, bnd, crs_2025, tf_2025)
        pin_limits(ax, data)                       # ← re-pin after boundary

        add_north_arrow(ax)
        add_scale_bar(ax, pixel_m_2025)

    patches = [mpatches.Patch(color=c, label=n) for n, c in COLORS.items()]
    fig.legend(handles=patches, loc='lower center', ncol=3,
               fontsize=10, frameon=True, bbox_to_anchor=(0.5, -0.01))
    plt.tight_layout(rect=[0, 0.06, 1, 1])
    fp = OUT / f"{site}_2035_vs_2025.png"
    fig.savefig(fp, dpi=200, bbox_inches='tight')
    plt.close(fig)
    print(f"  Saved: {fp.name}")

print()

# ═══════════════════════════════════════════════════════════════════════════
# FIX 3 — keta_landscape_metrics.png (uniform bar colour)
# ═══════════════════════════════════════════════════════════════════════════
print("=" * 60)
print("FIX 3: keta_landscape_metrics.png")
print("=" * 60)

site = 'keta'
bnd  = load_boundary(site)

def shannon_di(arr):
    v = arr[~np.isnan(arr)].astype(int)
    v = v[np.isin(v, [1, 2, 3])]
    n = len(v)
    if n == 0: return 0.0
    s = 0.0
    for c in [1, 2, 3]:
        p = np.sum(v == c) / n
        if p > 0: s -= p * np.log(p)
    return s

def patch_metrics(arr):
    tot = int(np.sum(~np.isnan(arr)))
    out = {}
    for cls in [1, 2, 3]:
        labeled, n_pat = ndlabel((arr == cls).astype(int))
        if n_pat > 0:
            sizes = [int(np.sum(labeled == p)) for p in range(1, n_pat + 1)]
            lpi   = max(sizes) / tot * 100 if tot else 0.0
        else:
            lpi = 0.0
        out[CNAMES[cls]] = {'n_patches': n_pat, 'LPI': round(lpi, 4)}
    return out

LM = {}
for year in [1991, 2001, 2015, 2025]:
    arr, _, _ = load_raster(site, year, bnd)
    LM[year]  = {'SDI': shannon_di(arr), 'frag': patch_metrics(arr)}
    print(f"  {year}: SDI={LM[year]['SDI']:.4f}")

pred_keta   = load_pred_2035(site)
LM[2035]    = {'SDI': shannon_di(pred_keta), 'frag': patch_metrics(pred_keta)}
print(f"  2035: SDI={LM[2035]['SDI']:.4f}")

yr_list  = sorted(LM.keys())
sdi_vals = [LM[y]['SDI'] for y in yr_list]

fig, axes = plt.subplots(1, 3, figsize=(16, 5))

# All bars the same green — no special 2035 colour
GREEN = '#2d6a2d'
axes[0].bar([str(y) for y in yr_list], sdi_vals, color=GREEN, edgecolor='white')
axes[0].set_title('Shannon Diversity Index', fontweight='bold')
axes[0].set_xlabel('Year'); axes[0].set_ylabel('SDI')
for i, v in enumerate(sdi_vals):
    axes[0].text(i, v + 0.003, f'{v:.3f}', ha='center', va='bottom', fontsize=8)
axes[0].grid(axis='y', alpha=0.3)

for ci, cls in enumerate(CNAMES.values()):
    yrs  = [y for y in yr_list if cls in LM[y]['frag']]
    vals = [LM[y]['frag'][cls]['n_patches'] for y in yrs]
    axes[1].plot([str(y) for y in yrs], vals, marker='o',
                 color=list(COLORS.values())[ci], label=cls, lw=2)
axes[1].set_title('Number of Patches per Class', fontweight='bold')
axes[1].set_xlabel('Year'); axes[1].set_ylabel('Patch Count')
axes[1].legend(fontsize=8); axes[1].grid(alpha=0.3)

for ci, cls in enumerate(CNAMES.values()):
    yrs  = [y for y in yr_list if cls in LM[y]['frag']]
    vals = [LM[y]['frag'][cls]['LPI'] for y in yrs]
    axes[2].plot([str(y) for y in yrs], vals, marker='s',
                 color=list(COLORS.values())[ci], label=cls, lw=2)
axes[2].set_title('Largest Patch Index (%)', fontweight='bold')
axes[2].set_xlabel('Year'); axes[2].set_ylabel('LPI (%)')
axes[2].legend(fontsize=8); axes[2].grid(alpha=0.3)

fig.suptitle('Keta Lagoon Complex — Landscape Metrics',
             fontweight='bold', fontsize=12)
plt.tight_layout()
fp = OUT / 'keta_landscape_metrics.png'
fig.savefig(fp, dpi=200, bbox_inches='tight')
plt.close(fig)
print(f"\n  Saved: {fp.name}")

print("\n" + "=" * 60)
print("All 3 figures fixed and saved.")
print("=" * 60)
