"""
=============================================================================
SCATTER PLOT — Days Since Expiry vs Package ID, Coloured by Branch
=============================================================================

Plots:
  X axis  →  days_expired      (how long ago subscription ended)
  Y axis  →  package_id        (which plan they were on)
  Colour  →  member_branch_id  (which branch they belonged to)

Features:
  - Interactive HTML plot (hover to see customer details)
  - Static PNG version for presentations
  - Configurable filters (segment, tier, branch)
  - Jitter on Y axis so overlapping points are visible
  - Branch 1008 always highlighted if present

HOW TO RUN:
  pip install pandas plotly openpyxl kaleido matplotlib --break-system-packages
  python3 packages_scatter.py

OUTPUT:
  scatter_days_vs_package.html  ← interactive (open in browser)
  scatter_days_vs_package.png   ← static image for presentations
=============================================================================
"""

import pandas as pd
import numpy as np
import warnings
import os

warnings.filterwarnings('ignore')

# ── CONFIG ────────────────────────────────────────────────────────────────────
INPUT_FILE = "sample data expiry members.xlsx"
SHEET_NAME = 0          # ← 0 = first sheet (works regardless of sheet name)

# Filter options — set to None to include all
FILTER_SEGMENT  = None  # e.g. "A - Friction Churner"
FILTER_TIER     = None  # e.g. "Tier 1 - Personalised Outreach"
FILTER_BRANCHES = None  # e.g. [1008, 126, 45]

# Only show members expired 180+ days
MIN_DAYS_EXPIRED = 180

# Jitter — adds tiny random spread on Y so stacked points are visible
JITTER_AMOUNT = 0.3

# Max points to plot (None = all)
MAX_POINTS = None

# Output files
OUT_HTML = "scatter_days_vs_package.html"
OUT_PNG  = "scatter_days_vs_package.png"

# Highlight a specific branch with star markers (None to disable)
HIGHLIGHT_BRANCH = 1008
# ─────────────────────────────────────────────────────────────────────────────


def load_and_prepare():
    print("\n" + "=" * 60)
    print("LOADING AND PREPARING DATA")
    print("=" * 60)

    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(
            f"\n❌ '{INPUT_FILE}' not found.\n"
            f"   Make sure the Excel file is in the same folder as this script.\n"
        )

    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    print(f"✅ Loaded {len(df):,} rows from sheet index {SHEET_NAME}")
    print(f"   Columns: {list(df.columns)}")

    # Ensure required columns exist
    required = ["days_expired", "package_id", "member_branch_id"]
    missing  = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"❌ Missing required columns: {missing}\n"
                         f"   Available: {list(df.columns)}")

    # Clean types
    df["days_expired"]     = pd.to_numeric(df["days_expired"],     errors="coerce")
    df["member_branch_id"] = df["member_branch_id"].astype(str).str.strip()
    df["package_id"]       = df["package_id"].astype(str).str.strip()

    # Drop rows missing core values
    df = df.dropna(subset=["days_expired", "package_id", "member_branch_id"])

    # Apply 180+ days filter
    df = df[df["days_expired"] >= MIN_DAYS_EXPIRED]
    print(f"   After {MIN_DAYS_EXPIRED}+ days filter: {len(df):,} rows")

    # Apply optional filters
    if FILTER_SEGMENT and "churn_segment" in df.columns:
        df = df[df["churn_segment"] == FILTER_SEGMENT]
        print(f"   Segment filter '{FILTER_SEGMENT}': {len(df):,} rows")

    if FILTER_TIER and "outreach_tier" in df.columns:
        df = df[df["outreach_tier"].astype(str).str.contains(FILTER_TIER, na=False)]
        print(f"   Tier filter '{FILTER_TIER}': {len(df):,} rows")

    if FILTER_BRANCHES:
        df = df[df["member_branch_id"].isin([str(b) for b in FILTER_BRANCHES])]
        print(f"   Branch filter {FILTER_BRANCHES}: {len(df):,} rows")

    # Sample if too large
    if MAX_POINTS and len(df) > MAX_POINTS:
        df = df.sample(n=MAX_POINTS, random_state=42)
        print(f"   Sampled to {MAX_POINTS:,} points")

    # Y axis: numeric encoding of package_id (it's categorical)
    packages_ordered = sorted(df["package_id"].unique())
    pkg_to_num       = {pkg: i for i, pkg in enumerate(packages_ordered)}
    df["package_num"] = df["package_id"].map(pkg_to_num)

    # Add jitter to Y so overlapping points spread out
    if JITTER_AMOUNT > 0:
        np.random.seed(42)
        df["package_num_jittered"] = (
            df["package_num"]
            + np.random.uniform(-JITTER_AMOUNT, JITTER_AMOUNT, size=len(df))
        )
    else:
        df["package_num_jittered"] = df["package_num"]

    # Build hover tooltip text
    hover_parts = [
        "<b>Branch:</b> "       + df["member_branch_id"].astype(str),
        "<b>Package:</b> "      + df["package_id"].astype(str),
        "<b>Days Expired:</b> " + df["days_expired"].astype(int).astype(str),
    ]
    for col_name, label in [
        ("name",               "<b>Name:</b> "),
        ("membership_no",      "<b>Membership No:</b> "),
        ("membership_months",  "<b>Tenure (months):</b> "),
        ("renewal_count",      "<b>Renewals:</b> "),
        ("total_borrow_count", "<b>Total Borrows:</b> "),
        ("latest_amount_paid", "<b>Amount Paid (₹):</b> "),
        ("num_books",          "<b>Books at a Time:</b> "),
        ("expiry_date",        "<b>Expiry Date:</b> "),
    ]:
        if col_name in df.columns:
            hover_parts.append(label + df[col_name].fillna("—").astype(str))

    df["hover_text"] = ["<br>".join(parts) for parts in zip(*hover_parts)]

    print(f"\n✅ Ready: {len(df):,} members  ·  "
          f"{df['member_branch_id'].nunique()} branches  ·  "
          f"{len(packages_ordered)} packages")

    return df, packages_ordered, pkg_to_num


def build_interactive_plot(df, packages_ordered):
    """Interactive Plotly HTML — hover over any dot to see full member details."""
    try:
        import plotly.graph_objects as go
        import plotly.express as px
        import plotly.colors as pc
    except ImportError:
        print("❌ plotly not installed — run: pip install plotly --break-system-packages")
        return

    print("\n" + "=" * 60)
    print("BUILDING INTERACTIVE HTML PLOT")
    print("=" * 60)

    branches      = sorted(df["member_branch_id"].unique())
    n_branches    = len(branches)
    highlight_str = str(HIGHLIGHT_BRANCH) if HIGHLIGHT_BRANCH else None

    # Generate colour palette (handles 49 branches)
    if n_branches <= 10:
        palette = px.colors.qualitative.Set1[:n_branches]
    elif n_branches <= 24:
        palette = px.colors.qualitative.Dark24[:n_branches]
    else:
        palette = pc.sample_colorscale(
            "Turbo", [i / max(n_branches - 1, 1) for i in range(n_branches)]
        )

    branch_colour = {b: palette[i % len(palette)] for i, b in enumerate(branches)}

    fig = go.Figure()

    # Plot each branch (except highlighted — drawn on top)
    for branch in branches:
        if branch == highlight_str:
            continue
        sub = df[df["member_branch_id"] == branch]
        fig.add_trace(go.Scatter(
            x    = sub["days_expired"],
            y    = sub["package_num_jittered"],
            mode = "markers",
            name = f"Branch {branch}",
            marker = dict(
                color   = branch_colour[branch],
                size    = 7,
                opacity = 0.65,
                line    = dict(width=0.3, color="white"),
            ),
            text          = sub["hover_text"],
            hovertemplate = "%{text}<extra></extra>",
        ))

    # Highlighted branch on top with star markers
    if highlight_str and highlight_str in df["member_branch_id"].values:
        sub_hl = df[df["member_branch_id"] == highlight_str]
        fig.add_trace(go.Scatter(
            x    = sub_hl["days_expired"],
            y    = sub_hl["package_num_jittered"],
            mode = "markers",
            name = f"⭐ Branch {HIGHLIGHT_BRANCH} (highlighted)",
            marker = dict(
                symbol  = "star",
                color   = "#FF4444",
                size    = 11,
                opacity = 0.9,
                line    = dict(width=0.5, color="white"),
            ),
            text          = sub_hl["hover_text"],
            hovertemplate = "%{text}<extra></extra>",
        ))
        print(f"   ⭐ Branch {HIGHLIGHT_BRANCH} highlighted with stars")

    # Reference lines at 180d / 1yr / 2yr
    max_days = df["days_expired"].max()
    for xval, label, colour in [
        (180, "180 days", "rgba(255,165,0,0.7)"),
        (365, "1 year",   "rgba(220,50,50,0.6)"),
        (730, "2 years",  "rgba(150,0,0,0.5)"),
    ]:
        if max_days >= xval:
            fig.add_vline(
                x=xval, line_dash="dash", line_color=colour, line_width=1.5,
                annotation_text=label, annotation_position="top",
                annotation_font_size=11, annotation_font_color=colour,
            )

    # Y axis labels (package names)
    y_tickvals = list(range(len(packages_ordered)))
    y_ticktext = packages_ordered

    fig.update_layout(
        title=dict(
            text    = "Lapsed Members — Days Since Expiry  ×  Subscription Package  ×  Branch",
            font    = dict(size=16, family="Arial", color="#1F3864"),
            x=0.5, xanchor="center",
        ),
        xaxis=dict(
            title     = "Days Since Subscription Expired",
            tickfont  = dict(size=12, family="Arial"),
            gridcolor = "#E8E8E8",
            zeroline  = False,
            tickformat = ",d",
        ),
        yaxis=dict(
            title    = "Package / Subscription Plan",
            tickfont = dict(size=12, family="Arial"),
            tickmode = "array",
            tickvals = y_tickvals,
            ticktext = y_ticktext,
            gridcolor= "#E8E8E8",
        ),
        legend=dict(
            title=dict(text="Branch ID", font=dict(size=12)),
            font=dict(size=9),
            itemsizing="constant",
        ),
        plot_bgcolor  = "#FAFAFA",
        paper_bgcolor = "#FFFFFF",
        font          = dict(family="Arial"),
        height        = 700,
        width         = 1300,
        margin        = dict(l=80, r=200, t=80, b=80),
        hoverlabel    = dict(bgcolor="white", font_size=12, font_family="Arial"),
    )

    # Summary annotation (bottom-left)
    fig.add_annotation(
        x=0.01, y=0.02, xref="paper", yref="paper",
        text=(f"<b>Total members:</b> {len(df):,}<br>"
              f"<b>Branches:</b> {n_branches}<br>"
              f"<b>Packages:</b> {len(packages_ordered)}<br>"
              f"<b>Median days expired:</b> {int(df['days_expired'].median())}"),
        showarrow=False, align="left",
        bgcolor="rgba(255,255,255,0.85)",
        bordercolor="#CCCCCC", borderwidth=1,
        font=dict(size=11, family="Arial"),
    )

    fig.write_html(OUT_HTML)
    print(f"✅ Saved: '{OUT_HTML}'  ← open in any browser")


def build_static_plot(df, packages_ordered):
    """Static matplotlib PNG for presentations."""
    try:
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        from matplotlib.lines import Line2D
    except ImportError:
        print("❌ matplotlib not installed.")
        return

    print("\n" + "=" * 60)
    print("BUILDING STATIC PNG PLOT")
    print("=" * 60)

    branches      = sorted(df["member_branch_id"].unique())
    n_branches    = len(branches)
    highlight_str = str(HIGHLIGHT_BRANCH) if HIGHLIGHT_BRANCH else None

    cmap  = plt.cm.get_cmap("tab20",  min(n_branches, 20))
    cmap2 = plt.cm.get_cmap("tab20b", max(n_branches - 20, 1))

    def get_colour(i):
        return cmap(i) if i < 20 else cmap2(i - 20)

    branch_colour = {b: get_colour(i) for i, b in enumerate(branches)}

    fig, ax = plt.subplots(figsize=(16, 8))
    fig.patch.set_facecolor("#FAFAFA")
    ax.set_facecolor("#FAFAFA")

    for branch in branches:
        if branch == highlight_str:
            continue
        sub = df[df["member_branch_id"] == branch]
        ax.scatter(sub["days_expired"], sub["package_num_jittered"],
                   c=[branch_colour[branch]], s=18, alpha=0.55,
                   linewidths=0, zorder=2)

    if highlight_str and highlight_str in df["member_branch_id"].values:
        sub_hl = df[df["member_branch_id"] == highlight_str]
        ax.scatter(sub_hl["days_expired"], sub_hl["package_num_jittered"],
                   c="#FF2222", s=60, marker="*", alpha=0.9,
                   linewidths=0.3, edgecolors="white", zorder=5,
                   label=f"Branch {HIGHLIGHT_BRANCH} (★)")

    # Reference lines
    for xval, label, colour, ls in [
        (180, "180d", "orange",  "--"),
        (365, "1yr",  "red",     "--"),
        (730, "2yr",  "darkred", ":"),
    ]:
        if df["days_expired"].max() >= xval:
            ax.axvline(x=xval, color=colour, linestyle=ls, linewidth=1.2, alpha=0.7)
            ax.text(xval + 5, len(packages_ordered) - 0.3, label,
                    color=colour, fontsize=9, va="top")

    ax.set_yticks(range(len(packages_ordered)))
    ax.set_yticklabels(packages_ordered, fontsize=10)
    ax.set_ylim(-0.8, len(packages_ordered) - 0.2)

    # Legend — top 15 branches only
    top_branches   = df["member_branch_id"].value_counts().head(15).index.tolist()
    legend_handles = []
    if highlight_str and highlight_str in df["member_branch_id"].values:
        legend_handles.append(Line2D([0],[0], marker="*", color="w",
                                     markerfacecolor="#FF2222", markersize=12,
                                     label=f"★ Branch {HIGHLIGHT_BRANCH}"))
    for b in top_branches:
        if b == highlight_str:
            continue
        legend_handles.append(mpatches.Patch(color=branch_colour[b],
                                              label=f"Branch {b}"))
    if n_branches > 15:
        legend_handles.append(mpatches.Patch(color="lightgrey",
                                              label=f"+ {n_branches-15} more"))

    ax.legend(handles=legend_handles, title="Branch (top 15)",
              loc="upper right", fontsize=8, title_fontsize=9,
              framealpha=0.9, ncol=2 if len(legend_handles) > 10 else 1)

    ax.set_title(
        f"Lapsed Members — Days Since Expiry  ×  Package  ×  Branch\n"
        f"{len(df):,} members  ·  {n_branches} branches  ·  {len(packages_ordered)} packages  ·  180+ days expired only",
        fontsize=13, fontweight="bold", color="#1F3864", pad=14,
    )
    ax.set_xlabel("Days Since Subscription Expired", fontsize=12, labelpad=10)
    ax.set_ylabel("Package / Subscription Plan",     fontsize=12, labelpad=10)
    ax.xaxis.set_major_formatter(
        __import__("matplotlib.ticker", fromlist=["FuncFormatter"])
        .FuncFormatter(lambda v, _: f"{int(v):,}"))
    ax.grid(axis="both", color="#E0E0E0", linewidth=0.6, zorder=1)
    ax.spines[["top", "right"]].set_visible(False)

    plt.tight_layout()
    plt.savefig(OUT_PNG, dpi=150, bbox_inches="tight", facecolor="#FAFAFA")
    print(f"✅ Saved: '{OUT_PNG}'")
    plt.close()


# =============================================================================
# MAIN
# =============================================================================
def main():
    print("\n" + "█" * 60)
    print("  SCATTER — Days Expired  ×  Package ID  ×  Branch")
    print("█" * 60)

    try:
        import plotly
        has_plotly = True
    except ImportError:
        has_plotly = False
        print("⚠️  plotly not found — static plot only")
        print("   pip install plotly --break-system-packages")

    try:
        import matplotlib
        has_matplotlib = True
    except ImportError:
        has_matplotlib = False

    df, packages_ordered, pkg_to_num = load_and_prepare()

    if has_plotly:
        build_interactive_plot(df, packages_ordered)

    if has_matplotlib:
        build_static_plot(df, packages_ordered)

    print("\n" + "█" * 60)
    print("  ✅ DONE")
    if has_plotly:
        print(f"  🌐 {OUT_HTML}  ← open in browser, hover for member details")
    if has_matplotlib:
        print(f"  🖼️  {OUT_PNG}   ← for presentations")
    print("█" * 60)

    print("\nWHAT TO LOOK FOR:")
    print("  Vertical clusters  → old lapsed members concentrated on one package")
    print("  Branch dominance   → one branch churning more from a specific plan")
    print("  Points near 180d   → recent churn wave worth immediate outreach")
    print("  Points at 700d+    → long-unrecovered members needing nurture campaign")
    print("  ★ Branch 1008      → compare its spread vs other branches\n")


if __name__ == "__main__":
    main()