#!/usr/bin/env python3
"""Recompute archived screening summaries and draw submission figures.

This script reproduces *conditional* screening outputs from the archived
scenario files. It does not model a tBLG-coated sodium-ion battery and it
must not be used as evidence of electrochemical performance, manufacturing
feasibility, or commercial competitiveness.
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path

import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
FIGURES = ROOT / "figures"


def font(size: int, bold: bool = False):
    names = [
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for name in names:
        if Path(name).exists():
            return ImageFont.truetype(name, size)
    return ImageFont.load_default()


# Sized for legibility after a 4000 px figure is placed at journal-column width.
# The previous values were tuned against Pillow's bitmap fallback font.  Once a
# real Arial font was found, those values became far too large and caused
# collisions.  These sizes correspond to roughly 10.5--27 pt at 300 dpi.
F_TITLE = font(108, True)
F_SUBTITLE = font(58)
F_BODY = font(60)
F_SMALL = font(48)
F_LABEL = font(60, True)
F_TICK = font(44)
NAVY, BLUE, TEAL, ORANGE, GREY = "#19334d", "#3978b8", "#2aa981", "#df7b22", "#59636e"
PALE = ["#cfe0f5", "#c9f1dc", "#fff0bb", "#f7dce8", "#e7e3fa", "#e2e4e8"]
GRID = "#aeb8c2"


def canvas(size=(4000, 1600)):
    return Image.new("RGB", size, "white")


def wrap(draw, text, fnt, width):
    words, lines, line = text.split(), [], ""
    for word in words:
        test = (line + " " + word).strip()
        if draw.textlength(test, font=fnt) <= width:
            line = test
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines


def centered(draw, xy, text, fnt, fill=NAVY, width=None, spacing=8):
    x, y = xy
    lines = wrap(draw, text, fnt, width) if width else [text]
    block = "\n".join(lines)
    box = draw.multiline_textbbox((0, 0), block, font=fnt, spacing=spacing, align="center")
    draw.multiline_text(
        (x - (box[2] - box[0]) / 2 - box[0], y - box[1]),
        block,
        font=fnt,
        fill=fill,
        spacing=spacing,
        align="center",
    )
    return y + box[3] - box[1]


def arrow(draw, start, end, fill=NAVY, width=8, head=28):
    """Draw a correctly oriented arrow between two points."""
    x1, y1 = start
    x2, y2 = end
    dx, dy = x2 - x1, y2 - y1
    length = math.hypot(dx, dy)
    ux, uy = dx / length, dy / length
    px, py = -uy, ux
    draw.line((x1, y1, x2, y2), fill=fill, width=width)
    base_x, base_y = x2 - ux * head, y2 - uy * head
    draw.polygon(
        [
            (x2, y2),
            (base_x + px * head * 0.55, base_y + py * head * 0.55),
            (base_x - px * head * 0.55, base_y - py * head * 0.55),
        ],
        fill=fill,
    )


def rounded(draw, box, fill, outline=NAVY, radius=28, width=5):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def dashed_line(draw, start, end, fill=GREY, width=6, dash=28, gap=18):
    """Draw a non-directional dashed connector."""
    x1, y1 = start
    x2, y2 = end
    dx, dy = x2 - x1, y2 - y1
    length = math.hypot(dx, dy)
    if length == 0:
        return
    ux, uy = dx / length, dy / length
    position = 0.0
    while position < length:
        finish = min(position + dash, length)
        draw.line(
            (
                x1 + ux * position,
                y1 + uy * position,
                x1 + ux * finish,
                y1 + uy * finish,
            ),
            fill=fill,
            width=width,
        )
        position += dash + gap


def pill(draw, center, text, fill, text_fill="white", width=620, height=92):
    x, y = center
    rounded(
        draw,
        (x - width / 2, y - height / 2, x + width / 2, y + height / 2),
        fill,
        outline=fill,
        radius=int(height / 2),
        width=2,
    )
    centered(draw, (x, y - 28), text, font(43, True), text_fill, width=width - 40)


def draw_particle(draw, center, radius, *, coating=True, muted=False):
    """Draw a stylized hard-carbon particle and optional discontinuous coating."""
    cx, cy = center
    fill = "#d7dde3" if muted else "#8996a3"
    outline = "#7a858f" if muted else NAVY
    draw.ellipse((cx - radius, cy - radius, cx + radius, cy + radius), fill=fill, outline=outline, width=8)
    for dx, dy, scale in ((-0.35, -0.22, 0.18), (0.28, -0.32, 0.14), (-0.12, 0.28, 0.22), (0.38, 0.18, 0.12)):
        r = radius * scale
        draw.ellipse((cx + dx * radius - r, cy + dy * radius - r, cx + dx * radius + r, cy + dy * radius + r), outline="#f3f5f7", width=7)
    if coating:
        coat = "#2aa981" if not muted else "#98aaa4"
        for start, end in ((205, 305), (325, 68), (88, 178)):
            draw.arc((cx - radius - 28, cy - radius - 28, cx + radius + 28, cy + radius + 28), start=start, end=end, fill=coat, width=24)


def draw_document(draw, box, *, fill="white", outline=NAVY, folded=True, lines=True):
    x1, y1, x2, y2 = box
    fold = min(70, (x2 - x1) * 0.18)
    if folded:
        polygon = [(x1, y1), (x2 - fold, y1), (x2, y1 + fold), (x2, y2), (x1, y2)]
        draw.polygon(polygon, fill=fill, outline=outline)
        draw.line((x2 - fold, y1, x2 - fold, y1 + fold, x2, y1 + fold), fill=outline, width=5)
    else:
        draw.rectangle(box, fill=fill, outline=outline, width=6)
    if lines:
        for offset in (0.38, 0.55, 0.72):
            y = y1 + (y2 - y1) * offset
            draw.line((x1 + 45, y, x2 - 45, y), fill="#9ba6af", width=5)


def draw_lock(draw, center, scale=1.0, fill="#8e3028"):
    cx, cy = center
    w, h = 120 * scale, 105 * scale
    draw.arc((cx - w * 0.36, cy - h * 0.95, cx + w * 0.36, cy - h * 0.20), 180, 360, fill=fill, width=max(4, int(12 * scale)))
    rounded(draw, (cx - w / 2, cy - h * 0.35, cx + w / 2, cy + h * 0.45), "#f8e4e1", outline=fill, radius=max(8, int(14 * scale)), width=max(3, int(7 * scale)))
    draw.ellipse((cx - 9 * scale, cy - 2 * scale, cx + 9 * scale, cy + 16 * scale), fill=fill)


def save(img, name, dpi=600):
    FIGURES.mkdir(parents=True, exist_ok=True)
    img.save(FIGURES / name, dpi=(dpi, dpi))


def fig1():
    # Scientific evidence-boundary schematic; deliberately not a process flow.
    img = canvas((4500, 2640))
    d = ImageDraw.Draw(img)

    # Supported planar aqueous model system.
    d.rounded_rectangle((120, 170, 1510, 2240), radius=110, fill="#eef6ff", outline="#a9c5df", width=6)
    pill(d, (815, 245), "SUPPORTED MODEL SYSTEM", "#006B4F", width=770)
    # Beaker and aqueous phase.
    d.line((330, 600, 330, 1670, 1270, 1670, 1270, 600), fill=NAVY, width=13)
    d.line((280, 600, 1320, 600), fill=NAVY, width=13)
    d.rectangle((345, 925, 1255, 1655), fill="#d9effb")
    d.line((345, 925, 1255, 925), fill=BLUE, width=7)
    # Two twisted graphene sheets represented as offset lattices.  Draw each
    # lattice on a temporary layer and paste it through its polygon mask so no
    # line can escape the physical sheet or cross the beaker/evidence boundary.
    for layer_y, color, offset in ((1320, NAVY, 0), (1435, TEAL, 38)):
        polygon = (
            (430 + offset, layer_y),
            (1120 + offset, layer_y - 110),
            (1190 + offset, layer_y + 80),
            (500 + offset, layer_y + 190),
        )
        layer = img.copy()
        layer_draw = ImageDraw.Draw(layer)
        layer_draw.polygon(polygon, fill="#ffffff")
        for i in range(7):
            x = 500 + offset + i * 100
            layer_draw.line((x, layer_y + 45, x + 500, layer_y - 35), fill=color, width=5)
        for i in range(4):
            y = layer_y - 45 + i * 60
            layer_draw.line((510 + offset, y, 1140 + offset, y + 100), fill=color, width=5)
        mask = Image.new("L", img.size, 0)
        ImageDraw.Draw(mask).polygon(polygon, fill=255)
        img.paste(layer, (0, 0), mask)
        d.line((*polygon, polygon[0]), fill=color, width=4, joint="curve")
    # Restore the vessel outline above the clipped sheets for a clean edge.
    d.line((330, 600, 330, 1670, 1270, 1670, 1270, 600), fill=NAVY, width=13)
    d.line((280, 600, 1320, 600), fill=NAVY, width=13)
    # Outer-sphere probes remain in solution rather than intercalating.
    probes = [(520, 820, "Ox", ORANGE), (825, 755, "Red", BLUE), (1080, 850, "Ox", ORANGE)]
    for x, y, label, color in probes:
        d.ellipse((x - 74, y - 74, x + 74, y + 74), fill="white", outline=color, width=10)
        centered(d, (x, y - 28), label, font(43, True), color)
        dashed_line(d, (x, y + 86), (x, 1220), fill=color, width=5, dash=18, gap=14)
    centered(d, (815, 1785), "Planar tBLG in aqueous solution", font(56, True), NAVY, width=1180)
    centered(d, (815, 1885), "Twist-sensitive heterogeneous k0", font(51), NAVY, width=1180)
    centered(d, (815, 1980), "for reported outer-sphere probes", font(48), GREY, width=1180)

    # Explicit boundary between observation and hypothesis.  The label starts
    # exactly at the right edge of the supported-system panel, so it neither
    # overlaps that panel nor reads as a second header inside it.
    boundary_x = 1815
    dashed_line(d, (boundary_x, 330), (boundary_x, 2240), fill="#9a541f", width=10, dash=42, gap=28)
    pill(d, (boundary_x, 245), "EVIDENCE BOUNDARY", "#fff4e6", text_fill="#8a4d17", width=600)

    # Proposed particle-level system, rendered as an unresolved physical object.
    # Center this label over the complete untested domain (particle plus cell,
    # process, commercial, and policy contexts), not only over the particle.
    pill(d, (3050, 245), "PROPOSED • NOT TESTED", "#F2B544", text_fill="#2b2b2b", width=760)
    draw_particle(d, (2600, 1170), 520, coating=True)
    for x, y in ((2060, 720), (2110, 1580), (3030, 680), (3170, 1410), (2870, 1760)):
        d.ellipse((x - 52, y - 52, x + 52, y + 52), fill="#f3f7ff", outline=BLUE, width=7)
        centered(d, (x, y - 24), "Na+", font(36, True), BLUE)
    centered(d, (2600, 1840), "tBLG-coated hard-carbon particle", font(55, True), NAVY, width=1200)
    centered(d, (2600, 1940), "coverage • adhesion • pore access", font(47), GREY, width=1250)
    centered(d, (2600, 2025), "desolvation • SEI • Na storage", font(47), GREY, width=1250)
    centered(d, (2600, 2115), "all unmeasured", font(48, True), "#8e3028")

    # Downstream application context: pictorial icons, not a sequence.
    x0 = 3570
    # Pouch cell.
    rounded(d, (3470, 450, 4200, 930), "#f4f5f6", outline="#7c8791", radius=45, width=8)
    d.rectangle((3640, 360, 3770, 455), fill="#7c8791")
    d.rectangle((3900, 360, 4030, 455), fill="#7c8791")
    d.line((3665, 690, 4005, 690), fill="#7c8791", width=8)
    d.line((3835, 555, 3835, 825), fill="#7c8791", width=8)
    centered(d, (3835, 970), "Cell performance", font(49, True), NAVY)
    centered(d, (3835, 1055), "not established", font(43), "#8e3028")
    # Factory.
    d.rectangle((3470, 1240, 4200, 1700), fill="#f4f5f6", outline="#7c8791", width=8)
    d.polygon(((3470, 1340), (3710, 1200), (3710, 1340), (3950, 1200), (3950, 1340), (4200, 1210), (4200, 1700), (3470, 1700)), fill="#e2e5e8", outline="#7c8791")
    d.rectangle((3550, 1440, 3700, 1600), fill="white", outline="#7c8791", width=5)
    d.rectangle((3820, 1440, 3970, 1600), fill="white", outline="#7c8791", width=5)
    centered(d, (3835, 1760), "Process feasibility", font(49, True), NAVY)
    centered(d, (3835, 1845), "not established", font(43), "#8e3028")
    # Policy / market records.
    draw_document(d, (3560, 1900, 4110, 2230), fill="#f4f5f6", outline="#7c8791")
    centered(d, (3835, 2270), "Commercial and policy evidence", font(45, True), NAVY, width=800)
    centered(d, (3835, 2365), "not established", font(43), "#8e3028")

    # The permitted-inference statement belongs in the manuscript caption;
    # omitting it here avoids a title-like footer inside the artwork.
    save(img, "Figure_1_evidence_gates.png")


def fig2(tea):
    img = canvas()
    d = ImageDraw.Draw(img)

    # Panel A: time series with explicit ticks and a legend that sits above the
    # data rather than over the first observations.
    d.rectangle((135, 205, 1970, 1400), outline=GRID, width=3)
    d.text((175, 225), "(a)", font=F_LABEL, fill=NAVY)
    d.text((320, 300), "Illustrative scenario cost (USD/kWh)", font=F_SMALL, fill=GREY)
    left, top, right, bottom = 320, 390, 1900, 1235
    years = tea["Year"].to_numpy()
    hc = tea["HC_SIB_USD_kWh"].to_numpy()
    tb = tea["tBLG_SIB_Baseline_USD_kWh"].to_numpy()
    ymax = 400

    def pt(i, value):
        return (
            left + i * (right - left) / (len(years) - 1),
            bottom - value * (bottom - top) / ymax,
        )

    for value in range(0, 401, 100):
        y = pt(0, value)[1]
        d.line((left, y, right, y), fill="#c4ccd4", width=2)
        d.text((left - 28, y), str(value), font=F_TICK, fill=GREY, anchor="rm")
    for idx in (0, len(years) // 2, len(years) - 1):
        x = pt(idx, 0)[0]
        d.line((x, bottom, x, bottom + 10), fill=GREY, width=3)
        d.text((x, bottom + 18), str(years[idx]), font=F_TICK, fill=GREY, anchor="ma")
    for values, color in ((hc, BLUE), (tb, ORANGE)):
        points = [pt(i, value) for i, value in enumerate(values)]
        d.line(points, fill=color, width=7)
        for point in points:
            d.ellipse((point[0] - 10, point[1] - 10, point[0] + 10, point[1] + 10), fill=color)
    legend_x = 1030
    for y, color, name in (
        (410, BLUE, "Conventional HC-SIB reference"),
        (473, ORANGE, "tBLG-SIB scenario input"),
    ):
        d.line((legend_x, y + 21, legend_x + 95, y + 21), fill=color, width=7)
        d.text((legend_x + 120, y), name, font=font(43), fill=color)
    centered(d, ((left + right) / 2, 1315), "Year", F_SMALL, GREY)

    # Panel B: horizontal rank associations.  The label and data regions are
    # separated so negative bars cannot collide with long category names.
    d.rectangle((2070, 205, 3865, 1400), outline=GRID, width=3)
    d.text((2110, 225), "(b)", font=F_LABEL, fill=NAVY)
    corr_cols = ["CVD_Coating_USD_kWh","Raw_Material_USD_kWh","Cell_Assembly_USD_kWh","BOS_USD_kWh","Overhead_USD_kWh","Cumulative_Production_GWh","Learning_Rate"]
    labels = ["CVD coating","Raw materials","Cell assembly","BOS","Overhead","Cumulative production","Learning rate"]
    mc = pd.read_csv(DATA / "Monte_Carlo_TEA_2035_Cost.csv")
    cost = mc["Total_System_Cost_USD_kWh"]
    rho = [mc[c].rank().corr(cost.rank()) for c in corr_cols]
    plot_left, plot_right = 2780, 3785
    x_min, x_max = -0.25, 0.80

    def corr_x(value):
        return plot_left + (value - x_min) * (plot_right - plot_left) / (x_max - x_min)

    for tick in (-0.2, 0.0, 0.4, 0.8):
        x = corr_x(tick)
        d.line((x, 345, x, 1215), fill="#c4ccd4", width=2)
        d.text((x, 1230), f"{tick:.1f}", font=F_TICK, fill=GREY, anchor="ma")
    zero = corr_x(0)
    d.line((zero, 335, zero, 1215), fill=GREY, width=4)
    for i, (lab, val) in enumerate(zip(labels, rho)):
        y = 380 + i * 119
        end = corr_x(val)
        d.text((2120, y + 25), lab, font=font(46), fill=NAVY, anchor="lm")
        d.rectangle((min(zero, end), y, max(zero, end), y + 52), fill=ORANGE if val >= 0 else BLUE)
        if val >= 0:
            d.text((end + 16, y + 26), f"{val:.2f}", font=F_TICK, fill=NAVY, anchor="lm")
        else:
            d.text((end - 16, y + 26), f"{val:.2f}", font=F_TICK, fill=NAVY, anchor="rm")
    centered(d, (2990, 1322), "Spearman rank correlation with 2035 scenario cost", font(44), GREY, width=1600)
    # Remove the former title zone and pad to the exact manuscript frame ratio.
    content = img.crop((0, 130, 4000, 1480))
    output = canvas((4000, 1675))
    output.paste(content, (0, 120))
    save(output,"Figure_2_cost_trajectories_and_rank_correlations.png")


def fig3(mc):
    img = canvas()
    d = ImageDraw.Draw(img)
    cost = mc["Total_System_Cost_USD_kWh"].to_numpy()
    bins = np.linspace(cost.min(), cost.max(), 31)
    hist, edges = np.histogram(cost, bins)

    # Panel A: histogram with a separate statistic key so the near-coincident
    # mean and median annotations cannot overlap.
    d.rectangle((135, 205, 1970, 1400), outline=GRID, width=3)
    d.text((175, 225), "(a)", font=F_LABEL, fill=NAVY)
    d.text((320, 300), "Frequency", font=F_SMALL, fill=GREY)
    left, top, right, bottom = 320, 455, 1900, 1235
    hmax = hist.max()

    def hist_x(value):
        return left + (value - edges[0]) * (right - left) / (edges[-1] - edges[0])

    for i, count in enumerate(hist):
        x1 = left + i * (right - left) / len(hist)
        x2 = left + (i + 1) * (right - left) / len(hist) - 3
        y = bottom - count * (bottom - top) / hmax
        d.rectangle((x1, y, x2, bottom), fill="#739bc5")
    for tick in (0, round(hmax / 2), hmax):
        y = bottom - tick * (bottom - top) / hmax
        d.line((left, y, right, y), fill="#c4ccd4", width=2)
        d.text((left - 24, y), str(tick), font=F_TICK, fill=GREY, anchor="rm")
    for tick in np.linspace(edges[0], edges[-1], 5):
        x = hist_x(tick)
        d.line((x, bottom, x, bottom + 10), fill=GREY, width=3)
        d.text((x, bottom + 18), f"{tick:.0f}", font=F_TICK, fill=GREY, anchor="ma")
    stats = [
        ("P5", np.percentile(cost, 5), ORANGE),
        ("Median", np.median(cost), "#111111"),
        ("Mean", cost.mean(), "#c84630"),
        ("P95", np.percentile(cost, 95), ORANGE),
    ]
    for _, value, color in stats:
        x = hist_x(value)
        d.line((x, top, x, bottom), fill=color, width=5)
    for x, (label, value, color) in zip((330, 720, 1100, 1485), stats):
        d.line((x, 374, x + 72, 374), fill=color, width=7)
        d.text((x + 90, 374), f"{label} {value:.1f}", font=font(42), fill=color, anchor="lm")
    centered(d, ((left + right) / 2, 1315), "2035 scenario cost (USD/kWh)", F_SMALL, GREY)

    # Panel B: horizontal bars provide sufficient room for component names.
    d.rectangle((2070, 205, 3865, 1400), outline=GRID, width=3)
    d.text((2110, 225), "(b)", font=F_LABEL, fill=NAVY)
    comps=["Raw_Material_USD_kWh","CVD_Coating_USD_kWh","Cell_Assembly_USD_kWh","BOS_USD_kWh","Overhead_USD_kWh"]
    labs=["Raw materials","CVD coating","Cell assembly","BOS","Overhead"]
    cols=[BLUE,ORANGE,TEAL,"#b977a6","#d5be22"]
    adjusted=[mc[c]*mc["LR_Multiplier"] for c in comps]
    means=[series.mean() for series in adjusted]
    intervals=[np.percentile(series,[5,95]) for series in adjusted]
    xmax=math.ceil(max(interval[1] for interval in intervals)/10)*10
    plot_left, plot_right = 2810, 3780

    def component_x(value):
        return plot_left + value * (plot_right - plot_left) / xmax

    for tick in np.linspace(0, xmax, 5):
        x = component_x(tick)
        d.line((x, 350, x, 1200), fill="#c4ccd4", width=2)
        d.text((x, 1220), f"{tick:.0f}", font=F_TICK, fill=GREY, anchor="ma")
    for i,(lab,mean,interval,color) in enumerate(zip(labs,means,intervals,cols)):
        y = 390 + i * 160
        d.text((2125, y + 28), lab, font=font(48), fill=NAVY, anchor="lm")
        end = component_x(mean)
        d.rectangle((plot_left, y, end, y + 56), fill=color)
        low, high = map(component_x, interval)
        d.line((low, y + 28, high, y + 28), fill="#222222", width=5)
        d.line((low, y + 13, low, y + 43), fill="#222222", width=4)
        d.line((high, y + 13, high, y + 43), fill="#222222", width=4)
        d.text((end + 14, y + 28), f"{mean:.1f}", font=F_TICK, fill=NAVY, anchor="lm")
    centered(d, ((plot_left + plot_right) / 2, 1315), "Learning-adjusted cost (USD/kWh)", F_SMALL, GREY)
    content = img.crop((0, 130, 4000, 1480))
    output = canvas((4000, 1540))
    output.paste(content, (0, 80))
    save(output,"Figure_3_monte_carlo_cost_envelope.png")


def fig4():
    # Non-sequential evidence-domain schematic arranged around the concept.
    img = canvas((4500, 1886))
    d = ImageDraw.Draw(img)
    center = (2250, 890)
    domain_centers = [
        (700, 390),
        (2250, 255),
        (3800, 390),
        (3800, 1300),
        (2250, 1535),
        (700, 1300),
    ]
    domains = [
        ("Electrolyte", "", "electrolyte"),
        ("Particle coating", "", "coating"),
        ("Half-cell", "", "halfcell"),
        ("Full cell", "", "fullcell"),
        ("Process / LCA", "", "factory"),
        ("Innovation system", "", "people"),
    ]

    # Non-directional dashed links denote interfaces that require calibration.
    for node in domain_centers:
        dx, dy = node[0] - center[0], node[1] - center[1]
        length = math.hypot(dx, dy)
        start = (center[0] + dx * 355 / length, center[1] + dy * 355 / length)
        end = (node[0] - dx * 235 / length, node[1] - dy * 235 / length)
        dashed_line(d, start, end, fill="#9aa5af", width=7, dash=25, gap=18)

    # Central proposed material system.
    draw_particle(d, center, 280, coating=True)
    for start in range(0, 360, 35):
        d.arc((center[0] - 350, center[1] - 350, center[0] + 350, center[1] + 350), start=start, end=start + 19, fill="#F2B544", width=12)
    centered(d, (center[0], center[1] - 92), "Proposed tBLG coating", font(40, True), "white", width=500)
    centered(d, (center[0], center[1] - 25), "on hard carbon", font(40, True), "white", width=500)
    pill(d, (center[0], center[1] + 92), "NOT VALIDATED", "#F2B544", text_fill="#2b2b2b", width=430, height=72)

    def domain(node, title, subtitle, kind, fill):
        x, y = node
        radius = 220
        d.ellipse((x - radius, y - radius, x + radius, y + radius), fill=fill, outline=NAVY, width=7)
        iy = y - 62
        if kind == "electrolyte":
            d.polygon(((x - 68, iy - 70), (x + 68, iy - 70), (x + 105, iy + 75), (x - 105, iy + 75)), fill="white", outline=NAVY)
            d.polygon(((x - 88, iy + 12), (x + 88, iy + 12), (x + 105, iy + 75), (x - 105, iy + 75)), fill="#a9d8f0")
            for px, py in ((x - 42, iy - 2), (x + 28, iy - 25), (x + 58, iy + 30)):
                d.ellipse((px - 16, py - 16, px + 16, py + 16), fill=ORANGE)
        elif kind == "coating":
            draw_particle(d, (x, iy), 100, coating=True)
        elif kind == "halfcell":
            d.ellipse((x - 95, iy - 68, x + 95, iy + 68), fill="#d9dee3", outline=NAVY, width=8)
            d.ellipse((x - 75, iy - 52, x + 75, iy + 52), fill="white", outline=TEAL, width=7)
            d.line((x - 70, iy, x + 70, iy), fill=ORANGE, width=9)
        elif kind == "fullcell":
            rounded(d, (x - 105, iy - 95, x + 105, iy + 90), "white", outline=NAVY, radius=24, width=8)
            d.rectangle((x - 68, iy - 132, x - 18, iy - 91), fill=NAVY)
            d.rectangle((x + 18, iy - 132, x + 68, iy - 91), fill=NAVY)
            d.text((x, iy - 5), "Na+", font=font(47, True), fill=BLUE, anchor="mm")
        elif kind == "factory":
            d.polygon(((x - 125, iy + 75), (x - 125, iy - 10), (x - 30, iy - 70), (x - 30, iy - 10), (x + 55, iy - 70), (x + 55, iy - 10), (x + 125, iy - 60), (x + 125, iy + 75)), fill="white", outline=NAVY)
            for px in (x - 95, x - 10, x + 78):
                d.rectangle((px, iy + 18, px + 38, iy + 60), fill="#d9effb", outline=NAVY, width=4)
        else:
            for px, py in ((x - 72, iy), (x, iy - 45), (x + 72, iy)):
                d.ellipse((px - 30, py - 30, px + 30, py + 30), fill="white", outline=NAVY, width=6)
                d.arc((px - 55, py + 10, px + 55, py + 105), 190, 350, fill=NAVY, width=7)
        centered(d, (x, y + 82), title, font(60, True), NAVY, width=390, spacing=3)

    for node, (title, subtitle, kind), fill in zip(domain_centers, domains, PALE):
        domain(node, title, subtitle, kind, fill)

    # Interpretation is carried by the manuscript caption, not repeated as an image title.
    save(img,"Figure_4_integrated_assessment_chain.png")


def draw_cell(draw, box, value, cell_font, fill=NAVY, padding=22, spacing=5):
    """Draw wrapped, vertically centered text within a table-like cell."""
    x1, y1, x2, y2 = box
    lines = wrap(draw, str(value), cell_font, x2 - x1 - 2 * padding)
    block = "\n".join(lines)
    bounds = draw.multiline_textbbox((0, 0), block, font=cell_font, spacing=spacing)
    block_height = bounds[3] - bounds[1]
    y = y1 + max(padding, (y2 - y1 - block_height) / 2) - bounds[1]
    draw.multiline_text((x1 + padding, y), block, font=cell_font, fill=fill, spacing=spacing)


def fig5(audit):
    # Pictorial archive audit; deliberately avoids a table or quantitative chart.
    img = Image.new("RGB", (4500, 2180), "white")
    d = ImageDraw.Draw(img)
    statuses = set(audit["Status"].astype(str))
    if "Available" not in statuses or "Available as legacy file" not in statuses:
        raise ValueError("Evidence audit does not contain the expected archived records.")

    # Archive folder with the two retained record types.
    d.polygon(((180, 520), (650, 520), (820, 690), (1720, 690), (1720, 1760), (180, 1760)), fill="#d9e8f6", outline=NAVY)
    d.line((180, 690, 1720, 690), fill=NAVY, width=10)
    centered(d, (950, 390), "Archive contents", font(70, True), NAVY)
    draw_document(d, (350, 790, 1480, 1230), fill="#e4f5ee", outline="#006B4F")
    pill(d, (915, 875), "AVAILABLE", "#006B4F", width=430, height=78)
    centered(d, (915, 1030), "TIS function definitions", font(66, True), NAVY, width=900)
    draw_document(d, (440, 1270, 1570, 1670), fill="#fff1c5", outline="#b77a15")
    pill(d, (1005, 1375), "LEGACY FILE", "#F2B544", text_fill="#2b2b2b", width=450, height=78)
    centered(d, (1005, 1475), "Aggregate author-coded labels", font(58, True), NAVY, width=1040)
    centered(d, (1005, 1575), "retained • not inferential", font(52), "#8a4d17", width=1000)

    # Magnifier highlights the absence of reconstructable source records.
    d.ellipse((1880, 690, 2760, 1570), fill="#f7f9fa", outline=BLUE, width=24)
    d.line((2590, 1440, 2940, 1750), fill=BLUE, width=58)
    d.ellipse((2190, 1010, 2450, 1270), outline="#9ba6af", width=8)
    dashed_line(d, (2070, 880), (2570, 1380), fill="#b7c0c8", width=5, dash=22, gap=14)
    centered(d, (2320, 1110), "?", font(150, True), "#8e3028")
    centered(d, (2320, 1605), "Missing source records", font(62, True), "#8e3028")

    missing_items = [
        "Raw indicators",
        "Search protocol",
        "Codebook & weights",
        "Independent coding",
        "Stakeholder evidence",
    ]
    positions = [(3160, 360), (3910, 600), (3690, 1040), (3160, 1460), (3980, 1490)]
    for (x, y), label in zip(positions, missing_items):
        draw_document(d, (x - 155, y - 145, x + 155, y + 105), fill="#fbf1ef", outline="#8e3028")
        d.line((x - 65, y - 55, x + 65, y + 75), fill="#8e3028", width=14)
        d.line((x + 65, y - 55, x - 65, y + 75), fill="#8e3028", width=14)
        centered(d, (x, y + 155), label, font(58, True), NAVY, width=650)

    # Locked inference types make the permitted interpretation explicit.
    lock_y = 1840
    for x, label in ((1850, "No ranking"), (2650, "No commercialization claim"), (3650, "No policy recommendation")):
        draw_lock(d, (x, lock_y - 20), scale=0.75)
        centered(d, (x, lock_y + 85), label, font(54, True), "#8e3028", width=820)
    # The provenance-only conclusion is stated in the caption rather than as a
    # title-like footer inside the image.
    save(img, "Figure_5_tis_rubric_and_provenance.png")


def main():
    parser=argparse.ArgumentParser(); parser.add_argument("--no-figures",action="store_true"); args=parser.parse_args()
    mc=pd.read_csv(DATA / "Monte_Carlo_TEA_2035_Cost.csv")
    tea=pd.read_csv(DATA / "TEA_Cost_Trajectories.csv")
    audit=pd.read_csv(DATA / "TIS_Evidence_Audit.csv", keep_default_na=False)
    total=mc["Total_System_Cost_USD_kWh"]
    summary=pd.DataFrame({"statistic":["mean","median","P5","P95","standard_deviation"],"USD_kWh":[total.mean(),total.median(),total.quantile(.05),total.quantile(.95),total.std(ddof=1)]})
    summary.to_csv(DATA / "recomputed_cost_summary.csv",index=False)
    corr_cols=["CVD_Coating_USD_kWh","Raw_Material_USD_kWh","Cell_Assembly_USD_kWh","BOS_USD_kWh","Overhead_USD_kWh","Cumulative_Production_GWh","Learning_Rate"]
    pd.DataFrame({"input":corr_cols,"spearman_rho":[mc[c].rank().corr(total.rank()) for c in corr_cols]}).to_csv(DATA / "recomputed_spearman_correlations.csv",index=False)
    with open(DATA / "Monte_Carlo_TEA_Summary.json",encoding="utf-8") as fh: archived=json.load(fh)
    assert archived["N_simulations"] == len(mc) == 10000
    if not args.no_figures:
        fig1(); fig2(tea); fig3(mc); fig4(); fig5(audit)
    print("Recomputed conditional screening summaries and wrote five PNG figures; Figure 5 is a provenance audit.")


if __name__ == "__main__":
    main()
