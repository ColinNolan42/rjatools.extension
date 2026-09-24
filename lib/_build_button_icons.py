# -*- coding: ascii -*-
"""Regenerate the Size Gas and Size Water button icons.

Build-time tool, never imported by pyRevit. Run with CPython (Pillow needed):

    python _build_button_icons.py

Both buttons share one design: a discipline glyph over the same sizing-arrow
bar, on a flat Material field. Size Gas is Deep Orange 800 with a flame, Size
Water is Light Blue 800 with a droplet, which matches the existing One-Line
(Blue Grey 800) and Duct Velocity (Teal 900) buttons.

Drawn at 8x and downsampled with LANCZOS so the curves stay clean at 32 px.
The arrowheads are deliberately chunky: a first attempt with fine arrows
dissolved into the bar at button size.
"""

import os

from PIL import Image, ImageDraw

S = 8                       # supersample factor
N = 32 * S
ORANGE = (230, 81, 0, 255)  # Material Deep Orange 800, gas
BLUE = (2, 119, 189, 255)   # Material Light Blue 800, water
WHITE = (255, 255, 255, 255)

HERE = os.path.dirname(os.path.abspath(__file__))
PANEL = os.path.normpath(
    os.path.join(HERE, "..", "Developer Tools.tab", "Pipe Tools.panel"))


def sizing_bar(d):
    """Shared bottom element: a pipe bar with outward size arrows."""
    bar_h = N * 0.10
    bar_y = N * 0.795
    d.rectangle([N * 0.26, bar_y - bar_h / 2.0, N * 0.74, bar_y + bar_h / 2.0],
                fill=WHITE)
    a = N * 0.145
    for direction in (-1, 1):
        tip_x = N * 0.5 + direction * N * 0.47
        base_x = tip_x - direction * a
        d.polygon([(tip_x, bar_y),
                   (base_x, bar_y - a * 0.80),
                   (base_x, bar_y + a * 0.80)], fill=WHITE)


def droplet(d, cx, cy, w, h):
    """Water: circular bottom tapering to a point at the top."""
    r = w / 2.0
    bottom_cy = cy + h / 2.0 - r
    d.ellipse([cx - r, bottom_cy - r, cx + r, bottom_cy + r], fill=WHITE)
    d.polygon([(cx, cy - h / 2.0),
               (cx - r, bottom_cy + r * 0.10),
               (cx + r, bottom_cy + r * 0.10)], fill=WHITE)


# Asymmetric flame outline, normalised to a 1x1 box centred on (0,0).
# The asymmetry is the point: a symmetric teardrop reads as the water droplet,
# which is the icon this one has to be told apart from at 32 px.
_FLAME = [
    (0.16, -0.50),
    (0.30, -0.28), (0.40, -0.10), (0.46, 0.08),
    (0.46, 0.26), (0.34, 0.42), (0.14, 0.50),
    (-0.10, 0.50), (-0.30, 0.41), (-0.41, 0.24),
    (-0.44, 0.04), (-0.34, -0.14),
    (-0.14, -0.26), (-0.04, -0.14),
    (0.02, 0.02), (0.10, -0.22),
]

_FLAME_INNER = [
    (0.04, 0.04), (0.20, 0.20), (0.15, 0.36),
    (-0.02, 0.41), (-0.18, 0.33), (-0.18, 0.16),
]


def flame(d, cx, cy, w, h, field):
    """Gas: asymmetric flame with an inner tongue knocked out in the field
    colour, which is what keeps it reading as fire once scaled down."""
    d.polygon([(cx + px * w * 1.06, cy + py * h) for px, py in _FLAME],
              fill=WHITE)
    d.polygon([(cx + px * w * 1.06, cy + py * h) for px, py in _FLAME_INNER],
              fill=field)


def table_rows(d):
    """Bottom element for the report button: stacked table rows.

    Sits where the sizing bar sits on the two sizing buttons, so Water Report
    reads as the same family of tool while saying it produces paperwork rather
    than changing the model.
    """
    row_h = N * 0.055
    gap = N * 0.038
    y = N * 0.700
    for index in range(3):
        right = N * 0.80 if index else N * 0.70   # a header row, then two data rows
        d.rectangle([N * 0.20, y, right, y + row_h], fill=WHITE)
        y += row_h + gap


# Glyph geometry for the two SIZING buttons. Do not change these: the flame
# and droplet produced by them are the icons Colin reviewed and accepted, and
# regenerating with different numbers silently replaces an approved icon.
SIZING_GLYPH = (0.36, 0.40, 0.52)      # cy, w, h as fractions of the canvas
# The report button gives up a little glyph height to the table rows below it.
REPORT_GLYPH = (0.32, 0.38, 0.46)


def build(field, glyph, out_path, bottom=None, geometry=SIZING_GLYPH):
    cy, w, h = geometry
    img = Image.new("RGBA", (N, N), field)
    d = ImageDraw.Draw(img)
    glyph(d, N * 0.50, N * cy, N * w, N * h)
    (bottom or sizing_bar)(d)
    img.resize((32, 32), Image.LANCZOS).save(out_path)
    print("wrote " + out_path)


if __name__ == "__main__":
    build(BLUE, lambda d, cx, cy, w, h: droplet(d, cx, cy, w, h),
          os.path.join(PANEL, "Size Water.pushbutton", "icon.png"))
    build(ORANGE, lambda d, cx, cy, w, h: flame(d, cx, cy, w, h, ORANGE),
          os.path.join(PANEL, "Size Gas.pushbutton", "icon.png"))
    build(BLUE, lambda d, cx, cy, w, h: droplet(d, cx, cy, w, h),
          os.path.join(PANEL, "Water Report.pushbutton", "icon.png"),
          bottom=table_rows, geometry=REPORT_GLYPH)
