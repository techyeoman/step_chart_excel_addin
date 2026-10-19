"""Generate the add-in icons (a small CDF step chart) into assets/.

Run with: python3 scripts/make_icons.py   (requires Pillow)
"""
from pathlib import Path
from PIL import Image, ImageDraw

ASSETS = Path(__file__).resolve().parent.parent / "assets"
BLUE = (31, 111, 178, 255)
ORANGE = (230, 126, 34, 255)
STEPS = [(0.0, 0.0), (0.3, 0.0), (0.3, 0.4), (0.6, 0.4), (0.6, 0.7), (0.85, 0.7), (0.85, 1.0), (1.0, 1.0)]


def step_points(size, pad, inset):
    span = size - 2 * pad
    return [(pad + x * span, size - pad - inset - y * (span - 2 * inset)) for x, y in STEPS]


def draw_icon(size, background=True, color=BLUE, outline_only=False):
    scale = 4  # supersample for smooth lines
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    pad = round(s * 0.12)
    if background and not outline_only:
        d.rounded_rectangle([0, 0, s - 1, s - 1], radius=round(s * 0.18), fill=(255, 255, 255, 255),
                            outline=(200, 210, 220, 255), width=max(scale, round(s * 0.02)))
    width = max(scale * 1, round(s * 0.085))
    axis_w = max(scale, round(s * 0.04))
    axis_color = (255, 255, 255, 255) if outline_only else (90, 90, 90, 255)
    d.line([(pad, pad), (pad, s - pad), (s - pad, s - pad)], fill=axis_color, width=axis_w)
    pts = step_points(s, pad, inset=round(s * 0.06))
    d.line(pts, fill=(255, 255, 255, 255) if outline_only else color, width=width, joint="curve")
    return img.resize((size, size), Image.LANCZOS)


def main():
    ASSETS.mkdir(exist_ok=True)
    for size in (16, 32, 64, 80, 128):
        draw_icon(size).save(ASSETS / f"icon-{size}.png")
    # Unified manifest app-package icons: color 192x192, outline 32x32 white on transparent.
    draw_icon(192).save(ASSETS / "color.png")
    draw_icon(32, background=False, outline_only=True).save(ASSETS / "outline.png")
    print("Icons written to", ASSETS)


if __name__ == "__main__":
    main()
