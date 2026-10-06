"""Generate a built-in fake screenshot used by the live preview pane."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets" / "sample_screenshot.png"


def main() -> None:
    # Roughly a 16:10 desktop screenshot at 1600x1000 for preview demos.
    w, h = 1600, 1000
    img = Image.new("RGB", (w, h), "#1f2933")
    draw = ImageDraw.Draw(img)

    # Title bar
    draw.rectangle((0, 0, w, 48), fill="#111827")
    draw.ellipse((18, 16, 34, 32), fill="#ef4444")
    draw.ellipse((44, 16, 60, 32), fill="#f59e0b")
    draw.ellipse((70, 16, 86, 32), fill="#22c55e")
    draw.text((110, 14), "Sample App — Screenshot Preview", fill="#e5e7eb")

    # Sidebar
    draw.rectangle((0, 48, 280, h), fill="#111827")
    for i, label in enumerate(["Inbox", "Docs", "Images", "Settings"]):
        y = 80 + i * 56
        fill = "#2563eb" if i == 1 else "#1f2937"
        draw.rounded_rectangle((16, y, 264, y + 40), radius=8, fill=fill)
        draw.text((36, y + 12), label, fill="#f9fafb")

    # Content cards
    draw.rectangle((280, 48, w, h), fill="#f3f4f6")
    draw.text((312, 72), "Project Dashboard", fill="#111827")
    draw.text((312, 108), "This built-in sample shows how paste-fit styles affect size.", fill="#4b5563")

    cards = [
        ("#dbeafe", "Width limited"),
        ("#dcfce7", "Height limited"),
        ("#fef3c7", "Table cell fit"),
        ("#fce7f3", "No upscale"),
    ]
    for idx, (color, title) in enumerate(cards):
        x = 312 + (idx % 2) * 620
        y = 160 + (idx // 2) * 340
        draw.rounded_rectangle((x, y, x + 580, y + 300), radius=16, fill=color)
        draw.text((x + 24, y + 24), title, fill="#111827")
        draw.rectangle((x + 24, y + 80, x + 556, y + 250), fill="#ffffff")
        for line in range(6):
            yy = y + 100 + line * 22
            draw.rectangle((x + 48, yy, x + 520, yy + 10), fill="#d1d5db")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    img.save(OUT, format="PNG")
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
