"""
make_icon.py  —  one-time helper
=================================
Converts logo.png → icon.ico for use as the AHA Bot window icon.

Usage:
    1. Save the CPR Lifeline logo as  logo.png  in this folder.
    2. Run:  python make_icon.py
    3. icon.ico will appear in this folder — that's all you need.
"""

from pathlib import Path
from PIL import Image

SRC  = Path(__file__).parent / "logo.png"
DEST = Path(__file__).parent / "icon.ico"

if not SRC.exists():
    raise FileNotFoundError(
        f"Could not find '{SRC}'.\n"
        "Save the CPR Lifeline logo as 'logo.png' in the same folder as this script."
    )

img = Image.open(SRC).convert("RGBA")

# Windows ICO needs square dimensions — pad if necessary
w, h = img.size
side  = max(w, h)
if w != h:
    padded = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    padded.paste(img, ((side - w) // 2, (side - h) // 2))
    img = padded

img.save(DEST, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (256, 256)])
print(f"Saved: {DEST}")
