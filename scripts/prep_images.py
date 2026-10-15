"""Prepare slide images for the green roof talk.

Reads originals in place (OneDrive NovaTech deck and draft manuscript, plus sources/), writes
working images to BUILD (outside the project, so the project folder stays clean).
Run from the project folder:  python scripts/prep_images.py
"""
import io
import json
import os
import zipfile
from pathlib import Path

from PIL import Image, ImageOps

ROOT = Path(__file__).resolve().parents[1]
ONEDRIVE = Path(r"C:\Users\yonatanz\OneDrive - Technion\מחקרים\תואר שלישי")
NOVATECH = ONEDRIVE / "כנסים" / "NovaTech 2026" / ("Remote Monitoring of Water-Efficient Blue–Green Roofs Using "
                                                  "Satellite-Derived Evapotranspiration in a Mediterranean Climate.pptx")
DRAFT = ONEDRIVE / "my articles" / "Remote sensing and green roof" / "Can Green Roof Be Monitored from Space .docx"
USER = ROOT / "sources" / "user_images"
WEB = ROOT / "sources" / "web_photos"
BUILD = Path(os.environ.get("TALK_BUILD", Path(os.environ["TEMP"]) / "green_roof_talk_build"))
IMG = BUILD / "img"
IMG.mkdir(parents=True, exist_ok=True)
MAX_SIDE = 1800


def from_zip(zpath, member):
    with zipfile.ZipFile(zpath) as z:
        return Image.open(io.BytesIO(z.read(member)))


def save_photo(im, name, box=None, max_side=MAX_SIDE):
    im = ImageOps.exif_transpose(im).convert("RGB")
    if box:
        w, h = im.size
        im = im.crop((int(box[0] * w), int(box[1] * h), int(box[2] * w), int(box[3] * h)))
    im.thumbnail((max_side, max_side))
    im.save(IMG / name, quality=88)


def nt(name):
    return from_zip(NOVATECH, f"ppt/media/{name}")


V2 = ROOT / "output" / "green_roof_talk_v2.pptx"


def dr(name, fallback=None):
    """Image from the draft .docx. If the draft is locked (open in Word), use the copy embedded in the v2 deck
    ("v2:imageN.ext") or the NovaTech equivalent ("imageN.ext")."""
    try:
        return from_zip(DRAFT, f"word/media/{name}")
    except PermissionError:
        if fallback is None:
            raise
        print(f"draft locked: using {fallback} instead of draft {name}")
        if fallback.startswith("v2:"):
            return from_zip(V2, "ppt/media/" + fallback[3:])
        return nt(fallback)


def stack(im, name, bands, x_range=(0.0, 1.0)):
    im = im.convert("RGB")
    w, h = im.size
    x0, x1 = int(x_range[0] * w), int(x_range[1] * w)
    parts = [im.crop((x0, int(t * h), x1, int(b * h))) for t, b in bands]
    out = Image.new("RGB", (x1 - x0, sum(p.height for p in parts)), "white")
    y = 0
    for p in parts:
        out.paste(p, (0, y))
        y += p.height
    out.save(IMG / name)


# Technion logo without "leave a mark" (separator at x~3140 of 3907)
logo = nt("image44.png").convert("RGBA").crop((0, 0, 3030, 946))
logo.crop(logo.getbbox()).save(IMG / "technion_logo.png")

# Own photos
save_photo(nt("image17.JPG"), "drone_clean.jpg")
save_photo(nt("image4.JPG"), "title_roof.jpg")            # v1 title photo
save_photo(nt("image22.jpeg"), "construction.jpg")        # substrate being laid
save_photo(nt("image47.png"), "roof_lush.jpg")
save_photo(nt("image15.png"), "aerial_city.jpg")
save_photo(nt("image26.jpeg"), "storage_module.jpg")
save_photo(nt("image28.jpeg"), "sentinel2.jpg")
save_photo(nt("image45.jpeg"), "team.jpg")
save_photo(dr("image2.png", "v2:image8.png"), "drone_plots.png")
save_photo(dr("image6.jpeg", "v2:image12.jpg"), "station.jpg")
save_photo(dr("image7.jpeg", "v2:image10.jpg"), "valve.jpg")
save_photo(dr("image5.jpeg", "image27.jpeg"), "aptenia.jpg")
save_photo(dr("image3.png", "v2:image13.png"), "pixel_grid.png")
for src, name in [("image41.png", "logo_water4all.png"), ("image42.png", "logo_rainwater.png"),
                  ("image49.jpeg", "logo_smartflow.png"), ("image50.jpeg", "logo_gwri.png")]:
    nt(src).convert("RGBA").save(IMG / name)

# Images you sent (kept at full quality)
for f in USER.glob("*.png"):
    Image.open(f).save(IMG / f.name)
# Left half only: the green-blue roof with arrows down
gb = Image.open(USER / "green_vs_bluegreen.png").convert("RGBA")
w, h = gb.size
gb = gb.crop((0, 0, w // 2, h)).convert("RGB")
ink = Image.eval(gb.convert("L"), lambda v: 255 if v < 235 else 0)   # anything not near-white
x0, y0, x1, y1 = ink.getbbox()
gb.crop((max(0, x0 - 6), max(0, y0 - 6), x1 + 6, y1 + 6)).save(IMG / "green_blue_roof.png")
# Satellite figure without its suptitle and footnote, so the plots can be shown larger
sat = Image.open(USER / "sentinel_segmented.png").convert("RGB")
w, h = sat.size
sat.crop((0, int(0.035 * h), w, int(0.975 * h))).save(IMG / "sentinel_segmented_trim.png")

# Open-licence web photos (credits in sources/web_photos/credits.json)
for f in WEB.glob("*.jpg"):
    save_photo(Image.open(f), f.name)
(IMG / "credits.json").write_text((WEB / "credits.json").read_text(encoding="utf-8"), encoding="utf-8")

# Field-vs-satellite figure from the draft: legend + GNDVI row
try:
    fig_jj = from_zip(DRAFT, "word/media/image15.png")
except PermissionError:
    fig_jj = None
    print("draft locked: field-vs-satellite figure taken from the v2 deck")
    from_zip(V2, "ppt/media/image16.png").save(IMG / "fig_field_vs_sat_gndvi.png")
if fig_jj is not None:
    stack(fig_jj, "fig_field_vs_sat_gndvi.png", [(0.083, 0.106), (0.530, 0.968)], x_range=(0.165, 0.86))
print("images written to", IMG)
