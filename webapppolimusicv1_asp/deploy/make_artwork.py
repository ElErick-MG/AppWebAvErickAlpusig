"""Generate placeholder artwork for PoliMusic (album covers, artist photos, avatars).

Uses only the project palette: charcoal backgrounds + terracotta / jade accents.
Run:  python deploy/make_artwork.py
"""
import math
import os
import random
from PIL import Image, ImageDraw, ImageFilter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

TERRACOTTA = [(176, 58, 46), (192, 57, 43), (146, 43, 33)]
JADE = [(0, 168, 107), (20, 143, 119), (26, 188, 156)]
DARK = [(13, 13, 13), (22, 22, 22), (31, 31, 31)]
LIGHT = (242, 244, 244)


def lerp(a, b, t):
    return tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))


def gradient(size, c1, c2, angle=135):
    w, h = size
    img = Image.new("RGB", size, c1)
    px = img.load()
    rad = math.radians(angle)
    dx, dy = math.cos(rad), math.sin(rad)
    proj = [x * dx + y * dy for x, y in ((0, 0), (w, 0), (0, h), (w, h))]
    lo, hi = min(proj), max(proj)
    for y in range(h):
        for x in range(w):
            t = ((x * dx + y * dy) - lo) / (hi - lo)
            px[x, y] = lerp(c1, c2, max(0.0, min(1.0, t)))
    return img


def radial_glow(img, center, radius, color, strength=0.7):
    glow = Image.new("RGB", img.size, (0, 0, 0))
    d = ImageDraw.Draw(glow)
    d.ellipse([center[0] - radius, center[1] - radius,
               center[0] + radius, center[1] + radius], fill=color)
    glow = glow.filter(ImageFilter.GaussianBlur(radius * 0.5))
    return Image.blend(img, Image.blend(img, glow, 0.9), strength)


def noise(img, amount=4):
    w, h = img.size
    px = img.load()
    rnd = random.Random(3)
    for y in range(h):
        for x in range(w):
            n = rnd.randint(-amount, amount)
            r, g, b = px[x, y]
            px[x, y] = (max(0, min(255, r + n)), max(0, min(255, g + n)),
                        max(0, min(255, b + n)))
    return img


def vignette(img, strength=0.6):
    w, h = img.size
    mask = Image.new("L", (w, h), 0)
    d = ImageDraw.Draw(mask)
    d.ellipse([-w * 0.3, -h * 0.3, w * 1.3, h * 1.3], fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(w * 0.18))
    dark = Image.new("RGB", (w, h), (4, 4, 4))
    return Image.composite(img, Image.blend(img, dark, strength), mask)


def album_cover(idx, accent, second, seed):
    """Bold geometric cover: soft glow + concentric rings + jade rule."""
    rnd = random.Random(seed)
    size = (700, 700)
    img = gradient(size, DARK[0], lerp(DARK[0], accent, 0.30), angle=115 + idx * 22)
    cx, cy = 350, 330
    img = radial_glow(img, (cx - 40, cy - 60), 330, accent, 0.75)

    layer = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)

    # soft disc
    d.ellipse([cx - 210, cy - 210, cx + 210, cy + 210], fill=accent + (58,))

    # concentric rings
    for i in range(6):
        r = 120 + i * 62
        d.ellipse([cx - r, cy - r, cx + r, cy + r],
                  outline=LIGHT + (max(10, 46 - i * 6),), width=2)

    # jade accent rule
    d.rounded_rectangle([cx - 250, cy + 268, cx - 250 + 170, cy + 274], radius=3,
                        fill=second + (215,))
    d.ellipse([cx + 196, cy + 258, cx + 216, cy + 278], fill=LIGHT + (200,))

    # small counter weight
    if idx % 2 == 0:
        d.ellipse([cx - 268, cy - 268, cx - 212, cy - 212], outline=second + (170,), width=2)

    img = Image.alpha_composite(img.convert("RGBA"), layer).convert("RGB")
    return noise(vignette(img, 0.62), 4)


def artist_photo(idx, accent, second):
    """Stylised portrait silhouette with rim light."""
    size = (560, 560)
    img = gradient(size, DARK[0], lerp(DARK[0], accent, 0.26), angle=80 + idx * 26)
    img = radial_glow(img, (280, 200), 250, accent, 0.55)

    layer = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)

    # arc behind the head
    d.arc([90, 60, 470, 440], start=200, end=340, fill=second + (120,), width=3)

    # shoulders + head
    d.ellipse([120, 350, 440, 700], fill=(38, 38, 38, 255))
    d.ellipse([196, 132, 364, 330], fill=(46, 46, 46, 255))

    # rim light
    d.arc([196, 132, 364, 330], start=250, end=360, fill=accent + (150,), width=4)
    d.arc([120, 350, 440, 700], start=250, end=340, fill=accent + (110,), width=4)

    img = Image.alpha_composite(img.convert("RGBA"), layer).convert("RGB")
    return noise(vignette(img, 0.5), 3)


def avatar(accent, second):
    size = (360, 360)
    img = gradient(size, DARK[1], lerp(DARK[1], accent, 0.40), angle=140)
    img = radial_glow(img, (130, 110), 170, accent, 0.6)
    layer = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    d.ellipse([126, 78, 234, 186], fill=LIGHT + (96,))
    d.ellipse([80, 208, 280, 420], fill=LIGHT + (96,))
    d.arc([60, 60, 300, 300], start=150, end=330, fill=second + (90,), width=3)
    img = Image.alpha_composite(img.convert("RGBA"), layer).convert("RGB")
    return vignette(img, 0.3)


def save(img, folder, filename, quality=92):
    path = os.path.join(ROOT, folder)
    os.makedirs(path, exist_ok=True)
    out = os.path.join(path, filename)
    if filename.lower().endswith(".png"):
        img.save(out, "PNG", optimize=True)
    else:
        img.save(out, "JPEG", quality=quality, optimize=True)
    print("wrote", os.path.relpath(out, ROOT))


if __name__ == "__main__":
    combos = [
        (TERRACOTTA[1], JADE[0]),
        (JADE[1], TERRACOTTA[1]),
        (TERRACOTTA[2], JADE[2]),
        (JADE[0], TERRACOTTA[0]),
        (TERRACOTTA[0], JADE[1]),
    ]
    for i in range(5):
        accent, second = combos[i]
        save(album_cover(i, accent, second, 21 + i), "albumImages", f"album{i + 1}.png")

    artist_files = ["artist1.jpg", "artist2.jpg", "artist3.jpg", "artist4.jpg", "artist5.png"]
    for i, fn in enumerate(artist_files):
        accent, second = combos[i]
        save(artist_photo(i, accent, second), "ArtistImages", fn)

    save(avatar(TERRACOTTA[1], JADE[0]), "userImages", "username1.png")
