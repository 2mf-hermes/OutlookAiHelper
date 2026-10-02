from PIL import Image, ImageDraw
import struct, os, io


def make_icon(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    pad = max(1, size // 16)
    d.rounded_rectangle(
        (pad, pad, size - pad, size - pad),
        radius=max(4, size // 5),
        fill=(10, 132, 255, 255),
    )
    highlight = Image.new("RGBA", img.size, (0, 0, 0, 0))
    hd = ImageDraw.Draw(highlight)
    hd.rounded_rectangle(
        (pad + 2, pad + 2, size - pad - 2, size // 2),
        radius=max(4, size // 5),
        fill=(255, 255, 255, 55),
    )
    img = Image.alpha_composite(img, highlight)
    d = ImageDraw.Draw(img)
    m = size * 0.22
    top = size * 0.30
    bottom = size * 0.70
    left, right = m, size - m
    d.rounded_rectangle(
        (left, top, right, bottom),
        radius=max(2, size // 20),
        fill=(255, 255, 255, 255),
    )
    mid_x = size // 2
    d.polygon(
        [(left + 2, top + 3), (mid_x, top + size * 0.18), (right - 2, top + 3)],
        fill=(10, 132, 255, 255),
    )
    dot_r = max(1, size // 18)
    colors = [(255, 69, 58), (255, 159, 10), (10, 132, 255), (142, 142, 147)]
    centers = [
        (size * 0.34, size * 0.55),
        (size * 0.52, size * 0.55),
        (size * 0.34, size * 0.66),
        (size * 0.52, size * 0.66),
    ]
    for (cx, cy), col in zip(centers, colors):
        d.ellipse([cx - dot_r, cy - dot_r, cx + dot_r, cy + dot_r], fill=col)
    return img


def png_to_ico(png_images, ico_path):
    count = len(png_images)
    offset = 6 + 16 * count
    entries = []
    blobs = []
    for im in png_images:
        w, h = im.size
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        data = buf.getvalue()
        blobs.append(data)
        ww = 0 if w >= 256 else w
        hh = 0 if h >= 256 else h
        entries.append((ww, hh, len(data), offset))
        offset += len(data)
    with open(ico_path, "wb") as f:
        f.write(struct.pack("<HHH", 0, 1, count))
        for ww, hh, size, off in entries:
            f.write(struct.pack("<BBBBHHII", ww, hh, 0, 0, 1, 32, size, off))
        for data in blobs:
            f.write(data)


root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
out_dir = os.path.join(root, "assets")
os.makedirs(out_dir, exist_ok=True)
sizes = [16, 24, 32, 48, 64, 128, 256]
imgs = [make_icon(s) for s in sizes]
ico_path = os.path.join(out_dir, "app.ico")
png_to_ico(imgs, ico_path)
imgs[-1].save(os.path.join(out_dir, "app-256.png"))
print("wrote", ico_path, "size", os.path.getsize(ico_path))
