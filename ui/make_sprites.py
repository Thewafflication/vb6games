#!/usr/bin/env python3
"""Generates the sprite sets of the games (needs Pillow).

Each game that has bitmap sprites carries them at four sizes
(sprites/64, 128, 192, 256: GIF files), and picks the set that suits the
screen at run time (ThSprite in ui/modTheme.bas). Run from the repository
root:   python ui/make_sprites.py

* restaum: a socket with a glossy ball and its shadow, drawn at any size;
* pipes:   the original pipe pictures, redrawn smooth in the new colors;
* captureking: the original chess pieces, smoothed.
"""
import io
import os
import re
import sys
from PIL import Image, ImageDraw, ImageFilter, ImageChops

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SIZES = (64, 128, 192, 256)

# ------------------------------------------------------------------ colors
CARD = (27, 32, 51)            # ThPanel
SOCKET = (14, 16, 25)
PAGE_Y = (255, 226, 74)        # Pipes: sunny yellow
CELL = (233, 248, 253)         # Pipes: board cell
CELL_LINE = (150, 210, 232)
PIPE_OUT = (14, 69, 105)
PIPE_MID = (111, 181, 214)
PIPE_BODY = (255, 255, 255)


def save_gif(im, path, transparent=None):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    rgb = im.convert('RGB')
    pal = rgb.quantize(colors=255 if transparent else 256, method=Image.MEDIANCUT, dither=Image.NONE)
    if transparent is not None:
        # the transparent color is palette entry 255
        data = pal.getpalette()[:255 * 3]
        data += [0, 0, 0] * (256 - len(data) // 3)
        pal.putpalette(data)
        mask = transparent
        px = pal.load()
        mp = mask.load()
        for y in range(pal.height):
            for x in range(pal.width):
                if mp[x, y] < 128:
                    px[x, y] = 255
        pal.save(path, transparency=255, optimize=True)
    else:
        pal.save(path, optimize=True)


# ------------------------------------------------------------------ restaum
def restaum_ball(size):
    k = 4
    s = size * k
    base = Image.new('RGB', (s, s), CARD)
    cx = cy = s / 2
    r_sock = s * 0.392
    # socket
    sock = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(sock)
    d.ellipse([cx - r_sock, cy - r_sock, cx + r_sock, cy + r_sock], fill=SOCKET + (255,), outline=(53, 60, 92, 255), width=max(2, s // 128))
    base.paste(sock, (0, 0), sock)
    # soft shadow of the ball inside the socket
    sh = Image.new('L', (s, s), 0)
    ImageDraw.Draw(sh).ellipse([cx - s * 0.33 + s * 0.025, cy - s * 0.33 + s * 0.06, cx + s * 0.33 + s * 0.025, cy + s * 0.33 + s * 0.06], fill=200)
    sh = sh.filter(ImageFilter.GaussianBlur(s * 0.03))
    base.paste(Image.new('RGB', (s, s), (0, 0, 0)), (0, 0), sh)
    # the ball: a radial gradient
    r = s * 0.335
    ball = Image.new('RGB', (s, s), (0, 0, 0))
    px = ball.load()
    hx, hy = cx - r * 0.38, cy - r * 0.42
    for y in range(int(cy - r) - 1, int(cy + r) + 2):
        for x in range(int(cx - r) - 1, int(cx + r) + 2):
            dx, dy = x - cx, y - cy
            if dx * dx + dy * dy > r * r:
                continue
            t = (((x - hx) ** 2 + (y - hy) ** 2) ** 0.5) / (r * 1.5)
            t = min(1.0, t)
            c0, c1, c2 = (255, 246, 170), (247, 190, 36), (150, 84, 8)
            if t < 0.55:
                u = t / 0.55
                c = tuple(c0[i] + (c1[i] - c0[i]) * u for i in range(3))
            else:
                u = (t - 0.55) / 0.45
                c = tuple(c1[i] + (c2[i] - c1[i]) * u for i in range(3))
            px[x, y] = tuple(int(v) for v in c)
    m = Image.new('L', (s, s), 0)
    ImageDraw.Draw(m).ellipse([cx - r, cy - r, cx + r, cy + r], fill=255)
    base.paste(ball, (0, 0), m)
    # specular highlight
    hl = Image.new('L', (s, s), 0)
    ImageDraw.Draw(hl).ellipse([cx - r * 0.62, cy - r * 0.68, cx - r * 0.12, cy - r * 0.28], fill=210)
    hl = hl.filter(ImageFilter.GaussianBlur(s * 0.012))
    base.paste(Image.new('RGB', (s, s), (255, 255, 235)), (0, 0), hl)
    return base.resize((size, size), Image.LANCZOS)


def make_restaum():
    for size in SIZES:
        save_gif(restaum_ball(size), os.path.join(ROOT, 'restaum', 'source', 'sprites', str(size), 'ball.gif'))


# -------------------------------------------------------------------- pipes
def recolor(im):
    """The original pipe picture (grayscale pipes with a black outline) in
    the new colors, smooth, on a transparent background."""
    im = im.convert('RGBA')
    gray = im.convert('L')
    a = im.getchannel('A')
    lut_r, lut_g, lut_b = [], [], []
    for v in range(256):
        if v < 128:
            u = v / 128
            c = tuple(PIPE_OUT[i] + (PIPE_MID[i] - PIPE_OUT[i]) * u for i in range(3))
        else:
            u = (v - 128) / 127
            c = tuple(PIPE_MID[i] + (PIPE_BODY[i] - PIPE_MID[i]) * u for i in range(3))
        lut_r.append(int(c[0])); lut_g.append(int(c[1])); lut_b.append(int(c[2]))
    r = gray.point(lut_r); g = gray.point(lut_g); b = gray.point(lut_b)
    return Image.merge('RGBA', (r, g, b, a))


def smooth_up(img, size):
    """Smooth enlargement of pixel art: scale up large, blur a little, then
    sharpen the alpha edge, then reduce."""
    k = 4
    big = img.resize((size * k, size * k), Image.BICUBIC)
    a = big.getchannel('A').filter(ImageFilter.GaussianBlur(size * k / 64))
    a = a.point(lambda v: 255 if v > 127 else 0)
    big.putalpha(a)
    return big.resize((size, size), Image.LANCZOS)


def tile_on(color, size, line=None):
    t = Image.new('RGBA', (size, size), color + (255,))
    if line:
        d = ImageDraw.Draw(t)
        d.rectangle([0, 0, size - 1, size - 1], outline=line + (255,), width=max(1, size // 64))
    return t


def pipes_pictures():
    frm = open(os.path.join(ROOT, 'pipes', 'FORM1.FRM'), 'rb').read().decode('latin-1').replace('\r\n', '\n')
    frx = open(os.path.join(ROOT, 'pipes', 'FORM1.FRX'), 'rb').read()
    out = {}
    for m in re.finditer(r"   Begin VB\.Image (\w+) \n(.*?)\n   End", frm, flags=re.S):
        name, b = m.group(1), m.group(2)
        pic = re.search(r'Picture\s+=\s+"Form1.frx":([0-9A-F]+)', b)
        if not pic:
            continue
        idx = re.search(r'Index\s+=\s+(\d+)', b)
        off = int(pic.group(1), 16)
        total = int.from_bytes(frx[off:off + 4], 'little')
        data = frx[off + 12:off + 4 + total]
        key = '%s%s' % (name, idx.group(1) if idx else '')
        out[key] = Image.open(io.BytesIO(data)).convert('RGBA')
    return out


def make_pipes():
    pics = pipes_pictures()
    for size in SIZES:
        base = os.path.join(ROOT, 'pipes', 'sprites', str(size))
        # the empty cell is fully transparent (the water shows through)
        blank = Image.new('RGB', (size, size), PIPE_OUT)
        save_gif(blank, os.path.join(base, 'blank.gif'), transparent=Image.new('L', (size, size), 0))
        for i in range(28):
            key = 'piece%d' % i
            if key not in pics:
                continue
            p = smooth_up(recolor(pics[key]), size)
            save_gif(p, os.path.join(base, 'piece%d.gif' % i), transparent=p.getchannel('A'))
        # the pipe stub and the arrow beside the board
        for key, name in (('out1', 'stub'), ('out0', 'arrow')):
            if key in pics:
                if name == 'arrow':
                    arrow = pics[key].convert('RGBA')
                    ar = Image.new('RGBA', arrow.size, PIPE_OUT + (255,))
                    ar.putalpha(arrow.getchannel('A'))
                    p = smooth_up(ar, size)
                else:
                    p = smooth_up(recolor(pics[key]), size)
                save_gif(p, os.path.join(base, name + '.gif'), transparent=p.getchannel('A'))


# -------------------------------------------------------------- captureking
def make_captureking():
    src = os.path.join(ROOT, 'captureking', 'Img')
    for size in SIZES:
        base = os.path.join(ROOT, 'captureking', 'sprites', str(size))
        for n in [c + str(i) for c in 'bw' for i in range(1, 7)]:
            im = Image.open(os.path.join(src, n + '.ico')).convert('RGBA')
            k = 4
            big = im.resize((size * k, size * k), Image.BICUBIC)
            a = big.getchannel('A').filter(ImageFilter.GaussianBlur(size * k / 80)).point(lambda v: 255 if v > 110 else 0)
            rgb = big.convert('RGB').filter(ImageFilter.UnsharpMask(radius=size * k / 40, percent=60, threshold=2))
            out = rgb.resize((size, size), Image.LANCZOS)
            mask = a.resize((size, size), Image.LANCZOS)
            save_gif(out, os.path.join(base, n + '.gif'), transparent=mask)


if __name__ == '__main__':
    which = sys.argv[1:] or ['restaum', 'pipes', 'captureking']
    for w in which:
        globals()['make_' + w]()
        print('made', w)
