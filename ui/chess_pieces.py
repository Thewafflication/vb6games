"""Chess pieces drawn as shaded shapes at any size (used by make_sprites.py)."""
from PIL import Image, ImageDraw, ImageFilter, ImageChops

K = 4  # supersampling


def _mask(size, drawer):
    s = size * K
    m = Image.new('L', (s, s), 0)
    d = ImageDraw.Draw(m)
    u = s / 100.0

    def P(pts):
        return [(x * u, y * u) for x, y in pts]

    def E(x0, y0, x1, y1):
        return [x0 * u, y0 * u, x1 * u, y1 * u]

    drawer(d, P, E, u)
    return m


def _plinth(d, P, E, u):
    d.rounded_rectangle(E(20, 84, 80, 94), radius=5 * u, fill=255)
    d.polygon(P([(26, 84), (74, 84), (68, 74), (32, 74)]), fill=255)


def _stem(d, P, E, u, top_y, top_w, bottom_w=22):
    d.polygon(P([(50 - bottom_w, 78), (50 + bottom_w, 78),
                 (50 + top_w, top_y), (50 - top_w, top_y)]), fill=255)


def pawn(d, P, E, u):
    _plinth(d, P, E, u)
    _stem(d, P, E, u, 52, 9, 18)
    d.ellipse(E(34, 46, 66, 56), fill=255)
    d.ellipse(E(37, 24, 63, 50), fill=255)


def rook(d, P, E, u):
    _plinth(d, P, E, u)
    d.polygon(P([(31, 78), (69, 78), (64, 36), (36, 36)]), fill=255)
    d.rectangle(E(28, 30, 72, 40), fill=255)
    for x in (28, 44, 60):
        d.rectangle(E(x, 16, x + 12, 32), fill=255)


def bishop(d, P, E, u):
    _plinth(d, P, E, u)
    _stem(d, P, E, u, 52, 9, 18)
    d.ellipse(E(34, 46, 66, 56), fill=255)
    d.ellipse(E(37, 16, 63, 52), fill=255)
    d.ellipse(E(44, 8, 56, 20), fill=255)


def queen(d, P, E, u):
    _plinth(d, P, E, u)
    d.polygon(P([(30, 78), (70, 78), (62, 40), (38, 40)]), fill=255)
    d.ellipse(E(30, 34, 70, 46), fill=255)
    d.polygon(P([(30, 38), (26, 16), (39, 28), (50, 12), (61, 28), (74, 16), (70, 38)]), fill=255)
    for cx in (26, 50, 74):
        d.ellipse(E(cx - 5, 8 if cx == 50 else 12, cx + 5, 18 if cx == 50 else 22), fill=255)


def king(d, P, E, u):
    _plinth(d, P, E, u)
    d.polygon(P([(30, 78), (70, 78), (62, 40), (38, 40)]), fill=255)
    d.ellipse(E(30, 34, 70, 46), fill=255)
    d.pieslice(E(31, 18, 69, 58), 180, 360, fill=255)
    d.rectangle(E(46, 4, 54, 28), fill=255)
    d.rectangle(E(40, 10, 60, 17), fill=255)


def knight(d, P, E, u):
    _plinth(d, P, E, u)
    d.polygon(P([(30, 78), (72, 78), (70, 58), (66, 44), (68, 30), (60, 18), (50, 12), (46, 6),
                 (40, 14), (28, 26), (20, 40), (22, 50), (32, 52), (40, 44), (44, 42),
                 (40, 54), (34, 66)]), fill=255)


PIECES = {1: pawn, 2: rook, 3: knight, 4: bishop, 5: queen, 6: king}


def piece_image(kind, white, size):
    """RGBA picture of a piece at size x size pixels."""
    s = size * K
    mask = _mask(size, PIECES[kind])
    # knight details are cut out afterwards
    # fills: a shaded gradient (lighter at the left and top)
    if white:
        light, dark, rim = (255, 252, 240), (190, 180, 160), (70, 60, 50)
    else:
        light, dark, rim = (110, 118, 140), (26, 28, 38), (232, 238, 250)
    grad = Image.new('RGB', (s, s))
    gp = grad.load()
    for y in range(s):
        for x in range(s):
            t = min(1.0, max(0.0, (x / s) * 0.75 + (y / s) * 0.35))
            gp[x, y] = tuple(int(light[i] + (dark[i] - light[i]) * t) for i in range(3))
    img = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    # outline: the mask grown a little
    grown = mask.filter(ImageFilter.MaxFilter(2 * (K + 1) // 2 * 2 + 1))
    img.paste(Image.new('RGBA', (s, s), rim + (255,)), (0, 0), grown)
    img.paste(grad, (0, 0), mask.filter(ImageFilter.MinFilter(3)))
    # a soft highlight on the upper left
    hl = Image.new('L', (s, s), 0)
    ImageDraw.Draw(hl).ellipse([s * 0.30, s * 0.18, s * 0.50, s * 0.70], fill=70)
    hl = ImageChops.multiply(hl.filter(ImageFilter.GaussianBlur(s * 0.03)), mask)
    img.paste(Image.new('RGB', (s, s), (255, 255, 255) if white else (190, 205, 235)), (0, 0), hl)
    if kind == 3:
        d = ImageDraw.Draw(img)
        u = s / 100.0
        d.ellipse([34 * u, 30 * u, 40 * u, 36 * u], fill=rim + (255,))
        d.line([(50 * u, 20 * u), (58 * u, 34 * u), (60 * u, 50 * u)], fill=rim + (255,), width=int(1.6 * u))
    # fill the square: crop to the piece and enlarge it a little
    box = img.getbbox()
    img = img.crop(box)
    scale = min(0.88 * s / img.width, 0.94 * s / img.height)
    img = img.resize((max(1, round(img.width * scale)), max(1, round(img.height * scale))), Image.LANCZOS)
    out = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    out.paste(img, ((s - img.width) // 2, round(0.97 * s) - img.height))
    return out.resize((size, size), Image.LANCZOS)
