"""Draws the icon of each game and writes it as <project name>.ico beside the
game's .vbp (wfc build stamps that file into the program when the project's
IconForm has no icon of its own).

    python ui/make_icons.py

Needs Pillow. The drawings use the colors of ui/modTheme.bas.
"""
import os
import struct
import sys

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = 512
SIZES = (16, 24, 32, 48, 64, 128, 256)

BACK = (17, 20, 31)       # window background
PANEL = (27, 32, 51)      # raised surface
TEAL = (61, 214, 198)
AMBER = (255, 180, 84)
YELLOW = (255, 211, 77)
CORAL = (255, 92, 122)
BLUE = (45, 62, 143)
RED = (232, 76, 80)
GREEN = (47, 163, 108)
VIOLET = (145, 77, 199)
SKY = (64, 104, 224)
LIGHT = (232, 235, 245)


def canvas():
    image = Image.new("RGBA", (BASE, BASE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((16, 16, BASE - 16, BASE - 16), 88, fill=BACK)
    return image, draw


def ball(draw, cx, cy, r, color):
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=color)
    shine = tuple(min(255, c + 90) for c in color)
    draw.ellipse((cx - r * 0.55, cy - r * 0.6, cx - r * 0.2, cy - r * 0.25),
                 fill=shine)


def connect4():
    image, draw = canvas()
    draw.rounded_rectangle((56, 72, 456, 440), 40, fill=BLUE)
    colors = {(0, 3): YELLOW, (1, 3): CORAL, (1, 2): YELLOW, (2, 3): CORAL,
              (2, 2): YELLOW, (3, 3): YELLOW, (2, 1): YELLOW}
    for row in range(4):
        for col in range(4):
            cx = 120 + col * 90
            cy = 136 + row * 90
            color = colors.get((col, row))
            if color is None:
                draw.ellipse((cx - 36, cy - 36, cx + 36, cy + 36), fill=BACK)
            else:
                ball(draw, cx, cy, 38, color)
    return image


def clix():
    image, draw = canvas()
    palette = [RED, SKY, GREEN, YELLOW, VIOLET]
    for row in range(3):
        for col in range(3):
            ball(draw, 120 + col * 136, 120 + row * 136, 58,
                 palette[(row * 2 + col * 3) % 5])
    return image


def crown(draw, fill):
    draw.polygon([(96, 360), (96, 190), (190, 270), (256, 150), (322, 270),
                  (416, 190), (416, 360)], fill=fill)
    draw.rounded_rectangle((96, 372, 416, 420), 16, fill=fill)
    for x, y in ((96, 190), (256, 150), (416, 190)):
        draw.ellipse((x - 26, y - 26, x + 26, y + 26), fill=fill)


def captureking():
    image, draw = canvas()
    for row in range(8):
        for col in range(8):
            if (row + col) % 2 == 0:
                draw.rectangle((32 + col * 56, 32 + row * 56,
                                32 + col * 56 + 56, 32 + row * 56 + 56),
                               fill=PANEL)
    crown(draw, AMBER)
    for x in (176, 256, 336):
        draw.ellipse((x - 14, 330 - 14, x + 14, 330 + 14), fill=CORAL)
    return image


def executioner():
    image, draw = canvas()
    brown = (176, 122, 72)
    draw.rectangle((96, 420, 336, 450), fill=brown)
    draw.rectangle((140, 90, 170, 430), fill=brown)
    draw.rectangle((140, 90, 360, 120), fill=brown)
    draw.line((140, 180, 220, 120), fill=brown, width=22)
    draw.line((330, 120, 330, 190), fill=LIGHT, width=10)
    draw.ellipse((292, 190, 368, 266), outline=LIGHT, width=12)
    draw.line((330, 266, 330, 350), fill=LIGHT, width=12)
    draw.line((330, 290, 286, 330), fill=LIGHT, width=12)
    draw.line((330, 290, 374, 330), fill=LIGHT, width=12)
    draw.line((330, 350, 296, 410), fill=LIGHT, width=12)
    draw.line((330, 350, 364, 410), fill=LIGHT, width=12)
    return image


def pipes():
    image, draw = canvas()
    width = 70
    draw.line((60, 150, 256, 150), fill=TEAL, width=width)
    draw.line((256, 150, 256, 360), fill=TEAL, width=width)
    draw.line((256, 360, 452, 360), fill=TEAL, width=width)
    for x, y in ((256, 150), (256, 360)):
        draw.ellipse((x - 35, y - 35, x + 35, y + 35), fill=TEAL)
    inner = 26
    draw.line((60, 150, 256, 150), fill=AMBER, width=inner)
    draw.line((256, 150, 256, 360), fill=AMBER, width=inner)
    draw.line((256, 360, 452, 360), fill=AMBER, width=inner)
    for x, y in ((256, 150), (256, 360)):
        draw.ellipse((x - 13, y - 13, x + 13, y + 13), fill=AMBER)
    return image


def restaum():
    image, draw = canvas()
    draw.ellipse((40, 40, 472, 472), fill=(94, 62, 36))
    for gx in (-1.5, -0.5, 0.5, 1.5):
        for gy in (-1.5, -0.5, 0.5, 1.5):
            cx = 256 + gx * 92
            cy = 256 + gy * 92
            if (gx, gy) in ((0.5, 0.5),):
                draw.ellipse((cx - 30, cy - 30, cx + 30, cy + 30),
                             fill=(54, 34, 18))
            else:
                ball(draw, cx, cy, 34, AMBER)
    return image


def sumboku():
    image, draw = canvas()
    draw.rounded_rectangle((60, 60, 452, 452), 28, fill=PANEL)
    for i in range(1, 3):
        pos = 60 + i * 392 / 3
        draw.line((pos, 60, pos, 452), fill=TEAL, width=8)
        draw.line((60, pos, 452, pos), fill=TEAL, width=8)
    try:
        font = ImageFont.truetype("arialbd.ttf", 100)
    except OSError:
        font = ImageFont.load_default()
    cells = [(0, 0, "5"), (1, 0, "3"), (2, 1, "7"), (0, 2, "9"), (1, 1, "1"),
             (2, 2, "4")]
    for col, row, text in cells:
        cx = 60 + (col + 0.5) * 392 / 3
        cy = 60 + (row + 0.5) * 392 / 3
        draw.text((cx, cy), text, font=font, fill=AMBER if text in "579"
                  else LIGHT, anchor="mm")
    return image


GAMES = {
    "connect4/PConnect4": connect4,
    "clix/src/Clix": clix,
    "captureking/CaptureKing": captureking,
    "executioner/Executioner": executioner,
    "pipes/pipes": pipes,
    "restaum/source/BrainVita": restaum,
    "sumboku/v1.4.2/SuMBOku_V_1_4/Sudoku": sumboku,
    "sumboku/PSuMBOku": sumboku,
}


def frame(image, size):
    """One 32-bit BMP icon frame (DIB, bottom-up, with an empty AND mask)."""
    small = image.resize((size, size), Image.LANCZOS).convert("RGBA")
    pixels = small.load()
    rows = bytearray()
    for y in range(size - 1, -1, -1):
        for x in range(size):
            r, g, b, a = pixels[x, y]
            rows += bytes((b, g, r, a))
    mask_row = ((size + 31) // 32) * 4
    mask = bytes(mask_row * size)
    header = struct.pack("<IiiHHIIiiII", 40, size, size * 2, 1, 32, 0,
                         len(rows) + len(mask), 0, 0, 0, 0)
    return header + bytes(rows) + mask


def write_icon(path, image):
    frames = [frame(image, size) for size in SIZES]
    out = bytearray(struct.pack("<HHH", 0, 1, len(frames)))
    offset = 6 + 16 * len(frames)
    for size, data in zip(SIZES, frames):
        out += struct.pack("<BBBBHHII", size % 256, size % 256, 0, 0, 1, 32,
                           len(data), offset)
        offset += len(data)
    for data in frames:
        out += data
    with open(path, "wb") as stream:
        stream.write(out)


def main():
    for name, draw_game in GAMES.items():
        path = os.path.join(ROOT, name + ".ico")
        write_icon(path, draw_game())
        print(path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
