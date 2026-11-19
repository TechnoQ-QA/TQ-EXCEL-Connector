"""Writes the add-in's icons: the main one (icon-16/32/64/80.png) and one per ribbon button
(icon-<name>-16/32/80.png).

Each button icon is drawn from simple shapes (rounded boxes, circles, arcs, strokes) on a 16 x 16 grid,
sampled 4 x 4 per pixel so the edges are smooth at every size, on a clear background.

    python scripts/make_icons.py
"""

import math
import os
import struct
import zlib

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "addin", "assets")
BLUE = (11, 92, 173)
WHITE = (255, 255, 255)

C = {"blue": (16, 110, 190), "navy": (31, 58, 95), "green": (16, 124, 65), "red": (196, 43, 28),
     "amber": (234, 145, 0), "grey": (110, 120, 132), "light": (222, 233, 245), "white": (255, 255, 255),
     "paper": (250, 251, 252)}


# ------------------------------------------------------------------ shapes --
# Each shape: (test(x, y) -> bool, colour). Later shapes paint over earlier ones.

def box(x0, y0, x1, y1, colour, r=0.0):
    def test(x, y):
        dx = max(x0 + r - x, 0, x - (x1 - r))
        dy = max(y0 + r - y, 0, y - (y1 - r))
        return x0 <= x <= x1 and y0 <= y <= y1 and dx * dx + dy * dy <= r * r + 1e-9
    return test, colour


def disc(cx, cy, r, colour):
    return (lambda x, y: (x - cx) ** 2 + (y - cy) ** 2 <= r * r), colour


def ring(cx, cy, r, width, colour, start=0, end=360):
    """An arc of a circle, angles in degrees, 0 = right, counter-clockwise (y up)."""
    def test(x, y):
        d = math.hypot(x - cx, y - cy)
        if abs(d - r) > width / 2:
            return False
        a = math.degrees(math.atan2(cy - y, x - cx)) % 360
        return (start <= a <= end) if start <= end else (a >= start or a <= end)
    return test, colour


def stroke(points, width, colour):
    """A thick line through the points, with round ends."""
    segments = list(zip(points, points[1:]))

    def test(x, y):
        for (ax, ay), (bx, by) in segments:
            vx, vy = bx - ax, by - ay
            t = max(0.0, min(1.0, ((x - ax) * vx + (y - ay) * vy) / (vx * vx + vy * vy or 1)))
            if (x - ax - t * vx) ** 2 + (y - ay - t * vy) ** 2 <= (width / 2) ** 2:
                return True
        return False
    return test, colour


def poly(points, colour):
    def test(x, y):
        inside = False
        for (ax, ay), (bx, by) in zip(points, points[1:] + points[:1]):
            if (ay > y) != (by > y) and x < (bx - ax) * (y - ay) / (by - ay) + ax:
                inside = not inside
        return inside
    return test, colour


def star(cx, cy, outer, inner, colour, points=4):
    pts = []
    for n in range(points * 2):
        a = math.pi / points * n - math.pi / 2
        r = outer if n % 2 == 0 else inner
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return poly(pts, colour)


def refresh_badge(cx, cy, r, colour):
    """A small circular arrow (a refresh sign) in a white disc."""
    return [disc(cx, cy, r + 1.0, C["white"]),
            ring(cx, cy, r - 0.6, 1.3, colour, 120, 60),
            poly([(cx + r * 0.95, cy - r * 1.25), (cx + r * 1.15, cy + 0.15), (cx - 0.05, cy - r * 0.55)], colour)]


# ------------------------------------------------------------------- icons --

TQ_ORANGE = (244, 121, 32)   # Techno Q's orange, from the company logo
TQ_MARK = [box(0.4, 0.4, 15.6, 15.6, TQ_ORANGE, 3.2),
           box(1.9, 3.6, 8.0, 5.5, C["white"], 0.5), box(4.0, 3.6, 5.9, 12.6, C["white"], 0.5),
           ring(11.0, 8.1, 2.95, 1.75, C["white"]), stroke([(11.6, 9.4), (14.1, 12.6)], 1.75, C["white"])]


ICONS = {
    "tq": TQ_MARK,
    "panel": [box(2.5, 10, 5, 13.5, C["blue"], 0.6), box(6.75, 6, 9.25, 13.5, C["blue"], 0.6),
              box(11, 2.5, 13.5, 13.5, C["blue"], 0.6), box(1.5, 13.6, 14.5, 14.6, C["grey"], 0.4)],
    "wizard": [stroke([(3, 13.5), (10.2, 6.3)], 2.2, C["navy"]), stroke([(9, 7.5), (10.6, 5.9)], 2.2, C["amber"]),
               star(11.8, 4.2, 3.2, 1.0, C["amber"]), star(5, 3.5, 1.5, 0.45, C["amber"]),
               star(13.6, 10, 1.3, 0.4, C["amber"])],
    "dashboard": [box(1.5, 1.5, 14.5, 14.5, C["light"], 1.6),
                  ring(5.6, 5.6, 2.4, 2.6, C["blue"], 0, 270), ring(5.6, 5.6, 2.4, 2.6, C["amber"], 270, 360),
                  box(9.3, 4.2, 13, 5.1, C["grey"], 0.4), box(9.3, 6.2, 12, 7.1, C["grey"], 0.4),
                  box(3, 11, 4.8, 13, C["green"], 0.4), box(5.6, 9.4, 7.4, 13, C["green"], 0.4),
                  box(8.2, 10.4, 10, 13, C["green"], 0.4), box(10.8, 8.4, 12.6, 13, C["green"], 0.4)],
    "chat": [box(1.5, 2.5, 14.5, 11, C["blue"], 2.2), poly([(4, 10), (8, 10), (3.5, 14.2)], C["blue"]),
             disc(5, 6.75, 1.0, C["white"]), disc(8, 6.75, 1.0, C["white"]), disc(11, 6.75, 1.0, C["white"])],
    "refresh": [ring(8, 8, 5.2, 1.9, C["green"], 110, 50),
                poly([(12.4, 1.6), (14.4, 7.4), (8.6, 5.2)], C["green"])],
    "sheet": [box(1.5, 1.5, 12.5, 12.5, C["paper"], 1.0), box(1.5, 1.5, 12.5, 4.5, C["light"], 1.0)]
             + [box(1.5, 1.5, 12.5, 1.5 + 0.9, C["grey"], 0.4)]
             + [box(x, 1.5, x + 0.8, 12.5, C["grey"]) for x in (1.5, 5.2, 8.9, 11.7)]
             + [box(1.5, y, 12.5, y + 0.8, C["grey"]) for y in (4.4, 8.0, 11.7)]
             + refresh_badge(11.8, 11.8, 3.1, C["green"]),
    "selection": [box(x, 2, x + 1.6, 2.9, C["blue"]) for x in (2, 4.8, 7.6, 10.4)]
                 + [box(x, 11.1, x + 1.6, 12, C["blue"]) for x in (2, 4.8)]
                 + [box(2, y, 2.9, y + 1.6, C["blue"]) for y in (2, 4.8, 7.6, 10.4)]
                 + [box(11.1, y, 12, y + 1.6, C["blue"]) for y in (2, 4.8)]
                 + [box(4.6, 4.6, 9.4, 9.4, C["light"], 0.6)]
                 + refresh_badge(11.8, 11.8, 3.1, C["green"]),
    "drill": [disc(6.6, 6.6, 4.6, C["light"]), ring(6.6, 6.6, 4.6, 1.8, C["blue"]),
              box(4.4, 6.4, 5.6, 8.6, C["blue"], 0.3), box(6.1, 4.6, 7.3, 8.6, C["blue"], 0.3),
              box(7.8, 5.6, 9.0, 8.6, C["blue"], 0.3), stroke([(10.2, 10.2), (14, 14)], 2.4, C["navy"])],
    "formula": [stroke([(9.2, 2.2), (7.6, 2.2), (6.6, 3.4), (5.2, 13.8)], 1.6, C["blue"]),
                stroke([(3.6, 6.6), (8.6, 6.6)], 1.5, C["blue"]),
                stroke([(9.4, 8.4), (13.6, 13.6)], 1.6, C["green"]), stroke([(13.6, 8.4), (9.4, 13.6)], 1.6, C["green"])],
    "values": [box(1.5, 1.5, 9.5, 11, C["white"], 1.0), box(1.5, 1.5, 9.5, 11, C["grey"], 1.0),
               box(2.4, 2.4, 8.6, 10.1, C["white"], 0.5),
               box(5.5, 5, 14.5, 14.5, C["blue"], 1.0), box(6.4, 5.9, 13.6, 13.6, C["white"], 0.5),
               box(7.6, 7.6, 12.4, 8.5, C["blue"], 0.4), box(7.6, 9.5, 12.4, 10.4, C["blue"], 0.4),
               box(7.6, 11.4, 11, 12.3, C["blue"], 0.4)],
    "pdf": [poly([(3, 1), (10.5, 1), (13.5, 4), (13.5, 15), (3, 15)], C["grey"]),
            poly([(3.9, 1.9), (10.1, 1.9), (12.6, 4.4), (12.6, 14.1), (3.9, 14.1)], C["white"]),
            poly([(10.1, 1.9), (10.1, 4.4), (12.6, 4.4)], C["grey"]),
            box(1, 7, 15, 12.6, C["red"], 1.0),
            stroke([(2.9, 11.4), (2.9, 8.2), (4.4, 8.2), (4.9, 8.8), (4.4, 9.6), (2.9, 9.6)], 0.9, C["white"]),
            stroke([(6.6, 8.2), (6.6, 11.4), (7.8, 11.4), (8.8, 10.4), (8.8, 9.2), (7.8, 8.2), (6.6, 8.2)], 0.9, C["white"]),
            stroke([(12.8, 8.2), (10.6, 8.2), (10.6, 11.4)], 0.9, C["white"]), stroke([(10.6, 9.8), (12.2, 9.8)], 0.9, C["white"])],
}


# --------------------------------------------------------------------- PNG --

def chunk(kind, data):
    body = kind + data
    return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)


def png(width, height, pixel, alpha):
    rows = bytearray()
    for y in range(height):
        rows.append(0)
        for x in range(width):
            rows.extend(pixel(x, y))
    header = struct.pack(">IIBBBBB", width, height, 8, 6 if alpha else 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header)
            + chunk(b"IDAT", zlib.compress(bytes(rows), 9)) + chunk(b"IEND", b""))


def icon(size):
    """The main icon: the TQ mark (white T and Q on Techno Q's orange)."""
    return button_icon(TQ_MARK, size)


def old_icon(size):
    """The first placeholder icon: a blue square with a white bar chart."""
    unit = size / 16.0
    bars = [(3, 9), (7, 6), (11, 3)]  # (left, top) in 16ths; each bar is 2 wide, down to row 13

    def pixel(x, y):
        gx, gy = x / unit, y / unit
        return WHITE if any(left <= gx < left + 2 and top <= gy < 13 for left, top in bars) else BLUE
    return png(size, size, pixel, False)


def button_icon(shapes, size, samples=4):
    unit = 16.0 / size

    def pixel(px, py):
        r = g = b = a = 0.0
        for sy in range(samples):
            for sx in range(samples):
                x = (px + (sx + 0.5) / samples) * unit
                y = (py + (sy + 0.5) / samples) * unit
                for test, colour in reversed(shapes):
                    if test(x, y):
                        r, g, b, a = r + colour[0], g + colour[1], b + colour[2], a + 1
                        break
        n = samples * samples
        if not a:
            return (0, 0, 0, 0)
        return (int(r / a), int(g / a), int(b / a), int(255 * a / n))
    return png(size, size, pixel, True)


def ico(sizes=(16, 32, 48, 256)):
    """tq.ico for the setup file and its uninstaller: the main icon at several sizes, each stored as a PNG."""
    images = [icon(z) for z in sizes]
    out = struct.pack("<HHH", 0, 1, len(images))
    offset = 6 + 16 * len(images)
    for z, data in zip(sizes, images):
        out += struct.pack("<BBBBHHII", z % 256, z % 256, 0, 0, 1, 32, len(data), offset)
        offset += len(data)
    return out + b"".join(images)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    for size in (16, 32, 64, 80):
        with open(os.path.join(OUT, "icon-%d.png" % size), "wb") as handle:
            handle.write(icon(size))
    for name, shapes in ICONS.items():
        for size in (16, 32, 80):
            with open(os.path.join(OUT, "icon-%s-%d.png" % (name, size)), "wb") as handle:
                handle.write(button_icon(shapes, size))
    with open(os.path.join(OUT, "tq.ico"), "wb") as handle:
        handle.write(ico())
    print("Icons written to", OUT)
