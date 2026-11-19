"""Draws the graphPaper logo: the name written on graph paper, one letter to a
square, as plain SVG paths (no font needed to show it).

    python3 scripts/make_logo.py <out-dir>

Writes logo.svg (the word on one row, a letter to a square), icon.svg (a G
in the middle of a 3 x 3 sheet, for the app icon and favicon) and
logo-candidates.html to compare the variants.
The letters are Noto Sans Bold (SIL Open Font License), turned into outlines.
"""
import sys
from pathlib import Path
from fontTools.ttLib import TTFont
from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.boundsPen import BoundsPen

FONT = '/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf'
GRID = '#bcd3f5'
INK = '#1d2733'
BLUE = '#1f6feb'

font = TTFont(FONT)
glyphs = font.getGlyphSet()
cmap = font.getBestCmap()
upm = font['head'].unitsPerEm


def letter(ch, cx, cy, size, color, fill=0.72):
    """One letter centred in a square of `size` at (cx, cy), as a path."""
    name = cmap[ord(ch)]
    pen = SVGPathPen(glyphs)
    glyphs[name].draw(pen)
    bounds = BoundsPen(glyphs)
    glyphs[name].draw(bounds)
    x0, y0, x1, y1 = bounds.bounds
    scale = size * fill / upm
    # Centre on the letter's ink horizontally and on the cap height vertically,
    # so a 'p' hangs below the line the way handwriting in a square would.
    cap = font['OS/2'].sCapHeight
    tx = cx - (x0 + x1) / 2 * scale
    ty = cy + cap / 2 * scale
    return f'<path fill="{color}" transform="translate({tx:.2f} {ty:.2f}) scale({scale:.5f} {-scale:.5f})" d="{pen.getCommands()}"/>'


def sheet(rows, cell=40, pad=1, colors=None, stroke=2, frame=False, fill=0.72):
    """The words in `rows`, a letter to a square, on a grid `pad` squares wider all round."""
    cols = max(len(r) for r in rows) + 2 * pad
    height = len(rows) + 2 * pad
    w, h = cols * cell, height * cell
    lines = [f'<path stroke="{GRID}" stroke-width="{stroke}" fill="none" d="' +
             ''.join(f'M{x * cell} 0V{h}' for x in range(cols + 1)) +
             ''.join(f'M0 {y * cell}H{w}' for y in range(height + 1)) + '"/>']
    marks = []
    for r, word in enumerate(rows):
        start = pad + (cols - 2 * pad - len(word)) // 2
        for i, ch in enumerate(word):
            color = (colors or {}).get((r, i), INK)
            marks.append(letter(ch, (start + i + 0.5) * cell, (pad + r + 0.5) * cell, cell, color, fill))
    rx = cell * 0.4 if frame else 0
    border = f'<rect x="{stroke}" y="{stroke}" width="{w - 2 * stroke}" height="{h - 2 * stroke}" rx="{rx}" fill="none" stroke="{BLUE}" stroke-width="{stroke * 2}"/>' if frame else ''
    key = 'clip-' + ''.join(rows).lower()
    clip = f'<clipPath id="{key}"><rect width="{w}" height="{h}" rx="{rx}"/></clipPath>' if frame else ''
    body = f'<rect width="{w}" height="{h}" fill="#fff"/>' + ''.join(lines) + ''.join(marks)
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" width="{w}" height="{h}">' + clip +
            (f'<g clip-path="url(#{key})">{body}</g>' if frame else body) + border + '</svg>')


out = Path(sys.argv[1] if len(sys.argv) > 1 else '.')
out.mkdir(parents=True, exist_ok=True)
caps = {(0, 0): BLUE, (0, 5): BLUE}
variants = {
    'logo': sheet(['GraphPaper'], pad=0, colors=caps),
    'icon': sheet(['G'], cell=40, pad=1, colors={(0, 0): BLUE}, frame=True, fill=0.82),
    'two-rows': sheet(['Graph', 'Paper'], pad=1, colors={(0, 0): BLUE, (1, 0): BLUE}, frame=True),
}
(out / 'logo.svg').write_text(variants['logo'])
(out / 'icon.svg').write_text(variants['icon'])
html = ['<!doctype html><meta charset="utf-8"><body style="font-family:sans-serif;background:#f3f5f7;padding:24px">']
for k, v in variants.items():
    html.append(f'<div style="margin:0 0 28px"><div style="color:#666;margin-bottom:6px">{k.upper()}</div>{v}</div>')
html.append('<div style="color:#666;margin:8px 0">小さく（ヘッダー・ファビコン）</div>')
html.append(f'<div style="display:flex;gap:24px;align-items:center"><img style="height:28px" src="data:image/svg+xml;utf8,{variants["logo"].replace("#", "%23").replace(chr(34), chr(39))}">'
            f'<img style="height:64px" src="data:image/svg+xml;utf8,{variants["icon"].replace("#", "%23").replace(chr(34), chr(39))}">'
            f'<img style="height:16px" src="data:image/svg+xml;utf8,{variants["icon"].replace("#", "%23").replace(chr(34), chr(39))}"></div>')
(out / 'logo-candidates.html').write_text(''.join(html))
print('wrote', out)
