"""
map_icons.py - everything to do with the small brand icons mapmaking.py
draws on fuel stations: the image files themselves (in icons/fuel and
icons/food), loading them into the generated SVG, the colour-coded text
fallback for a brand with no icon file yet, and parsing the free-text
'Fuel brand' / 'Food brand(s)' columns (including the "X east; Y west"
convention for a service area with a different operator on each side).

These are real brand logos, used only to show which fuel and food/coffee
brand is actually present at an actual real-world location - not
decorative or promotional use. Keep that in mind before adding a new one:
source it from the operator's own site for that same location (as the
existing icons/fuel and icons/food files are, from serways.de) or from a
well-documented source such as Wikimedia Commons.
"""
import base64
import html as html_lib
import mimetypes
import os
import re

ICONS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icons')

_NARROW = set("iljtfrI.,:;'|!()[]- ")
_WIDE = set("mwMW@%&")


def text_width(text, font_size):
    """Conservative Arial width estimate (per-character classes), so a
    placed label is never wider on screen than the box reserved for it.
    (Shared with mapmaking.py, which imports it from here.)"""
    w = 0.0
    for ch in str(text or ''):
        if ch in _NARROW:
            w += 0.32
        elif ch in _WIDE:
            w += 0.86
        elif ch.isupper():
            w += 0.70
        else:
            w += 0.57
    return w * font_size * 1.03

# brand name (as it appears in the 'Fuel brand' / 'Food brand(s)' / free-text
# 'Facilities' columns) -> icon filename under icons/fuel or icons/food.
# Add a line here (and the image file) to show a real logo for another brand;
# a brand with a colour entry below but no line here still gets the sturdier
# coloured-text badge as a fallback.
FUEL_ICON_FILES = {
    "Shell": "shell.png", "Aral": "aral.png", "Esso": "esso.png",
    "Total": "total.png", "TotalEnergies": "totalenergies.svg",
    "Eni": "eni.svg", "OMV": "omv.svg", "AVIA": "avia.svg", "Avia": "avia.svg",
    "Westfalen": "westfalen.png", "Q8": "q8.svg", "Circle K": "circlek.svg",
}
FOOD_ICON_FILES = {
    "McDonald's": "mcdonalds.png", "Burger King": "burgerking.png",
    "Starbucks": "starbucks.png", "NORDSEE": "nordsee.png",
    "Dallmayr": "dallmayr.png", "Segafredo": "segafredo.jpg",
    "BrotZeit": "brotzeit.png", "Coffee Fellows": "coffeefellows.png",
    "Tabilo": "tabilo.png", "Lavazza": "lavazza.png",
}

# Fallback for a brand with no icon file (yet): background, text colour,
# short code - a sturdier stand-in than plain text, not logo artwork.
FUEL_BRAND_COLOURS = [
    ("Shell", "#FBCE07", "#D7000F", "Shell"),
    ("Aral", "#0A6EB4", "#FFFFFF", "Aral"),
    ("Esso", "#ED1B2E", "#FFFFFF", "Esso"),
    ("TotalEnergies", "#EE2E24", "#FFFFFF", "Total"),
    ("Total", "#EE2E24", "#FFFFFF", "Total"),
    ("Circle K", "#D0271D", "#FFFFFF", "CircleK"),
    ("Eni", "#FFCC00", "#1A1A1A", "Eni"),
    ("OMV", "#004F9F", "#FFFFFF", "OMV"),
    ("AVIA", "#003DA5", "#FFFFFF", "Avia"),
    ("Avia", "#003DA5", "#FFFFFF", "Avia"),
    ("Westfalen", "#004F9F", "#FFFFFF", "Westf."),
    ("Q8", "#F7941D", "#1A1A1A", "Q8"),
    ("BP", "#00843D", "#FFFFFF", "BP"),
    ("Texaco", "#ED1C24", "#FFFFFF", "Texaco"),
    ("Tinq", "#E2001A", "#FFFFFF", "Tinq"),
    ("Tango", "#E2001A", "#FFFFFF", "Tango"),
    ("AS24", "#1A1A1A", "#FFFFFF", "AS24"),
    ("OK", "#F39200", "#1A1A1A", "OK"),
    ("Samba Oil", "#1A1A1A", "#FFC600", "Samba"),
    ("Tamoil", "#004B87", "#FFFFFF", "Tamoil"),
]
FOOD_BRAND_COLOURS = [
    ("McDonald's", "#DA291C", "#FFC72C", "M"),
    ("Burger King", "#D62300", "#FFFFFF", "BK"),
    ("Starbucks", "#00704A", "#FFFFFF", "SB"),
    ("Subway", "#008938", "#FFC600", "Subway"),
    ("NORDSEE", "#0066B1", "#FFFFFF", "Nordsee"),
    ("Dallmayr", "#6F3B20", "#FFFFFF", "Dallmayr"),
    ("Segafredo", "#6F1D1D", "#FFFFFF", "Segafr."),
    ("BrotZeit", "#8B5A2B", "#FFFFFF", "BrotZ."),
    ("Coffee Fellows", "#4A2E19", "#FFFFFF", "CoffF."),
    ("Tabilo", "#8B5A2B", "#FFFFFF", "Tabilo"),
    ("Lavazza", "#6F1D1D", "#FFFFFF", "Lavazza"),
    ("La Place", "#006241", "#FFFFFF", "La Place"),
    ("Spar", "#006838", "#FFFFFF", "Spar"),
    ("Albert Heijn to go", "#00A0DC", "#FFFFFF", "AH to go"),
]

ICON_SIZE = 8.5     # map: brand icon tile side length, in SVG units
ICON_GAP = 1.2
BADGE_PAD_X = 1.4
BADGE_GAP = 1.0

_data_uri_cache = {}
_registered_icons = {}  # icon_id -> (subfolder, filename), for the current document


def _icon_data_uri(subfolder, filename):
    """Reads icons/<subfolder>/<filename> once and caches it as a data: URI,
    so the generated HTML stays a single portable file - no separate image
    requests, nothing that can go missing when the page is moved."""
    key = (subfolder, filename)
    if key in _data_uri_cache:
        return _data_uri_cache[key]
    path = os.path.join(ICONS_DIR, subfolder, filename)
    mime = mimetypes.guess_type(filename)[0] or 'application/octet-stream'
    with open(path, 'rb') as f:
        data = f.read()
    uri = f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"
    _data_uri_cache[key] = uri
    return uri


def reset_icon_registry():
    """Call once at the start of building a document (the map tab, or one
    route tab) - each is a separate <svg> in its own iframe, so each needs
    its own self-contained <defs> of just the icons it actually uses."""
    _registered_icons.clear()


def _register_icon(subfolder, filename):
    """Registers icons/<subfolder>/<filename> for the current document's
    <defs> (once, however many times it's actually used) and returns its
    element id, stable for a given subfolder+filename."""
    icon_id = f"icon-{subfolder}-{os.path.splitext(filename)[0]}"
    _registered_icons[icon_id] = (subfolder, filename)
    return icon_id


def icon_defs_svg():
    """<defs> of every icon registered since the last reset_icon_registry(),
    each wrapped in a <symbol> (a fixed 0 0 100 100 viewBox) so a <use> can
    size it to any badge/marker just by giving its own width/height - the
    actual image bytes appear exactly once however many times it's used."""
    if not _registered_icons:
        return ''
    parts = ['<defs>']
    for icon_id, (subfolder, filename) in _registered_icons.items():
        uri = _icon_data_uri(subfolder, filename)
        parts.append(f'<symbol id="{icon_id}" viewBox="0 0 100 100">'
                     f'<image width="100" height="100" href="{uri}" '
                     f'preserveAspectRatio="xMidYMid meet"/></symbol>')
    parts.append('</defs>')
    return ''.join(parts)


def use_icon_svg(icon_id, x, y, w, h):
    return f'<use href="#{icon_id}" x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}"/>'


def split_brand_list(s):
    """'Circle K / TotalEnergies' / 'Shell, Eni' / 'Tabilo and Dallmayr'
    -> ['Circle K', 'TotalEnergies'] etc."""
    return [p.strip() for p in re.split(r'/|,|\band\b', s, flags=re.IGNORECASE) if p.strip()]


def parse_sided_brands(text):
    """Parses the 'Fuel brand' / 'Food brand(s)' free-text convention:
    - 'Aral both directions' / plain 'Shell' -> one list, same both sides
    - 'Shell east; Eni west' -> a brand list per compass side
    - 'Circle K/TotalEnergies toward Antwerp; Esso toward Breda' -> a
      direction named by place can't be resolved to a map side from text
      alone, so it's treated as one combined list (every brand mentioned)

    Returns ('single', [brand, ...]) or ('sided', {'east': [...], ...}).
    """
    if not text:
        return 'single', []
    clauses = [c.strip() for c in re.split(r';', str(text)) if c.strip()]
    sided, single = {}, []
    for clause in clauses:
        clause = re.sub(r'\s+toward\s+.+$', '', clause, flags=re.IGNORECASE).strip()
        clause = re.sub(r'\s+both directions\s*$', '', clause, flags=re.IGNORECASE).strip()
        m = re.match(r'^(.*?)\s+(east|west|north|south)$', clause, re.IGNORECASE)
        if m:
            brand_part, side = m.group(1).strip(), m.group(2).lower()
            sided.setdefault(side, []).extend(split_brand_list(brand_part))
        elif clause:
            single.extend(split_brand_list(clause))
    if sided:
        for side_list in sided.values():
            side_list.extend(single)
        return 'sided', sided
    return 'single', single


def _entries_for_brands(brands, table, icon_files, subfolder, seen):
    """Matches each name in `brands` against `table`, preferring an exact
    (case-insensitive) match over a substring one - so free text naming the
    plain 'Total' brand doesn't pick up the 'TotalEnergies' entry just
    because one name contains the other."""
    out = []
    for b in brands:
        key = b.lower()
        match = next((row for row in table if row[0].lower() == key), None)
        if not match:
            match = next((row for row in table if row[0].lower() in key or key in row[0].lower()), None)
        if not match:
            continue
        canon, bg, fg, code = match
        if code in seen:
            continue
        seen.add(code)
        icon_file = icon_files.get(canon)
        icon_id = _register_icon(subfolder, icon_file) if icon_file else None
        out.append(dict(code=code, bg=bg, fg=fg, icon=icon_id))
    return out


def brand_badge_entries(fuel_text, food_text, facilities_text, max_food=4):
    """Combines the 'Fuel brand' and 'Food brand(s)' (or, lacking that,
    free-text 'Facilities') columns into badge entries for a fuel station.

    Returns ('single', [entry, ...]) when one badge row covers the whole
    station, or ('sided', {'east': [entry, ...], ...}) when the fuel
    and/or food brand differs per side of the road - each entry is
    dict(code, bg, fg, img) where `img` is a data: URI (real logo) or None
    (use the coloured bg/fg/code fallback)."""
    fuel_mode, fuel_data = parse_sided_brands(fuel_text)
    if food_text:
        food_mode, food_data = parse_sided_brands(food_text)
    else:
        food_mode = 'single'
        food_data = [b for b, *_ in FOOD_BRAND_COLOURS
                     if facilities_text and b.lower() in str(facilities_text).lower()][:max_food]

    sides = set()
    if fuel_mode == 'sided':
        sides |= set(fuel_data)
    if food_mode == 'sided':
        sides |= set(food_data)

    if not sides:
        seen = set()
        entries = _entries_for_brands(fuel_data, FUEL_BRAND_COLOURS, FUEL_ICON_FILES, 'fuel', seen)
        entries += _entries_for_brands(food_data[:max_food], FOOD_BRAND_COLOURS, FOOD_ICON_FILES, 'food', seen)
        return 'single', entries

    def side_brands(mode, data, side):
        return data.get(side, []) if mode == 'sided' else data  # single: same list every side

    out = {}
    for side in sides:
        seen = set()
        fuel_here = side_brands(fuel_mode, fuel_data, side)
        food_here = side_brands(food_mode, food_data, side)
        entries = _entries_for_brands(fuel_here, FUEL_BRAND_COLOURS, FUEL_ICON_FILES, 'fuel', seen)
        entries += _entries_for_brands(food_here[:max_food], FOOD_BRAND_COLOURS, FOOD_ICON_FILES, 'food', seen)
        out[side] = entries
    return 'sided', out


def badge_row_metrics(entries, size=ICON_SIZE):
    """Width/height of a row of brand badges (icon tiles, or - for a brand
    with no icon file - a coloured text chip), and each one's width."""
    if not entries:
        return 0.0, 0.0, []
    widths = []
    for e in entries:
        if e['icon']:
            widths.append(size)
        else:
            widths.append(text_width(e['code'], 4.2) + 2 * BADGE_PAD_X)
    total_w = sum(widths) + BADGE_GAP * (len(widths) - 1)
    height = size
    return total_w, height, widths


def badge_row_svg(entries, left_x, top_y, size=ICON_SIZE):
    """Left-aligned row of brand badges starting at (left_x, top_y): a
    small square tile with the real logo where we have one on file, or a
    coloured rounded-rect + short code where we don't."""
    if not entries:
        return ''
    _total_w, height, widths = badge_row_metrics(entries, size)
    out = ['<g class="brand-badge">']
    x = left_x
    for e, w in zip(entries, widths):
        if e['icon']:
            out.append(f'<rect x="{x:.1f}" y="{top_y:.1f}" width="{w:.1f}" height="{height:.1f}" '
                       f'rx="1.2" fill="white" stroke="#B5B5AE" stroke-width="0.4"/>')
            pad = w * 0.08
            out.append(use_icon_svg(e['icon'], x + pad, top_y + pad, w - 2*pad, height - 2*pad))
        else:
            out.append(f'<rect x="{x:.1f}" y="{top_y:.1f}" width="{w:.1f}" height="{height:.1f}" '
                       f'rx="1.4" fill="{e["bg"]}" stroke="white" stroke-width="0.4"/>')
            out.append(f'<text x="{x + w/2:.1f}" y="{top_y + height/2 + 1.5:.1f}" '
                       f'text-anchor="middle" font-size="4.2" font-weight="700" '
                       f'fill="{e["fg"]}">{html_lib.escape(e["code"])}</text>')
        x += w + BADGE_GAP
    out.append('</g>')
    return ''.join(out)


def primary_fuel_icon(brand_text):
    """The first matching fuel brand's (icon id, bg colour) for a
    single-icon marker (the route tab's square, which - unlike the map
    tab - doesn't try to place a badge on each physical side of the
    road). (None, a neutral orange) when nothing in `brand_text` matches
    a known brand. Pass the icon id to use_icon_svg() to draw it."""
    mode, data = parse_sided_brands(brand_text)
    brands = data if mode == 'single' else next(iter(data.values()), [])
    entries = _entries_for_brands(brands, FUEL_BRAND_COLOURS, FUEL_ICON_FILES, 'fuel', set())
    if entries:
        return entries[0]['icon'], entries[0]['bg']
    return None, "#E8871E"


def mountain_icon_svg(cx, cy, size=10, colour="#4B7A3C"):
    """Small two-peak mountain pictogram for a natural-region label (not a
    brand logo - an original shape, free to draw)."""
    h, w = size, size * 1.15
    return (f'<path d="M {cx-w:.1f} {cy+h*0.5:.1f} L {cx-w*0.25:.1f} {cy-h*0.55:.1f} '
            f'L {cx+w*0.05:.1f} {cy-h*0.1:.1f} L {cx+w*0.35:.1f} {cy-h*0.75:.1f} '
            f'L {cx+w:.1f} {cy+h*0.5:.1f} Z" fill="{colour}" opacity="0.9"/>')
