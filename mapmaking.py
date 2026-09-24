#!/usr/bin/env python3
"""
mapmaking.py - builds one HTML page with two tabs from the route workbook:

  * Kaart  - the geography-preserving schematic map of the whole network
             (junctions near their real lat/lon, important roads pulled
             into straighter regional "spines"; see Part 1 below)
  * Route  - an NS-style "metro board" diagram of the route between two
             places with its alternative branches (ordinal positions, no
             geography; see Part 2 below)

Bridge, tunnel and fuel-station labels in both views are placed by
LabelPlacer: each label is tried at increasing distances and angles around
its marker until it overlaps no other label, junction label or marker
(and, close to the marker, no road line either). Leader lines connect
labels that had to move away. In the (vertical) route diagram, labels are
first stacked in ordered columns beside each line, so their leaders don't
cross. Because each view is one SVG that scales as a whole, a layout
without overlaps in SVG units has no overlaps at any zoom level, and the
maximum zoom is set so the text is fully readable.

Usage (reads junctions_topology_v5.xlsx, writes holidays.html, both
next to this script):
    python3 mapmaking.py
    python3 mapmaking.py --from Vught --to Berwang --via "Venlo,Koblenz"
"""
import argparse
import colorsys
import hashlib
import html as html_lib
import json
import math
import os
import re
from collections import defaultdict

import networkx as nx
from openpyxl import load_workbook

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSX = os.path.join(SCRIPT_DIR, 'junctions_topology_v5.xlsx')
DEFAULT_OUTPUT = os.path.join(SCRIPT_DIR, 'holidays.html')


# ======================================================================
# Shared: point-label placement (bridges, tunnels, fuel stations)
# ======================================================================

POINT_FONT = 4.5          # map: point name size, in SVG units
POINT_SUB_FONT = 3.8      # map: second line (bridge length / fuel brand)
MAP_POI_FONT = 5.5        # map: point-of-interest name size
MAP_POI_R = 6.0           # map: point-of-interest star radius
MAP_MAX_ZOOM = 16         # map: max zoom - POINT_FONT * 16 = 72px text
GRAPH_POINT_FONT = 10     # route diagram: point label size, in SVG units
GRAPH_DIST_FONT = 10      # route diagram: segment distance label size
GRAPH_JUNCTION_FONT = 14  # route diagram: junction name size
GRAPH_ROAD_FONT = 13      # route diagram: road number size
GRAPH_FUEL_HALF = 5.0     # route diagram: half the size of a fuel-station square
GRAPH_BRIDGE_R = 4.0      # route diagram: bridge/tunnel dot radius
GRAPH_POI_R = 9.0         # route diagram: point-of-interest star radius
GRAPH_MAX_ZOOM = 12
LINE_HEIGHT = 1.2         # line box height as a multiple of the font size
BASELINE = 0.93           # baseline offset from the top of a line box

GRAPH_BRIDGE_CATEGORIES = ("brug (rivier)", "brug (dal)", "Brug", "tunnel", "ecoduct")
# placement order: the first placed get the closest spots
LABEL_PRIORITY = {'poi': -1, 'tankstation': 0, 'autohof': 0, 'tunnel': 1, 'brug (rivier)': 2,
                  'ecoduct': 3, 'brug': 4, 'brug (dal)': 5}

_NARROW = set("iljtfrI.,:;'|!()[]- ")
_WIDE = set("mwMW@%&")


def text_width(text, font_size):
    """Conservative Arial width estimate (per-character classes), so a
    placed label is never wider on screen than the box reserved for it."""
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


def nearest_on_box(x, y, box):
    return min(max(x, box[0]), box[2]), min(max(y, box[1]), box[3])


def _seg_hits_box(seg, box):
    """Liang-Barsky: does line segment (x1, y1, x2, y2) touch the box?"""
    x1, y1, x2, y2 = seg
    dx, dy = x2 - x1, y2 - y1
    t0, t1 = 0.0, 1.0
    for p, q in ((-dx, x1 - box[0]), (dx, box[2] - x1), (-dy, y1 - box[1]), (dy, box[3] - y1)):
        if p == 0:
            if q < 0:
                return False
        else:
            t = q / p
            if p < 0:
                if t > t1:
                    return False
                t0 = max(t0, t)
            else:
                if t < t0:
                    return False
                t1 = min(t1, t)
    return True


class LabelPlacer:
    """Greedy collision-free label placement on a uniform grid.

    Hard obstacles (boxes): junction circles and labels, point markers and
    every label already placed - a label never overlaps any of these.
    Soft obstacles (lines): road and river lines - avoided while the label
    is still close to its marker; further out a label may cross a line
    (its halo keeps it readable) rather than drift ever further away.
    Leader lines: a leader avoids running through any text (point or
    junction label) where possible, and no later label is ever placed on
    top of an earlier leader.
    """
    CELL = 24.0
    DISTANCES = (0, 1.5, 3, 5, 8, 12, 17, 23, 30, 40, 52, 66, 82, 100, 125, 155, 190, 230, 280)
    ANGLES = (0, 25, -25, 50, -50, 75, -75, 90, -90)

    def __init__(self, gap=0.6):
        self.gap = gap
        self._boxes = defaultdict(list)
        self._lines = defaultdict(list)
        self._texts = defaultdict(list)     # text boxes only: what a leader must not cross
        self._leaders = defaultdict(list)
        self.stats = dict(placed=0, with_leader=0, crossing_line=0, leader_through_text=0,
                          forced_overlap=0, max_leader=0.0)

    def _cells(self, box):
        c = self.CELL
        for i in range(int(math.floor(box[0] / c)), int(math.floor(box[2] / c)) + 1):
            for j in range(int(math.floor(box[1] / c)), int(math.floor(box[3] / c)) + 1):
                yield i, j

    def add_box(self, box, text=False):
        for k in self._cells(box):
            self._boxes[k].append(box)
            if text:
                self._texts[k].append(box)

    @staticmethod
    def rotated_boxes(ox, oy, angle_deg, lx, ly, w, h):
        """A rotated rectangle ([lx, lx+w] x [ly, ly+h] in a frame rotated by
        angle_deg around (ox, oy)), as a chain of small axis-aligned boxes."""
        a = math.radians(angle_deg)
        ca, sa = math.cos(a), math.sin(a)
        n = max(1, int(math.ceil(w / max(h, 1.0))))
        boxes = []
        for i in range(n):
            x0, x1 = lx + w * i / n, lx + w * (i + 1) / n
            pts = [(ox + x * ca - y * sa, oy + x * sa + y * ca) for x in (x0, x1) for y in (ly, ly + h)]
            boxes.append((min(p[0] for p in pts), min(p[1] for p in pts),
                          max(p[0] for p in pts), max(p[1] for p in pts)))
        return boxes

    def add_rotated_box(self, ox, oy, angle_deg, lx, ly, w, h, text=True):
        for box in self.rotated_boxes(ox, oy, angle_deg, lx, ly, w, h):
            self.add_box(box, text=text)

    def boxes_free(self, boxes, avoid_lines=False):
        """None of these boxes overlaps anything placed (nor, optionally, a line)?"""
        return not any(self._hits_box(b, None) or (avoid_lines and self._hits_line(b)) for b in boxes)

    def add_polyline(self, poly):
        for (x1, y1), (x2, y2) in zip(poly, poly[1:]):
            seg = (x1, y1, x2, y2)
            for k in self._cells((min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))):
                self._lines[k].append(seg)

    def _hits_box(self, box, ignore):
        g = self.gap
        padded = (box[0] - g, box[1] - g, box[2] + g, box[3] + g)
        for k in self._cells(padded):
            for o in self._boxes[k]:
                if o is ignore:
                    continue
                if not (padded[2] <= o[0] or o[2] <= padded[0] or padded[3] <= o[1] or o[3] <= padded[1]):
                    return True
        return False

    def _hits_line(self, box, grid=None):
        grid = self._lines if grid is None else grid
        for k in self._cells(box):
            for seg in grid[k]:
                if _seg_hits_box(seg, box):
                    return True
        return False

    def _leader_blocked(self, ax, ay, box, clear, h):
        """Does this candidate's leader line run through any text? Boxes
        containing the anchor itself don't count."""
        ex, ey = nearest_on_box(ax, ay, box)
        if math.hypot(ex - ax, ey - ay) <= clear + 0.3 * h:
            return False
        seg = (ax, ay, ex, ey)
        seen = set()
        for k in self._cells((min(ax, ex), min(ay, ey), max(ax, ex), max(ay, ey))):
            for o in self._texts[k]:
                if id(o) in seen:
                    continue
                seen.add(id(o))
                if o[0] <= ax <= o[2] and o[1] <= ay <= o[3]:
                    continue
                if _seg_hits_box(seg, o):
                    return True
        return False

    def place(self, ax, ay, w, h, normal_deg, pref_side=1, clear=4.0, ignore=None):
        """Box (x0, y0, x1, y1) for a w*h label belonging to the marker at
        (ax, ay), plus whether it needs a leader line. Candidates go outward
        in rings; within a ring the preferred side and the road normal come
        first. `clear` is how far the marker itself reaches from (ax, ay)."""
        for strict in (True, False):
            found = self._search(ax, ay, w, h, normal_deg, pref_side, clear, ignore, strict)
            if found:
                box, crossing = found
                if not strict:
                    self.stats['leader_through_text'] += 1
                return self._accept(box, ax, ay, clear, h, crossing)
        # nothing free at all: closest spot, overlapping
        self.stats['forced_overlap'] += 1
        a = math.radians(normal_deg)
        vx, vy = math.cos(a) * pref_side, math.sin(a) * pref_side
        reach = clear + abs(vx) * w / 2 + abs(vy) * h / 2
        cx, cy = ax + vx * reach, ay + vy * reach
        return self._accept((cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2), ax, ay, clear, h, False)

    def _search(self, ax, ay, w, h, normal_deg, pref_side, clear, ignore, strict):
        line_free_until = 3 * h
        for d in self.DISTANCES:
            free_crossing = None
            for side in (pref_side, -pref_side):
                for off in self.ANGLES:
                    a = math.radians(normal_deg + off)
                    vx, vy = math.cos(a) * side, math.sin(a) * side
                    reach = clear + d + abs(vx) * w / 2 + abs(vy) * h / 2
                    cx, cy = ax + vx * reach, ay + vy * reach
                    box = (cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)
                    if self._hits_box(box, ignore) or self._hits_line(box, self._leaders):
                        continue
                    if strict and self._leader_blocked(ax, ay, box, clear, h):
                        continue
                    if not self._hits_line(box):
                        return box, False
                    if free_crossing is None:
                        free_crossing = box
            if free_crossing is not None and d >= line_free_until:
                return free_crossing, True
        return None

    def fits(self, box, ignore=None, anchor=None, clear=0.0):
        """Free of labels, markers, junctions, leaders and lines - and, given
        the marker's anchor, is its leader free of text?"""
        if self._hits_box(box, ignore) or self._hits_line(box, self._leaders) or self._hits_line(box):
            return False
        return anchor is None or not self._leader_blocked(anchor[0], anchor[1], box, clear,
                                                          box[3] - box[1])

    def _accept(self, box, ax, ay, clear, h, crossing, leader=None):
        self.add_box(box, text=True)
        ex, ey = nearest_on_box(ax, ay, box)
        dist = math.hypot(ex - ax, ey - ay)
        if leader is None:
            leader = dist > clear + 0.3 * h
        self.stats['placed'] += 1
        self.stats['crossing_line'] += int(crossing)
        if leader:
            for k in self._cells((min(ax, ex), min(ay, ey), max(ax, ex), max(ay, ey))):
                self._leaders[k].append((ax, ay, ex, ey))
            self.stats['with_leader'] += 1
            self.stats['max_leader'] = round(max(self.stats['max_leader'], dist), 1)
        return box, leader


def stack_beside_vertical_lines(items, placer, side=-1, only=None, line_clearance=15.0, gap=1.0):
    """Route diagram: the labels of all points on one vertical line are
    stacked in a column beside it (side -1: left, right-aligned; +1: right,
    left-aligned - either way the column has a straight edge towards the
    line), in the same order as their markers - so leaders never
    cross - and each as close to its own marker as the stack allows
    (1-D cluster packing: a run of labels that would overlap is centred on
    the mean of their markers). A label whose spot is taken by something
    else (a junction name, another line) slides up or down to the nearest
    free spot that still keeps the order. `line_clearance` keeps the column clear of
    the largest junction circle. items: dicts with x, y, w, h, ignore;
    `only`: indices to consider (default all). Returns {index: (box,
    leader)}; labels whose spot is taken by something else are left out,
    for another column or LabelPlacer.place."""
    columns = defaultdict(list)
    for i, it in enumerate(items):
        if only is None or i in only:
            columns[round(it['x'], 1)].append(i)
    placed = {}
    for col in columns.values():
        col.sort(key=lambda i: items[i]['y'])
        clusters = []
        for i in col:
            it = items[i]
            clusters.append(dict(idx=[i], ideal=[it['y'] - it['h'] / 2], hs=[it['h']]))
            while len(clusters) > 1:
                prev, cur = clusters[-2], clusters[-1]
                if prev['top'] + sum(prev['hs']) + gap * len(prev['hs']) <= cur_top(cur, gap):
                    break
                clusters[-2:] = [dict(idx=prev['idx'] + cur['idx'], ideal=prev['ideal'] + cur['ideal'],
                                      hs=prev['hs'] + cur['hs'])]
                cur_top(clusters[-1], gap)
            cur_top(clusters[-1], gap)
        floor = -math.inf     # bottom of the previous label in this column
        for c in clusters:
            target = c['top']
            for i, hgt in zip(c['idx'], c['hs']):
                it = items[i]
                step = hgt / 3
                found = None
                for j in range(0, 31):
                    for top in ((target,) if j == 0 else (target + j * step, target - j * step)):
                        if top < floor:
                            continue
                        if side < 0:
                            box = (it['x'] - line_clearance - it['w'], top, it['x'] - line_clearance, top + hgt)
                        else:
                            box = (it['x'] + line_clearance, top, it['x'] + line_clearance + it['w'], top + hgt)
                        if placer.fits(box, it['ignore'], anchor=(it['x'], it['y']), clear=line_clearance):
                            found = (top, box)
                            break
                    if found:
                        break
                if found:
                    top, box = found
                    leader = abs(top + hgt / 2 - it['y']) > 0.35 * hgt
                    placed[i] = placer._accept(box, it['x'], it['y'], line_clearance, hgt, False, leader=leader)
                    floor = top + hgt + gap
                    target = max(target + hgt + gap, floor)
                else:
                    target += hgt + gap
    return placed


def cur_top(cluster, gap):
    """Best top for a packed run of labels: the mean of each label's ideal
    top minus its offset within the run (least-squares displacement)."""
    offsets, o = [], 0.0
    for hgt in cluster['hs']:
        offsets.append(o)
        o += hgt + gap
    cluster['top'] = sum(i - off for i, off in zip(cluster['ideal'], offsets)) / len(offsets)
    return cluster['top']


# ======================================================================
# Shared: points of interest (sheet 'Points of Interest')
# ======================================================================

POI_FILL = '#F5B301'      # gold star
POI_STROKE = '#7A4F00'
POI_TEXT = '#7A4F00'


def load_pois(wb, junctions, seg_ends):
    """Points of interest with a known segment. Each gets `frac`: where its
    lat/lon projects onto the straight line between the segment's From
    and To junction (0 = From, 1 = To), so it can be drawn on that
    segment's line in either view. seg_ends: edge id -> (from id, to id)."""
    if 'Points of Interest' not in wb.sheetnames:
        return []
    ws = wb['Points of Interest']
    header = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(header)}
    if 'Segment' not in idx:
        return []
    pois = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name, eid = row[idx['Name']], row[idx['Segment']]
        lat, lon = row[idx.get('Latitude')], row[idx.get('Longitude')]
        if not name or eid not in seg_ends or lat is None or lon is None:
            continue
        a, b = junctions.get(seg_ends[eid][0]), junctions.get(seg_ends[eid][1])
        if not a or not b or a.get('lat') is None or b.get('lat') is None:
            continue
        cos_lat = math.cos(math.radians(lat))
        ax, ay, bx, by = a['lon'] * cos_lat, a['lat'], b['lon'] * cos_lat, b['lat']
        dx, dy = bx - ax, by - ay
        t = ((lon * cos_lat - ax) * dx + (lat - ay) * dy) / ((dx * dx + dy * dy) or 1.0)
        pois.append(dict(name=str(name), category=row[idx['Category']] if 'Category' in idx else None,
                         notes=row[idx['Notes']] if 'Notes' in idx else None,
                         edge_id=eid, frac=min(0.95, max(0.05, t))))
    return pois


def star_svg(cx, cy, r, cls):
    """A five-pointed star centred on (cx, cy)."""
    pts = []
    for k in range(10):
        rad = r if k % 2 == 0 else r * 0.45
        a = math.radians(-90 + k * 36)
        pts.append(f"{cx + rad * math.cos(a):.1f},{cy + rad * math.sin(a):.1f}")
    return (f'<polygon class="{cls}" points="{" ".join(pts)}" fill="{POI_FILL}" '
            f'stroke="{POI_STROKE}" stroke-width="{max(0.5, r * 0.12):.2f}" stroke-linejoin="round"/>')


def poi_title(poi):
    parts = [poi['name']] + [str(v) for v in (poi.get('category'), poi.get('notes')) if v]
    return html_lib.escape(' - '.join(parts), quote=False)


def marker_extent(sx, sy, ang, cat, sides, side, length_m):
    """Bounding box of a map point marker (see bar_svg / fuel_marker_svg)
    and how far it reaches from the road; (None, small) if no marker."""
    a = math.radians(ang)
    tx, ty = math.cos(a), math.sin(a)
    nx_, ny_ = -ty, tx
    if cat in BAR_CATEGORIES:
        half_len, half_w = 6.5, 3.0
        pts = [(sx + nx_ * s * half_len + tx * t * half_w, sy + ny_ * s * half_len + ty * t * half_w)
               for s in (-1, 1) for t in (-1, 1)]
        reach = half_len + 0.8
    elif cat in ('tankstation', 'autohof'):
        length, base = 10.0, 4.5
        pts = [(sx + nx_ * side * length, sy + ny_ * side * length),
               (sx + tx * base, sy + ty * base), (sx - tx * base, sy - ty * base)]
        if sides and int(sides) >= 2:
            pts.append((sx - nx_ * side * length, sy - ny_ * side * length))
        reach = length + 0.8
    else:
        return None, 3.0
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return (min(xs), min(ys), max(xs), max(ys)), reach


def label_block_svg(group_class, lines, box, ax, ay, leader, halo, leader_width):
    """One placed label: 1-2 text lines, each with a halo in the background
    colour so it stays readable over lines, and its optional leader line.
    Returned separately so all leaders can be drawn under all labels."""
    leader_svg = ''
    if leader:
        ex, ey = nearest_on_box(ax, ay, box)
        leader_svg = (f'<line class="{group_class}" x1="{ax:.1f}" y1="{ay:.1f}" x2="{ex:.1f}" y2="{ey:.1f}" '
                      f'stroke="#8A8A85" stroke-width="{leader_width}"/>')
    out = [f'<g class="{group_class}">']
    y = box[1]
    for text, size, fill, cls in lines:
        cls_attr = f' class="{cls}"' if cls else ''
        out.append(f'<text{cls_attr} x="{box[0]:.1f}" y="{y + size * BASELINE:.1f}" font-size="{size}" '
                   f'fill="{fill}" stroke="{halo}" stroke-width="{size * 0.3:.2f}" stroke-linejoin="round" '
                   f'paint-order="stroke">{html_lib.escape(str(text), quote=False)}</text>')
        y += size * LINE_HEIGHT
    out.append('</g>')
    return leader_svg, ''.join(out)


# =====================================================================
# Part 1: geography-preserving schematic map (was generate_map_v2.py)
# =====================================================================

UNIT_PER_DEGREE = 10.0     # normalised-unit scale for Geography Lock / Max Move
CANVAS_SCALE = 60.0        # px per normalised unit, for the final SVG

TIER_STYLE = {
    'Small':  dict(radius_px=5.0, outline_px=1.2),
    'Medium': dict(radius_px=8.0, outline_px=1.8),
    'Large':  dict(radius_px=12.0, outline_px=2.6),
}

ROAD_WIDTH = {'Primary': 4.5, 'Secondary': 3.4, 'Connector': 2.5, 'Local': 2.0}

PALETTE = [
    "#7F77DD", "#1D9E75", "#D85A30", "#D4537E", "#639922", "#BA7517",
    "#993C1D", "#0F6E56", "#993556", "#3B6D11", "#854F0B", "#26215C",
    "#04342C", "#4A1B0C", "#72243E", "#27500A", "#633806", "#5F5E5A",
    "#712B13", "#4B1528", "#173404", "#412402", "#7A4E9E", "#1E6B6B",
    "#B5442E", "#8A6D3B", "#3E7C4A", "#9B2D5C", "#7C5A2E",
    "#8C3E3E", "#5E7C3E", "#6B3E8C", "#3E8C7C",
]
NAVY_BLUE = "#1B3A6B"
PINNED_COLOURS = {"A2": NAVY_BLUE, "A61": NAVY_BLUE, "A7": NAVY_BLUE}
LOCAL_ROAD_COLOUR = "#8A8A85"


# --------------------------------------------------------------- loading ---

def load_all(xlsx_path):
    wb = load_workbook(xlsx_path, data_only=True)

    # Junctions
    ws = wb['Junctions']
    header = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(header)}
    junctions = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        jid = row[0]
        if not jid:
            continue
        junctions[jid] = {
            'name': row[idx['Name']] if 'Name' in idx else row[1],
            'lat': row[idx['Latitude']],
            'lon': row[idx['Longitude']],
            'region': row[idx.get('Layout region', -1)] if 'Layout region' in idx else None,
            'geo_lock': (row[idx.get('Geography lock', -1)] if 'Geography lock' in idx else 0.78) or 0.78,
            'max_move': (row[idx.get('Max move', -1)] if 'Max move' in idx else 0.85) or 0.85,
            'tier': row[idx.get('Junction tier', -1)] if 'Junction tier' in idx else 'Medium',
            'jtype': row[idx.get('Junction type', -1)] if 'Junction type' in idx else 'Junction',
        }

    # Roads (canonical -> hierarchy/layout params)
    ws = wb['Roads']
    header = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(header)}
    roads = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        cid = row[idx.get('Canonical Road ID', 0)]
        if not cid:
            continue
        roads[cid] = {
            'hierarchy': row[idx.get('Hierarchy')] or 'Connector',
            'layout_priority': row[idx.get('Layout priority')] or 25,
            'bend_penalty': row[idx.get('Bend penalty')] or 2,
            'region_bend': row[idx.get('Region bend allowance')] or 6,
            'connector_km': row[idx.get('Short connector direct threshold (km)')] or 0,
        }

    # Segments
    ws = wb['Segments']
    header = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(header)}
    segments = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        fid, tid = row[idx['From ID']], row[idx['To ID']]
        if not fid or not tid or fid not in junctions or tid not in junctions:
            continue
        segments.append({
            'edge_id': row[idx.get('Edge ID')],
            'from': fid, 'to': tid,
            'road': row[idx['Road']],
            'road_section': row[idx.get('RoadSectionID')],
            'order': row[idx.get('OrderOnRoad')] or 0,
            'dist_km': row[idx.get('Distance (km)')],
            'hierarchy_override': row[idx.get('Hierarchy Override')] if 'Hierarchy Override' in idx else None,
        })

    return wb, junctions, roads, segments


def canonical_road_id(road):
    if not road:
        return "UNKNOWN"
    r = road.strip()
    if r.lower().startswith('lokale weg'):
        return 'LOKALE_WEG'
    base = r.split('(')[0].strip()
    return base if base else r


VALID_HIERARCHIES = {'Primary', 'Secondary', 'Connector', 'Local'}


def segment_hierarchy(seg, roads):
    """A segment's effective Hierarchy: its own 'Hierarchy Override' if set
    (a per-segment escape hatch for the odd short/local-feeling stretch of
    an otherwise Primary/Secondary road - e.g. a Kreuz-to-Kreuz hop that
    shouldn't be pulled onto the regional spine with the rest of the road),
    falling back to its road's default Hierarchy otherwise."""
    override = seg.get('hierarchy_override')
    if override in VALID_HIERARCHIES:
        return override
    cid = canonical_road_id(seg['road'])
    return roads.get(cid, {}).get('hierarchy', 'Connector')


# ------------------------------------------------------------ projection ---

def project_all(junctions):
    """Geographic projection only - kept completely separate from any
    later generated/adjusted position, per spec section 10."""
    lats = [j['lat'] for j in junctions.values() if j['lat'] is not None]
    lons = [j['lon'] for j in junctions.values() if j['lon'] is not None]
    centre_lat = sum(lats) / len(lats)
    centre_lon = sum(lons) / len(lons)
    cos_lat = math.cos(math.radians(centre_lat))
    geo_pos = {}
    for jid, j in junctions.items():
        if j['lat'] is None or j['lon'] is None:
            continue
        x = (j['lon'] - centre_lon) * cos_lat * UNIT_PER_DEGREE
        y = (centre_lat - j['lat']) * UNIT_PER_DEGREE  # north = up
        geo_pos[jid] = [x, y]
    return geo_pos, centre_lat, centre_lon, cos_lat


# ------------------------------------------------------- soft-spine solve ---

def build_road_sections(segments, roads, hierarchies):
    """Group segments into ordered chains per RoadSectionID, restricted to
    the given set of Hierarchy values (e.g. just ['Primary'])."""
    sections = defaultdict(list)
    for seg in segments:
        cid = canonical_road_id(seg['road'])
        hier = segment_hierarchy(seg, roads)
        if hier not in hierarchies:
            continue
        key = seg['road_section'] or cid
        sections[key].append(seg)
    for key in sections:
        sections[key].sort(key=lambda s: s['order'])
    return sections


def ordered_junction_chain(section_segs):
    """Best-effort linear list of junction IDs in OrderOnRoad sequence for
    one road section (already sorted by 'order')."""
    chain = []
    seen = set()
    for seg in section_segs:
        for jid in (seg['from'], seg['to']):
            if jid not in seen:
                chain.append(jid)
                seen.add(jid)
    return chain


def fit_line_direction(points):
    """Principal direction of a set of 2D points via simple covariance
    (equivalent to 1-component PCA) - returns a unit vector and the
    centroid. Falls back to the endpoint-to-endpoint vector for 2 points."""
    n = len(points)
    cx = sum(p[0] for p in points) / n
    cy = sum(p[1] for p in points) / n
    if n < 2:
        return (1.0, 0.0), (cx, cy)
    sxx = sum((p[0]-cx)**2 for p in points)
    syy = sum((p[1]-cy)**2 for p in points)
    sxy = sum((p[0]-cx)*(p[1]-cy) for p in points)
    # principal eigenvector of [[sxx,sxy],[sxy,syy]]
    theta = 0.5 * math.atan2(2*sxy, sxx - syy)
    dx, dy = math.cos(theta), math.sin(theta)
    return (dx, dy), (cx, cy)


def snap_angle_if_close(dx, dy, tolerance_deg=12):
    """Snap a direction to the nearest of 0/45/90/135 degrees only if it's
    already within tolerance - a soft preference, never forced."""
    ang = math.degrees(math.atan2(dy, dx)) % 180
    candidates = [0, 45, 90, 135]
    best = min(candidates, key=lambda c: min(abs(ang-c), abs(ang-c+180), abs(ang-c-180)))
    diff = min(abs(ang-best), abs(ang-best+180), abs(ang-best-180))
    if diff <= tolerance_deg:
        rad = math.radians(best)
        return math.cos(rad), math.sin(rad)
    return dx, dy


def apply_regional_spines(geo_pos, junctions, segments, roads, hierarchies, schematic_scale=1.0):
    """Sections 11-12: for each road section restricted to `hierarchies`,
    walk its ordered junction chain, split into contiguous same-Layout
    region runs, fit a soft regional line through each run's *geographic*
    positions, and nudge those junctions towards that line - reduced by
    each junction's own Geography Lock, capped by its Max Move, and
    further scaled by `schematic_scale` (used to give Secondary roads a
    weaker pull than Primary, per spec section 12)."""
    sections = build_road_sections(segments, roads, hierarchies)
    proposals = defaultdict(list)  # jid -> list of (delta_x, delta_y, layout_priority)

    for road_section, segs in sections.items():
        chain = ordered_junction_chain(segs)
        if len(chain) < 3:
            continue
        cid = canonical_road_id(segs[0]['road'])
        priority = roads.get(cid, {}).get('layout_priority', 25)
        connected_pairs = set()
        for s in segs:
            connected_pairs.add((s['from'], s['to']))
            connected_pairs.add((s['to'], s['from']))

        # split into contiguous same-region runs - also splits wherever the
        # naive chain jumps to a junction that isn't actually directly
        # connected to the previous one (a branching road section produces
        # more than one real path; without this check a branch's junctions
        # get appended after the main line and corrupt the line fit for
        # both, per the Het Vonderen/Roermond/Tiglia case)
        runs = []
        current = [chain[0]]
        for jid in chain[1:]:
            same_region = junctions[jid]['region'] == junctions[current[-1]]['region']
            actually_connected = (current[-1], jid) in connected_pairs
            if same_region and actually_connected:
                current.append(jid)
            else:
                runs.append(current)
                current = [jid]
        runs.append(current)

        for run in runs:
            if len(run) < 3:
                continue
            pts = [geo_pos[j] for j in run if j in geo_pos]
            if len(pts) < 3:
                continue
            (dx, dy), (cx, cy) = fit_line_direction(pts)
            dx, dy = snap_angle_if_close(dx, dy)
            for jid in run:
                if jid not in geo_pos:
                    continue
                px, py = geo_pos[jid]
                # project (px,py) onto the line through (cx,cy) direction (dx,dy)
                t = (px-cx)*dx + (py-cy)*dy
                target_x, target_y = cx + t*dx, cy + t*dy
                proposals[jid].append((target_x - px, target_y - py, priority))

    # combine competing proposals (weighted by layout priority) and apply
    # geography lock + max move
    new_pos = {jid: list(p) for jid, p in geo_pos.items()}
    moved_to_cap = []
    for jid, props in proposals.items():
        total_w = sum(pr for _, _, pr in props)
        if total_w <= 0:
            continue
        avg_dx = sum(dx*pr for dx, dy, pr in props) / total_w
        avg_dy = sum(dy*pr for dx, dy, pr in props) / total_w
        j = junctions[jid]
        schematic_strength = (1 - j['geo_lock']) * schematic_scale
        delta_x = avg_dx * schematic_strength
        delta_y = avg_dy * schematic_strength
        dist = math.hypot(delta_x, delta_y)
        max_move = j['max_move']
        if dist > max_move and dist > 0:
            delta_x *= max_move / dist
            delta_y *= max_move / dist
            moved_to_cap.append(jid)
        new_pos[jid][0] += delta_x
        new_pos[jid][1] += delta_y
    return new_pos, moved_to_cap


# ------------------------------------------------------- light relaxation ---

def light_relaxation_pass(pos, junctions, connected_pairs, min_dist=0.35, snap=0.06, passes=2):
    """Section 17: a light grid snap plus a couple of deterministic
    collision-avoidance passes - never the unrestricted force-directed
    layout the spec explicitly rules out. A proposed move is accepted only
    if it stays within the junction's own Max Move from its geographic
    start, and never applied to two junctions that are directly connected
    (that would distort real adjacency)."""
    start = {jid: list(p) for jid, p in pos.items()}
    ids = list(pos.keys())
    for _ in range(passes):
        moved_any = False
        for i in range(len(ids)):
            for k in range(i+1, len(ids)):
                a, b = ids[i], ids[k]
                if (a, b) in connected_pairs:
                    continue
                ax, ay = pos[a]; bx, by = pos[b]
                d = math.hypot(bx-ax, by-ay)
                if 0 < d < min_dist:
                    push = (min_dist - d) / 2
                    ux, uy = (bx-ax)/d, (by-ay)/d
                    for jid, sign in ((a, -1), (b, 1)):
                        nx = pos[jid][0] + sign*ux*push*0.5
                        ny = pos[jid][1] + sign*uy*push*0.5
                        sx, sy = start[jid]
                        if math.hypot(nx-sx, ny-sy) <= junctions[jid]['max_move']:
                            pos[jid][0], pos[jid][1] = nx, ny
                            moved_any = True
        if not moved_any:
            break
    # light grid snap, only if it doesn't exceed max move
    for jid in ids:
        sx, sy = start[jid]
        gx = round(pos[jid][0] / snap) * snap
        gy = round(pos[jid][1] / snap) * snap
        if math.hypot(gx-sx, gy-sy) <= junctions[jid]['max_move']:
            pos[jid][0], pos[jid][1] = gx, gy
    return pos


# ------------------------------------------------------------- polylines ---

def build_road_polylines(junctions, segments, pos, roads, tube_style_hierarchies=None, angles_deg=None):
    """Straight line between each segment's two (already spine-adjusted)
    junction positions by default - bends come entirely from junction
    movement, not from artificial angle-snapping of a line between two
    fixed points.

    If `tube_style_hierarchies` is given (e.g. {'Primary', 'Secondary'}),
    segments on a road of one of those hierarchies are instead routed with
    London-Underground-style angle snapping (an elbow bend onto the two
    nearest allowed angles) - Connector and Local roads always stay as
    direct lines regardless, since they're meant to read as flexible local
    links, not schematic corridors."""
    tube_style_hierarchies = tube_style_hierarchies or set()
    angles_deg = angles_deg or [0, 30, 45, 60, 90, 120, 135, 150, 180, 210, 225, 240, 270, 300, 315, 330]
    angle_units = {a: (math.cos(math.radians(a)), math.sin(math.radians(a))) for a in sorted(set(angles_deg))}
    angle_list = sorted(angle_units)

    def octilinear(p1, p2):
        x1, y1 = p1; x2, y2 = p2
        dx, dy = x2 - x1, y2 - y1
        if math.hypot(dx, dy) < 1e-6:
            return [p1, p2]
        bearing = math.degrees(math.atan2(dy, dx)) % 360
        n = len(angle_list)
        a1 = a2 = None
        for i in range(n):
            c1, c2 = angle_list[i], angle_list[(i + 1) % n]
            span = (c2 - c1) % 360
            rel = (bearing - c1) % 360
            if -1e-6 <= rel <= span + 1e-6:
                a1, a2 = c1, c2
                break
        v1, v2 = angle_units[a1], angle_units[a2]
        det = v1[0]*v2[1] - v2[0]*v1[1]
        if abs(det) < 1e-9:
            return [p1, p2]
        a = (dx*v2[1] - dy*v2[0]) / det
        b = (v1[0]*dy - v1[1]*dx) / det
        a, b = max(a, 0), max(b, 0)
        mx, my = x1 + a*v1[0], y1 + a*v1[1]
        if math.hypot(mx-x1, my-y1) < 1e-6 or math.hypot(x2-mx, y2-my) < 1e-6:
            return [p1, p2]
        return [p1, (mx, my), p2]

    polylines = {}
    for seg in segments:
        a, b = pos.get(seg['from']), pos.get(seg['to'])
        if a is None or b is None:
            continue
        hier = segment_hierarchy(seg, roads)
        if hier in tube_style_hierarchies:
            polylines[seg['edge_id']] = octilinear(tuple(a), tuple(b))
        else:
            polylines[seg['edge_id']] = [tuple(a), tuple(b)]
    return polylines


# --------------------------------------------------------- border layout ---

def build_border_polylines(wb, pos, junctions, river_pos=None):
    """Section 5/9: borders as an independent orthogonal (90-degree-only)
    network, drawn between Border Node positions (reusing a road Junction
    or River Junction position wherever the node is linked to one)."""
    river_pos = river_pos or {}
    if 'Border Nodes' not in wb.sheetnames or 'Border Segments' not in wb.sheetnames:
        return {}, {}
    wsn = wb['Border Nodes']
    header = [c.value for c in wsn[1]]
    idx = {h: i for i, h in enumerate(header)}
    node_pos = {}
    node_info = {}
    for row in wsn.iter_rows(min_row=2, values_only=True):
        bnid = row[idx['BorderNodeID']]
        if not bnid:
            continue
        linked_j = row[idx.get('LinkedJunctionID')]
        linked_r = row[idx.get('LinkedRiverJunctionID')]
        if linked_j and linked_j in pos:
            node_pos[bnid] = pos[linked_j]
        elif linked_r and linked_r in river_pos:
            node_pos[bnid] = river_pos[linked_r]
        else:
            # dedicated border-only node (e.g. a tripoint) - project its own lat/lon
            lat, lon = row[idx.get('Latitude')], row[idx.get('Longitude')]
            if lat is not None and lon is not None:
                node_pos[bnid] = None  # filled by caller once centre/scale known
        node_info[bnid] = {
            'name': row[idx['Name']], 'type': row[idx['BorderNodeType']],
            'lat': row[idx.get('Latitude')], 'lon': row[idx.get('Longitude')],
        }

    wss = wb['Border Segments']
    header = [c.value for c in wss[1]]
    idx = {h: i for i, h in enumerate(header)}
    segs = []
    for row in wss.iter_rows(min_row=2, values_only=True):
        fbn, tbn = row[idx['FromBorderNodeID']], row[idx['ToBorderNodeID']]
        if fbn and tbn:
            segs.append({'from': fbn, 'to': tbn, 'pair': row[idx['CountryPair']]})
    return node_pos, {'segments': segs, 'info': node_info}


def segments_intersect(p1, p2, p3, p4):
    def ccw(a, b, c):
        return (c[1]-a[1])*(b[0]-a[0]) > (b[1]-a[1])*(c[0]-a[0])
    return ccw(p1,p3,p4) != ccw(p2,p3,p4) and ccw(p1,p2,p3) != ccw(p1,p2,p4)


def path_crosses_roads(path, road_polylines, skip_near, tol=0.15):
    """True if any leg of `path` crosses any road polyline leg, away from
    the border segment's own two endpoints (a real crossing point is
    allowed to touch a road there - that's the whole point of a border
    crossing junction)."""
    for i in range(len(path) - 1):
        a, b = path[i], path[i+1]
        for poly in road_polylines:
            for j in range(len(poly) - 1):
                c, d = poly[j], poly[j+1]
                if segments_intersect(a, b, c, d):
                    # find rough intersection point (midpoint of the shared span)
                    mx, my = (a[0]+b[0])/2, (a[1]+b[1])/2
                    if all(math.hypot(mx-sx, my-sy) > tol for sx, sy in skip_near):
                        return True
    return False


def orthogonal_route(a, b, road_polylines=None, skip_near=None, max_tries=8):
    """Border lines only ever bend at right angles. Tries the two natural
    single-elbow corners first; if both illegally cross a road away from
    the segment's own endpoints, inserts a small LOCAL notch around the
    actual crossing point (not a full-length parallel shift of the whole
    segment, which would read as an unrelated extra line cutting across
    unrelated clusters) - grown only as far as needed to clear."""
    ax, ay = a; bx, by = b
    road_polylines = road_polylines or []
    skip_near = skip_near or [a, b]

    if abs(bx-ax) < 1e-9 or abs(by-ay) < 1e-9:
        return [a, b]

    # default: bend at the midpoint - a Z-shape with two 90-degree turns,
    # which reliably clears more obstacles than a single-elbow L-shape
    # (per project convention, this is the preferred default, not just a
    # fallback)
    mid_x, mid_y = ax + (bx-ax)/2, ay + (by-ay)/2
    z_vertical_first = [a, (ax, mid_y), (bx, mid_y), b]
    z_horizontal_first = [a, (mid_x, ay), (mid_x, by), b]
    for candidate in (z_vertical_first, z_horizontal_first):
        if not path_crosses_roads(candidate, road_polylines, skip_near):
            return candidate

    corner1 = (bx, ay)  # horizontal first, then vertical
    corner2 = (ax, by)  # vertical (south/north) first, then horizontal
    for corner in (corner2, corner1):
        candidate = [a, corner, b]
        if not path_crosses_roads(candidate, road_polylines, skip_near):
            return candidate

    # neither simple elbow is clear - insert a small local notch around the
    # actual crossing point (using corner1's path as the base to notch)
    base = [a, corner2, b]
    for attempt in range(1, max_tries + 1):
        notch = 0.10 * attempt
        # notch near the vertical leg (a -> corner1)
        mid_v = ((ax + corner1[0]) / 2, (ay + corner1[1]) / 2)
        for sign in (1, -1):
            nx = mid_v[0] + notch * sign
            candidate = [a, (nx, ay), (nx, corner1[1]), corner1, b]
            if not path_crosses_roads(candidate, road_polylines, skip_near):
                return candidate
        # notch near the horizontal leg (corner1 -> b)
        mid_h = ((corner1[0] + bx) / 2, (corner1[1] + by) / 2)
        for sign in (1, -1):
            ny = mid_h[1] + notch * sign
            candidate = [a, corner1, (corner1[0], ny), (bx, ny), b]
            if not path_crosses_roads(candidate, road_polylines, skip_near):
                return candidate
    return base


# --------------------------------------------------------- road colours ---

def road_key(road):
    return canonical_road_id(road)


def is_local_road(road):
    return road_key(road) == 'LOKALE_WEG'


def assign_road_colours(segments, palette):
    road_junctions = defaultdict(set)
    for seg in segments:
        key = road_key(seg['road'])
        road_junctions[key].update([seg['from'], seg['to']])
    conflicts = defaultdict(set)
    for jid in set().union(*road_junctions.values()) if road_junctions else set():
        touching = [k for k, js in road_junctions.items() if jid in js]
        for i in range(len(touching)):
            for j in range(i+1, len(touching)):
                conflicts[touching[i]].add(touching[j])
                conflicts[touching[j]].add(touching[i])
    colours = {}
    for key, colour in PINNED_COLOURS.items():
        if key in road_junctions:
            colours[key] = colour
    if 'LOKALE_WEG' in road_junctions:
        colours['LOKALE_WEG'] = LOCAL_ROAD_COLOUR
    remaining = sorted((k for k in road_junctions if k not in colours),
                       key=lambda k: -len(conflicts[k]))
    for key in remaining:
        used_nearby = {colours[n] for n in conflicts[key] if n in colours}
        for c in palette:
            if c == NAVY_BLUE:
                continue
            if c not in used_nearby:
                colours[key] = c
                break
        else:
            colours[key] = palette[len(colours) % len(palette)]
    return colours


# ------------------------------------------------------------ points ---

def load_points_v2(wb, junctions):
    if 'Points' not in wb.sheetnames:
        return []
    ws = wb['Points']
    header = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(header)}
    points = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        pid = row[idx.get('Point ID')]
        if not pid:
            continue
        eid = row[idx.get('Edge ID')]
        pos_on_edge = row[idx.get('PositionOnEdge')]
        order_on_seg = row[idx.get('OrderOnSegment')]
        if not eid or pos_on_edge is None:
            continue  # unmatched - already reported in Validation, skip silently here
        notes = row[idx.get('Notes')]
        length_m = None
        m = re.search(r'lengte:\s*([\d.,]+)\s*m', str(notes or ''))
        if m:
            try:
                length_m = float(m.group(1).replace(',', '.'))
            except ValueError:
                length_m = None
        points.append({
            'id': pid, 'name': row[idx.get('Name')], 'category': row[idx.get('Category')],
            'sides': row[idx.get('Sides (1 or 2)')] or 1,
            'edge_id': eid, 'pos_on_edge': pos_on_edge,
            'order_on_segment': order_on_seg or 0,
            'notes': notes, 'length_m': length_m,
            'lat': row[idx.get('Latitude')], 'lon': row[idx.get('Longitude')],
            'brand': row[idx.get('Fuel brand')] if 'Fuel brand' in idx else None,
            'facilities': row[idx.get('Facilities')] if 'Facilities' in idx else None,
        })
    return points


def real_side_sign(from_lat, from_lon, to_lat, to_lon, pt_lat, pt_lon):
    """Which side of the road a one-sided point is really on, as a
    left(+1)/right(-1) sign relative to travel from the segment's real
    From junction to its real To junction - computed once from true
    lat/lon (independent of anything the schematic layout does to the
    junction positions), so a station real drivers can only reach heading
    in one direction renders on that same relative side on the map, no
    matter how the schematic line happens to be angled."""
    cos_lat = math.cos(math.radians((from_lat + to_lat) / 2))
    fx, fy = from_lon * cos_lat, -from_lat
    tx, ty = to_lon * cos_lat, -to_lat
    px, py = pt_lon * cos_lat, -pt_lat
    road_dx, road_dy = tx - fx, ty - fy
    pt_dx, pt_dy = px - fx, py - fy
    cross = road_dx * pt_dy - road_dy * pt_dx
    return 1 if cross >= 0 else -1


def point_at_fraction_on_polyline(poly, frac):
    seg_lens = [math.hypot(poly[i+1][0]-poly[i][0], poly[i+1][1]-poly[i][1]) for i in range(len(poly)-1)]
    total = sum(seg_lens)
    if total <= 0:
        return poly[0], 0.0
    target = frac * total
    travelled = 0.0
    for i, leglen in enumerate(seg_lens):
        if travelled + leglen >= target or i == len(seg_lens)-1:
            t = 0.0 if leglen < 1e-9 else (target-travelled)/leglen
            ax, ay = poly[i]; bx, by = poly[i+1]
            x, y = ax + t*(bx-ax), ay + t*(by-ay)
            ang = math.degrees(math.atan2(by-ay, bx-ax))
            return (x, y), ang
        travelled += leglen
    return poly[-1], 0.0


# ------------------------------------------------------------- symbols ---

MARKER_COLOUR = {'tankstation': '#A83232', 'autohof': '#A83232',
                  'brug (dal)': '#B0B0B0', 'brug (rivier)': '#185FA5',
                  'tunnel': '#5F5E5A', 'ecoduct': '#3B6D11', 'poi': '#3B6D11'}
BAR_CATEGORIES = {'brug (dal)', 'brug (rivier)', 'tunnel', 'ecoduct'}


def estimate_text_width(text, font_size):
    return len(text or '') * font_size * 0.56


def boxes_overlap(a, b):
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def place_label(name, cx, cy, junction_r, occupied, font_size=9, close_dist=18):
    """Try a handful of candidate offsets around the junction (varying both
    angle and distance) and return the first that doesn't collide with an
    already-placed label. Labels are NOT all forced to the same distance -
    a busy spot may need to reach further out than a quiet one. If even the
    furthest candidate collides, fall back to the least-bad option and draw
    a thin leader line back to the junction so it's still clear which point
    the label belongs to."""
    w = estimate_text_width(name, font_size)
    h = font_size * 1.15
    candidates = []
    for dist in (junction_r + 6, junction_r + 14, junction_r + 24, junction_r + 38, junction_r + 56):
        for angle_deg in (0, 45, -45, 90, -90, 135, -135, 180):
            a = math.radians(angle_deg)
            lx = cx + math.cos(a) * dist
            ly = cy + math.sin(a) * dist
            box = (lx, ly - h, lx + w, ly)
            candidates.append((dist, box, lx, ly))

    for dist, box, lx, ly in candidates:
        if not any(boxes_overlap(box, o) for o in occupied):
            occupied.append(box)
            return lx, ly, dist > close_dist

    dist, box, lx, ly = candidates[0]
    occupied.append(box)
    return lx, ly, True


def real_world_side(junctions, fid, tid, lat, lon):
    """Which side of the road (in real geography) a point sits on: +1 or
    -1, via the sign of the cross product between the real From->To
    direction and the vector from From to the point. Used so a one-sided
    tankstation's marker points to its actual real-world side rather than
    a fixed default - otherwise roughly half of them end up mirrored."""
    a, b = junctions.get(fid), junctions.get(tid)
    if not a or not b or a.get('lat') is None or lat is None:
        return 1
    cos_lat = math.cos(math.radians((a['lat'] + b['lat']) / 2))
    ax, ay = a['lon'] * cos_lat, a['lat']
    bx, by = b['lon'] * cos_lat, b['lat']
    px, py = lon * cos_lat, lat
    cross = (bx - ax) * (py - ay) - (by - ay) * (px - ax)
    return 1 if cross >= 0 else -1


def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi, dl = math.radians(lat2-lat1), math.radians(lon2-lon1)
    a = math.sin(dphi/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2 * R * math.asin(math.sqrt(a))


def fuel_brand_badge(brand, facilities):
    """Short text badge combining the fuel brand with up to 2 recognisable
    amenities (fast food / coffee chains) found in the free-text Facilities
    field - a lightweight text stand-in for a brand logo, since embedding
    real logo images in a generated SVG isn't practical here."""
    parts = []
    if brand:
        parts.append(str(brand).strip())
    if facilities:
        text = str(facilities)
        known = ["McDonald's", 'Burger King', 'Starbucks', 'Subway', 'NORDSEE',
                 "Dallmayr", 'Segafredo', 'BrotZeit', 'Coffee Fellows']
        found = [k for k in known if k.lower() in text.lower()]
        parts.extend(found[:2])
    return ' · '.join(parts) if parts else None


def fuel_marker_svg(cx, cy, angle_deg, sides, colour, side=1, length=10, base=4.5):
    """Triangle (one side) or diamond (two sides) sticking out from the
    road. `length` controls how far it reaches away from the road;
    `base` controls its width along the road - kept independent so the
    marker can be stretched outward without also getting wider. `side`
    (+1/-1) flips which perpendicular direction a one-sided marker points,
    matching its real-world side of the road - ignored for two-sided
    (diamond) markers, which straddle the line either way."""
    a = math.radians(angle_deg); perp = a + math.pi/2
    px, py = math.cos(perp) * side, math.sin(perp) * side
    lx, ly = math.cos(a), math.sin(a)
    if sides and int(sides) >= 2:
        pts = [(cx+px*length, cy+py*length), (cx+lx*base, cy+ly*base),
               (cx-px*length, cy-py*length), (cx-lx*base, cy-ly*base)]
    else:
        pts = [(cx+px*length, cy+py*length), (cx-lx*base, cy-ly*base),
               (cx+lx*base, cy+ly*base)]
    s = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
    return f'<polygon points="{s}" fill="{colour}" stroke="white" stroke-width="0.5"/>'


def bar_svg(cx, cy, angle_deg, colour, length_m=None, bar_len=13):
    """Bridge/tunnel bar, perpendicular to the road. Width uses 4 length
    tiers so a short culvert and a kilometre-long viaduct read as visibly
    different weights, capped so it never overwhelms the road line."""
    if length_m is None:
        width = 3.0
    elif length_m < 150:
        width = 2.2
    elif length_m < 400:
        width = 3.2
    elif length_m < 800:
        width = 4.4
    else:
        width = 5.8
    a = math.radians(angle_deg) + math.pi/2
    dx, dy = math.cos(a)*bar_len/2, math.sin(a)*bar_len/2
    rx, ry = math.cos(math.radians(angle_deg))*width/2, math.sin(math.radians(angle_deg))*width/2
    pts = [(cx-dx-rx, cy-dy-ry), (cx+dx-rx, cy+dy-ry), (cx+dx+rx, cy+dy+ry), (cx-dx+rx, cy-dy+ry)]
    s = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
    return f'<polygon points="{s}" fill="{colour}" stroke="#2C2C2A" stroke-width="0.3"/>'


def junction_circle_svg(cx, cy, jtype, tier, colours_here=None):
    """Sections 4 and 8: tiered hollow circle for a normal Junction/Exit,
    or a thick split-colour circle for a BorderCrossing (tier still sets
    diameter), or a small neutral circle for a Tripoint (own fixed size,
    ignores tier entirely). Sized in fixed screen pixels - deliberately
    decoupled from the Geography Lock / Max Move normalised-unit scale,
    which operates at a completely different order of magnitude (a small
    fraction of a degree) than a legible on-screen circle needs to be."""
    if jtype == 'Tripoint':
        r = 3.0
        return (f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{r:.1f}" '
                f'fill="#EDEDE8" stroke="#3A3A38" stroke-width="0.6"/>')
    style = TIER_STYLE.get(tier, TIER_STYLE['Medium'])
    r = style['radius_px']
    ow = style['outline_px']
    if jtype == 'BorderCrossing':
        col_a, col_b = (colours_here or ('#999999', '#CCCCCC'))
        parts = [f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{r:.1f}" fill="white" stroke="#1A1A18" stroke-width="{ow*1.6:.1f}"/>']
        parts.append(f'<path d="M {cx:.1f} {cy-r:.1f} A {r:.1f} {r:.1f} 0 0 1 {cx:.1f} {cy+r:.1f} Z" fill="{col_a}"/>')
        parts.append(f'<path d="M {cx:.1f} {cy+r:.1f} A {r:.1f} {r:.1f} 0 0 1 {cx:.1f} {cy-r:.1f} Z" fill="{col_b}"/>')
        parts.append(f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{r:.1f}" fill="none" stroke="#1A1A18" stroke-width="{ow*1.6:.1f}"/>')
        return "\n".join(parts)
    return (f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{r:.1f}" '
            f'fill="white" stroke="#2C2C2A" stroke-width="{ow:.1f}"/>')


def border_crossing_colours(jid, segments, road_colours):
    """Pick the two colours for a split border-crossing circle: the (up to)
    two distinct road colours actually touching this junction, ranked by
    how many segments use them here (a reasonable proxy for 'the
    international continuation' without needing manual override data).
    Falls back to a lighter tint of the same colour if only one road
    touches the crossing."""
    counts = defaultdict(int)
    for seg in segments:
        if seg['from'] == jid or seg['to'] == jid:
            counts[road_key(seg['road'])] += 1
    ranked = sorted(counts, key=lambda k: -counts[k])
    cols = [road_colours.get(k, '#999999') for k in ranked[:2]]
    if len(cols) == 0:
        return ('#999999', '#CCCCCC'), True
    if len(cols) == 1:
        c = cols[0]
        # lighter tint of the same colour for the second half
        r = int(c[1:3], 16); g = int(c[3:5], 16); b = int(c[5:7], 16)
        r2 = min(255, r + (255-r)//2); g2 = min(255, g + (255-g)//2); b2 = min(255, b + (255-b)//2)
        return (c, f'#{r2:02X}{g2:02X}{b2:02X}'), True
    return (cols[0], cols[1]), (len(ranked) > 2)


# ------------------------------------------------------------- render ---

def resolve_road_crossings(segments, polylines, roads, max_passes=3):
    """Sections 14/spec-wide rule: a segment may only cross another via a
    shared junction. Where two unrelated segments' straight lines happen to
    cross on the schematic, nudge the LOWER-hierarchy one with a small
    perpendicular jog at the crossing point - Primary/Secondary stay
    perfectly straight (their importance means other roads move out of
    their way, never the reverse), Connector/Local may bend, matching the
    hierarchy already granted them elsewhere in the spec."""
    hier_rank = {'Primary': 3, 'Secondary': 2, 'Connector': 1, 'Local': 0}
    seg_by_eid = {s['edge_id']: s for s in segments if s['edge_id'] in polylines}

    def hier_of(eid):
        s = seg_by_eid[eid]
        return segment_hierarchy(s, roads)

    for _ in range(max_passes):
        eids = list(polylines.keys())
        moved_any = False
        for i in range(len(eids)):
            e1 = eids[i]
            s1 = seg_by_eid[e1]
            for k in range(i+1, len(eids)):
                e2 = eids[k]
                s2 = seg_by_eid[e2]
                if {s1['from'], s1['to']} & {s2['from'], s2['to']}:
                    continue  # shares a real junction - a legitimate meeting point
                p1, p2 = polylines[e1], polylines[e2]
                crossing_pt = None
                for a in range(len(p1)-1):
                    for b in range(len(p2)-1):
                        if segments_intersect(p1[a], p1[a+1], p2[b], p2[b+1]):
                            ax, ay = p1[a]; bx, by = p1[a+1]
                            crossing_pt = ((ax+bx)/2, (ay+by)/2)
                            break
                    if crossing_pt:
                        break
                if not crossing_pt:
                    continue
                h1, h2 = hier_of(e1), hier_of(e2)
                lower_eid, lower_poly = (e2, p2) if hier_rank[h2] <= hier_rank[h1] else (e1, p1)
                # jog the lower-hierarchy line's nearest vertex pair with a
                # small perpendicular offset at the crossing point
                lx, ly = crossing_pt
                a0, b0 = lower_poly[0], lower_poly[-1]
                dx, dy = b0[0]-a0[0], b0[1]-a0[1]
                length = math.hypot(dx, dy) or 1
                perp = (-dy/length, dx/length)
                offset = 0.12
                new_pt = (lx + perp[0]*offset, ly + perp[1]*offset)
                polylines[lower_eid] = [a0, new_pt, b0]
                moved_any = True
        if not moved_any:
            break
    return polylines
def sample_quadratic_bezier(p1, control, p2, n=14):
    pts = []
    for i in range(n + 1):
        t = i / n
        mt = 1 - t
        x = mt*mt*p1[0] + 2*mt*t*control[0] + t*t*p2[0]
        y = mt*mt*p1[1] + 2*mt*t*control[1] + t*t*p2[1]
        pts.append((x, y))
    return pts


def river_bow_control(a, b, road_polylines, max_tries=12):
    """A river never runs exactly straight (real rivers curve) and must
    never illegally cross a road except at its own two endpoints. Starts
    with a mild default bow for the organic look; if that crosses a road,
    grows the bow (alternating sides) until clear."""
    dx, dy = b[0]-a[0], b[1]-a[1]
    length = math.hypot(dx, dy) or 1
    perp = (-dy/length, dx/length)
    mx, my = (a[0]+b[0])/2, (a[1]+b[1])/2
    base_bow = max(0.08, length * 0.08)
    skip_near = [a, b]
    for sign in (1, -1):
        control = (mx + perp[0]*base_bow*sign, my + perp[1]*base_bow*sign)
        if not path_crosses_roads(sample_quadratic_bezier(a, control, b), road_polylines, skip_near, tol=0.1):
            return control
    for attempt in range(1, max_tries+1):
        for sign in (1, -1):
            bow = base_bow + 0.1*attempt
            control = (mx + perp[0]*bow*sign, my + perp[1]*bow*sign)
            if not path_crosses_roads(sample_quadratic_bezier(a, control, b), road_polylines, skip_near, tol=0.1):
                return control
    return (mx + perp[0]*base_bow, my + perp[1]*base_bow)


def render_map(junctions, segments, pos, roads, points, border_data, border_node_pos,
                river_data=None, ambiguous_crossings=None, tube_style_hierarchies=None, pois=None):
    """Assembles the SVG following the render order in spec section 16:
    background, borders, rivers, local -> connector -> secondary -> primary
    roads, road points, normal junctions/exits, border crossings,
    tripoints, labels."""
    ambiguous_crossings = ambiguous_crossings or []
    xs = [p[0] for p in pos.values()] + [p[0] for p in border_node_pos.values() if p]
    ys = [p[1] for p in pos.values()] + [p[1] for p in border_node_pos.values() if p]
    margin = 2.5
    minx, miny = min(xs)-margin, min(ys)-margin
    maxx, maxy = max(xs)+margin, max(ys)+margin

    def X(x): return (x - minx) * CANVAS_SCALE
    def Y(y): return (y - miny) * CANVAS_SCALE

    width = (maxx-minx) * CANVAS_SCALE
    height = (maxy-miny) * CANVAS_SCALE

    road_colours = assign_road_colours(segments, PALETTE)
    polylines = build_road_polylines(junctions, segments, pos, roads, tube_style_hierarchies=tube_style_hierarchies)
    polylines = resolve_road_crossings(segments, polylines, roads)

    parts = [f'<svg width="{width:.0f}" height="{height:.0f}" viewBox="0 0 {width:.0f} {height:.0f}" '
             f'xmlns="http://www.w3.org/2000/svg" font-family="Arial, sans-serif">',
             f'<rect x="0" y="0" width="{width:.0f}" height="{height:.0f}" fill="#FAFAF7"/>']

    # Compute river paths first (avoiding roads, which are already final),
    # then compute borders LAST of all three - now that both the road and
    # river geometry are fully known, borders can properly dodge everything
    # except their own crossing-point endpoints. Both are still inserted
    # into the SVG near the start, so they render underneath roads/points.
    river_svg_paths = []
    river_sampled_polys = []
    road_polylines_list = list(polylines.values())
    if river_data:
        for seg in river_data.get('segments', []):
            a, b = river_data['pos'].get(seg['from']), river_data['pos'].get(seg['to'])
            if not a or not b:
                continue
            control = river_bow_control(a, b, road_polylines_list)
            sampled = sample_quadratic_bezier(a, control, b)
            river_sampled_polys.append(sampled)
            river_svg_paths.append(
                f'<path d="M {X(a[0]):.1f} {Y(a[1]):.1f} Q {X(control[0]):.1f} {Y(control[1]):.1f} '
                f'{X(b[0]):.1f} {Y(b[1]):.1f}" stroke="#7FB8E0" stroke-width="3.5" '
                f'fill="none" stroke-linecap="round" opacity="0.85"/>')

    avoid_polys = road_polylines_list + river_sampled_polys
    border_svg_paths = []
    tripoint_ids = {bnid for bnid, info in border_data.get('info', {}).items() if info['type'] == 'Tripoint'}

    # Group every segment touching a tripoint by that tripoint, so all of
    # its 2-3 lines can be assigned distinct first directions (up/down/
    # left/right) together - assigning them independently risks two picking
    # the same axis and overlapping until they diverge further away, which
    # looks like the tripoint isn't really a 3-way meeting point.
    by_tripoint = defaultdict(list)
    other_segs = []
    for seg in border_data.get('segments', []):
        a, b = border_node_pos.get(seg['from']), border_node_pos.get(seg['to'])
        if not a or not b:
            continue
        if seg['from'] in tripoint_ids:
            by_tripoint[seg['from']].append((seg, a, b, True))
        elif seg['to'] in tripoint_ids:
            by_tripoint[seg['to']].append((seg, b, a, False))
        else:
            other_segs.append((seg, a, b))

    def cardinal_distance(bearing_deg, cardinal):
        target = {'E': 0, 'S': 90, 'W': 180, 'N': -90}[cardinal]
        diff = abs((bearing_deg - target + 180) % 360 - 180)
        return diff

    def ranked_cardinals(origin, target):
        dx, dy = target[0]-origin[0], target[1]-origin[1]
        bearing = math.degrees(math.atan2(dy, dx))
        return sorted(['E', 'S', 'W', 'N'], key=lambda c: cardinal_distance(bearing, c))

    tripoint_routes = {}  # id(seg) -> (route, origin)
    for tp_id, entries in by_tripoint.items():
        # rank every entry by how confident its own best cardinal match is
        # (smallest angular distance to its nearest cardinal), then assign
        # directions most-confident-first - a segment that's nearly due
        # south should get South before a segment that's more ambiguous
        # between two directions has to fight it out
        ranked_entries = []
        for seg, origin, target, from_is_tp in entries:
            prefs = ranked_cardinals(origin, target)
            best_dist = cardinal_distance(
                math.degrees(math.atan2(target[1]-origin[1], target[0]-origin[0])), prefs[0])
            ranked_entries.append((best_dist, seg, origin, target, prefs))
        ranked_entries.sort(key=lambda r: r[0])

        used_dirs = set()
        for _dist, seg, origin, target, prefs in ranked_entries:
            direction = next((c for c in prefs if c not in used_dirs), prefs[0])
            used_dirs.add(direction)
            ox, oy = origin
            tx, ty = target
            elbow = (tx, oy) if direction in ('E', 'W') else (ox, ty)
            candidate = [origin, elbow, target]
            if path_crosses_roads(candidate, avoid_polys, [origin, target]):
                candidate = orthogonal_route(origin, target, road_polylines=avoid_polys, skip_near=[origin, target])
            tripoint_routes[id(seg)] = (candidate, origin)

    for seg in border_data.get('segments', []):
        a, b = border_node_pos.get(seg['from']), border_node_pos.get(seg['to'])
        if not a or not b:
            continue
        if id(seg) in tripoint_routes:
            route, origin = tripoint_routes[id(seg)]
            if origin != a:
                route = list(reversed(route))
        else:
            route = orthogonal_route(a, b, road_polylines=avoid_polys, skip_near=[a, b])
        d = "M " + " L ".join(f"{X(x):.1f} {Y(y):.1f}" for x, y in route)
        border_svg_paths.append(f'<path d="{d}" stroke="#3A3A38" stroke-width="2" fill="none" '
                                 f'stroke-dasharray="6,4" stroke-linecap="round" opacity="0.7"/>')

    # 2. country borders (background, computed last, drawn first)
    parts.extend(border_svg_paths)
    # 3. rivers
    parts.extend(river_svg_paths)

    # --- label pre-pass ---------------------------------------------------
    # Junction labels are placed first (Large tier first, as before), then
    # every bridge/tunnel/fuel-station label is fitted around them by
    # LabelPlacer. Everything is in SVG units and the whole SVG scales
    # together, so a layout without overlaps here has none at any zoom.
    placer = LabelPlacer()
    for poly in polylines.values():
        placer.add_polyline([(X(x), Y(y)) for x, y in poly])
    for poly in river_sampled_polys:
        placer.add_polyline([(X(x), Y(y)) for x, y in poly])

    # point-of-interest stars first, so junction names are placed around them
    occupied_label_boxes = []
    poi_positions = []
    for poi in pois or []:
        poly = polylines.get(poi['edge_id'])
        if not poly:
            continue
        (px, py), ang = point_at_fraction_on_polyline(poly, poi['frac'])
        sx, sy, r = X(px), Y(py), MAP_POI_R
        poi_positions.append((poi, sx, sy, ang))
        occupied_label_boxes.append((sx - r, sy - r, sx + r, sy + r))

    tier_order = {'Large': 0, 'Medium': 1, 'Small': 2}
    normal_junctions = [(jid, j) for jid, j in junctions.items()
                        if jid in pos and j['jtype'] not in ('BorderCrossing', 'Tripoint')]
    normal_junctions.sort(key=lambda kv: tier_order.get(kv[1]['tier'], 1))
    border_junctions = [(jid, j) for jid, j in junctions.items()
                        if jid in pos and j['jtype'] == 'BorderCrossing']
    junction_labels = {}
    for jid, j in normal_junctions + border_junctions:
        sx, sy = X(pos[jid][0]), Y(pos[jid][1])
        r = TIER_STYLE.get(j['tier'], TIER_STYLE['Medium'])['radius_px']
        junction_labels[jid] = place_label(j['name'], sx, sy, r, occupied_label_boxes)
        placer.add_box((sx - r, sy - r, sx + r, sy + r))
        # the label as drawn (font-size 9, baseline at ly), with the same
        # conservative width estimate the point labels use
        lx, ly, _leader = junction_labels[jid]
        placer.add_box((lx, ly - 9 * 0.95, lx + text_width(j['name'], 9), ly + 9 * 0.25), text=True)

    via_by_edge = defaultdict(list)
    for p in points:
        via_by_edge[p['edge_id']].append(p)
    for eid in via_by_edge:
        via_by_edge[eid].sort(key=lambda p: (p['order_on_segment'], p['pos_on_edge']))

    # every point's marker position/side, and its marker as an obstacle
    point_marks = []
    for seg in segments:
        poly = polylines.get(seg['edge_id'])
        if not poly:
            continue
        fj, tj = junctions.get(seg['from']), junctions.get(seg['to'])
        for p in via_by_edge.get(seg['edge_id'], []):
            (px, py), ang = point_at_fraction_on_polyline(poly, p['pos_on_edge'])
            sx, sy = X(px), Y(py)
            cat = (p['category'] or '').lower()
            side = 1
            if cat in ('tankstation', 'autohof') and p['sides'] == 1 and p.get('lat') is not None \
                    and p.get('lon') is not None and fj and tj \
                    and fj['lat'] is not None and tj['lat'] is not None:
                side = real_side_sign(fj['lat'], fj['lon'], tj['lat'], tj['lon'], p['lat'], p['lon'])
            box, reach = marker_extent(sx, sy, ang, cat, p['sides'], side, p.get('length_m'))
            if box:
                placer.add_box(box)
            point_marks.append(dict(p=p, sx=sx, sy=sy, ang=ang, cat=cat, side=side,
                                    box=box, reach=reach))

    # points of interest: a star on their segment's line (drawn on top of
    # the roads), labelled first so they get the closest spots
    poi_star_svg = []
    for poi, sx, sy, ang in poi_positions:
        r = MAP_POI_R
        box = (sx - r, sy - r, sx + r, sy + r)
        placer.add_box(box)
        point_marks.append(dict(p=dict(name=poi['name']), sx=sx, sy=sy, ang=ang, cat='poi', side=1,
                                box=box, reach=r + 0.8))
        poi_star_svg.append(f'<g class="poi"><title>{poi_title(poi)}</title>{star_svg(sx, sy, r, "poi-star")}</g>')

    point_label_svg = []
    point_leader_svg = []
    for pm in sorted(point_marks, key=lambda pm: (LABEL_PRIORITY.get(pm['cat'], 9),
                                                  -len(pm['p']['name'] or ''))):
        p, cat = pm['p'], pm['cat']
        if not p['name']:
            continue
        lines = [(p['name'], POINT_FONT, '#4A4A47', '')]
        if cat == 'poi':
            group_class = 'poi'
            lines = [(p['name'], MAP_POI_FONT, POI_TEXT, 'poi-name')]
        elif cat in BAR_CATEGORIES:
            group_class = 'zoom-point-label bridge-label'
            if p.get('length_m'):
                lines.append((f"{p['length_m']:.0f}m", POINT_SUB_FONT, '#7A7A76', 'bridge-length-label'))
        else:
            group_class = 'zoom-point-label'
            if cat in ('tankstation', 'autohof'):
                badge = fuel_brand_badge(p.get('brand'), p.get('facilities'))
                if badge:
                    lines.append((badge, POINT_SUB_FONT, '#6B6B67', ''))
        w = max(text_width(t, size) for t, size, _c, _k in lines)
        h = sum(size * LINE_HEIGHT for _t, size, _c, _k in lines)
        # one-sided fuel stations: label on the side the marker points to
        pref = pm['side'] if cat in ('tankstation', 'autohof') else 1
        box, leader = placer.place(pm['sx'], pm['sy'], w, h, pm['ang'] + 90, pref_side=pref,
                                   clear=pm['reach'], ignore=pm['box'])
        leader_svg, text_svg = label_block_svg(group_class, lines, box, pm['sx'], pm['sy'], leader,
                                               halo='#FAFAF7', leader_width=0.35)
        point_leader_svg.append(leader_svg)
        point_label_svg.append(text_svg)

    # 4-7. roads, local -> connector -> secondary -> primary (so primary
    # draws on top), each with a white casing underneath for separation
    order = ['Local', 'Connector', 'Secondary', 'Primary']
    segs_by_hier = defaultdict(list)
    for seg in segments:
        hier = segment_hierarchy(seg, roads)
        segs_by_hier[hier].append(seg)

    for hier in order:
        for seg in segs_by_hier.get(hier, []):
            poly = polylines.get(seg['edge_id'])
            if not poly:
                continue
            screen_poly = [(X(x), Y(y)) for x, y in poly]
            colour = LOCAL_ROAD_COLOUR if is_local_road(seg['road']) else road_colours.get(road_key(seg['road']), '#888')
            width_px = ROAD_WIDTH[hier]
            dash = ' stroke-dasharray="3,5"' if is_local_road(seg['road']) else ''

            # 8. road points on this edge, drawn UNDER the road line so the
            # road overlays them (per user preference) - only the part that
            # sticks out past the road's casing width remains visible
            fj, tj = junctions.get(seg['from']), junctions.get(seg['to'])
            for p in via_by_edge.get(seg['edge_id'], []):
                (px, py), ang = point_at_fraction_on_polyline(poly, p['pos_on_edge'])
                sx, sy = X(px), Y(py)
                cat = (p['category'] or '').lower()
                if cat in BAR_CATEGORIES:
                    parts.append(bar_svg(sx, sy, ang, MARKER_COLOUR.get(cat, '#999'), length_m=p.get('length_m')))
                elif cat in ('tankstation', 'autohof'):
                    side = 1
                    if p['sides'] == 1 and p.get('lat') is not None and p.get('lon') is not None:
                        if fj and tj and fj['lat'] is not None and tj['lat'] is not None:
                            side = real_side_sign(fj['lat'], fj['lon'], tj['lat'], tj['lon'], p['lat'], p['lon'])
                    parts.append(fuel_marker_svg(sx, sy, ang, p['sides'], MARKER_COLOUR.get(cat, '#A83232'), side=side))
                # distance from this segment's From junction to this point,
                # in real km (haversine on true lat/lon) - own colour/toggle,
                # independent of the point-name labels above
                if fj and fj.get('lat') is not None and p.get('lat') is not None:
                    d_km = haversine_km(fj['lat'], fj['lon'], p['lat'], p['lon'])
                    parts.append(f'<text class="dist-label dist-point-label" x="{sx-6:.1f}" y="{sy-8:.1f}" '
                                 f'font-size="6" fill="#B5442E" text-anchor="end">{d_km:.1f}km</text>')
                # debug: this point's Point ID, only shown with debug mode on
                parts.append(f'<text class="debug-label" x="{sx+6:.1f}" y="{sy+8:.1f}" '
                             f'font-size="6" fill="#C0392B">{p["id"]}</text>')

            d = "M " + " L ".join(f"{x:.1f} {y:.1f}" for x, y in screen_poly)
            parts.append(f'<path d="{d}" stroke="{colour}" stroke-width="{width_px:.1f}" fill="none"{dash} '
                         f'stroke-linecap="round" opacity="0.95"/>')
            # segment's total real distance, own colour/toggle
            if seg.get('dist_km') is not None and len(screen_poly) >= 2:
                midx = sum(x for x, y in screen_poly) / len(screen_poly)
                midy = sum(y for x, y in screen_poly) / len(screen_poly)
                parts.append(f'<text class="dist-label dist-segment-label" x="{midx:.1f}" y="{midy-6:.1f}" '
                             f'font-size="7" fill="#185FA5" font-weight="bold" text-anchor="middle">{seg["dist_km"]:.0f}km</text>')
            # debug: this segment's Edge ID, placed at its midpoint
            if len(screen_poly) >= 2:
                midx = sum(x for x, y in screen_poly) / len(screen_poly)
                midy = sum(y for x, y in screen_poly) / len(screen_poly)
                parts.append(f'<text class="debug-label" x="{midx:.1f}" y="{midy:.1f}" '
                             f'font-size="7" fill="#1A6B3C" font-weight="bold">{seg["edge_id"]}</text>')

    # 9-11. junctions/exits, then border crossings, then tripoints (on top)
    # labels placed in tier order (Large first) so the most important names
    # get first pick of a clear spot; a leader line is drawn whenever a
    # label had to move further than a "close" distance to avoid a collision
    for jid, j in normal_junctions:
        x, y = pos[jid]
        sx, sy = X(x), Y(y)
        parts.append(junction_circle_svg(sx, sy, j['jtype'], j['tier']))
        lx, ly, needs_leader = junction_labels[jid]
        if needs_leader:
            parts.append(f'<line x1="{sx:.1f}" y1="{sy:.1f}" x2="{lx:.1f}" y2="{ly:.1f}" '
                         f'stroke="#999" stroke-width="0.6" stroke-dasharray="2,2"/>')
        parts.append(f'<text x="{lx:.1f}" y="{ly:.1f}" font-size="9" fill="#2C2C2A">{j["name"]}</text>')

    for jid, j in junctions.items():
        if jid not in pos or j['jtype'] != 'BorderCrossing':
            continue
        x, y = pos[jid]
        sx, sy = X(x), Y(y)
        (col_a, col_b), was_ambiguous = border_crossing_colours(jid, segments, road_colours)
        if was_ambiguous:
            ambiguous_crossings.append(jid)
        parts.append(junction_circle_svg(sx, sy, 'BorderCrossing', j['tier'], (col_a, col_b)))
        lx, ly, needs_leader = junction_labels[jid]
        if needs_leader:
            parts.append(f'<line x1="{sx:.1f}" y1="{sy:.1f}" x2="{lx:.1f}" y2="{ly:.1f}" '
                         f'stroke="#999" stroke-width="0.6" stroke-dasharray="2,2"/>')
        parts.append(f'<text x="{lx:.1f}" y="{ly:.1f}" font-size="9" fill="#2C2C2A">{j["name"]}</text>')

    for jid, j in junctions.items():
        if jid not in pos or j['jtype'] != 'Tripoint':
            continue
        x, y = pos[jid]
        parts.append(junction_circle_svg(X(x), Y(y), 'Tripoint', j['tier']))

    # point labels on top of everything else (their halo keeps them
    # readable where they have to cross a line)
    parts.extend(poi_star_svg)
    parts.extend(point_leader_svg)
    parts.extend(point_label_svg)

    # export every segment's real endpoints + final schematic screen
    # polyline, so client-side JS can match a live GPS fix to the nearest
    # real road and place a marker at the equivalent point on the drawn
    # (schematic) line - this is the only data GPS-locate needs
    segments_geo = []
    for seg in segments:
        poly = polylines.get(seg['edge_id'])
        fj, tj = junctions.get(seg['from']), junctions.get(seg['to'])
        if not poly or not fj or not tj or fj.get('lat') is None or tj.get('lat') is None:
            continue
        segments_geo.append({
            'eid': seg['edge_id'],
            'flat': fj['lat'], 'flon': fj['lon'], 'tlat': tj['lat'], 'tlon': tj['lon'],
            'poly': [[round(X(x), 1), round(Y(y), 1)] for x, y in poly],
        })

    parts.append('</svg>')
    return "\n".join(parts), road_colours, width, height, segments_geo, placer.stats


# ------------------------------------------------------------- generate ---

def build_map(xlsx_path=DEFAULT_XLSX, title="Geographic Spine Map", tube_style_hierarchies=None):
    """Builds the geography-preserving map page. Returns (html, stats)."""
    if tube_style_hierarchies is None:
        tube_style_hierarchies = {'Primary', 'Secondary'}  # tube-style is now the default
    wb, junctions, roads, segments = load_all(xlsx_path)
    geo_pos, centre_lat, centre_lon, cos_lat = project_all(junctions)

    pos, moved_primary = apply_regional_spines(geo_pos, junctions, segments, roads, ['Primary'], schematic_scale=1.0)
    pos, moved_secondary = apply_regional_spines(pos, junctions, segments, roads, ['Secondary'], schematic_scale=0.5)

    connected_pairs = set()
    for seg in segments:
        connected_pairs.add((seg['from'], seg['to']))
        connected_pairs.add((seg['to'], seg['from']))
    pos = light_relaxation_pass(pos, junctions, connected_pairs)

    points = load_points_v2(wb, junctions)
    pois = load_pois(wb, junctions, {s['edge_id']: (s['from'], s['to']) for s in segments})

    # river positions: project River Junctions geographically the same way,
    # snap bridge-point positions from the already-built road polylines
    river_pos = {}
    river_segs = []
    if 'River Junctions' in wb.sheetnames and 'River Segments' in wb.sheetnames:
        wsrj = wb['River Junctions']
        rj_header = [c.value for c in wsrj[1]]
        ridx = {h: i for i, h in enumerate(rj_header)}
        for row in wsrj.iter_rows(min_row=2, values_only=True):
            rid = row[ridx.get('River Junction ID', 0)]
            lat, lon = row[ridx.get('Latitude')], row[ridx.get('Longitude')]
            if rid and lat is not None:
                x = (lon - centre_lon) * cos_lat * UNIT_PER_DEGREE
                y = (centre_lat - lat) * UNIT_PER_DEGREE
                river_pos[rid] = (x, y)
        polylines_tmp = build_road_polylines(junctions, segments, pos, roads)
        pts_by_edge = defaultdict(list)
        for p in points:
            pts_by_edge[p['edge_id']].append(p)
        point_screen_pos = {}
        for eid, poly in polylines_tmp.items():
            for p in pts_by_edge.get(eid, []):
                (x, y), _ang = point_at_fraction_on_polyline(poly, p['pos_on_edge'])
                point_screen_pos[p['id']] = (x, y)
        for pid, (x, y) in point_screen_pos.items():
            river_pos[pid] = (x, y)

        wsrs = wb['River Segments']
        rs_header = [c.value for c in wsrs[1]]
        sidx = {h: i for i, h in enumerate(rs_header)}
        for row in wsrs.iter_rows(min_row=2, values_only=True):
            fid, tid = row[sidx.get('From ID', 0)], row[sidx.get('To ID', 2)]
            if fid and tid:
                river_segs.append({'from': fid, 'to': tid})

    border_node_pos, border_data = build_border_polylines(wb, pos, junctions, river_pos)
    # fill any still-unresolved dedicated border-only nodes (tripoints) from their own lat/lon
    if border_data:
        for bnid, info in border_data['info'].items():
            if border_node_pos.get(bnid) is None and info['lat'] is not None:
                x = (info['lon'] - centre_lon) * cos_lat * UNIT_PER_DEGREE
                y = (centre_lat - info['lat']) * UNIT_PER_DEGREE
                border_node_pos[bnid] = (x, y)

    ambiguous = []
    svg, road_colours, width, height, segments_geo, label_stats = render_map(
        junctions, segments, pos, roads, points, border_data or {}, border_node_pos,
        river_data={'segments': river_segs, 'pos': river_pos} if river_segs else None,
        ambiguous_crossings=ambiguous, tube_style_hierarchies=tube_style_hierarchies, pois=pois,
    )
    segments_geo_json = json.dumps(segments_geo)

    html = f'''<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>{title}</title>
<style>
  body {{ margin:0; background:#f0efe9; overflow:hidden; }}
  #wrap {{ width:100vw; height:100vh; overflow:hidden; cursor:grab; touch-action:none; }}
  #wrap.grabbing {{ cursor:grabbing; }}
  #stage {{ transform-origin: 0 0; }}
  #controls {{ position:fixed; top:12px; right:12px; z-index:10; display:flex; flex-direction:column; gap:6px; }}
  #controls button {{ width:36px; height:36px; font-size:20px; border:1px solid #999; background:white;
                       border-radius:6px; cursor:pointer; box-shadow:0 1px 3px rgba(0,0,0,0.2); }}
  #controls button:active {{ background:#eee; }}
  #controls button.active {{ background:#2C6BD1; color:white; }}
  #gpsMarker {{ display:none; }}
  /* point (tankstation/bridge) names only appear once zoomed in far enough -
     toggled by JS adding/removing 'labels-visible' on #stage */
  .zoom-point-label {{ display:none; }}
  #stage.labels-visible .zoom-point-label {{ display:block; }}
  /* bridge/tunnel names have their own 3-state toggle (off/on/on+length),
     independent of the zoom-based fuel-station names above */
  #stage.labels-visible.bridges-off .bridge-label {{ display:none; }}
  #stage.labels-visible.bridges-off .bridge-length-label {{ display:none; }}
  #stage.labels-visible.bridges-on .bridge-length-label {{ display:none; }}
  .debug-label {{ display:none; }}
  #stage.debug-mode .debug-label {{ display:block; }}
  /* segment/point distances - own zoom-gated toggle, off by default */
  .dist-label {{ display:none; }}
  #stage.labels-visible.distances-on .dist-label {{ display:block; }}
  /* points of interest: own toggle, visible at every zoom level */
  .poi {{ display:none; }}
  #stage.pois-on .poi {{ display:block; }}
  .poi-name {{ font-weight:bold; }}
</style>
</head><body>
<div id="controls">
  <button id="zoomIn" title="Inzoomen">+</button>
  <button id="zoomOut" title="Uitzoomen">&minus;</button>
  <button id="zoomReset" title="Reset">&#8634;</button>
  <button id="debugToggle" title="Debug aan/uit">D</button>
  <button id="bridgeToggle" title="Brug/tunnel namen">B</button>
  <button id="distToggle" title="Afstanden aan/uit">Km</button>
  <button id="poiToggle" class="active" title="Bezienswaardigheden aan/uit">&#9733;</button>
  <button id="gpsToggle" title="Mijn locatie volgen">&#128205;</button>
</div>
<div id="wrap"><div id="stage" class="pois-on">{svg}</div></div>
<script>
(function() {{
  const wrap = document.getElementById('wrap');
  const stage = document.getElementById('stage');
  const LABEL_ZOOM_THRESHOLD = 2.5;  // scale at which point names appear
  const MAP_MAX_ZOOM = {MAP_MAX_ZOOM};
  let scale = 1, panX = 0, panY = 0;
  let dragging = false, lastX = 0, lastY = 0;
  // bridge/tunnel name toggle: 0 = off, 1 = on (name only), 2 = on + length
  const bridgeStates = ['bridges-off', 'bridges-on', 'bridges-full'];
  let bridgeState = 2;

  function applyBridgeState() {{
    stage.classList.remove('bridges-off', 'bridges-on', 'bridges-full');
    stage.classList.add(bridgeStates[bridgeState]);
    const btn = document.getElementById('bridgeToggle');
    btn.classList.toggle('active', bridgeState !== 0);
    btn.title = ['Brugnamen: uit', 'Brugnamen: aan', 'Brugnamen: aan + lengte'][bridgeState];
  }}

  function updateLabelVisibility() {{
    stage.classList.toggle('labels-visible', scale >= LABEL_ZOOM_THRESHOLD);
  }}

  function apply() {{
    stage.style.transform = `translate(${{panX}}px, ${{panY}}px) scale(${{scale}})`;
    updateLabelVisibility();
  }}

  function zoomAt(factor, cx, cy) {{
    const newScale = Math.min(MAP_MAX_ZOOM, Math.max(0.15, scale * factor));
    panX = cx - (cx - panX) * (newScale / scale);
    panY = cy - (cy - panY) * (newScale / scale);
    scale = newScale;
    apply();
  }}

  wrap.addEventListener('wheel', function(e) {{
    e.preventDefault();
    const rect = wrap.getBoundingClientRect();
    const cx = e.clientX - rect.left, cy = e.clientY - rect.top;
    zoomAt(e.deltaY < 0 ? 1.15 : 1/1.15, cx, cy);
  }}, {{ passive: false }});

  wrap.addEventListener('mousedown', function(e) {{
    dragging = true; lastX = e.clientX; lastY = e.clientY; wrap.classList.add('grabbing');
  }});
  window.addEventListener('mousemove', function(e) {{
    if (!dragging) return;
    panX += e.clientX - lastX; panY += e.clientY - lastY;
    lastX = e.clientX; lastY = e.clientY;
    apply();
  }});
  window.addEventListener('mouseup', function() {{ dragging = false; wrap.classList.remove('grabbing'); }});

  // touch support (pinch + drag)
  let pinchDist = null;
  wrap.addEventListener('touchstart', function(e) {{
    if (e.touches.length === 1) {{ dragging = true; lastX = e.touches[0].clientX; lastY = e.touches[0].clientY; }}
    else if (e.touches.length === 2) {{
      pinchDist = Math.hypot(e.touches[0].clientX-e.touches[1].clientX, e.touches[0].clientY-e.touches[1].clientY);
    }}
  }});
  wrap.addEventListener('touchmove', function(e) {{
    e.preventDefault();
    if (e.touches.length === 1 && dragging) {{
      panX += e.touches[0].clientX - lastX; panY += e.touches[0].clientY - lastY;
      lastX = e.touches[0].clientX; lastY = e.touches[0].clientY;
      apply();
    }} else if (e.touches.length === 2 && pinchDist) {{
      const d = Math.hypot(e.touches[0].clientX-e.touches[1].clientX, e.touches[0].clientY-e.touches[1].clientY);
      const rect = wrap.getBoundingClientRect();
      const cx = (e.touches[0].clientX+e.touches[1].clientX)/2 - rect.left;
      const cy = (e.touches[0].clientY+e.touches[1].clientY)/2 - rect.top;
      zoomAt(d/pinchDist, cx, cy);
      pinchDist = d;
    }}
  }}, {{ passive: false }});
  wrap.addEventListener('touchend', function() {{ dragging = false; pinchDist = null; }});

  document.getElementById('zoomIn').onclick = () => zoomAt(1.3, wrap.clientWidth/2, wrap.clientHeight/2);
  document.getElementById('zoomOut').onclick = () => zoomAt(1/1.3, wrap.clientWidth/2, wrap.clientHeight/2);
  document.getElementById('zoomReset').onclick = () => {{ scale=1; panX=0; panY=0; apply(); }};
  document.getElementById('debugToggle').onclick = function() {{
    stage.classList.toggle('debug-mode');
    this.classList.toggle('active');
  }};
  document.getElementById('bridgeToggle').onclick = function() {{
    bridgeState = (bridgeState + 1) % 3;
    applyBridgeState();
  }};
  document.getElementById('distToggle').onclick = function() {{
    stage.classList.toggle('distances-on');
    this.classList.toggle('active');
  }};
  document.getElementById('poiToggle').onclick = function() {{
    stage.classList.toggle('pois-on');
    this.classList.toggle('active');
  }};

  // --- GPS: live-locate the user on the nearest real road, mapped onto
  // that road's schematic line - meant to be used dynamically during an
  // actual trip, not just as a one-off lookup.
  const SEGMENTS_GEO = {segments_geo_json};
  let watchId = null;
  let gpsMarker = null;

  function ensureGpsMarker() {{
    if (gpsMarker) return gpsMarker;
    const svgEl = stage.querySelector('svg');
    const NS = 'http://www.w3.org/2000/svg';
    gpsMarker = document.createElementNS(NS, 'circle');
    gpsMarker.setAttribute('id', 'gpsMarker');
    gpsMarker.setAttribute('r', '9');
    gpsMarker.setAttribute('fill', '#1E90FF');
    gpsMarker.setAttribute('stroke', 'white');
    gpsMarker.setAttribute('stroke-width', '2.5');
    gpsMarker.setAttribute('opacity', '0.95');
    svgEl.appendChild(gpsMarker);
    return gpsMarker;
  }}

  function nearestPointOnSegmentLine(px, py, ax, ay, bx, by) {{
    const abx = bx - ax, aby = by - ay;
    const len2 = abx*abx + aby*aby;
    let t = len2 < 1e-12 ? 0 : ((px-ax)*abx + (py-ay)*aby) / len2;
    t = Math.max(0, Math.min(1, t));
    const qx = ax + t*abx, qy = ay + t*aby;
    const dx = px - qx, dy = py - qy;
    return {{ t: t, distSq: dx*dx + dy*dy }};
  }}

  function pointAtFractionOnPoly(poly, frac) {{
    let total = 0;
    const legLens = [];
    for (let i = 0; i < poly.length - 1; i++) {{
      const dx = poly[i+1][0]-poly[i][0], dy = poly[i+1][1]-poly[i][1];
      const l = Math.hypot(dx, dy);
      legLens.push(l);
      total += l;
    }}
    if (total <= 0) return poly[0];
    let target = frac * total, travelled = 0;
    for (let i = 0; i < legLens.length; i++) {{
      if (travelled + legLens[i] >= target || i === legLens.length - 1) {{
        const t = legLens[i] < 1e-9 ? 0 : (target - travelled) / legLens[i];
        const ax = poly[i][0], ay = poly[i][1], bx = poly[i+1][0], by = poly[i+1][1];
        return [ax + t*(bx-ax), ay + t*(by-ay)];
      }}
      travelled += legLens[i];
    }}
    return poly[poly.length-1];
  }}

  function updateGpsPosition(lat, lon) {{
    // find the nearest real road (equirectangular approx - fine at this
    // regional scale) by projecting onto each segment's real from->to line
    const cosLat = Math.cos(lat * Math.PI / 180);
    let best = null;
    for (const seg of SEGMENTS_GEO) {{
      const ax = seg.flon * cosLat, ay = -seg.flat;
      const bx = seg.tlon * cosLat, by = -seg.tlat;
      const px = lon * cosLat, py = -lat;
      const r = nearestPointOnSegmentLine(px, py, ax, ay, bx, by);
      if (!best || r.distSq < best.distSq) best = {{ seg: seg, t: r.t, distSq: r.distSq }};
    }}
    if (!best) return;
    const [sx, sy] = pointAtFractionOnPoly(best.seg.poly, best.t);
    const marker = ensureGpsMarker();
    marker.setAttribute('cx', sx);
    marker.setAttribute('cy', sy);
    marker.style.display = 'block';
  }}

  document.getElementById('gpsToggle').onclick = function() {{
    if (watchId !== null) {{
      navigator.geolocation.clearWatch(watchId);
      watchId = null;
      this.classList.remove('active');
      if (gpsMarker) gpsMarker.style.display = 'none';
      return;
    }}
    if (!navigator.geolocation) {{
      alert('Locatie wordt niet ondersteund door deze browser.');
      return;
    }}
    this.classList.add('active');
    watchId = navigator.geolocation.watchPosition(
      (pos) => updateGpsPosition(pos.coords.latitude, pos.coords.longitude),
      (err) => {{ alert('Kon locatie niet ophalen: ' + err.message); this.classList.remove('active'); watchId = null; }},
      {{ enableHighAccuracy: true, maximumAge: 5000, timeout: 15000 }}
    );
  }};

  applyBridgeState();
  apply();
}})();
</script>
</body></html>'''

    return html, {
        'junctions': len(junctions), 'segments': len(segments), 'points': len(points),
        'roads': len(road_colours), 'borders': len(border_data.get('segments', [])) if border_data else 0,
        'primary_moved_to_cap': len(moved_primary), 'secondary_moved_to_cap': len(moved_secondary),
        'ambiguous_border_crossings': ambiguous,
        'width': width, 'height': height, 'pois': len(pois), 'point_labels': label_stats,
    }


# =====================================================================
# Part 2: NS-style route diagram (was generate_graph.py)
# =====================================================================

DEFAULT_FROM = "Vught"
DEFAULT_TO = "Berwang"
DEFAULT_MARGIN = 0.5      # allow alternate branches up to 50% longer than the shortest path between the same two nodes
MAX_PATHS = 12            # hard cap on number of accepted alternate routes (readability)
MAX_CANDIDATES = 400      # safety cap on how many candidate paths Yen's algorithm may examine
MIN_NOVEL_KM = 3.0        # a candidate path must add at least this much genuinely new road to be worth a branch
MAX_LOCAL_SHARE = 0.25    # at most this share of a branch's detour may be Local road (e.g. a B-road)
# branches always drawn in the route diagram, before the automatically found ones: each
# is the list of junction names it must pass, in order. --branch replaces these,
# --no-default-branches turns them off.
DEFAULT_BRANCHES = [
    ["Kerpen", "Köln-West", "Frankfurter Kreuz", "Biebelried", "Feuchtwangen/Crailsheim", "Ulm/Elchingen"],  # A3/A7
]

HIERARCHY_WIDTH = {"Primary": 5.0, "Secondary": 4.0, "Connector": 3.0, "Local": 2.2}
TIER_RADIUS = {"Small": 5.0, "Medium": 8.0, "Large": 12.0}
LANE_HEIGHT = 90.0   # px per lane, downward
STEP_X = 110.0       # px per ordinal hop
MARGIN_PX = 60

ROAD_PALETTE = [
    "#1B3A6B", "#C0392B", "#1E8449", "#B9770E", "#6C3483", "#117864",
    "#A93226", "#2471A3", "#B7950B", "#7D3C98", "#229954", "#CA6F1E",
    "#2E86C1", "#943126",
]


# --------------------------------------------------------------------------
# Data loading
# --------------------------------------------------------------------------

def load_data(xlsx_path):
    wb = load_workbook(xlsx_path, data_only=True)

    junctions = {}
    for row in wb["Junctions"].iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        junctions[row[0]] = dict(
            id=row[0], name=row[1], roads=row[2], country=row[3],
            lat=row[5], lon=row[6], tier=row[11] or "Medium", jtype=row[12] or "Junction",
        )

    roads_lookup = {}
    for row in wb["Roads"].iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        roads_lookup[row[0]] = dict(canonical=row[8] or row[0], hierarchy=row[9] or "Connector")

    def canonical_and_hierarchy(road_field):
        if not road_field:
            return ("?", "Connector")
        if road_field in roads_lookup:
            r = roads_lookup[road_field]
            return (r["canonical"], r["hierarchy"])
        base = re.sub(r"\s*\([^)]*\)", "", road_field).strip()
        if base in roads_lookup:
            r = roads_lookup[base]
            return (r["canonical"], r["hierarchy"])
        for r in roads_lookup.values():
            if r["canonical"] == base:
                return (r["canonical"], r["hierarchy"])
        if "lokale" in road_field.lower():
            return (base or road_field, "Local")
        return (base or road_field, "Connector")

    segments = []
    seg_by_pair = {}
    for row in wb["Segments"].iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        edge_id, ffrom, _fn, tto, _tn, road, dist = row[0], row[1], row[2], row[3], row[4], row[5], row[6]
        override = row[11]
        canonical, hierarchy = canonical_and_hierarchy(road)
        if override in ("Primary", "Secondary", "Connector", "Local"):
            hierarchy = override
        ju1, ju2 = junctions.get(ffrom), junctions.get(tto)
        if not ju1 or not ju2:
            continue
        if dist and dist > 0:
            distance_km = float(dist)
        else:
            distance_km = haversine_km(ju1["lat"], ju1["lon"], ju2["lat"], ju2["lon"])
        rec = dict(edge_id=edge_id, from_id=ffrom, to_id=tto, road=road,
                   canonical=canonical, hierarchy=hierarchy, distance_km=distance_km)
        segments.append(rec)
        key = frozenset((ffrom, tto))
        existing = seg_by_pair.get(key)
        if existing is None or distance_km < existing["distance_km"]:
            seg_by_pair[key] = rec

    points_by_edge = defaultdict(list)
    for row in wb["Points"].iter_rows(min_row=2, values_only=True):
        if not row[0] or not row[9]:
            continue
        length_m = None
        notes = row[8] or ""
        m = re.search(r"lengte:\s*(\d+)\s*m", notes)
        if m:
            length_m = int(m.group(1))
        points_by_edge[row[9]].append(dict(
            id=row[0], name=row[1], category=row[2], sides=row[5],
            lat=row[6], lon=row[7], edge_id=row[9], pos=row[10] or 0.5,
            order=row[11] or 0, brand=row[12], facilities=row[13], length_m=length_m,
        ))
    for eid in points_by_edge:
        points_by_edge[eid].sort(key=lambda p: (p["pos"], p["order"] or 0))

    # points of interest ride along as points with category 'poi'
    for poi in load_pois(wb, junctions, {sg["edge_id"]: (sg["from_id"], sg["to_id"]) for sg in segments}):
        points_by_edge[poi["edge_id"]].append(dict(
            id=None, name=poi["name"], category="poi", sides=None, lat=None, lon=None,
            edge_id=poi["edge_id"], pos=poi["frac"], order=0, brand=None, facilities=None,
            length_m=None, poi=poi))

    return junctions, segments, seg_by_pair, points_by_edge


def find_junction_by_name(junctions, name):
    name_l = name.strip().lower()
    for j in junctions.values():
        if (j["name"] or "").strip().lower() == name_l:
            return j["id"]
    for j in junctions.values():
        if name_l in (j["name"] or "").strip().lower():
            return j["id"]
    raise SystemExit(f"Kon geen junction vinden met naam '{name}'")


# --------------------------------------------------------------------------
# Graph + path search
# --------------------------------------------------------------------------

def build_graph(junctions, segments):
    G = nx.Graph()
    for jid in junctions:
        G.add_node(jid)
    for seg in segments:
        # keep the shortest of any parallel segments between the same pair
        if G.has_edge(seg["from_id"], seg["to_id"]):
            if seg["distance_km"] >= G[seg["from_id"]][seg["to_id"]]["distance"]:
                continue
        G.add_edge(seg["from_id"], seg["to_id"], distance=seg["distance_km"], hierarchy=seg["hierarchy"])
    return G


def build_main_route(G, waypoint_ids):
    """The main trunk is the concatenation of the shortest path between
    each consecutive pair of user-specified waypoints, guaranteeing the
    trunk passes through the real-world route the user actually drives,
    even where a globally 'shortest by km' search would wander off
    through an unrelated, only-slightly-cheaper detour."""
    route = [waypoint_ids[0]]
    total = 0.0
    for a, b in zip(waypoint_ids, waypoint_ids[1:]):
        sub = nx.shortest_path(G, a, b, weight="distance")
        total += nx.shortest_path_length(G, a, b, weight="distance")
        route.extend(sub[1:])
    return route, total


def find_leg_alternates(G, a, b, margin, max_branches, min_novel_km, drawn_edges=None):
    """Alternative routes between two consecutive waypoints only (a 'leg'
    of the main trunk), so that alternatives are judged against the
    length of that leg, not the whole multi-hundred-km trip. This is what
    lets genuinely distinct regional alternatives (a different city, a
    different pass) surface, instead of the search budget being spent on
    countless near-duplicate detours in whichever region happens to have
    the densest local road network."""
    accepted, _ = find_routes(G, a, b, margin, max_paths=max_branches + 1, min_novel_km=min_novel_km,
                              drawn_edges=drawn_edges)
    return accepted[1:]  # [0] is the leg's own shortest path, already part of the main trunk


def find_routes(G, source, target, margin, max_paths=MAX_PATHS, max_candidates=MAX_CANDIDATES,
                 min_novel_km=MIN_NOVEL_KM, drawn_edges=None):
    """Picks the most distinct alternative routes. Yen's k-shortest simple
    paths supplies up to `max_candidates` candidates within `margin` of the
    shortest route; then, one branch at a time, the candidate adding the
    MOST road not yet drawn wins (ties: the shorter one). Picking by length
    alone lets dozens of near-identical local variants (a different slip
    road, a neighbouring junction) fill every slot before a genuinely
    different corridor gets a turn. A branch must be ONE detour - leave the
    drawn routes once and rejoin once, so no branch stacks unrelated
    detours - of at least `min_novel_km`, with at most MAX_LOCAL_SHARE of
    it on Local roads (a country road is not a motorway alternative).
    `drawn_edges`: edges already drawn by other routes (e.g. fixed
    branches), which count as not new."""
    candidates = []
    shortest = None
    for i, path in enumerate(nx.shortest_simple_paths(G, source, target, weight="distance")):
        if i >= max_candidates:
            break
        length = sum(G[u][v]["distance"] for u, v in zip(path, path[1:]))
        if shortest is None:
            shortest = length
        if length > shortest * (1 + margin):
            break
        candidates.append((path, length))
    if not candidates:
        return [], shortest

    accepted = [candidates[0]]
    drawn = set(drawn_edges or ())
    drawn.update(frozenset(e) for e in zip(candidates[0][0], candidates[0][0][1:]))
    remaining = candidates[1:]
    while remaining and len(accepted) < max_paths:
        best = None
        for k, (path, length) in enumerate(remaining):
            detours = _detours(path, drawn)
            if len(detours) != 1:
                continue
            novel_km = sum(G[u][v]["distance"] for u, v in detours[0])
            local_km = sum(G[u][v]["distance"] for u, v in detours[0] if G[u][v].get("hierarchy") == "Local")
            if novel_km < min_novel_km or local_km > MAX_LOCAL_SHARE * novel_km:
                continue
            if best is None or (novel_km, -length) > (best[1], -best[2]):
                best = (k, novel_km, length)
        if best is None:
            break
        path, length = remaining.pop(best[0])
        accepted.append((path, length))
        drawn.update(frozenset(e) for e in zip(path, path[1:]))
    return accepted, shortest


def _detours(path, drawn):
    """The stretches of `path` that leave the already-drawn edges, each as a
    list of (u, v) edges."""
    runs, cur = [], []
    for u, v in zip(path, path[1:]):
        if frozenset((u, v)) in drawn:
            if cur:
                runs.append(cur)
                cur = []
        else:
            cur.append((u, v))
    if cur:
        runs.append(cur)
    return runs


# --------------------------------------------------------------------------
# Schematic (ordinal) layout
# --------------------------------------------------------------------------

def order_alternates_for_layout(alternates):
    """Group alternates that share the same divergence/convergence anchor
    pair (i.e. every alternate within one leg, which all leave and rejoin
    the trunk at the identical two junctions) and sort each group by
    ascending hop count. Combined with the lane assignment in
    layout_routes below, this guarantees non-crossing geometry: within a
    shared-anchor group, the alternate with fewer hops (a bigger, steeper
    diagonal fraction) always gets the lane closest to the trunk, and each
    further-out lane on the same side gets a path with *more* hops (a
    smaller diagonal fraction) — so its transition line always stays
    inside the arc of the lane(s) nested within it, and lines only ever
    touch at the shared junctions, never in between."""
    groups = {}
    order = []
    for path, length in alternates:
        key = (path[0], path[-1])
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append((path, length))
    result = []
    for key in order:
        result.extend(sorted(groups[key], key=lambda pl: len(pl[0])))
    return result


def layout_routes(accepted):
    """Assign ordinal (x=hop position, y=signed offset) coordinates to
    every node, and record every directed edge to draw (u, v) exactly
    once.

    Two rules make the result crossing-free (lines only ever touch at a
    shared junction, per the project's requirement):

    1. Corridors: interior nodes of a branch are linearly interpolated
       between the x of its (already-drawn) attach junction and the x of
       its (already-drawn) rejoin junction, so a branch never drifts past
       the point where it reconnects to the rest of the graph.

    2. Side/depth nesting: a branch that attaches to the *trunk* (offset
       0) may go to either side. A branch that attaches to a junction
       that is itself already on a branch (offset != 0) is forced onto
       that SAME side and to a strictly larger depth (further from the
       trunk) — it can only nest outward from its parent, never cross
       back over it or over the trunk. Within a side/depth, lanes are
       only reused for x-ranges that provably don't overlap (with a
       safety gap). Combined with `order_alternates_for_layout` sorting
       same-anchor alternates by ascending hop count (so tighter/steeper
       transitions always nest inside gentler/shallower ones), no two
       branch lines cross except by sharing an actual junction node."""
    node_x, node_lane = {}, {}
    edges_drawn = {}          # frozenset({u,v}) -> (u, v) in the order first drawn
    # side_intervals[side][depth] = list of (x_min, x_max) already used at that depth
    side_intervals = {1: defaultdict(list), -1: defaultdict(list)}
    GAP = 0.35

    main_path = accepted[0][0]
    for i, node in enumerate(main_path):
        node_x[node] = float(i)
        node_lane[node] = 0
    for u, v in zip(main_path, main_path[1:]):
        edges_drawn[frozenset((u, v))] = (u, v)

    def free(x_lo, x_hi, occ):
        return all(x_hi < s - GAP or x_lo > e + GAP for (s, e) in occ)

    def get_offset(x_lo, x_hi, forced_side, min_depth):
        sides = [forced_side] if forced_side else [1, -1]
        depth = min_depth
        while True:
            for side in sides:
                if free(x_lo, x_hi, side_intervals[side][depth]):
                    side_intervals[side][depth].append((x_lo, x_hi))
                    return side * depth
            depth += 1

    ordered_alternates = order_alternates_for_layout(accepted[1:])
    runs = []                 # interior nodes of each detour, sharing one offset

    for path, _length in ordered_alternates:
        n = len(path)
        i = 0
        while i < n - 1:
            u = path[i]
            key = frozenset((u, path[i + 1]))
            if key in edges_drawn:
                i += 1
                continue
            if path[i + 1] in node_x:
                # both endpoints already exist but this exact edge is new (rare
                # shortcut/chord) — draw it directly, no interior nodes involved
                edges_drawn[key] = (u, path[i + 1])
                i += 1
                continue

            # start of a new run of not-yet-drawn interior nodes, bounded by
            # the attach junction (path[i], already drawn) and the next
            # already-drawn node further along the path (the rejoin junction)
            attach_idx = i
            j = i + 1
            while j < n and path[j] not in node_x:
                j += 1
            j = min(j, n - 1)  # the target is always drawn (part of the main path)
            rejoin_idx = j

            x_attach = node_x[path[attach_idx]]
            x_rejoin = node_x[path[rejoin_idx]]
            hop_count = rejoin_idx - attach_idx
            parent_offset = node_lane[path[attach_idx]]
            if parent_offset == 0:
                forced_side, min_depth = None, 1
            else:
                forced_side = 1 if parent_offset > 0 else -1
                min_depth = abs(parent_offset) + 1
            offset = get_offset(min(x_attach, x_rejoin), max(x_attach, x_rejoin), forced_side, min_depth)

            for k in range(attach_idx + 1, rejoin_idx):
                frac = (k - attach_idx) / hop_count
                node_x[path[k]] = x_attach + frac * (x_rejoin - x_attach)
                node_lane[path[k]] = offset
            if rejoin_idx > attach_idx + 1:
                runs.append(path[attach_idx + 1:rejoin_idx])
            for k in range(attach_idx, rejoin_idx):
                edges_drawn[frozenset((path[k], path[k + 1]))] = (path[k], path[k + 1])
            i = rejoin_idx

    untangle_lanes(runs, node_x, node_lane, edges_drawn)
    return node_x, node_lane, edges_drawn


LANE_CHOICES = sorted({s * d / 2 for s in (1, -1) for d in range(1, 9)}, key=lambda o: (abs(o), -o))


def untangle_lanes(runs, node_x, node_lane, edges_drawn, passes=6):
    """The lane rules above only look at where a branch LEAVES; one that
    rejoins a line on the other side can still cross. So each detour's lane
    is re-chosen (whole and half lanes, e.g. between the trunk and a branch)
    to minimise, in order: crossings, then total distance from the trunk -
    one detour at a time, repeated until nothing improves."""
    edges = list(edges_drawn.values())

    def cost():
        return count_crossings(node_x, node_lane, edges), sum(abs(node_lane[r[0]]) for r in runs)

    best = cost()
    for _ in range(passes):
        improved = False
        for run in runs:
            current = node_lane[run[0]]
            for lane in LANE_CHOICES:
                if lane == current:
                    continue
                for n in run:
                    node_lane[n] = lane
                c = cost()
                if c < best:
                    best, current, improved = c, lane, True
                else:
                    for n in run:
                        node_lane[n] = current
        if not improved:
            break

    # half lanes are cramped: renumber each side's lanes to 1, 2, 3... in the
    # same order, and keep that only if it adds no crossings
    before = {n: node_lane[n] for run in runs for n in run}
    for side in (1, -1):
        used = sorted({node_lane[r[0]] for r in runs if node_lane[r[0]] * side > 0}, key=abs)
        remap = {lane: side * (i + 1) for i, lane in enumerate(used)}
        for run in runs:
            target = remap.get(node_lane[run[0]])
            if target is not None:
                for n in run:
                    node_lane[n] = target
    if count_crossings(node_x, node_lane, edges) > best[0]:
        node_lane.update(before)
    return count_crossings(node_x, node_lane, edges)


def count_crossings(node_x, node_lane, edges):
    """Pairs of drawn edges that cross, or run on top of each other, anywhere
    other than at a junction they share."""
    pts = {n: (node_x[n], node_lane[n]) for e in edges for n in e}

    def orient(a, b, c):
        v = (b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])
        return 0 if abs(v) < 1e-9 else (1 if v > 0 else -1)

    def on_seg(a, b, c):
        return (min(a[0], b[0]) - 1e-9 <= c[0] <= max(a[0], b[0]) + 1e-9
                and min(a[1], b[1]) - 1e-9 <= c[1] <= max(a[1], b[1]) + 1e-9)

    n = 0
    for i in range(len(edges)):
        u1, v1 = edges[i]
        a, b = pts[u1], pts[v1]
        for j in range(i + 1, len(edges)):
            u2, v2 = edges[j]
            c, d = pts[u2], pts[v2]
            if max(a[0], b[0]) < min(c[0], d[0]) or max(c[0], d[0]) < min(a[0], b[0]):
                continue
            o1, o2, o3, o4 = orient(a, b, c), orient(a, b, d), orient(c, d, a), orient(c, d, b)
            shared = {u1, v1} & {u2, v2}
            if shared:
                if o1 == 0 and o2 == 0:          # collinear from the shared junction: overlap?
                    s = next(iter(shared))
                    p = pts[s]
                    q1 = pts[v1 if u1 == s else u1]
                    q2 = pts[v2 if u2 == s else u2]
                    if (q1[0] - p[0]) * (q2[0] - p[0]) + (q1[1] - p[1]) * (q2[1] - p[1]) > 0:
                        n += 1
                continue
            if (o1 != o2 and o3 != o4) or (o1 == 0 and on_seg(a, b, c)) or (o2 == 0 and on_seg(a, b, d)) \
                    or (o3 == 0 and on_seg(c, d, a)) or (o4 == 0 and on_seg(c, d, b)):
                n += 1
    return n


# --------------------------------------------------------------------------
# Rendering
# --------------------------------------------------------------------------

def assign_colours(edges_drawn, seg_by_pair):
    colours = {}
    palette_i = 0
    order = []
    for key in edges_drawn:
        u, v = edges_drawn[key]
        seg = seg_by_pair.get(key)
        canon = seg["canonical"] if seg else "?"
        if canon not in colours:
            colours[canon] = ROAD_PALETTE[palette_i % len(ROAD_PALETTE)]
            palette_i += 1
            order.append(canon)
    return colours


def fuel_side_visible(point, from_j, to_j):
    """Direction-aware visibility for one-sided fuel stations: only show a
    one-sided station if it lies on the right-hand side of the direction
    of travel from from_j to to_j (right-hand traffic countries)."""
    if point["sides"] != 1:
        return True
    if point["lat"] is None or point["lon"] is None:
        return True
    dx = to_j["lon"] - from_j["lon"]
    dy = to_j["lat"] - from_j["lat"]
    # right-hand normal of travel direction (rotate -90deg in lon(x)/lat(y) plane)
    rnx, rny = dy, -dx
    vx = point["lon"] - from_j["lon"]
    vy = point["lat"] - from_j["lat"]
    side_val = vx * rnx + vy * rny
    return side_val > 0


def render_graph(junctions, seg_by_pair, points_by_edge, node_x, node_lane, edges_drawn, source, target,
           orientation="vertical"):
    colours = assign_colours(edges_drawn, seg_by_pair)

    # primary = position along the route (hop index); secondary = branch offset (zig-zag lane)
    primary = {n: node_x[n] for n in node_x}
    secondary = {n: node_lane[n] for n in node_x}

    min_p, max_p = min(primary.values()), max(primary.values())
    min_s, max_s = min(secondary.values()) - 1, max(secondary.values()) + 1

    if orientation == "vertical":
        def px(n):
            return MARGIN_PX + 300 + (secondary[n] - min_s) * LANE_HEIGHT

        def py(n):
            return MARGIN_PX + (primary[n] - min_p) * STEP_X

        width = MARGIN_PX * 2 + 300 + (max_s - min_s) * LANE_HEIGHT + 260
        height = MARGIN_PX * 2 + (max_p - min_p) * STEP_X
    else:
        def px(n):
            return MARGIN_PX + (primary[n] - min_p) * STEP_X

        def py(n):
            return MARGIN_PX + 60 + (secondary[n] - min_s) * LANE_HEIGHT

        width = MARGIN_PX * 2 + (max_p - min_p) * STEP_X + 300
        height = MARGIN_PX * 2 + 60 + (max_s - min_s) * LANE_HEIGHT

    # which canonical road "arrives" at each node, to detect where a road-number label is needed
    incoming_canonical = {}
    for key, (u, v) in edges_drawn.items():
        seg = seg_by_pair.get(key)
        incoming_canonical[v] = seg["canonical"] if seg else "?"

    placer = LabelPlacer()
    pending_points = []   # (x, y, line angle, point) - labels placed after all obstacles are known
    svg_lines = []
    svg_points = []
    svg_pois = []
    svg_nodes = []
    svg_labels = []
    svg_road_labels = []

    # --- road lines ---
    for key, (u, v) in edges_drawn.items():
        ju, jv = junctions[u], junctions[v]
        seg = seg_by_pair.get(key)
        canon = seg["canonical"] if seg else "?"
        hierarchy = seg["hierarchy"] if seg else "Connector"
        colour = colours.get(canon, "#888")
        width_px = HIERARCHY_WIDTH.get(hierarchy, 3.0)
        x1, y1, x2, y2 = px(u), py(u), px(v), py(v)
        placer.add_polyline([(x1, y1), (x2, y2)])
        dist_label = f"{seg['distance_km']:.0f} km" if seg else ""
        svg_lines.append(
            f'<line class="road-line" x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" '
            f'stroke="{colour}" stroke-width="{width_px}" stroke-linecap="round">'
            f'<title>{canon} ({hierarchy}) {ju["name"]} → {jv["name"]} – {dist_label}</title></line>'
        )

        is_run_start = incoming_canonical.get(u) != canon
        if is_run_start:
            mx, my = (x1 + x2) / 2, (y1 + y2) / 2
            angle = math.degrees(math.atan2(y2 - y1, x2 - x1))
            if angle > 90 or angle < -90:
                angle += 180
            rw, rh = 6 + len(canon) * GRAPH_ROAD_FONT * 0.62, GRAPH_ROAD_FONT * 1.25
            placer.add_rotated_box(mx, my, angle, -1, -rh * 0.7, rw, rh)
            svg_road_labels.append(
                f'<g transform="translate({mx:.1f} {my:.1f}) rotate({angle:.1f})">'
                f'<rect x="-1" y="{-rh * 0.7:.1f}" width="{rw:.0f}" height="{rh:.1f}" rx="3" '
                f'fill="#fbfaf6" fill-opacity="0.88"/>'
                f'<text x="3" y="{rh * 0.3 - 1:.1f}" font-size="{GRAPH_ROAD_FONT}" font-weight="700" fill="{colour}">{_esc(canon)}</text>'
                f'</g>'
            )

        # --- points along this edge ---
        if seg:
            reversed_dir = seg["from_id"] != u
            for pt in points_by_edge.get(seg["edge_id"], []):
                frac = pt["pos"] if not reversed_dir else (1 - pt["pos"])
                px_pt = x1 + frac * (x2 - x1)
                py_pt = y1 + frac * (y2 - y1)
                line_angle = math.degrees(math.atan2(y2 - y1, x2 - x1))
                if pt["category"] == "tankstation":
                    if not fuel_side_visible(pt, ju if not reversed_dir else jv, jv if not reversed_dir else ju):
                        continue
                    placer.add_box((px_pt - GRAPH_FUEL_HALF, py_pt - GRAPH_FUEL_HALF,
                                    px_pt + GRAPH_FUEL_HALF, py_pt + GRAPH_FUEL_HALF))
                    pending_points.append((px_pt, py_pt, line_angle, pt))
                elif pt["category"] == "poi":
                    placer.add_box((px_pt - GRAPH_POI_R, py_pt - GRAPH_POI_R,
                                    px_pt + GRAPH_POI_R, py_pt + GRAPH_POI_R))
                    pending_points.append((px_pt, py_pt, line_angle, pt))
                elif pt["category"] in GRAPH_BRIDGE_CATEGORIES:
                    placer.add_box((px_pt - GRAPH_BRIDGE_R, py_pt - GRAPH_BRIDGE_R,
                                    px_pt + GRAPH_BRIDGE_R, py_pt + GRAPH_BRIDGE_R))
                    pending_points.append((px_pt, py_pt, line_angle, pt))

    # --- junction circles ---
    for n in node_x:
        j = junctions[n]
        r = TIER_RADIUS.get(j["tier"], 6.0)
        cx, cy = px(n), py(n)
        is_endpoint = n in (source, target)
        fill = "#1B3A6B" if is_endpoint else "#ffffff"
        stroke = "#1B3A6B"
        dash = ' stroke-dasharray="3,2"' if j["jtype"] == "BorderCrossing" else ""
        placer.add_box((cx - r, cy - r, cx + r, cy + r))
        svg_nodes.append(
            f'<circle class="junction" cx="{cx:.1f}" cy="{cy:.1f}" r="{r}" '
            f'fill="{fill}" stroke="{stroke}" stroke-width="2"{dash}>'
            f'<title>{j["name"]} ({j["id"]})</title></circle>'
        )

    # --- junction names: endpoints and the main route first, each on the
    # first free spot of: angled up-right (the classic look), horizontal
    # right, angled down-left, horizontal left ---
    f = GRAPH_JUNCTION_FONT
    for n in sorted(node_x, key=lambda n: (n not in (source, target), node_lane[n] != 0, node_x[n])):
        j = junctions[n]
        r = TIER_RADIUS.get(j["tier"], 6.0)
        cx, cy = px(n), py(n)
        weight = "700" if n in (source, target) else "600" if node_lane[n] == 0 else "500"
        w = text_width(j["name"], f) * 1.08          # a little wider: bold
        gap = max(9, r + 4)
        options = [  # (angle, text-anchor, dx, dy in em, box x0, box y0) in the rotated frame
            (-28, "start", gap, -0.3, gap, -1.2 * f),
            (0, "start", r + 4, 0.35, r + 4, -0.55 * f),
            (-28, "end", -gap, 0.9, -gap - w, 0.0),
            (0, "end", -(r + 4), 0.35, -(r + 4) - w, -0.55 * f),
        ]
        chosen = None
        for avoid_lines in (True, False):
            for opt in options:
                boxes = placer.rotated_boxes(cx, cy, opt[0], opt[4], opt[5], w, f * 1.15)
                if placer.boxes_free(boxes, avoid_lines):
                    chosen = (opt, boxes)
                    break
            if chosen:
                break
        if chosen is None:
            opt = options[0]
            chosen = (opt, placer.rotated_boxes(cx, cy, opt[0], opt[4], opt[5], w, f * 1.15))
        (angle, anchor, dx, dy, _bx, _by), boxes = chosen
        for b in boxes:
            placer.add_box(b, text=True)
        rot = f' transform="rotate({angle} {cx:.1f} {cy:.1f})"' if angle else ""
        svg_labels.append(
            f'<text class="junction-label" x="{cx:.1f}" y="{cy:.1f}" dx="{dx:.1f}" dy="{dy}em" '
            f'text-anchor="{anchor}"{rot} font-weight="{weight}">{_esc(j["name"])}</text>'
        )

    # label texts and sizes
    items = []
    for x, y, line_angle, pt in pending_points:
        if pt["category"] == "poi":
            text = pt["name"]
            reach = GRAPH_POI_R + 0.5
        elif pt["category"] == "tankstation":
            text = (pt["name"] or "") + (f' [{pt["brand"]}]' if pt.get("brand") else "")
            reach = GRAPH_FUEL_HALF + 0.5
        else:
            text = (pt["name"] or "") + (f" ({pt['length_m']}m)" if pt.get("length_m") else "")
            reach = GRAPH_BRIDGE_R + 0.5
        bold = 1.08 if pt["category"] == "poi" else 1.0
        items.append(dict(x=x, y=y, line_angle=line_angle, pt=pt, reach=reach,
                          w=text_width(text, GRAPH_POINT_FONT) * bold, h=GRAPH_POINT_FONT * LINE_HEIGHT,
                          ignore=(x - reach, y - reach, x + reach, y + reach)))

    # vertical diagram: ordered columns beside each line - on the left first
    # (junction names run up and to the right), then on the right for what
    # didn't fit
    placed = {}
    if orientation == "vertical":
        placed = stack_beside_vertical_lines(items, placer, side=-1)
        left_over = {i for i in range(len(items)) if i not in placed}
        placed.update(stack_beside_vertical_lines(items, placer, side=1, only=left_over))
    # whatever didn't fit in a column (or a horizontal diagram): free placement,
    # fuel stations first, then tunnels/river bridges, then valley bridges
    rest = [i for i in range(len(items)) if i not in placed]
    rest.sort(key=lambda i: (LABEL_PRIORITY.get((items[i]["pt"]["category"] or "").lower(), 9),
                             -items[i]["w"]))
    for i in rest:
        it = items[i]
        normal = it["line_angle"] + 90
        nxc, nyc = math.cos(math.radians(normal)), math.sin(math.radians(normal))
        pref = (1 if nxc < 0 else -1) if abs(nxc) > abs(nyc) else (1 if nyc > 0 else -1)
        placed[i] = placer.place(it["x"], it["y"], it["w"], it["h"], normal, pref_side=pref,
                                 clear=it["reach"], ignore=it["ignore"])
    for i, it in enumerate(items):
        box, leader = placed[i]
        if it["pt"]["category"] == "poi":
            svg_pois.append(_poi_marker(it["x"], it["y"], it["pt"], box, leader))
        elif it["pt"]["category"] == "tankstation":
            svg_points.append(_fuel_marker(it["x"], it["y"], it["pt"], box, leader))
        else:
            svg_points.append(_bridge_marker(it["x"], it["y"], it["pt"], box, leader))

    # --- segment distances (toggle "Afstanden"): placed last, beside the
    # middle of each line on the side away from the point-label columns ---
    svg_dists = []
    for key, (u, v) in edges_drawn.items():
        seg = seg_by_pair.get(key)
        if not seg or seg["distance_km"] is None:
            continue
        x1, y1, x2, y2 = px(u), py(u), px(v), py(v)
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        km = seg["distance_km"]
        text = f"{km:.1f} km" if km < 10 else f"{km:.0f} km"
        w, h = text_width(text, GRAPH_DIST_FONT), GRAPH_DIST_FONT * LINE_HEIGHT
        normal = math.degrees(math.atan2(y2 - y1, x2 - x1)) + 90
        nxc, nyc = math.cos(math.radians(normal)), math.sin(math.radians(normal))
        pref = (1 if nxc > 0 else -1) if abs(nxc) > abs(nyc) else (1 if nyc < 0 else -1)
        box, leader = placer.place(mx, my, w, h, normal, pref_side=pref, clear=3.0)
        line = ""
        if leader:
            ex, ey = nearest_on_box(mx, my, box)
            line = f'<line class="leader" x1="{mx:.1f}" y1="{my:.1f}" x2="{ex:.1f}" y2="{ey:.1f}"/>'
        svg_dists.append(
            f'<g class="dist-label">{line}<text x="{box[0]:.1f}" y="{box[1] + GRAPH_DIST_FONT * BASELINE:.1f}">'
            f'{text}</text></g>')

    svg = (
        f'<svg id="mapsvg" width="{width:.0f}" height="{height:.0f}" '
        f'viewBox="0 0 {width:.0f} {height:.0f}" xmlns="http://www.w3.org/2000/svg">'
        f'<rect width="100%" height="100%" fill="#fbfaf6"/>'
        + "".join(svg_lines) + "".join(svg_road_labels) + "".join(svg_points)
        + "".join(svg_nodes) + "".join(svg_labels) + "".join(svg_dists) + "".join(svg_pois)
        + "</svg>"
    )
    return svg, width, height, placer.stats


def _esc(s):
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _label_pos(x, y, box):
    """Text anchor (relative to the marker at x, y) for a label box: the
    right edge for labels left of their marker (right-aligned, so the text
    meets its leader), else the left edge."""
    anchor_x = box[2] if box[2] <= x else box[0]
    return anchor_x - x, box[1] - y + GRAPH_POINT_FONT * BASELINE, ("end" if box[2] <= x else "start")


def _leader(x, y, box, leader):
    if not leader:
        return ""
    ex, ey = nearest_on_box(x, y, box)
    return f'<line class="leader" x1="0" y1="0" x2="{ex - x:.1f}" y2="{ey - y:.1f}"/>'


def _fuel_marker(x, y, pt, box, leader):
    label = _esc(pt["name"])
    brand = f' <tspan class="brand">[{_esc(pt["brand"])}]</tspan>' if pt.get("brand") else ""
    lx, ly, anchor = _label_pos(x, y, box)
    return (
        f'<g class="fuel-marker" transform="translate({x:.1f} {y:.1f})">'
        + _leader(x, y, box, leader) +
        f'<rect x="{-GRAPH_FUEL_HALF}" y="{-GRAPH_FUEL_HALF}" width="{2 * GRAPH_FUEL_HALF}" '
        f'height="{2 * GRAPH_FUEL_HALF}" fill="#E8871E" stroke="#7a4400" stroke-width="1"/>'
        f'<text class="fuel-label zoom-label" x="{lx:.1f}" y="{ly:.1f}" text-anchor="{anchor}">{label}{brand}</text>'
        f'</g>'
    )


def _poi_marker(x, y, pt, box, leader):
    lx, ly, anchor = _label_pos(x, y, box)
    return (
        f'<g class="poi-marker" transform="translate({x:.1f} {y:.1f})">'
        f'<title>{poi_title(pt["poi"])}</title>'
        + _leader(x, y, box, leader) + star_svg(0, 0, GRAPH_POI_R, "poi-star") +
        f'<text class="poi-label" x="{lx:.1f}" y="{ly:.1f}" text-anchor="{anchor}">{_esc(pt["name"])}</text>'
        f'</g>'
    )


def _bridge_marker(x, y, pt, box, leader):
    kind = "tunnel" if pt["category"] == "tunnel" else ("ecoduct" if pt["category"] == "ecoduct" else "brug")
    colour = {"tunnel": "#555", "ecoduct": "#3a7d33", "brug": "#2471A3"}[kind]
    length_txt = f" ({pt['length_m']}m)" if pt.get("length_m") else ""
    lx, ly, anchor = _label_pos(x, y, box)
    return (
        f'<g class="bridge-marker" transform="translate({x:.1f} {y:.1f})">'
        + _leader(x, y, box, leader) +
        f'<circle r="{GRAPH_BRIDGE_R}" fill="{colour}" stroke="#222" stroke-width="0.8"/>'
        f'<text class="bridge-label bridges-on" x="{lx:.1f}" y="{ly:.1f}" text-anchor="{anchor}">{_esc(pt["name"])}</text>'
        f'<text class="bridge-label bridges-full" x="{lx:.1f}" y="{ly:.1f}" text-anchor="{anchor}">{_esc(pt["name"])}{length_txt}</text>'
        f'</g>'
    )


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="utf-8">
<title>{title}</title>
<style>
  html, body {{ margin:0; padding:0; height:100%; background:#e9e6dd; font-family: Arial, Helvetica, sans-serif; overflow:hidden; }}
  #wrap {{ position:absolute; inset:0; overflow:hidden; cursor:grab; }}
  #wrap.dragging {{ cursor:grabbing; }}
  #stage {{ position:absolute; left:0; top:0; transform-origin: 0 0; }}
  .junction-label {{ font-size:{junction_font}px; fill:#1a1a1a; }}
  .fuel-label, .bridge-label {{ font-size:{point_font}px; fill:#333; display:none;
                                 paint-order:stroke; stroke:#fbfaf6; stroke-width:2.8px; stroke-linejoin:round; }}
  .leader {{ stroke:#888; stroke-width:0.6; }}
  #stage.points-visible .fuel-label {{ display:block; }}
  #stage.bridges-on .bridge-label.bridges-on {{ display:block; }}
  #stage.bridges-full .bridge-label.bridges-full {{ display:block; }}
  #stage.bridges-full .bridge-label.bridges-on {{ display:none; }}
  .fuel-marker, .bridge-marker {{ display:none; }}
  #stage.points-visible .fuel-marker {{ display:block; }}
  #stage.bridges-on .bridge-marker, #stage.bridges-full .bridge-marker {{ display:block; }}
  .brand {{ fill:#a35b00; font-style:italic; }}
  .poi-marker {{ display:none; }}
  #stage.pois-on .poi-marker {{ display:block; }}
  .poi-label {{ font-size:{point_font}px; font-weight:700; fill:{poi_text};
                paint-order:stroke; stroke:#fbfaf6; stroke-width:2.8px; stroke-linejoin:round; }}
  .dist-label {{ display:none; }}
  #stage.distances-on .dist-label {{ display:block; }}
  .dist-label text {{ font-size:{dist_font}px; font-weight:700; fill:#185FA5;
                      paint-order:stroke; stroke:#fbfaf6; stroke-width:2.8px; stroke-linejoin:round; }}
  #controls {{ position:absolute; top:12px; right:12px; z-index:5; display:flex; flex-direction:column; gap:6px; }}
  #controls button {{ font-size:13px; padding:6px 10px; border-radius:6px; border:1px solid #999; background:#fff; cursor:pointer; }}
  #controls button.active {{ background:#1B3A6B; color:#fff; border-color:#1B3A6B; }}
  #info {{ position:absolute; bottom:10px; left:12px; z-index:5; font-size:12px; color:#444; background:rgba(255,255,255,.85);
           padding:4px 8px; border-radius:6px; max-width:60vw; }}
  #title {{ position:absolute; left:12px; top:12px; z-index:5; font-size:14px; font-weight:700; color:#1B3A6B;
            background:rgba(255,255,255,.88); padding:4px 10px; border-radius:6px; }}
</style>
</head>
<body>
<div id="title">{title}</div>
<div id="controls">
  <button id="zoomIn">+</button>
  <button id="zoomOut">&minus;</button>
  <button id="zoomReset">reset</button>
  <button id="toggleFuel" class="active">Tankstations</button>
  <button id="toggleBridges">Bruggen/tunnels</button>
  <button id="toggleDist" class="active">Afstanden</button>
  <button id="togglePoi" class="active">&#9733; Bezienswaardigheden</button>
</div>
<div id="info">hoofdroute ca. {shortest_km:.0f} km &middot; {route_count} route-varianten getoond (marge {margin_pct:.0f}% per traject)</div>
<div id="wrap"><div id="stage" class="points-visible distances-on pois-on">{svg}</div></div>
<script>
(function() {{
  const stage = document.getElementById('stage');
  const wrap = document.getElementById('wrap');
  let scale = 1, tx = 0, ty = 0;
  let dragging = false, lastX = 0, lastY = 0;

  function apply() {{
    stage.style.transform = `translate(${{tx}}px, ${{ty}}px) scale(${{scale}})`;
  }}
  function fitInitial() {{
    const svg = document.getElementById('mapsvg');
    const w = svg.width.baseVal.value, h = svg.height.baseVal.value;
    const ww = wrap.clientWidth, wh = wrap.clientHeight;
    scale = Math.min(ww / w, wh / h, 1) * 0.94;
    tx = (ww - w * scale) / 2; ty = 20;
    apply();
  }}
  wrap.addEventListener('wheel', (e) => {{
    e.preventDefault();
    const rect = wrap.getBoundingClientRect();
    const mx = e.clientX - rect.left, my = e.clientY - rect.top;
    const before = scale;
    scale *= (e.deltaY < 0) ? 1.12 : 1/1.12;
    scale = Math.min(Math.max(scale, 0.15), {max_zoom});
    tx = mx - (mx - tx) * (scale / before);
    ty = my - (my - ty) * (scale / before);
    apply();
  }}, {{ passive:false }});
  wrap.addEventListener('mousedown', (e) => {{ dragging = true; lastX=e.clientX; lastY=e.clientY; wrap.classList.add('dragging'); }});
  window.addEventListener('mouseup', () => {{ dragging=false; wrap.classList.remove('dragging'); }});
  window.addEventListener('mousemove', (e) => {{
    if (!dragging) return;
    tx += e.clientX - lastX; ty += e.clientY - lastY;
    lastX = e.clientX; lastY = e.clientY; apply();
  }});
  let lastTouchDist = null;
  wrap.addEventListener('touchstart', (e) => {{
    if (e.touches.length === 1) {{ dragging=true; lastX=e.touches[0].clientX; lastY=e.touches[0].clientY; }}
    else if (e.touches.length === 2) {{ lastTouchDist = Math.hypot(e.touches[0].clientX-e.touches[1].clientX, e.touches[0].clientY-e.touches[1].clientY); }}
  }}, {{passive:true}});
  wrap.addEventListener('touchmove', (e) => {{
    if (e.touches.length === 1 && dragging) {{
      tx += e.touches[0].clientX - lastX; ty += e.touches[0].clientY - lastY;
      lastX = e.touches[0].clientX; lastY = e.touches[0].clientY; apply();
    }} else if (e.touches.length === 2) {{
      const d = Math.hypot(e.touches[0].clientX-e.touches[1].clientX, e.touches[0].clientY-e.touches[1].clientY);
      if (lastTouchDist) {{ scale *= d / lastTouchDist; scale = Math.min(Math.max(scale,0.15),{max_zoom}); apply(); }}
      lastTouchDist = d;
    }}
  }}, {{passive:true}});
  wrap.addEventListener('touchend', () => {{ dragging=false; lastTouchDist=null; }});

  document.getElementById('zoomIn').onclick = () => {{ scale = Math.min(scale*1.25, {max_zoom}); apply(); }};
  document.getElementById('zoomOut').onclick = () => {{ scale = Math.max(scale/1.25, 0.15); apply(); }};
  document.getElementById('zoomReset').onclick = fitInitial;

  const fuelBtn = document.getElementById('toggleFuel');
  fuelBtn.onclick = () => {{
    stage.classList.toggle('points-visible');
    fuelBtn.classList.toggle('active');
  }};
  const poiBtn = document.getElementById('togglePoi');
  poiBtn.onclick = () => {{
    stage.classList.toggle('pois-on');
    poiBtn.classList.toggle('active');
  }};
  const distBtn = document.getElementById('toggleDist');
  distBtn.onclick = () => {{
    stage.classList.toggle('distances-on');
    distBtn.classList.toggle('active');
  }};
  const bridgeBtn = document.getElementById('toggleBridges');
  let bridgeState = 0; // 0 off, 1 on, 2 on+length
  bridgeBtn.onclick = () => {{
    bridgeState = (bridgeState + 1) % 3;
    stage.classList.remove('bridges-on','bridges-full');
    bridgeBtn.classList.remove('active');
    if (bridgeState === 1) {{ stage.classList.add('bridges-on'); bridgeBtn.classList.add('active'); bridgeBtn.textContent='Bruggen/tunnels: aan'; }}
    else if (bridgeState === 2) {{ stage.classList.add('bridges-full'); bridgeBtn.classList.add('active'); bridgeBtn.textContent='Bruggen/tunnels: +lengte'; }}
    else {{ bridgeBtn.textContent='Bruggen/tunnels'; }}
  }};

  window.addEventListener('resize', fitInitial);
  fitInitial();
}})();
</script>
</body>
</html>
"""


def build_graph_page(xlsx_path=DEFAULT_XLSX, start_name=DEFAULT_FROM,
                     end_name=DEFAULT_TO, margin=DEFAULT_MARGIN, title=None,
                     branches_per_leg=4, min_novel_km=MIN_NOVEL_KM, via=None, orientation="vertical",
                     explicit_branches=None):
    """Builds the route-diagram page. Returns (html, stats). explicit_branches:
    lists of junction names; None means DEFAULT_BRANCHES."""
    if explicit_branches is None:
        explicit_branches = DEFAULT_BRANCHES
    junctions, segments, seg_by_pair, points_by_edge = load_data(xlsx_path)
    source = find_junction_by_name(junctions, start_name)
    target = find_junction_by_name(junctions, end_name)
    via_ids = [find_junction_by_name(junctions, v) for v in (via or [])]
    waypoint_ids = [source] + via_ids + [target]

    G = build_graph(junctions, segments)
    main_route, main_length = build_main_route(G, waypoint_ids)

    accepted = [(main_route, main_length)]
    leg_count = 0
    # explicit, user-specified branches (named waypoint chains) first — real
    # alternatives (e.g. a well-known corridor via a different city) that should
    # always be drawn; their road then counts as drawn, so the automatic search
    # below doesn't add a near-copy of them
    drawn = {frozenset(e) for e in zip(main_route, main_route[1:])}
    for chain_names in (explicit_branches or []):
        chain_ids = [find_junction_by_name(junctions, n) for n in chain_names]
        branch_path, branch_length = build_main_route(G, chain_ids)
        accepted.append((branch_path, branch_length))
        drawn.update(frozenset(e) for e in zip(branch_path, branch_path[1:]))
        leg_count += 1

    for a, b in zip(waypoint_ids, waypoint_ids[1:]):
        leg_alts = find_leg_alternates(G, a, b, margin, branches_per_leg, min_novel_km, drawn_edges=drawn)
        accepted.extend(leg_alts)
        leg_count += len(leg_alts)
        for path, _length in leg_alts:
            drawn.update(frozenset(e) for e in zip(path, path[1:]))

    node_x, node_lane, edges_drawn = layout_routes(accepted)
    svg, width, height, label_stats = render_graph(
        junctions, seg_by_pair, points_by_edge, node_x, node_lane, edges_drawn, source, target,
        orientation=orientation,
    )

    title = title or f"Route: {junctions[source]['name']} → {junctions[target]['name']}"
    html = HTML_TEMPLATE.format(
        title=_esc(title), svg=svg,
        route_count=leg_count + 1, shortest_km=main_length, margin_pct=margin * 100,
        point_font=GRAPH_POINT_FONT, dist_font=GRAPH_DIST_FONT, junction_font=GRAPH_JUNCTION_FONT,
        poi_text=POI_TEXT,
        max_zoom=GRAPH_MAX_ZOOM,
    )
    return html, dict(
        title=title, main_route_km=round(main_length, 1), branches=leg_count,
        nodes=len(node_x), edges=len(edges_drawn), width=width, height=height,
        point_labels=label_stats,
    )


# =====================================================================
# Part 3: one page, two tabs
# =====================================================================

def combine_pages(map_html, graph_html, map_tab, graph_tab, page_title):
    """Both views as separate documents in one file: each sits in its own
    iframe (srcdoc), so their ids, styles and scripts can't collide. Both
    frames keep their full size (hidden with visibility, not display), so
    each view's initial fit and zoom state survive switching tabs."""
    esc = lambda s: html_lib.escape(s, quote=True)
    return f'''<!DOCTYPE html>
<html lang="nl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(page_title)}</title>
<style>
  html, body {{ margin:0; height:100%; overflow:hidden; background:#f0efe9;
               font-family: Arial, Helvetica, sans-serif; }}
  #tabs {{ position:absolute; top:0; left:0; right:0; height:40px; display:flex; gap:4px;
           padding:6px 8px 0; box-sizing:border-box; background:#1B3A6B; }}
  #tabs button {{ border:0; border-radius:6px 6px 0 0; padding:0 16px; font-size:14px; cursor:pointer;
                  background:#3A5A8C; color:#DDE6F3; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
  #tabs button[aria-selected="true"] {{ background:#f0efe9; color:#1B3A6B; font-weight:700; }}
  .view {{ position:absolute; top:40px; left:0; width:100%; height:calc(100% - 40px); border:0;
           visibility:hidden; }}
  .view.active {{ visibility:visible; }}
</style>
</head><body>
<div id="tabs" role="tablist">
  <button role="tab" id="tab-map" data-view="view-map" aria-selected="true">{esc(map_tab)}</button>
  <button role="tab" id="tab-route" data-view="view-route" aria-selected="false">{esc(graph_tab)}</button>
</div>
<iframe id="view-map" class="view active" title="{esc(map_tab)}" allow="geolocation" srcdoc="{esc(map_html)}"></iframe>
<iframe id="view-route" class="view" title="{esc(graph_tab)}" srcdoc="{esc(graph_html)}"></iframe>
<script>
(function() {{
  const tabs = Array.from(document.querySelectorAll('#tabs button'));
  function show(viewId) {{
    tabs.forEach(t => {{
      const on = t.dataset.view === viewId;
      t.setAttribute('aria-selected', on ? 'true' : 'false');
      document.getElementById(t.dataset.view).classList.toggle('active', on);
    }});
    try {{ history.replaceState(null, '', '#' + viewId.replace('view-', '')); }} catch (e) {{}}
    const frame = document.getElementById(viewId);
    try {{ frame.contentWindow.focus(); }} catch (e) {{}}
  }}
  tabs.forEach(t => t.addEventListener('click', () => show(t.dataset.view)));
  if (location.hash === '#route') show('view-route');
}})();
</script>
</body></html>
'''


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--xlsx', default=DEFAULT_XLSX, help='input workbook')
    ap.add_argument('--output', default=DEFAULT_OUTPUT, help='output HTML file (overwritten)')

    mg = ap.add_argument_group('map tab')
    mg.add_argument('--map-title', default='Geographic Spine Map')
    mg.add_argument('--no-tube-primary', action='store_true',
                    help='disable London-Underground-style angle snapping for Primary roads (on by default)')
    mg.add_argument('--no-tube-secondary', action='store_true',
                    help='disable London-Underground-style angle snapping for Secondary roads (on by default)')

    rg = ap.add_argument_group('route tab')
    rg.add_argument('--from', dest='start_name', default=DEFAULT_FROM)
    rg.add_argument('--to', dest='end_name', default=DEFAULT_TO)
    rg.add_argument('--margin', type=float, default=DEFAULT_MARGIN,
                    help='max. relatieve omrijding voor alternatieve takken t.o.v. de kortste route PER TRAJECT (0.5 = 50%%)')
    rg.add_argument('--branches-per-leg', type=int, default=4,
                    help='max. aantal alternatieve takken per traject tussen twee waypoints (leesbaarheid)')
    rg.add_argument('--min-novel-km', type=float, default=MIN_NOVEL_KM,
                    help='minimale hoeveelheid nieuwe/afwijkende km die een tak moet toevoegen om mee te tellen')
    rg.add_argument('--via', default=None,
                    help="komma-gescheiden lijst van junction-namen die de hoofdroute verplicht moet passeren, "
                         "in volgorde (bijv. 'Venlo,Koblenz,Hockenheim,Karlsruhe,Ulm/Elchingen,Grenztunnel Fussen')")
    rg.add_argument('--orientation', choices=['vertical', 'horizontal'], default='vertical')
    rg.add_argument('--branch', action='append', default=None,
                    help="komma-gescheiden waypoint-keten voor een expliciete extra tak, bijv. "
                         "'Hockenheim,Weinsberg,Würzburg-West,Feuchtwangen/Crailsheim,Ulm/Elchingen'. "
                         "Mag meerdere keren opgegeven worden; vervangt DEFAULT_BRANCHES.")
    rg.add_argument('--no-default-branches', action='store_true',
                    help='teken de vaste takken uit DEFAULT_BRANCHES niet')
    rg.add_argument('--route-title', default=None)
    args = ap.parse_args()

    tube_style = set()
    if not args.no_tube_primary:
        tube_style.add('Primary')
    if not args.no_tube_secondary:
        tube_style.add('Secondary')
    map_html, map_stats = build_map(xlsx_path=args.xlsx, title=args.map_title,
                                    tube_style_hierarchies=tube_style)

    via = [v.strip() for v in args.via.split(',')] if args.via else []
    if args.branch is not None:
        explicit_branches = [[n.strip() for n in b.split(',')] for b in args.branch]
    elif args.no_default_branches:
        explicit_branches = []
    else:
        explicit_branches = None   # DEFAULT_BRANCHES
    graph_html, graph_stats = build_graph_page(
        args.xlsx, args.start_name, args.end_name, args.margin, args.route_title,
        args.branches_per_leg, args.min_novel_km, via, args.orientation, explicit_branches)

    page = combine_pages(map_html, graph_html, map_tab='Kaart', graph_tab=graph_stats['title'],
                         page_title=args.map_title)
    with open(args.output, 'w', encoding='utf-8') as f:
        f.write(page)

    print(json.dumps({'output': args.output, 'map': map_stats, 'route': graph_stats},
                     indent=2, default=str, ensure_ascii=False))


if __name__ == '__main__':
    main()
