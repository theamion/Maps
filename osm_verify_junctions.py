#!/usr/bin/env python3
"""
osm_verify_junctions.py — Verify junction coordinates against real
OpenStreetMap motorway-junction nodes (Overpass API).

WHY THIS EXISTS
----------------
The workbook's Latitude/Longitude for a junction should be the real
interchange ("Kreuz"/"Knooppunt"/"Dreieck") node, not a geocoded village
or town centroid a few km away. This script never invents or guesses a
coordinate. For every junction, in order:

  1. HOLIDAYS SOURCE (status MATCHED_HOLIDAYS_SOURCE) — the original
     Holidays.xlsx "Database" sheet is this whole map's primary source
     and already has hand-curated Locatie/OL/NB rows for junctions,
     exits and border crossings named exactly like this workbook (e.g.
     'Kreuz Walldorf', 'Frankfurter Kreuz'). Matching is approximate on
     purpose (a junction-type word like 'Kreuz'/'Knooppunt'/'Dreieck' is
     stripped from both sides before comparing, so 'Kreuz Aachen' and
     'Aachen' are treated as the same place), with exact/substring/fuzzy
     tiers reported so you can judge confidence. If Holidays.xlsx isn't
     found or gives no match, it falls through to:
  2. OPENSTREETMAP BY NAME (status MATCHED) — search OSM for an actual
     `highway=motorway_junction` node whose name/ref plausibly matches,
     close to the coordinate you already have. If nothing matches:
  3. OPENSTREETMAP BY ROAD CROSSING (status MATCHED_VIA_CROSSING) — read
     the junction's "Roads meeting here" column (e.g. "A2, A27 (NL)"),
     fetch the real road geometry for those road numbers from OSM, and
     compute the actual point where those roads cross.

Only if none of the three find anything is a junction marked NOT_FOUND —
it never falls back to a geocoded town/village centroid at any point.

It only ADDS "OSM ..." columns to the Junctions tab (it never touches
Latitude/Longitude or any other existing column). You review the
proposed coordinates yourself and decide what to accept — consistent
with "never silently repair topology" for this project.

RATE LIMITING / RESUMABILITY
-----------------------------
This calls a free public Overpass API. To be a good citizen (and to
avoid getting temporarily blocked), it:
  - waits `--delay` seconds between requests (default 1.5s)
  - pauses `--batch-pause` seconds every `--batch-size` junctions (default
    20 junctions / 12s pause)
  - retries with backoff on rate-limit/server errors, rotating across a
    small list of public Overpass mirrors
  - writes results to a JSON cache after every junction, so if you
    interrupt the script (Ctrl+C, laptop sleeps, network drops) you can
    just run it again and it will skip everything already resolved
    (junctions that ended in ERROR, e.g. a server timeout, are retried)

By default it reads junctions_topology_v5.xlsx and Holidays.xlsx next to
this script and writes the OSM columns into the Junctions tab of
junctions_topology_v5.xlsx itself. Re-running refills the existing OSM
columns instead of adding new ones. Pass --holidays-xlsx '' to skip the
Holidays step and go straight to OSM.

Usage:
    pip3 install requests
    python3 osm_verify_junctions.py

    # do a small test run first, into a copy:
    python3 osm_verify_junctions.py --output test.xlsx --cache test_cache.json --limit 10
"""
import argparse
import difflib
import json
import math
import os
import re
import time
import unicodedata

import requests
from openpyxl import load_workbook

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSX = os.path.join(SCRIPT_DIR, "junctions_topology_v5.xlsx")
DEFAULT_HOLIDAYS = os.path.join(SCRIPT_DIR, "Holidays.xlsx")
DEFAULT_CACHE = os.path.join(SCRIPT_DIR, "osm_junction_cache.json")

OVERPASS_MIRRORS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
]

# Prefixes that name the *kind* of junction rather than the place — stripped
# off before matching against OSM name/ref tags, in the languages used
# across this map (NL/DE/AT/BE/LU).
KIND_PREFIXES = [
    "kreuz", "dreieck", "autobahnkreuz", "autobahndreieck",
    "ausfahrt", "anschlussstelle", "autobahnausfahrt",
    "knooppunt", "afrit", "aansluiting",
    "échangeur", "sortie",
]

HEADERS = {"User-Agent": "holiday-route-db-junction-check/1.0 (personal hobby project)"}

# Row types in Holidays.xlsx's 'Database' sheet worth matching junctions
# against (i.e. actual interchanges/exits/crossings, not scenery points).
HOLIDAYS_TYPE_WHITELIST = {
    "Knooppunt", "Afslag", "Grensovergang", "Grenspunt", "Bestemming",
    "Rivier Knooppunt", "Stad",
}

# Junction-kind words stripped from EITHER side before comparing (can sit
# at the start — 'Kreuz Aachen' — or the end — 'Frankfurter Kreuz').
JUNCTION_KIND_WORDS = {
    "kreuz", "dreieck", "autobahnkreuz", "autobahndreieck",
    "ausfahrt", "anschlussstelle", "autobahnausfahrt",
    "knooppunt", "afrit", "aansluiting", "afslag",
    "grensovergang", "grenspunt", "echangeur", "sortie",
}


def fold_accents(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def normalize_place(s):
    """Lowercase, strip accents/punctuation, and drop junction-kind words
    wherever they occur, so 'Kreuz Aachen', 'Aachen Kreuz' and 'Aachen'
    all normalize to the same 'aachen' for approximate matching."""
    s = fold_accents(s).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    words = [w for w in s.split() if w not in JUNCTION_KIND_WORDS]
    return " ".join(words).strip()


def load_holidays_index(path):
    """Load Holidays.xlsx's 'Database' sheet into a list of
    (normalized_name, raw_name, lat, lon, type, road) tuples for
    approximate matching. Returns None if the file/sheet isn't there."""
    if not path or not os.path.exists(path):
        return None
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb["Database"]
    except Exception:
        return None
    index = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        loc, lon, lat, _, _, _, _, typ = row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]
        road = row[10] if len(row) > 10 else None
        if not loc or lat is None or lon is None:
            continue
        if typ not in HOLIDAYS_TYPE_WHITELIST:
            continue
        norm = normalize_place(str(loc))
        if not norm:
            continue
        index.append((norm, str(loc), float(lat), float(lon), typ, road))
    return index


def match_holidays(name, cur_lat, cur_lon, holidays_index, fuzzy_threshold=0.84):
    """Approximate-match a junction name against the Holidays.xlsx index.
    Tries every '/'-separated part of the name too (so 'Feuchtwangen/
    Crailsheim' can match a Holidays row named just 'Crailsheim'). Reports
    a confidence tier; when several rows normalize the same, picks the
    one closest to the current stored coordinate."""
    if not holidays_index:
        return None
    terms = [normalize_place(p) for p in re.split(r"[/,]", name)]
    terms = [t for t in terms if t] or [normalize_place(name)]

    def score(norm_row):
        best = 0.0
        tier = None
        for term in terms:
            if term == norm_row:
                return 1.0, "EXACT"
            if term in norm_row or norm_row in term:
                if 0.9 > best:
                    best, tier = 0.9, "SUBSTRING"
            ratio = difflib.SequenceMatcher(None, term, norm_row).ratio()
            if ratio > best:
                best, tier = ratio, "FUZZY"
        return best, tier

    scored = []
    for norm_row, raw_row, lat, lon, typ, road in holidays_index:
        s, tier = score(norm_row)
        if s >= fuzzy_threshold or tier in ("EXACT", "SUBSTRING"):
            d = haversine_km(cur_lat, cur_lon, lat, lon) if (cur_lat is not None and cur_lon is not None) else 0.0
            scored.append((s, tier, d, raw_row, lat, lon, typ, road))
    if not scored:
        return None
    # best confidence first, then closest to the current stored coordinate
    scored.sort(key=lambda t: (-{"EXACT": 2, "SUBSTRING": 1, "FUZZY": 0}[t[1]], t[2]))
    best = scored[0]
    return dict(
        status="MATCHED_HOLIDAYS_SOURCE",
        match_tier=best[1], match_score=round(best[0], 3),
        osm_name=f"{best[3]} ({best[6]}" + (f", {best[7]}" if best[7] else "") + ")",
        osm_lat=best[4], osm_lon=best[5],
        distance_km=round(best[2], 3) if cur_lat is not None else None,
        candidates_found=len(scored),
    )


def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def core_name(name):
    """Strip a leading 'Kreuz '/'Knooppunt '/etc. so 'Kreuz Frankenthal'
    becomes 'Frankenthal' for matching against OSM name/ref tags. Also
    strips anything after a '/' (e.g. 'Feuchtwangen/Crailsheim' -> both
    halves are tried separately by the caller)."""
    n = name.strip()
    low = n.lower()
    for p in KIND_PREFIXES:
        if low.startswith(p + " "):
            n = n[len(p) + 1:]
            break
    return n.strip()


def name_candidates(raw_name):
    """A junction name can hide more than one usable search term
    ('Feuchtwangen/Crailsheim', 'Kreuz Mönchengladbach-Nord')."""
    parts = re.split(r"[/,]", raw_name)
    cands = []
    for p in parts:
        c = core_name(p)
        if c and c not in cands:
            cands.append(c)
    if not cands:
        cands = [raw_name.strip()]
    return cands


ROAD_REF_RE = re.compile(r"\b([A-Z]{1,2}\d{1,3})\b")

COUNTRY_WORD_TO_CODE = {
    "NL": "NL", "NETHERLANDS": "NL",
    "BE": "BE", "BELGIUM": "BE",
    "DE": "DE", "GERMANY": "DE",
    "AT": "AT", "AUSTRIA": "AT",
    "LU": "LU", "LUXEMBOURG": "LU",
    "FR": "FR", "FRANCE": "FR",
    "CH": "CH", "SWITZERLAND": "CH",
}


def parse_roads_meeting(raw, default_country=None):
    """Pull out (road-number token, country hint) pairs from a free-text
    'Roads meeting here' cell like 'A1 (DE), A44' or 'R0 (Brussels ring),
    A1 (BE, toward Antwerpen)'. The country hint comes from a '(DE)' /
    '(Austria)' annotation next to that specific token if present,
    otherwise falls back to `default_country` (normally the junction's
    own Country column) — needed so a German 'A3' can also be searched
    for as 'BAB3' (Bundesautobahn), a Dutch 'A2' is not. Descriptive-only
    cells ('grenspunt') simply yield no tokens."""
    results = []
    seen = set()
    # split into comma-separated pieces (but not on a comma that's inside
    # a parenthetical, e.g. '(BE, toward Antwerpen)') so a '(DE)' only
    # applies to the token(s) actually named in that piece
    for piece in re.split(r",(?![^(]*\))", raw):
        piece_country = None
        cm = re.search(r"\(([^)]*)\)", piece)
        if cm:
            for word in re.split(r"[,\s/]+", cm.group(1).upper()):
                if word in COUNTRY_WORD_TO_CODE:
                    piece_country = COUNTRY_WORD_TO_CODE[word]
                    break
        for m in ROAD_REF_RE.finditer(piece.upper()):
            t = m.group(1)
            if t in seen:
                continue
            seen.add(t)
            results.append((t, piece_country or default_country))
    return results


def expand_ref_variants(token, country):
    """Alternate ways OSM might record the same road, beyond the bare
    number — e.g. a German 'A3' is legally a 'Bundesautobahn 3' and is
    sometimes tagged/named as 'BAB3' / 'BAB 3' rather than plain 'A3'."""
    m = re.match(r"^([A-Z]{1,2})(\d{1,3})$", token)
    ref_variants = [token]
    name_variants = []
    if m:
        prefix, num = m.groups()
        ref_variants.append(f"{prefix}\\s*{num}")  # allow 'A 3' as well as 'A3'
        if country == "DE" and prefix == "A":
            ref_variants += [f"BAB\\s*{num}", f"BAB{num}"]
            name_variants.append(f"Bundesautobahn\\s*{num}\\b")
    else:
        ref_variants.append(re.escape(token))
    # dedupe while preserving order
    seen = set()
    ref_out = []
    for v in ref_variants:
        if v not in seen:
            seen.add(v)
            ref_out.append(v)
    return ref_out, name_variants


def _seg_intersection(p1, p2, p3, p4):
    """Classic 2D line-segment intersection (lat/lon treated as flat
    coordinates — fine at the few-km scale a single junction spans).
    Returns the (lat, lon) intersection point, or None."""
    x1, y1 = p1
    x2, y2 = p2
    x3, y3 = p3
    x4, y4 = p4
    d = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
    if abs(d) < 1e-12:
        return None
    t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / d
    u = ((x1 - x3) * (y1 - y2) - (y1 - y3) * (x1 - x2)) / d
    if -0.02 <= t <= 1.02 and -0.02 <= u <= 1.02:  # small tolerance past segment ends
        return (x1 + t * (x2 - x1), y1 + t * (y2 - y1))
    return None


def fetch_road_ways(token, country, lat, lon, radius_m, session, timeout=30):
    """Fetch geometry of motorway/trunk/primary/secondary/tertiary ways
    matching this road, near (lat, lon) — trying the bare number against
    `ref` and `int_ref` (European E-numbers often live in int_ref), plus
    country-specific alternate namings (e.g. German 'BAB3'/'Bundesautobahn
    3' for 'A3') against `ref`/`name`, all in one query."""
    ref_variants, name_variants = expand_ref_variants(token, country)
    highway_re = "^(motorway|trunk|primary|secondary|tertiary)(_link)?$"
    ref_alt = "|".join(ref_variants)
    clauses = [
        f'way["highway"~"{highway_re}"]["ref"~"(^|;)\\\\s*({ref_alt})(\\\\s*;|$)",i](around:{radius_m},{lat},{lon});',
        f'way["highway"~"{highway_re}"]["int_ref"~"(^|;)\\\\s*({ref_alt})(\\\\s*;|$)",i](around:{radius_m},{lat},{lon});',
    ]
    for nv in name_variants:
        clauses.append(f'way["highway"~"{highway_re}"]["name"~"{nv}",i](around:{radius_m},{lat},{lon});')
    query = f'[out:json][timeout:25];(' + "".join(clauses) + ');out geom;'
    for mirror in OVERPASS_MIRRORS:
        try:
            resp = session.post(mirror, data={"data": query}, headers=HEADERS, timeout=timeout)
            if resp.status_code == 200:
                return resp.json().get("elements", [])
            if resp.status_code in (429, 504, 502, 503):
                time.sleep(3)
                continue
        except requests.RequestException:
            time.sleep(2)
    return []


def find_road_crossing(roads_meeting_raw, default_country, cur_lat, cur_lon, session, radii_m):
    """Fallback for when no named motorway_junction node matches: fetch
    the real OSM geometry of the roads that are recorded as meeting at
    this junction, and compute where their lines actually cross. This is
    still a real, OSM-grounded coordinate — never a guessed town/village
    centroid — just derived geometrically instead of by name lookup."""
    refs = parse_roads_meeting(roads_meeting_raw or "", default_country)
    if len(refs) < 2:
        return dict(status="NOT_FOUND", detail="minder dan 2 herkenbare wegnummers in 'Roads meeting here'")

    for radius_m in radii_m:
        ways_by_ref = {}
        for token, country in refs:
            elements = fetch_road_ways(token, country, cur_lat, cur_lon, radius_m, session)
            time.sleep(0.4)
            ways = [el for el in elements if el.get("type") == "way" and el.get("geometry")]
            if ways:
                ways_by_ref[token] = ways

        if len(ways_by_ref) < 2:
            continue

        candidates = []
        ref_list = list(ways_by_ref.keys())
        for i in range(len(ref_list)):
            for k in range(i + 1, len(ref_list)):
                ref_a, ref_b = ref_list[i], ref_list[k]
                for way_a in ways_by_ref[ref_a]:
                    pts_a = [(p["lon"], p["lat"]) for p in way_a["geometry"]]
                    for way_b in ways_by_ref[ref_b]:
                        pts_b = [(p["lon"], p["lat"]) for p in way_b["geometry"]]
                        for j in range(len(pts_a) - 1):
                            for m in range(len(pts_b) - 1):
                                pt = _seg_intersection(pts_a[j], pts_a[j + 1], pts_b[m], pts_b[m + 1])
                                if pt:
                                    lon_c, lat_c = pt
                                    d = haversine_km(cur_lat, cur_lon, lat_c, lon_c)
                                    candidates.append((d, lat_c, lon_c, ref_a, ref_b))
        if candidates:
            candidates.sort(key=lambda c: c[0])
            best = candidates[0]
            return dict(
                status="MATCHED_VIA_CROSSING",
                osm_lat=best[1], osm_lon=best[2],
                osm_name=f"kruising {best[3]} × {best[4]}",
                distance_km=round(best[0], 3),
                candidates_found=len(candidates), search_radius_km=radius_m / 1000,
            )
    return dict(status="NOT_FOUND",
                detail=f"geen naam-match en geen wegkruising gevonden voor {refs}")


def overpass_query(lat, lon, radius_m, name_term, session, timeout=30):
    name_escaped = name_term.replace('"', '')
    query = (
        f'[out:json][timeout:25];'
        f'('
        f'node["highway"="motorway_junction"]["name"~"{name_escaped}",i](around:{radius_m},{lat},{lon});'
        f'node["highway"="motorway_junction"]["ref"~"{name_escaped}",i](around:{radius_m},{lat},{lon});'
        f')'
        f';out body;'
    )
    last_err = None
    for mirror in OVERPASS_MIRRORS:
        try:
            resp = session.post(mirror, data={"data": query}, headers=HEADERS, timeout=timeout)
            if resp.status_code == 200:
                return resp.json().get("elements", [])
            if resp.status_code in (429, 504, 502, 503):
                last_err = f"{mirror}: HTTP {resp.status_code} (busy), trying next mirror"
                time.sleep(3)
                continue
            last_err = f"{mirror}: HTTP {resp.status_code}"
        except requests.RequestException as e:
            last_err = f"{mirror}: {e}"
            time.sleep(2)
    raise RuntimeError(last_err or "all Overpass mirrors failed")


def verify_one(jid, name, roads_meeting, country, cur_lat, cur_lon, session, radii_m,
                holidays_index=None, holidays_fuzzy_threshold=0.84):
    """1) Try the Holidays.xlsx source first (the map's own original,
    hand-curated data). 2) If that has no match, try increasing search
    radii until a name-matching motorway_junction node is found on OSM.
    3) If nothing matches by name either, fall back to computing the real
    geometric crossing of the roads recorded as meeting here. At every
    step the result is a real, sourced coordinate — never a guessed
    town/village geocode."""
    holidays_match = match_holidays(name, cur_lat, cur_lon, holidays_index, holidays_fuzzy_threshold)
    if holidays_match:
        return holidays_match

    candidates_tried = name_candidates(name)
    for radius_m in radii_m:
        all_elements = []
        for term in candidates_tried:
            try:
                elements = overpass_query(cur_lat, cur_lon, radius_m, term, session)
            except RuntimeError as e:
                return dict(status="ERROR", detail=str(e))
            all_elements.extend(elements)
            time.sleep(0.4)  # be polite between the (few) sub-queries of one junction
        if not all_elements:
            continue
        # dedupe by node id, compute distance to current stored coordinate
        seen = {}
        for el in all_elements:
            seen[el["id"]] = el
        scored = []
        for el in seen.values():
            d = haversine_km(cur_lat, cur_lon, el["lat"], el["lon"])
            tag_name = el.get("tags", {}).get("name") or el.get("tags", {}).get("ref") or ""
            scored.append((d, el["id"], el["lat"], el["lon"], tag_name))
        scored.sort(key=lambda t: t[0])
        best = scored[0]
        return dict(
            status="MATCHED" if len(scored) == 1 else "MULTIPLE_CANDIDATES",
            osm_node_id=best[1], osm_lat=best[2], osm_lon=best[3],
            osm_name=best[4], distance_km=round(best[0], 3),
            candidates_found=len(scored), search_radius_km=radius_m / 1000,
        )
    # no named motorway_junction found anywhere in range — fall back to the
    # real crossing point of the roads recorded as meeting here
    return find_road_crossing(roads_meeting, country, cur_lat, cur_lon, session, radii_m)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", default=DEFAULT_XLSX)
    ap.add_argument("--output", default=None,
                    help="workbook to write the OSM columns to (default: the --xlsx workbook itself)")
    ap.add_argument("--cache", default=DEFAULT_CACHE)
    ap.add_argument("--radii-km", default="15,40", help="comma-separated search radii to try in order (km)")
    ap.add_argument("--delay", type=float, default=1.5, help="seconds between junctions")
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--batch-pause", type=float, default=12.0, help="seconds to pause after each batch")
    ap.add_argument("--limit", type=int, default=None, help="only process the first N junctions (for testing)")
    ap.add_argument("--holidays-xlsx", default=DEFAULT_HOLIDAYS,
                     help="path to the original Holidays.xlsx (primary source, tried before OSM); "
                          "pass --holidays-xlsx '' to skip this step entirely")
    ap.add_argument("--holidays-fuzzy-threshold", type=float, default=0.84,
                     help="minimum fuzzy-match ratio (0-1) to accept a Holidays.xlsx match")
    args = ap.parse_args()
    if args.output is None:
        args.output = args.xlsx

    radii_m = [int(float(r) * 1000) for r in args.radii_km.split(",")]

    holidays_index = load_holidays_index(args.holidays_xlsx) if args.holidays_xlsx else None
    if args.holidays_xlsx:
        if holidays_index is None:
            print(f"Let op: '{args.holidays_xlsx}' niet gevonden of geen 'Database'-sheet — "
                  f"sla deze stap over en ga direct naar OSM.")
        else:
            print(f"Holidays-bron geladen: {len(holidays_index)} bruikbare locaties "
                  f"(Knooppunt/Afslag/Grensovergang/...).")

    cache = {}
    if os.path.exists(args.cache):
        with open(args.cache, "r", encoding="utf-8") as f:
            cache = json.load(f)
        print(f"Cache geladen: {len(cache)} junctions al eerder opgezocht.")

    wb = load_workbook(args.xlsx, data_only=True)
    ws = wb["Junctions"]
    rows = list(ws.iter_rows(min_row=2, values_only=False))
    if args.limit:
        rows = rows[: args.limit]

    session = requests.Session()
    # junctions that failed with a network/server ERROR are retried on the next run
    todo = [r for r in rows if r[0].value and cache.get(r[0].value, {}).get("status", "ERROR") == "ERROR"]
    print(f"{len(rows)} junctions totaal, {len(todo)} nog te verifiëren via Overpass.")

    for idx, row in enumerate(todo, 1):
        jid = row[0].value
        name = row[1].value or ""
        roads_meeting = row[2].value or ""
        country = row[3].value
        lat = row[5].value
        lon = row[6].value
        if lat is None or lon is None:
            cache[jid] = dict(status="NO_COORDS")
        else:
            print(f"[{idx}/{len(todo)}] {jid} {name} ...", end=" ", flush=True)
            try:
                result = verify_one(jid, name, roads_meeting, country, lat, lon, session, radii_m,
                                     holidays_index, args.holidays_fuzzy_threshold)
            except Exception as e:
                result = dict(status="ERROR", detail=str(e))
            cache[jid] = result
            print(result.get("status"), f"(Δ {result.get('distance_km','-')} km)" if "distance_km" in result else "")

        with open(args.cache, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)

        time.sleep(args.delay)
        if idx % args.batch_size == 0 and idx < len(todo):
            print(f"--- batch klaar ({idx}/{len(todo)}), pauze {args.batch_pause}s ---")
            time.sleep(args.batch_pause)

    # --- write output workbook with new, additive columns ---
    wb_out = load_workbook(args.xlsx)
    ws_out = wb_out["Junctions"]
    headers = ["OSM status", "OSM naam", "OSM node", "OSM lat", "OSM lon",
               "OSM afstand tot huidige coord (km)", "OSM kandidaten gevonden", "OSM zoekradius (km)"]
    # reuse the columns from an earlier run, so re-running doesn't add a second set
    existing = {c.value: c.column for c in ws_out[1] if c.value}
    next_col = ws_out.max_column + 1
    cols = []
    for h in headers:
        if h not in existing:
            ws_out.cell(row=1, column=next_col, value=h)
            existing[h] = next_col
            next_col += 1
        cols.append(existing[h])

    for row in ws_out.iter_rows(min_row=2, max_row=ws_out.max_row):
        jid = row[0].value
        if not jid:
            continue
        r = cache.get(jid, dict(status="NOT_PROCESSED"))
        vals = [
            r.get("status"),
            r.get("osm_name"),
            f'https://www.openstreetmap.org/node/{r["osm_node_id"]}' if r.get("osm_node_id") else None,
            r.get("osm_lat"),
            r.get("osm_lon"),
            r.get("distance_km"),
            r.get("candidates_found"),
            r.get("search_radius_km"),
        ]
        for col, v in zip(cols, vals):
            ws_out.cell(row=row[0].row, column=col, value=v)

    wb_out.save(args.output)

    statuses = {}
    for r in cache.values():
        statuses[r.get("status")] = statuses.get(r.get("status"), 0) + 1
    print("\nKlaar. Overzicht:")
    for k, v in sorted(statuses.items()):
        print(f"  {k}: {v}")
    print(f"\nOutput: {args.output}")
    print("Let op: Latitude/Longitude in de originele kolommen zijn NIET aangepast.")
    print("Bekijk zelf de nieuwe 'OSM ...'-kolommen en beslis per junction of je de coördinaat wilt overnemen.")


if __name__ == "__main__":
    main()
