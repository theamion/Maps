#!/usr/bin/env python3
"""
osm_point_leg_distances.py — Verify the real driving distance from a
segment's From-junction to each point on it (tankstation by default),
between consecutive points, and from the last point to the To-junction —
so that From-junction -> point -> point -> To-junction adds up to a real,
checked number instead of an eyeballed PositionOnEdge fraction.

WHY THIS IS TRICKIER THAN JUST CALLING OSRM ONCE PER LEG
----------------------------------------------------------
Two failure modes showed up doing this by hand for Helenaveen/Illertal/
Innerbraz, and this script exists specifically to catch both:

  1. ONE-SIDED ACCESS. A `Sides=1` fuel station only has a driveway off
     ONE carriageway. Ask a router for the distance in the direction that
     doesn't have access, and it has to loop via a distant interchange to
     approach from the right side - a real, physically-correct distance,
     but not the corridor position you actually want. Symptom: the two
     directions of the SAME leg disagree by a lot, and the two legs of a
     point together add up to well more than the segment's real
     junction-to-junction distance.
  2. BAD SNAP. A point's stored lat/lon can sit a little off the actual
     motorway (an address geocode, a Holidays.xlsx coordinate) and OSRM
     silently snaps it onto whatever's nearest and drivable - a local
     street, a service road - which can make BOTH directions of a leg
     implausibly long, sometimes even longer than the whole segment.

So for every point this script:
  - first snaps the point's coordinate onto the ACTUAL named road geometry
    for that segment (same Overpass road-fetch this project's junction
    checker uses), not just wherever OSRM's nearest-neighbour lands it;
  - queries OSRM in BOTH directions for every leg and keeps the smaller;
  - checks the sum of all legs against the segment's real (OSRM) direct
    junction-to-junction distance;
  - if the sum blows that budget because of a direction-asymmetric leg,
    trusts the smaller (accessible) direction and derives the other leg
    by subtraction, tagging it ONE_SIDED_DERIVED;
  - if the overshoot can't be pinned on one clearly asymmetric leg, tags
    it NEEDS_REVIEW (with every measured number attached) instead of
    quietly writing a wrong one.

It only ADDS "OSM ..." columns to the Points tab - PositionOnEdge and every
other existing column are left untouched. By default it reads and writes
junctions_topology_v5.xlsx next to this script (in place); re-running
refills those columns instead of adding a second set. mapmaking.py uses
'OSM DistanceFromStart (km)' for the distances between fuel stations in
the route diagram.

RATE LIMITING / RESUMABILITY
-----------------------------
Same spirit as osm_verify_junctions.py / osm_segment_distances.py:
  - a small delay between the (cheap, high-limit) OSRM calls
  - the usual --delay / --batch-size / --batch-pause around the
    (heavier) Overpass road-snap call, one per point
  - a JSON cache written after every point, so Ctrl+C / a dropped
    connection just means re-running the script to pick up where it
    left off - already-cached points and edges are not re-queried, except
    those that ended in ERROR or SEGMENT_UNREACHABLE (retried)

Usage:
    pip3 install requests
    python3 osm_point_leg_distances.py

    # only fuel stations (the default), small test run into a copy first:
    python3 osm_point_leg_distances.py --output test.xlsx --cache test_cache.json --limit 5

    # every point type instead of just tankstations:
    python3 osm_point_leg_distances.py --cache osm_point_leg_cache_all.json --category ""
"""
import argparse
import json
import math
import os
import re
import time
from collections import defaultdict

import requests
from openpyxl import load_workbook

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSX = os.path.join(SCRIPT_DIR, "junctions_topology_v5.xlsx")
DEFAULT_CACHE = os.path.join(SCRIPT_DIR, "osm_point_leg_cache.json")
RETRY_STATUSES = ("ERROR", "SEGMENT_UNREACHABLE")

OVERPASS_MIRRORS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
]
OSRM_SERVERS = ["https://router.project-osrm.org"]

HEADERS = {"User-Agent": "holiday-route-db-point-leg-check/1.0 (personal hobby project)"}

# How much a leg-sum is allowed to exceed the real junction-to-junction
# distance before we start suspecting a problem (a real service area
# access ramp normally adds a modest, single-digit-percent detour).
DEFAULT_TOLERANCE_PCT = 15.0
# How lopsided the two directions of ONE leg have to be before we call it
# a one-sided-access asymmetry rather than routing noise.
ASYMMETRY_RATIO = 1.5


# --------------------------------------------------------------------------
# Small geometry / country helpers (same approach as osm_verify_junctions.py)
# --------------------------------------------------------------------------

def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


COUNTRY_WORD_TO_CODE = {
    "NL": "NL", "NETHERLANDS": "NL",
    "BE": "BE", "BELGIUM": "BE",
    "DE": "DE", "GERMANY": "DE",
    "AT": "AT", "AUSTRIA": "AT",
    "LU": "LU", "LUXEMBOURG": "LU",
    "FR": "FR", "FRANCE": "FR",
    "CH": "CH", "SWITZERLAND": "CH",
}


def road_token_from_field(road_field):
    """Segments.Road can be 'A67' or 'A67 (NL)' or similar - pull out just
    the bare road number/ref and any country annotation."""
    if not road_field:
        return None, None
    m = re.match(r"\s*([A-Za-z]{1,2}\d{1,3})\s*(?:\(([^)]*)\))?", str(road_field))
    if not m:
        return str(road_field).strip(), None
    token = m.group(1).upper()
    country = None
    if m.group(2):
        for word in re.split(r"[,\s/]+", m.group(2).upper()):
            if word in COUNTRY_WORD_TO_CODE:
                country = COUNTRY_WORD_TO_CODE[word]
                break
    return token, country


def expand_ref_variants(token, country):
    m = re.match(r"^([A-Z]{1,2})(\d{1,3})$", token)
    ref_variants = [token]
    name_variants = []
    if m:
        prefix, num = m.groups()
        ref_variants.append(f"{prefix}\\s*{num}")
        if country == "DE" and prefix == "A":
            ref_variants += [f"BAB\\s*{num}", f"BAB{num}"]
            name_variants.append(f"Bundesautobahn\\s*{num}\\b")
    else:
        ref_variants.append(re.escape(token))
    seen, ref_out = set(), []
    for v in ref_variants:
        if v not in seen:
            seen.add(v)
            ref_out.append(v)
    return ref_out, name_variants


def fetch_road_ways(token, country, lat, lon, radius_m, session, timeout=30):
    """Same query as osm_verify_junctions.py's road-crossing fallback, but
    used here to get the road's own geometry to snap a POINT onto, rather
    than to intersect two roads."""
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


def nearest_point_on_polyline(lat, lon, poly):
    """poly: list of (lat, lon). Returns (snap_lat, snap_lon, dist_km) for
    the closest point on the polyline to (lat, lon), projecting onto each
    segment in flat-earth-locally coordinates (fine at this scale)."""
    if len(poly) < 2:
        return None
    cos_lat = math.cos(math.radians(lat))
    px, py = lon * cos_lat, lat
    best = None
    for i in range(len(poly) - 1):
        alat, alon = poly[i]
        blat, blon = poly[i + 1]
        ax, ay = alon * cos_lat, alat
        bx, by = blon * cos_lat, blat
        dx, dy = bx - ax, by - ay
        if dx == 0 and dy == 0:
            continue
        t = ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)
        t = max(0.0, min(1.0, t))
        qx, qy = ax + t * dx, ay + t * dy
        qlat, qlon = qy, qx / cos_lat
        d = haversine_km(lat, lon, qlat, qlon)
        if best is None or d < best[2]:
            best = (qlat, qlon, d)
    return best


def snap_to_road(lat, lon, road_field, session, radii_m):
    """Try to snap (lat, lon) onto the actual road geometry for this
    segment's Road field, at increasing search radii. Returns a dict with
    status SNAPPED (+ snapped lat/lon/dist) or NO_ROAD_GEOMETRY (falls
    back to the original coordinate, unchanged, with a warning)."""
    token, country = road_token_from_field(road_field)
    if not token:
        return dict(status="NO_ROAD_FIELD", snap_lat=lat, snap_lon=lon, snap_dist_km=None)
    for radius_m in radii_m:
        elements = fetch_road_ways(token, country, lat, lon, radius_m, session)
        time.sleep(0.4)
        ways = [el for el in elements if el.get("type") == "way" and el.get("geometry")]
        if not ways:
            continue
        best = None
        for way in ways:
            poly = [(p["lat"], p["lon"]) for p in way["geometry"]]
            hit = nearest_point_on_polyline(lat, lon, poly)
            if hit and (best is None or hit[2] < best[2]):
                best = hit
        if best:
            return dict(status="SNAPPED", snap_lat=best[0], snap_lon=best[1],
                        snap_dist_km=round(best[2], 4), search_radius_km=radius_m / 1000)
    return dict(status="NO_ROAD_GEOMETRY", snap_lat=lat, snap_lon=lon, snap_dist_km=None)


# --------------------------------------------------------------------------
# OSRM
# --------------------------------------------------------------------------

def osrm_km(lat1, lon1, lat2, lon2, session, timeout=20):
    for base in OSRM_SERVERS:
        url = f"{base}/route/v1/driving/{lon1},{lat1};{lon2},{lat2}?overview=false"
        try:
            resp = session.get(url, headers=HEADERS, timeout=timeout)
            if resp.status_code == 200:
                data = resp.json()
                if data.get("code") == "Ok" and data.get("routes"):
                    return data["routes"][0]["distance"] / 1000.0, None
                return None, data.get("code", "NO_ROUTE")
            if resp.status_code in (429, 502, 503, 504):
                time.sleep(3)
                continue
        except requests.RequestException as e:
            return None, str(e)
    return None, "OSRM unreachable"


def leg_both_directions(latA, lonA, latB, lonB, session, osrm_delay):
    """Query OSRM in both directions for one leg, return the smaller as
    the leg distance plus both raw values for transparency."""
    fwd, err_fwd = osrm_km(latA, lonA, latB, lonB, session)
    time.sleep(osrm_delay)
    back, err_back = osrm_km(latB, lonB, latA, lonA, session)
    time.sleep(osrm_delay)
    vals = [v for v in (fwd, back) if v is not None]
    if not vals:
        return dict(status="ERROR", detail=err_fwd or err_back)
    chosen = min(vals)
    ratio = (max(vals) / min(vals)) if len(vals) == 2 and min(vals) > 0 else 1.0
    return dict(status="OK", km=round(chosen, 3), fwd_km=fwd, back_km=back, ratio=round(ratio, 2))


# --------------------------------------------------------------------------
# Main per-edge processing
# --------------------------------------------------------------------------

def process_edge(edge_id, seg, pts, junctions, session, radii_m, osrm_delay, tolerance_pct):
    """pts: list of point dicts on this edge, already sorted by their
    existing PositionOnEdge (our best prior guess at order along the
    road). Returns a dict {point_id: result_dict}."""
    from_j = junctions[seg["from_id"]]
    to_j = junctions[seg["to_id"]]

    cap = leg_both_directions(from_j["lat"], from_j["lon"], to_j["lat"], to_j["lon"], session, osrm_delay)
    if cap["status"] != "OK":
        return {p["id"]: dict(status="SEGMENT_UNREACHABLE", detail=cap.get("detail")) for p in pts}
    cap_km = cap["km"]

    # snap every point onto the segment's own road first
    for p in pts:
        p["_snap"] = snap_to_road(p["lat"], p["lon"], seg["road"], session, radii_m)

    # chain: [from-junction, point1, point2, ..., to-junction]
    anchors = (
        [dict(kind="junction", id=seg["from_id"], lat=from_j["lat"], lon=from_j["lon"])]
        + [dict(kind="point", id=p["id"], lat=p["_snap"]["snap_lat"], lon=p["_snap"]["snap_lon"]) for p in pts]
        + [dict(kind="junction", id=seg["to_id"], lat=to_j["lat"], lon=to_j["lon"])]
    )

    legs = []
    for i in range(len(anchors) - 1):
        a, b = anchors[i], anchors[i + 1]
        legs.append(leg_both_directions(a["lat"], a["lon"], b["lat"], b["lon"], session, osrm_delay))

    measured = [leg.get("km") for leg in legs]
    results = {}

    if all(v is not None for v in measured):
        total = sum(measured)
        over_pct = (total - cap_km) / cap_km * 100 if cap_km else 0

        if over_pct <= tolerance_pct:
            # every leg is trustworthy as measured
            cum = 0.0
            for i, p in enumerate(pts):
                leg_in, leg_out = legs[i], legs[i + 1]
                cum += leg_in["km"]
                results[p["id"]] = dict(
                    status="OK", distance_from_start_km=round(cum, 3),
                    leg_from_prev_km=leg_in["km"], leg_to_next_km=leg_out["km"],
                    snap=p["_snap"], cap_km=cap_km, sum_km=round(total, 3),
                    over_pct=round(over_pct, 1),
                    detail="beide benen binnen tolerantie t.o.v. junction-to-junction afstand",
                )
        else:
            # The sum overshoots the real junction-to-junction distance by
            # more than the tolerance. Two DIFFERENT things can cause
            # that, and they need different follow-up, so we distinguish
            # them instead of guessing:
            #
            #   (a) ONE leg alone is dramatically direction-asymmetric
            #       (its own fwd/back ratio >= ASYMMETRY_RATIO) while
            #       every other leg is close to symmetric. That is the
            #       one-sided-access signature: you can only reach the
            #       point from one carriageway, so the "wrong" direction
            #       has to loop via a distant interchange. In that
            #       specific, single-point, single-culprit-leg case the
            #       fix is unambiguous - trust the small/consistent
            #       direction of that leg and derive the other leg by
            #       subtraction so the two still add up to cap_km.
            #
            #   (b) Nothing is that clear-cut: no single leg shows a big
            #       fwd/back gap (both directions of both legs are
            #       roughly consistent with each other), yet the total
            #       still doesn't fit under the junction-to-junction
            #       distance. That is NOT the same failure as (a) - it
            #       showed up for real on Helenaveen, whose coordinate is
            #       a verified street address right by the A67 (snap
            #       distance a few dozen metres), not a bad geocode. The
            #       honest reading is: this one-sided station genuinely
            #       needs a bigger detour to enter/exit than a "normal"
            #       service area, on every approach, not just one. It
            #       could just as easily be a genuinely bad snap (a
            #       coordinate sitting on a parallel local road, as
            #       happened with Innerbraz) - the snap distance recorded
            #       alongside this result is the clue to tell those two
            #       apart, but it's a human call either way, so this is
            #       flagged rather than resolved automatically.
            first_leg, last_leg = legs[0], legs[-1]
            first_asym = first_leg.get("ratio", 1.0) >= ASYMMETRY_RATIO
            last_asym = last_leg.get("ratio", 1.0) >= ASYMMETRY_RATIO
            other_legs_symmetric = all(
                leg.get("ratio", 1.0) < ASYMMETRY_RATIO for leg in legs if leg not in (first_leg, last_leg)
            )
            any_leg_alone_too_big = any(v is not None and v > cap_km for v in measured)
            single_culprit = len(pts) == 1 and other_legs_symmetric and not any_leg_alone_too_big and (first_asym != last_asym)

            if single_culprit:
                p = pts[0]
                if first_asym:
                    leg_out = last_leg["km"]
                    leg_in = round(cap_km - leg_out, 3)
                else:
                    leg_in = first_leg["km"]
                    leg_out = round(cap_km - leg_in, 3)
                results[p["id"]] = dict(
                    status="ONE_SIDED_DERIVED", distance_from_start_km=round(leg_in, 3),
                    leg_from_prev_km=leg_in, leg_to_next_km=leg_out,
                    snap=p["_snap"], cap_km=cap_km, sum_km=round(total, 3), over_pct=round(over_pct, 1),
                    detail=(f"één been wijkt sterk per richting af (ratio been-in={first_leg.get('ratio')}, "
                            f"been-uit={last_leg.get('ratio')}), de andere niet - eenduidig eenzijdig "
                            f"(Sides=1) toegangspatroon; betrouwbare richting gebruikt, andere been "
                            f"afgeleid tot de som weer klopt met de junction-to-junction afstand ({cap_km} km)"),
                )
            else:
                for i, p in enumerate(pts):
                    leg_in, leg_out = legs[i], legs[i + 1]
                    results[p["id"]] = dict(
                        status="NEEDS_REVIEW", distance_from_start_km=None,
                        leg_from_prev_km=leg_in.get("km"), leg_to_next_km=leg_out.get("km"),
                        leg_from_prev_ratio=leg_in.get("ratio"), leg_to_next_ratio=leg_out.get("ratio"),
                        snap=p["_snap"], cap_km=cap_km, sum_km=round(total, 3), over_pct=round(over_pct, 1),
                        detail=(f"som van de benen ({round(total,3)} km) overschrijdt de junction-to-junction "
                                f"afstand ({cap_km} km) met {round(over_pct,1)}%, zonder één duidelijke "
                                f"schuldige richting - kan een verkeerde snap zijn (kijk naar de snap-afstand "
                                f"tot de weg: {p['_snap'].get('snap_dist_km')} km - groot = waarschijnlijk "
                                f"fout coordinaat) of een eenzijdig station dat op ELKE aanrijroute een "
                                f"grotere omweg nodig heeft dan de tolerantie toestaat (kleine snap-afstand, "
                                f"zoals bij Helenaveen). Geen positie automatisch bepaald - controleer zelf."),
                    )
    else:
        for p in pts:
            results[p["id"]] = dict(status="ERROR", detail="een van de benen kon niet bevraagd worden")

    return results


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", default=DEFAULT_XLSX)
    ap.add_argument("--output", default=None,
                    help="workbook to write the OSM columns to (default: the --xlsx workbook itself)")
    ap.add_argument("--cache", default=DEFAULT_CACHE)
    ap.add_argument("--category", default="tankstation",
                    help="only process Points whose Category matches this (case-insensitive); "
                         "pass --category \"\" to process every point type")
    ap.add_argument("--radii-km", default="0.3,0.8,1.5,3",
                    help="comma-separated search radii to try when snapping a point onto its road (km)")
    ap.add_argument("--delay", type=float, default=1.5, help="seconds between points (paces the Overpass snap call)")
    ap.add_argument("--osrm-delay", type=float, default=0.3, help="seconds between the (cheaper) OSRM calls")
    ap.add_argument("--batch-size", type=int, default=15)
    ap.add_argument("--batch-pause", type=float, default=10.0)
    ap.add_argument("--tolerance-pct", type=float, default=DEFAULT_TOLERANCE_PCT,
                    help="how many %% the leg-sum may exceed the real junction-to-junction distance "
                         "before it's treated as a problem rather than a normal service-area detour")
    ap.add_argument("--limit", type=int, default=None, help="only process the first N edges (for testing)")
    args = ap.parse_args()
    if args.output is None:
        args.output = args.xlsx

    radii_m = [int(float(r) * 1000) for r in args.radii_km.split(",")]

    cache = {}
    if os.path.exists(args.cache):
        with open(args.cache, "r", encoding="utf-8") as f:
            cache = json.load(f)
        print(f"Cache geladen: {len(cache)} punten al eerder verwerkt.")

    wb = load_workbook(args.xlsx, data_only=True)

    junctions = {}
    for row in wb["Junctions"].iter_rows(min_row=2, values_only=True):
        if row[0]:
            junctions[row[0]] = dict(lat=row[5], lon=row[6])

    segments = {}
    for row in wb["Segments"].iter_rows(min_row=2, values_only=True):
        if row[0]:
            segments[row[0]] = dict(from_id=row[1], to_id=row[3], road=row[5], distance_km=row[6])

    ws_pts = wb["Points"]
    header = [c.value for c in ws_pts[1]]
    idx = {h: i for i, h in enumerate(header)}
    category_filter = args.category.strip().lower() if args.category else None

    points_by_edge = defaultdict(list)
    for row in ws_pts.iter_rows(min_row=2, values_only=True):
        pid = row[idx["Point ID"]]
        eid = row[idx["Edge ID"]]
        if not pid or not eid or eid not in segments:
            continue
        cat = (row[idx["Category"]] or "").strip().lower()
        if category_filter and cat != category_filter:
            continue
        lat, lon = row[idx["Latitude"]], row[idx["Longitude"]]
        if lat is None or lon is None:
            continue
        points_by_edge[eid].append(dict(
            id=pid, lat=lat, lon=lon,
            pos=row[idx["PositionOnEdge"]] if row[idx["PositionOnEdge"]] is not None else 0.5,
        ))
    for eid in points_by_edge:
        points_by_edge[eid].sort(key=lambda p: p["pos"])

    edge_ids = list(points_by_edge.keys())
    if args.limit:
        edge_ids = edge_ids[: args.limit]

    all_point_ids = {p["id"] for eid in edge_ids for p in points_by_edge[eid]}
    def needs_work(pid):
        return pid not in cache or cache[pid].get("status") in RETRY_STATUSES
    todo_edges = [eid for eid in edge_ids if any(needs_work(p["id"]) for p in points_by_edge[eid])]
    print(f"{len(edge_ids)} segmenten met {len(all_point_ids)} punten (categorie: "
          f"{args.category or 'alle'}), {len(todo_edges)} segmenten nog (deels) te verwerken.")

    session = requests.Session()
    for idx_e, eid in enumerate(todo_edges, 1):
        pts = points_by_edge[eid]
        pending = [p for p in pts if needs_work(p["id"])]
        if not pending:
            continue
        seg = segments[eid]
        print(f"[{idx_e}/{len(todo_edges)}] {eid} ({seg['road']}, {len(pts)} punt(en)) ...", flush=True)
        try:
            results = process_edge(eid, seg, pts, junctions, session, radii_m, args.osrm_delay, args.tolerance_pct)
        except Exception as e:
            results = {p["id"]: dict(status="ERROR", detail=str(e)) for p in pts}
        for pid, r in results.items():
            cache[pid] = r
            print(f"    {pid}: {r.get('status')}"
                  + (f" ({r['distance_from_start_km']} km vanaf From-junction)" if r.get("distance_from_start_km") is not None else "")
                  + (f" - {r.get('detail')}" if r.get("status") not in ("OK",) and r.get("detail") else ""))

        with open(args.cache, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)

        time.sleep(args.delay)
        if idx_e % args.batch_size == 0 and idx_e < len(todo_edges):
            print(f"--- batch klaar ({idx_e}/{len(todo_edges)}), pauze {args.batch_pause}s ---")
            time.sleep(args.batch_pause)

    # --- write output workbook with new, additive columns ---
    wb_out = load_workbook(args.xlsx)
    ws_out = wb_out["Points"]
    headers = ["OSM DistanceFromStart (km)", "OSM leg from prev (km)", "OSM leg to next (km)",
               "OSM leg status", "OSM leg detail", "OSM snap status", "OSM snap lat", "OSM snap lon",
               "OSM snap afstand tot origineel (km)", "OSM segment cap (km)", "OSM som benen (km)",
               "OSM som t.o.v. cap (%)"]
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

    out_header = [c.value for c in ws_out[1]]
    pid_col = out_header.index("Point ID")
    for row in ws_out.iter_rows(min_row=2, max_row=ws_out.max_row):
        pid = row[pid_col].value
        if not pid or pid not in cache:
            continue
        r = cache[pid]
        snap = r.get("snap") or {}
        vals = [
            r.get("distance_from_start_km"), r.get("leg_from_prev_km"), r.get("leg_to_next_km"),
            r.get("status"), r.get("detail"),
            snap.get("status"), snap.get("snap_lat"), snap.get("snap_lon"), snap.get("snap_dist_km"),
            r.get("cap_km"), r.get("sum_km"), r.get("over_pct"),
        ]
        for col, v in zip(cols, vals):
            ws_out.cell(row=row[pid_col].row, column=col, value=v)

    wb_out.save(args.output)

    statuses = {}
    for r in cache.values():
        statuses[r.get("status")] = statuses.get(r.get("status"), 0) + 1
    print("\nKlaar. Overzicht:")
    for k, v in sorted(statuses.items()):
        print(f"  {k}: {v}")
    print(f"\nOutput: {args.output}")
    print("Let op: PositionOnEdge en alle andere bestaande kolommen zijn NIET aangepast - alleen de nieuwe")
    print("'OSM ...'-kolommen zijn toegevoegd. NEEDS_REVIEW-punten hebben geen DistanceFromStart gekregen;")
    print("bekijk die zelf (kijk naar de snap-afstand en de vier los gemeten richtingen) voordat je ze overneemt.")


if __name__ == "__main__":
    main()
