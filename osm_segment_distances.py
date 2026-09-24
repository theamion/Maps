#!/usr/bin/env python3
"""
osm_segment_distances.py — Verify Segments' Distance (km) against real
driving distances from OSRM (routing on OpenStreetMap road data).

Run this AFTER osm_verify_junctions.py, and after you've reviewed/applied
any junction coordinate corrections it suggested — this script just reads
whatever Latitude/Longitude is currently in the Junctions sheet of the
workbook you point it at, so it's only as accurate as those coordinates.

HOW A DISTANCE IS MEASURED
--------------------------
A junction's coordinate sits in the middle of an interchange, so routing
straight from it often snaps onto the wrong carriageway (or a ramp or a
crossing road) and the route then detours to turn around. Instead, for
each segment:
  - candidate points are placed around both junctions: on the junction
    itself and 150/400/800 m along the direction of the other junction
    (never more than 35% of the way), each also 30 m to the left and right
    so both carriageways are tried
  - one OSRM `table` request routes every From-candidate to every
    To-candidate; candidates that snap more than 120 m away (onto some
    other road) are dropped
  - for each pair the total is: straight-line distance From junction ->
    its snapped point + route + snapped point -> To junction
  - the shortest total wins: a candidate on the carriageway going the
    wrong way needs a turnaround, so it is never the shortest
  - a total longer than 2x the straight-line ("as the crow flies")
    distance between the junctions is never accepted: if no pair stays
    under that, the segment gets status TOO_LONG and no OSRM distance

It never overwrites the Distance (km) column: it only ADDS new columns
(OSRM route distance, delta vs. current, delta %) to the Segments tab of
the workbook (by default junctions_topology_v5.xlsx next to this script,
written in place), so you can review and decide which segments to
correct. Re-running refills those columns instead of adding new ones.

RATE LIMITING / RESUMABILITY
-----------------------------
Same approach as osm_verify_junctions.py: a small delay between requests,
a longer pause every batch, retries with backoff, and a JSON cache so an
interrupted run can simply be resumed by running the script again
(segments that ended in ERROR are retried).

Usage:
    pip3 install requests
    python3 osm_segment_distances.py

    # measure and copy the OSRM distances into Distance (km) in one go:
    python3 osm_segment_distances.py --apply

    # small test run first, into a copy:
    python3 osm_segment_distances.py --output test.xlsx --cache test_seg_cache.json --limit 10
"""
import argparse
import json
import math
import os
import time

import requests
from openpyxl import load_workbook

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSX = os.path.join(SCRIPT_DIR, "junctions_topology_v5.xlsx")
DEFAULT_CACHE = os.path.join(SCRIPT_DIR, "osm_segment_cache.json")

OSRM_SERVERS = [
    "https://router.project-osrm.org",
]

HEADERS = {"User-Agent": "holiday-route-db-segment-check/1.0 (personal hobby project)"}

METHOD = 2                       # cache entries from another method are recomputed
ALONG_M = (150, 400, 800)        # candidate distances along the segment direction
MAX_ALONG_FRACTION = 0.35        # ...but never further than this part of the segment
SIDE_M = (-30, 0, 30)            # left / on / right of the direction of travel
MAX_SNAP_M = 120                 # a candidate that snaps further away is on another road
MAX_CROW_RATIO = 2.0             # never accept a route this many times the straight line


def haversine_m(lat1, lon1, lat2, lon2):
    p1, p2 = math.radians(lat1), math.radians(lat2)
    a = (math.sin((p2 - p1) / 2) ** 2
         + math.cos(p1) * math.cos(p2) * math.sin(math.radians(lon2 - lon1) / 2) ** 2)
    return 2 * 6371000 * math.asin(math.sqrt(a))


def offset(lat, lon, east_m, north_m):
    return (lat + north_m / 111320.0,
            lon + east_m / (111320.0 * math.cos(math.radians(lat))))


def candidates(lat, lon, toward_lat, toward_lon, travel_e, travel_n, crow_m):
    """Points around a junction: the junction itself, plus points along the
    direction of the other junction, each on the left, the line and the
    right of the direction of travel (travel_e/n, unit vector)."""
    to_e = (toward_lon - lon) * 111320.0 * math.cos(math.radians(lat))
    to_n = (toward_lat - lat) * 111320.0
    norm = math.hypot(to_e, to_n) or 1.0
    to_e, to_n = to_e / norm, to_n / norm
    right_e, right_n = travel_n, -travel_e       # right-hand normal of the travel direction
    pts = [(lat, lon)]
    for along in ALONG_M:
        if along > crow_m * MAX_ALONG_FRACTION:
            break
        for side in SIDE_M:
            pts.append(offset(lat, lon, to_e * along + right_e * side, to_n * along + right_n * side))
    return pts


def osrm_table(sources, destinations, session, timeout=30):
    """Route every source to every destination in one OSRM table request.
    Returns (distances[i][j] in m or None, source snaps, destination snaps,
    error); a snap is (lat, lon, snap distance in m)."""
    coords = ";".join(f"{lon:.6f},{lat:.6f}" for lat, lon in sources + destinations)
    n = len(sources)
    params = (f"sources={';'.join(map(str, range(n)))}"
              f"&destinations={';'.join(map(str, range(n, n + len(destinations))))}"
              f"&annotations=distance")
    last_err = None
    for base in OSRM_SERVERS:
        for attempt in range(3):
            try:
                resp = session.get(f"{base}/table/v1/driving/{coords}?{params}", headers=HEADERS,
                                   timeout=timeout)
                if resp.status_code == 200:
                    data = resp.json()
                    if data.get("code") != "Ok":
                        return None, None, None, data.get("code", "NO_ROUTE")
                    snap = lambda w: (w["location"][1], w["location"][0], w.get("distance", 0.0))
                    return (data["distances"], [snap(w) for w in data["sources"]],
                            [snap(w) for w in data["destinations"]], None)
                last_err = f"HTTP {resp.status_code}"
                if resp.status_code in (429, 502, 503, 504):
                    time.sleep(3 * (attempt + 1))
                    continue
                break
            except requests.RequestException as e:
                last_err = str(e)
                time.sleep(2 * (attempt + 1))
    return None, None, None, last_err or "all OSRM servers failed"


def measure_segment(ju_from, ju_to, session):
    """Shortest plausible driving distance between two junctions (see the
    module docstring). Returns a cache entry (without the current-distance
    comparison)."""
    crow_m = haversine_m(ju_from["lat"], ju_from["lon"], ju_to["lat"], ju_to["lon"])
    if crow_m < 1:
        return dict(status="SAME_POINT", crow_km=0.0)
    travel_e = (ju_to["lon"] - ju_from["lon"]) * 111320.0 * math.cos(math.radians(ju_from["lat"]))
    travel_n = (ju_to["lat"] - ju_from["lat"]) * 111320.0
    norm = math.hypot(travel_e, travel_n)
    travel_e, travel_n = travel_e / norm, travel_n / norm

    src = candidates(ju_from["lat"], ju_from["lon"], ju_to["lat"], ju_to["lon"], travel_e, travel_n, crow_m)
    dst = candidates(ju_to["lat"], ju_to["lon"], ju_from["lat"], ju_from["lon"], travel_e, travel_n, crow_m)
    dist, src_snaps, dst_snaps, err = osrm_table(src, dst, session)
    if err:
        return dict(status="ERROR", detail=err, crow_km=round(crow_m / 1000, 2))

    best = None
    best_rejected = None
    for i, (slat, slon, ssnap) in enumerate(src_snaps):
        if ssnap > MAX_SNAP_M:
            continue
        lead_in = haversine_m(ju_from["lat"], ju_from["lon"], slat, slon)
        for j, (dlat, dlon, dsnap) in enumerate(dst_snaps):
            if dsnap > MAX_SNAP_M or dist[i][j] is None:
                continue
            total = lead_in + dist[i][j] + haversine_m(dlat, dlon, ju_to["lat"], ju_to["lon"])
            if total > MAX_CROW_RATIO * crow_m:
                if best_rejected is None or total < best_rejected:
                    best_rejected = total
                continue
            if best is None or total < best[0]:
                best = (total, i, j)

    entry = dict(crow_km=round(crow_m / 1000, 2), candidates=f"{len(src)}x{len(dst)}")
    if best is None:
        entry["status"] = "TOO_LONG" if best_rejected is not None else "NO_ROUTE"
        if best_rejected is not None:
            entry["detail"] = (f"kortste gevonden {best_rejected / 1000:.1f} km = "
                               f"{best_rejected / crow_m:.1f}x hemelsbreed")
        return entry
    total, i, j = best
    entry.update(status="OK", osrm_km=round(total / 1000, 1), crow_ratio=round(total / crow_m, 2),
                 detail=f"van-punt {i}, naar-punt {j}")
    return entry


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", default=DEFAULT_XLSX)
    ap.add_argument("--output", default=None,
                    help="workbook to write the OSRM columns to (default: the --xlsx workbook itself)")
    ap.add_argument("--cache", default=DEFAULT_CACHE)
    ap.add_argument("--delay", type=float, default=1.2, help="seconds between segments")
    ap.add_argument("--batch-size", type=int, default=25)
    ap.add_argument("--batch-pause", type=float, default=10.0)
    ap.add_argument("--limit", type=int, default=None, help="only process the first N segments (for testing)")
    ap.add_argument("--apply", action="store_true",
                    help="also copy the OSRM distance into 'Distance (km)' for every segment with status OK "
                         "(the column mapmaking.py uses)")
    args = ap.parse_args()
    if args.output is None:
        args.output = args.xlsx

    cache = {}
    if os.path.exists(args.cache):
        with open(args.cache, "r", encoding="utf-8") as f:
            cache = json.load(f)
        print(f"Cache geladen: {len(cache)} segmenten al eerder opgezocht.")

    wb = load_workbook(args.xlsx, data_only=True)
    jws = wb["Junctions"]
    junctions = {}
    for row in jws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            junctions[row[0]] = dict(name=row[1], lat=row[5], lon=row[6])

    sws = wb["Segments"]
    rows = list(sws.iter_rows(min_row=2, values_only=False))
    if args.limit:
        rows = rows[: args.limit]

    session = requests.Session()
    # redo segments that failed with ERROR, and results from an older method
    def needs_work(eid):
        r = cache.get(eid)
        return r is None or r.get("status") == "ERROR" or r.get("method") != METHOD
    todo = [r for r in rows if r[0].value and needs_work(r[0].value)]
    print(f"{len(rows)} segmenten totaal, {len(todo)} nog te verifiëren via OSRM.")

    for idx, row in enumerate(todo, 1):
        eid = row[0].value
        from_id, to_id = row[1].value, row[3].value
        cur_dist = row[6].value
        ju_from = junctions.get(from_id)
        ju_to = junctions.get(to_id)
        if not ju_from or not ju_to or ju_from["lat"] is None or ju_to["lat"] is None:
            cache[eid] = dict(status="MISSING_COORDS", method=METHOD)
        else:
            print(f"[{idx}/{len(todo)}] {eid} {ju_from['name']} -> {ju_to['name']} ...", end=" ", flush=True)
            entry = measure_segment(ju_from, ju_to, session)
            entry["method"] = METHOD
            entry["current_km"] = cur_dist
            if entry["status"] == "OK":
                km = entry["osrm_km"]
                entry["delta_km"] = None if cur_dist in (None, 0) else round(km - cur_dist, 1)
                entry["delta_pct"] = None if not cur_dist else round((km - cur_dist) / cur_dist * 100, 1)
                print(f"OSRM {km:.1f} km ({entry['crow_ratio']}x hemelsbreed {entry['crow_km']} km; "
                      f"huidig {cur_dist}, Δ {entry['delta_km']})")
            else:
                print(entry["status"], entry.get("detail", ""))
            cache[eid] = entry

        with open(args.cache, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)

        time.sleep(args.delay)
        if idx % args.batch_size == 0 and idx < len(todo):
            print(f"--- batch klaar ({idx}/{len(todo)}), pauze {args.batch_pause}s ---")
            time.sleep(args.batch_pause)

    # --- write output workbook with new, additive columns ---
    wb_out = load_workbook(args.xlsx)
    ws_out = wb_out["Segments"]
    headers = ["OSRM status", "OSRM route afstand (km)", "Delta vs huidige Distance (km)", "Delta (%)",
               "Hemelsbreed (km)", "OSRM / hemelsbreed", "OSRM opmerking"]
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

    dist_col = existing.get("Distance (km)")
    if args.apply and dist_col is None:
        raise SystemExit("--apply: column 'Distance (km)' not found in the Segments tab")
    applied = 0
    for row in ws_out.iter_rows(min_row=2, max_row=ws_out.max_row):
        eid = row[0].value
        if not eid:
            continue
        r = cache.get(eid, dict(status="NOT_PROCESSED"))
        if args.apply and r.get("status") == "OK" and r.get("osrm_km") is not None:
            ws_out.cell(row=row[0].row, column=dist_col, value=r["osrm_km"])
            applied += 1
        vals = [r.get("status"), r.get("osrm_km"), r.get("delta_km"), r.get("delta_pct"),
                r.get("crow_km"), r.get("crow_ratio"), r.get("detail")]
        for col, v in zip(cols, vals):
            ws_out.cell(row=row[0].row, column=col, value=v)

    wb_out.save(args.output)

    ok = [r for r in cache.values() if r.get("status") == "OK"]
    big_deltas = sorted(
        [r for r in ok if r.get("delta_pct") is not None and abs(r["delta_pct"]) >= 15],
        key=lambda r: -abs(r["delta_pct"]),
    )
    too_long = [r for r in cache.values() if r.get("status") == "TOO_LONG"]
    print(f"\nKlaar. {len(ok)} segmenten succesvol opgehaald.")
    print(f"{len(too_long)} segmenten TOO_LONG: geen route binnen {MAX_CROW_RATIO:g}x hemelsbreed "
          f"— controleer daar de junction-coördinaten of de topologie.")
    print(f"{len(big_deltas)} segmenten wijken >=15% af van de huidige Distance-waarde — bekijk deze het eerst.")
    print(f"\nOutput: {args.output}")
    if args.apply:
        print(f"--apply: Distance (km) bijgewerkt voor {applied} segmenten met status OK.")
    else:
        print("Let op: de Distance (km)-kolom zelf is NIET aangepast, alleen de nieuwe 'OSRM ...'-kolommen toegevoegd.")
        print("Gebruik --apply (of apply_osm_distances.py) om de OSRM-afstanden over te nemen.")


if __name__ == "__main__":
    main()
