# Maps: holiday route database tools

Scripts that build maps and route diagrams from an Excel workbook of motorway
junctions, roads and segments, plus bridges, fuel stations, rivers and borders.
The network covers the Netherlands, Belgium, Luxembourg, Germany and Austria.

Every script runs without arguments. By default each one reads and writes files
in the script's own folder, whichever directory you run it from.

## Files

| File | Role |
|------|------|
| `junctions_topology_v5.xlsx` | **Active database.** All scripts read it, and the OSM checks write their results into it |
| `Holidays.xlsx` | **Reference.** The original source (`Database` sheet), used by `osm_verify_junctions.py` as its first source for coordinates |
| `junctions_topology_v4.xlsx` | Previous workbook, which includes the research sheets `Tankstations` and `Nieuwe Points` |
| `mapmaking.py` | Map and route diagram in one HTML page with two tabs → `holidays.html` |
| `osm_verify_junctions.py` | Checks junction coordinates against Holidays.xlsx and OpenStreetMap |
| `osm_segment_distances.py` | Checks segment distances against OSRM driving distances |
| `osm_point_leg_distances.py` | Measures each fuel station's distance along its segment with OSRM → `OSM DistanceFromStart (km)` in the Points tab (used by the route diagram) |
| `apply_osm_coords.py`, `apply_osm_distances.py` | Copy reviewed OSM results into `Latitude`/`Longitude` and `Distance (km)` |
| `osm_junction_cache.json`, `osm_segment_cache.json` | Caches for the OSM checks, so runs can resume |
| `generate_map_v2.deprecatedpy`, `generate_graph.deprecatedpy` | Superseded by `mapmaking.py`. Kept for reference only |

## Publishing

The map is published with GitHub Pages at
**https://theamion.github.io/Maps/holidays.html**, from the `main` branch of
https://github.com/theamion/Maps.

A git hook in `githooks/pre-commit` runs `mapmaking.py` before every commit on
`main` and adds the fresh `holidays.html` to that commit. So a normal commit and
push (also from VS Code) updates the page, usually live within a minute or two.

- In a new clone, turn the hook on once with `git config core.hooksPath githooks`.
- Skip it for one commit with `git commit --no-verify`.
- If `mapmaking.py` fails (for example because the workbook is missing), the
  commit is stopped.

The workbooks and OSM caches are in `.gitignore`, so they stay local. Only the
generated page shows their data.

## Requirements

- Python 3
- `openpyxl` (all scripts)
- `networkx` (`mapmaking.py`)
- `requests` (the two `osm_*` scripts)

```bash
pip install openpyxl networkx requests
```

## Quick start

```bash
cd /home/theamion/Python/Maps
python3 osm_verify_junctions.py      # 1. check junction coordinates → Junctions tab
python3 osm_segment_distances.py     # 2. check segment distances → Segments tab
python3 mapmaking.py                 # map + route Vught → Berwang → holidays.html
```

The OSM scripts save into `junctions_topology_v5.xlsx`, so close the workbook in
Excel before running them.

## mapmaking.py: map and route diagram

Builds one HTML page with two tabs:

- **Kaart:** the whole network as a schematic map that stays close to real
  geography. It has zoom, a GPS button, and toggles for bridge names (B),
  distances (Km), points of interest (★) and debug IDs (D).
- **Route:** the route between two places as a "metro board" diagram, like the
  line diagrams on NS departure boards. The main route runs straight, and
  alternative branches fan out and rejoin it. Positions are ordinal hops, not
  geography. Buttons toggle fuel stations, bridges and tunnels (off, on, or on
  with length), segment distances (`Afstanden`, on by default) and points of
  interest (`★ Bezienswaardigheden`, on by default). The distances come from
  `Distance (km)`. On segments with fuel stations, the distance of each stretch
  (junction → station → station → junction, in travel direction) is shown in
  smaller italic teal text, while both `Afstanden` and `Tankstations` are on.
  These use the column `OSM DistanceFromStart (km)` in the Points tab when
  `osm_point_leg_distances.py` has filled it, and otherwise `PositionOnEdge` ×
  the segment's `Distance (km)`. Lines never cross: after the layout, each detour's lane is
  re-chosen to remove crossings.

```bash
python3 mapmaking.py
python3 mapmaking.py --from Vught --to Berwang --via "Venlo,Koblenz"
```

Open `holidays.html` in a browser. Adding `#route` to the address opens the
route tab directly.

| Option | Default | Description |
|--------|---------|-------------|
| `--xlsx` | `junctions_topology_v5.xlsx` | Input workbook |
| `--output` | `holidays.html` | Output HTML file (overwritten) |
| **Kaart tab** | | |
| `--map-title` | `Geographic Spine Map` | Page title |
| `--no-tube-primary` | off | Disable London-Underground-style angle snapping for Primary roads |
| `--no-tube-secondary` | off | Disable angle snapping for Secondary roads |
| **Route tab** | | |
| `--from` / `--to` | `Vught` / `Berwang` | Start and end junction names |
| `--via` | none | Comma-separated junctions the main route must pass, in order |
| `--branch` | `DEFAULT_BRANCHES` | Comma-separated waypoint chain for a fixed branch (can be repeated). Replaces `DEFAULT_BRANCHES` |
| `--no-default-branches` | off | Don't draw the fixed branches from `DEFAULT_BRANCHES` |
| `--margin` | `0.5` | Maximum detour for alternative branches, relative to the shortest route (0.5 = 50%) |
| `--branches-per-leg` | `4` | Maximum automatic branches between two waypoints (on top of the fixed ones) |
| `--min-novel-km` | `3.0` | Minimum km of new road a branch must add |
| `--orientation` | `vertical` | `vertical` or `horizontal` |
| `--route-title` | `Route: <from> → <to>` | Route tab title |

### Which branches the route diagram shows

1. **Main route:** the shortest route from `--from` to `--to` (through `--via`,
   if given).
2. **Fixed branches:** `DEFAULT_BRANCHES` at the top of `mapmaking.py`,
   currently the A3/A7 corridor Kerpen → Köln-West → Frankfurter Kreuz →
   Biebelried → Feuchtwangen/Crailsheim → Ulm/Elchingen, and the A81
   Weinsberg → Leonberg. These are always drawn.
   Add a line there for other routes you always want to see.
3. **Automatic branches:** up to `--branches-per-leg`. The script collects up to
   400 candidate routes within `--margin`. One at a time, the candidate adding the
   **most road that isn't drawn yet** wins, so each branch shows a genuinely
   different corridor rather than a local variant a few kilometres longer.
   Roads of the fixed branches already count as drawn. A branch must be **one
   detour**: it leaves the drawn routes once and rejoins once, so it never
   stacks unrelated detours. At most `MAX_LOCAL_SHARE` (25%) of that detour may
   be `Local` road, which keeps out routes over a B-road such as the B17.

### Points of interest

The `Points of Interest` sheet is shown in both tabs as a gold ★ with the name
in bold, and has its own toggle. Each point is drawn on the segment named in its
`Segment` column, at the spot where its lat/lon projects onto that segment.
Hovering over a star shows its category and notes. Points whose segment doesn't
exist in the Segments tab are skipped.

### How the map layout works

1. **Projection:** junction lat/lon are projected to normalised map units
   (1 degree = `UNIT_PER_DEGREE` = 10 units).
2. **Regional spines:** junctions on Primary roads, then on Secondary roads
   (at half strength), are pulled towards a straighter fitted line for each
   road section. How far they move is limited by `Geography lock` and `Max move`.
3. **Relaxation:** a light pass pushes apart junctions that are too close together.
4. **Rendering:** the script draws roads, borders, rivers, junction markers,
   bridges and tunnels, fuel-station badges, and labels as SVG. A2, A61 and A7
   are pinned to navy blue.

### Labels for bridges, tunnels and fuel stations

In both tabs these labels never overlap each other, junction names or markers,
and the maximum zoom is high enough for the text to be fully readable.

- Each view is one SVG that scales as a whole, so labels that don't overlap at
  one zoom level don't overlap at any.
- **Kaart:** each label is tried at increasing distances and angles around its
  marker until it finds a free spot. Close to the marker it also avoids road
  and river lines. Further out it may cross a line, but a halo in the
  background colour keeps it readable. Fuel stations are placed first, then
  tunnels, river bridges and valley bridges. Names appear from zoom 2.5×, and
  you can zoom to 16×.
- **Route:** labels are stacked in an ordered column beside each line, in the
  same order as their markers, so leader lines don't cross. Labels that don't
  fit there are placed freely, as on the map. You can zoom to 12×.
- A thin leader line connects a label that had to move away from its marker.
- A leader line may cross a junction name where branches are close together. The
  horizontal route layout (`--orientation horizontal`) doesn't use columns yet,
  so it has more of these crossings.

Tuning constants are at the top of the script: `POINT_FONT`, `POINT_SUB_FONT`,
`MAP_MAX_ZOOM`, `MAP_POI_FONT`, `MAP_POI_R`, `GRAPH_POINT_FONT`, `GRAPH_DIST_FONT`,
`GRAPH_JUNCTION_FONT`, `GRAPH_ROAD_FONT`, `GRAPH_POI_R`, `LANE_HEIGHT`, `STEP_X`, `GRAPH_MAX_ZOOM`, `LABEL_PRIORITY`,
`UNIT_PER_DEGREE`, `CANVAS_SCALE`, `TIER_STYLE`, `ROAD_WIDTH`, `PALETTE` and
`PINNED_COLOURS`.

## OpenStreetMap checks

Both scripts write their results as extra columns in
`junctions_topology_v5.xlsx`. They never change existing columns such as
`Latitude`, `Longitude` or `Distance (km)`, so you review the results and decide
what to correct.

They call free public APIs, so they wait between requests. Results go into a
JSON cache after every item, so you can stop with Ctrl+C and rerun the same
command to continue. On a rerun:

- items that failed with `ERROR` (for example a server timeout) are retried
- the existing result columns are refilled, not added a second time

To start completely fresh, delete the cache file.

Options for both scripts:

| Option | Default | Description |
|--------|---------|-------------|
| `--xlsx` | `junctions_topology_v5.xlsx` | Workbook to read |
| `--output` | same as `--xlsx` | Workbook to write to. Pass another name to write to a copy |
| `--cache` | `osm_junction_cache.json` / `osm_segment_cache.json` | Cache file |
| `--limit N` | all | Only process the first N items (for testing) |
| `--delay`, `--batch-size`, `--batch-pause` | | Rate limiting |

Test run that leaves the database untouched:

```bash
python3 osm_verify_junctions.py --output test.xlsx --cache test_cache.json --limit 10
```

### 1. osm_verify_junctions.py

For each junction it tries three sources, in this order:

1. **Holidays.xlsx** (`MATCHED_HOLIDAYS_SOURCE`): an approximate name match
   against the `Database` sheet. Words such as Kreuz, Knooppunt and Dreieck are
   ignored when comparing.
2. **OSM by name** (`MATCHED` / `MULTIPLE_CANDIDATES`): a
   `highway=motorway_junction` node with a matching name near the current
   coordinate.
3. **OSM road crossing** (`MATCHED_VIA_CROSSING`): the point where the roads in
   `Roads meeting here` cross.

If none of them finds anything, the status is `NOT_FOUND`. The results go into
the **Junctions** tab as columns `OSM status`, `OSM naam`, `OSM node`,
`OSM lat`, `OSM lon`, `OSM afstand tot huidige coord (km)`,
`OSM kandidaten gevonden` and `OSM zoekradius (km)`.

Extra options: `--holidays-xlsx` (default `Holidays.xlsx`; pass `''` to skip
that step), `--holidays-fuzzy-threshold` (default `0.84`) and `--radii-km`
(default `15,40`).

> Junctions already in the cache without an error aren't looked up again. If the
> cache was built before the Holidays step existed, delete
> `osm_junction_cache.json` so that every junction is checked against
> Holidays.xlsx.

### 2. osm_segment_distances.py

Run this after you've reviewed and corrected the junction coordinates. It
measures the driving distance for each segment with OSRM.

Routing straight from a junction's coordinate often lands on the wrong
carriageway, which adds a detour to turn around. So for each segment the
script:

1. places candidate points around both junctions: on the junction itself, and
   150, 400 and 800 m towards the other junction (at most 35% of the way), each
   also 30 m to the left and right, so both carriageways are tried
2. routes every From-candidate to every To-candidate in one OSRM `table`
   request, dropping candidates that snap more than 120 m away (onto another
   road)
3. adds the straight-line distance from each junction to its snapped point,
   and keeps the shortest total (a wrong carriageway needs a turnaround, so it
   never wins)
4. never accepts a total longer than 2× the straight-line distance between the
   junctions. If nothing fits, the status is `TOO_LONG`, and the coordinates or
   the topology of that segment need checking.

It adds `OSRM status`, `OSRM route afstand (km)`,
`Delta vs huidige Distance (km)`, `Delta (%)`, `Hemelsbreed (km)`,
`OSRM / hemelsbreed` and `OSRM opmerking` to the **Segments** tab. Segments
that differ by 15% or more are listed at the end of the run. Results from the
old method still in `osm_segment_cache.json` are recalculated automatically.

Because it takes the shortest route, a nearby parallel road can occasionally
win over the motorway itself. It's practically the same length, so the
distance is still a good approximation.

### 3. Applying the results

Once you've reviewed the OSM columns, two small scripts copy them into the
columns that `mapmaking.py` uses. Both save straight into
`junctions_topology_v5.xlsx`, so close it in Excel first.

```bash
python3 apply_osm_coords.py      # Junctions: OSM lat/lon → Latitude/Longitude (where both are filled)
python3 apply_osm_distances.py   # Segments: OSRM route afstand → Distance (km) (where OSRM status is OK)
```

Apply the coordinates first and run `osm_segment_distances.py` after that, since
the distances are measured from the junction coordinates.

To measure and apply the distances in one go, skipping the review step:

```bash
python3 osm_segment_distances.py --apply
```

## Workbook sheets (junctions_topology_v5.xlsx)

| Sheet | Contents | Used by |
|-------|----------|---------|
| `Junctions` | ID, name, roads meeting here, country, lat/lon, layout region, `Geography lock`, `Max move`, tier, type | all scripts |
| `Roads` | Road ID, number, hierarchy (Primary/Secondary/Connector/Local), layout parameters | mapmaking |
| `Segments` | Road edges between junctions: `Edge ID`, `From ID`, `To ID`, `Road`, `Distance (km)` | mapmaking, segment check |
| `Points` | Bridges, tunnels, fuel stations and rest areas placed along a segment (`Edge ID`, `PositionOnEdge`) | mapmaking |
| `River Junctions`, `River Segments` | River network | mapmaking (Kaart) |
| `Border Nodes`, `Border Segments` | Country borders | mapmaking (Kaart) |
| `Rivers`, `Points of Interest` | Reference data | none |
| `Map Model Notes` | Explains the fields | none |
