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
| `map_icons.py` | The brand-icon system `mapmaking.py` draws on fuel stations: loading `icons/fuel` and `icons/food`, the coloured-badge fallback, and the `Fuel brand`/`Food brand(s)` text parsing |
| `icons/fuel/`, `icons/food/` | The actual logo image files, one per brand |
| `serways_brands.py` | Finds each German station's fast-food/coffee brand(s) on serways.de → `Food brand(s)` in the Points tab |
| `osm_verify_junctions.py` | Checks junction coordinates against Holidays.xlsx and OpenStreetMap |
| `osm_segment_distances.py` | Checks segment distances against OSRM driving distances |
| `osm_point_leg_distances.py` | Measures each fuel station's distance along its segment with OSRM → `OSM DistanceFromStart (km)` in the Points tab (used by the route diagram) |
| `apply_osm_coords.py`, `apply_osm_distances.py` | Copy reviewed OSM results into `Latitude`/`Longitude` and `Distance (km)` |
| `osm_junction_cache.json`, `osm_segment_cache.json`, `serways_cache.json` | Caches for the OSM/serways checks, so runs can resume |
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
- `requests` (the two `osm_*` scripts; `serways_brands.py` only needs the
  standard library)

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

Builds one HTML page with a map tab and one tab per route in `ROUTE_TABS`:

- **Kaart:** the whole network as a schematic map that stays close to real
  geography. It has zoom, a GPS button, and toggles for bridge names (B),
  distances (Km), points of interest (★) and debug IDs (D).
- **Route tabs:** currently Vught → Berwang, Vught → Serfaus via Lindau, and
  the reverse of each. Each shows the route between two places as a "metro board" diagram, like the
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
first route tab directly, `#route2` the second, and so on.

The route tabs are defined in `ROUTE_TABS` at the top of `mapmaking.py`. Each
entry has a destination (`to`), optionally its own start (`from`, default
`DEFAULT_FROM`/`--from` - written as a plain dict literal rather than
`dict(...)`, since `from` can't be a keyword argument), optionally junctions
the main route must pass in order (`via`), and optionally its own fixed
branches (`branches`, default `DEFAULT_BRANCHES`). The Serfaus tabs add the
route via Füssen and the Fernpass as a fixed branch (`SERFAUS_BRANCHES`).
It's the shortest way to/from Serfaus, but because it never passes Lindau,
the automatic search would never find it. Add a line to `ROUTE_TABS` for
another destination (or another reverse pair).

Passing `--to`, `--via`, `--branch` or `--route-title` replaces `ROUTE_TABS`
with a single tab for just that route. The other route options (`--from`,
`--margin`, `--branches-per-leg`, `--min-novel-km`, `--orientation`,
`--no-default-branches`) apply to every route tab. Building takes about 25
seconds per route tab.

| Option | Default | Description |
|--------|---------|-------------|
| `--xlsx` | `junctions_topology_v5.xlsx` | Input workbook |
| `--output` | `holidays.html` | Output HTML file (overwritten) |
| **Kaart tab** | | |
| `--map-title` | `Geographic Spine Map` | Page title |
| `--no-tube-primary` | off | Disable London-Underground-style angle snapping for Primary roads |
| `--no-tube-secondary` | off | Disable angle snapping for Secondary roads |
| `--tube-relaxed` | off | Allow a looser 16-angle set for tube-style bends (every 30°, plus the diagonals) instead of the strict Underground convention (multiples of 45° only) |
| **Route tab** | | |
| `--from` / `--to` | `Vught` / `ROUTE_TABS` | Start and end junction names |
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

### Brand badges (fuel stations)

A tankstation/Autohof point gets a row of small badges under its name: one per
fuel brand (from the `Fuel brand` column) and up to 4 fast-food/coffee chains,
preferably from the structured `Food brand(s)` column (`serways_brands.py`
below), falling back to recognising a known chain's name in the free-text
`Facilities` column when that's empty. Where we have the real logo on file
(`icons/fuel`, `icons/food`, loaded by `map_icons.py`) it's shown as-is - these
are genuine brand logos, used only to show which fuel/food brand is actually
present at that real location, not decorative or promotional use. A brand
with no icon file yet falls back to a coloured badge (background colour + a
short code), still better than plain text. Add a brand to `FUEL_BRAND_COLOURS`/
`FOOD_BRAND_COLOURS` in `map_icons.py` to recognise more of them, and an entry
to `FUEL_ICON_FILES`/`FOOD_ICON_FILES` once you've added its image file.

On the Kaart tab the badges are a further zoom stage past the station name
(`BADGE_ZOOM_THRESHOLD` in `mapmaking.py`, names already show from 2.5×); on a
route tab the fuel-station square shows the fuel brand's icon (or, lacking
one, is tinted with its badge colour).

**Different brand per side:** when `Fuel brand` or `Food brand(s)` uses the
"X east; Y west" convention (also north/south - a service area with a
different operator on each physical side of the road), the Kaart tab draws
two separate one-sided markers instead of one, each with only that side's
badges - placed on its actual geographic side via `compass_side_sign()`
(`mapmaking.py`), regardless of which way the schematic line happens to run.
A direction given as a place instead ("toward Antwerp") can't be resolved
this way and falls back to one combined marker showing every brand mentioned.

**Which side a one-sided station's triangle points to** normally comes
straight from its own Latitude/Longitude (`real_side_sign()`): which side of
the straight line between its segment's two junctions that coordinate falls
on. That line can be a poor stand-in for the real road when the junctions are
far apart (the A8 Leonberg → Ulm/Elchingen segment is ~100km, and the real A8
doesn't run straight over that distance), so a station can end up on the
wrong side even with an accurate coordinate. The `Side direction` column in
the Points tab (north/south/east/west) overrides this when filled in - it's
pre-filled for every one-sided point whose official name or Notes already
state a direction (`Aichen Nord`, `Autohof Illertal-West`, ...), and used via
`compass_side_sign()` instead of the coordinate-based guess. Add it for
another station (or correct one) directly in Excel if its marker is on the
wrong side.

### serways_brands.py: food/coffee brands from serways.de

For a German Autobahn station, run `python3 serways_brands.py` to look up its
real fast-food/coffee brand(s) on serways.de (the consumer site of Tank &
Rast, which operates almost every Autobahn service area) and write them into
`Food brand(s)` in the Points tab. It matches a Points row to a serways.de
page by normalising its `Name` the same way serways.de names its pages
(umlauts expanded, lower-cased, direction suffix kept if the name already has
one); most Points rows won't match, since most of this network is Dutch,
Belgian or Austrian and outside Tank & Rast's network. A German-looking name
in the run's "not matched" list is worth a manual look - add a line to
`NAME_OVERRIDES` at the top of the script (or just say so in chat) once you
know its actual serways.de slug. Like the `osm_*.py` scripts, results are
cached (`serways_cache.json`) so an interrupted run can just be resumed, and
it never touches `Fuel brand` or `Facilities` - only `Food brand(s)`, and
only for a row it could match.

### Region blobs (Kaart only)

The workbook's **NaturalRegions** tab (`Region name`, `Junction name` - one
row per member junction) draws a soft green background shape behind a leisure
area - currently Ardennen & Eifel and Sauerland - with a mountain icon and
name, like the green "highlight" areas on a hand-drawn touring map. The blob
is that region's junctions' convex hull, padded outward and rounded into an
organic shape (`inflate_blob_points`/`smooth_closed_path` in `mapmaking.py`).
Add rows to NaturalRegions for another region - the junction names must exist
in the Junctions tab, and more of them (especially ones that outline the
area) give a better-fitting blob.

### Road-number labels (Kaart only)

Every non-local road gets its number ("A3", not the internal "A3 (DE)" used
to tell same-numbered roads in different countries apart - see
`Canonical Road ID` in the Roads tab) drawn on the line itself, aligned with
the line's direction, at the middle of a segment. A road broken into many
short consecutive segments doesn't repeat the label on each one:
`assign_road_labels()` groups segments by road and keeps only those whose
candidate label is more than `MAP_ROAD_LABEL_CLEARANCE` SVG units from
another kept label of the *same* road, so a long continuous stretch gets one
label every so often instead of one per segment, while a short stretch far
from the rest still gets its own. The label sits on a pill tinted lightly
with the road's own colour (`fill-opacity="0.22"`), but the text itself is
drawn in `complementary_colour()` of that colour (hue rotated 180 degrees) -
so the label always reads clearly against its pill regardless of how dark or
light the road's colour is, instead of risking text that blends into the
line it's labelling.

### Legend (Kaart only)

The "L" button toggles a fixed symbol legend in the bottom-left corner -
built once in `build_legend_html()` from the exact same drawing functions
the map itself uses (`junction_circle_svg`, `fuel_marker_svg`, `bar_svg`,
`star_svg`, ...), so it can't drift out of sync with what's actually drawn.
It explains symbol *shapes* (large/small junction, tankstation, brug, tunnel,
point of interest, rivierbrug, rivier, grensovergang, a generic weg, a
natuurgebied swatch) rather than the dozens of individual road colours. It's
its own small fixed-position SVG, not part of the zoomable `#stage`, so it
stays a constant, legible size and never blocks map panning/zooming
(`pointer-events:none`) regardless of the map's own zoom level.

### Controls panel (Kaart and Route)

All the per-view toggle/zoom buttons (zoom, legend, bridge/POI/distance
toggles, GPS, debug) live behind a single &#9776; button, collapsed by
default, instead of sitting on screen the whole time - every one of them
still works exactly the same once opened (same ids, same click handlers),
only the container around them is collapsible. This keeps them from
covering the route/map itself when you've panned or zoomed into a corner,
since a collapsed panel has nothing to overlap with. Route diagrams also get
extra blank margin around the canvas itself (`MARGIN_PX`), so there's more
natural empty space for the panel to sit over even before you open it.

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

With strict tube-style angle snapping (the default - see `--tube-relaxed`
below), two or more roads can end up snapped onto the exact same one of the
8 allowed angles right where they leave a junction, since there are only 8
to choose from. `debundle_junction_overlaps()` detects this (grouping
segment-ends by `(junction, angle)`) and nudges every line but the middle
one sideways for a short stretch near the junction, tapering back to the
real route just beyond it, so they fan out visibly instead of overlapping.

### Exit direction (Segments sheet, optional)

The angle a tube-style road leaves a junction at is normally picked purely
from the bearing to the next junction - which sometimes reads wrong even
after debundling, e.g. two genuinely different roads both legitimately
bending to leave a busy junction "northward" when one of them should
schematically read as going east. The Segments sheet has two optional
columns, `Exit direction (From)` and `Exit direction (To)`, to pin the
schematic compass direction a specific segment leaves its `From`/`To`
junction - a compass letter or word (`N`, `NE`, `oost`, `noordwest`, ...;
see `COMPASS_ANGLES` for every accepted spelling, Dutch and English). Only
tube-style segments use it; it's ignored for Connector/Local roads, which
are always drawn as direct lines. Leave both blank (the common case) and
the angle is picked automatically as before.

Example: at Ekkersweijer (A2/A50), A50 used to snap north in parallel with
A2 instead of reading as its own line. Setting `Exit direction (From)` = `E`
on the Ekkersweijer->Paalgraven (A50) segment, and `W`/`N` on the two A2
segments, makes the three roads fan out from Ekkersweijer the way they're
meant to be read, independent of their raw geographic bearing.

### Route diagrams reuse the Kaart's colour solver

`render_graph()` (the per-route pages, not the Kaart) used to pick colours
by cycling through a fixed 14-entry `ROAD_PALETTE` in first-seen order - a
route with more than 14 distinct roads (easily reached once branches are
included) would silently reuse a colour for an unrelated road (e.g. A3 and
A44 both landing on the same slot), making them look like the same line
where they crossed. It now builds the same synthetic segment list the Kaart's own colour
solver, `assign_road_colours()`, expects (reuse-first, CIE Lab-distance
constraints scaled by screen proximity and road importance - see its
docstring), keyed by the route diagram's own schematic node positions
instead of geographic ones, and reuses the same 87-colour `PALETTE` - so it
gets the same collision-aware assignment, including the `PINNED_COLOURS`
navy for A2/A61/A7.

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

### Marker declutter tiers (Kaart only)

Bridge/tunnel and fuel-station *markers* (not their names - those still follow
the zoom rule above) are visible well before zoom 2.5×, but not all at once:
fully zoomed out shows a sparse, still-lively sample instead of either
everything (illegible on a corridor with dozens of bridges a few hundred
metres apart) or nothing (an empty-looking map). `assign_declutter_tiers()`
greedily thins them into 5 tiers at build time - tier 0 is only markers that
are mutually more than `MARKER_TIER_CLEARANCES[0]` SVG units apart, tier 1
adds whatever's far enough from tier 0, and so on; whatever's left after the
last tier (the densest clusters) only appears once zoomed in enough that nothing
is that crowded any more. A bridge's priority (which one of several close
together wins an earlier tier) is its length - the biggest bridges show up
first. `MARKER_TIER_ZOOM` in the generated page's `<script>` sets the scale
each tier appears at.

Tuning constants are at the top of the script: `POINT_FONT`, `POINT_SUB_FONT`,
`MAP_MAX_ZOOM`, `MAP_POI_FONT`, `MAP_POI_R`, `GRAPH_POINT_FONT`, `GRAPH_DIST_FONT`,
`GRAPH_JUNCTION_FONT`, `GRAPH_ROAD_FONT`, `GRAPH_POI_R`, `LANE_HEIGHT`, `STEP_X`, `GRAPH_MAX_ZOOM`, `LABEL_PRIORITY`,
`UNIT_PER_DEGREE`, `CANVAS_SCALE`, `TIER_STYLE`, `ROAD_WIDTH`, `PALETTE` and
`PINNED_COLOURS`. For region blobs: `REGION_BLOB_FILL`, `REGION_BLOB_PAD` and
`REGION_LABEL_FONT`/`REGION_LABEL_COLOUR`. For brand badges: `BADGE_ZOOM_THRESHOLD`
in `mapmaking.py`, and `ICON_SIZE`, `FUEL_BRAND_COLOURS`, `FOOD_BRAND_COLOURS`,
`FUEL_ICON_FILES`, `FOOD_ICON_FILES` in `map_icons.py`. For marker declutter:
`MARKER_TIER_CLEARANCES`/`MARKER_TIER_ZOOM`. For road labels: `MAP_ROAD_FONT`,
`MAP_ROAD_LABEL_CLEARANCE`.

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
| `Segments` | Road edges between junctions: `Edge ID`, `From ID`, `To ID`, `Road`, `Distance (km)`, optional `Exit direction (From)`/`Exit direction (To)` | mapmaking, segment check |
| `Points` | Bridges, tunnels, fuel stations and rest areas placed along a segment (`Edge ID`, `PositionOnEdge`), incl. `Fuel brand`, `Food brand(s)` and `Side direction` | mapmaking, serways_brands |
| `River Junctions`, `River Segments` | River network | mapmaking (Kaart) |
| `Border Nodes`, `Border Segments` | Country borders | mapmaking (Kaart) |
| `NaturalRegions` | `Region name` + `Junction name`, one row per member junction | mapmaking (Kaart) |
| `Rivers`, `Points of Interest` | Reference data | none |
| `Map Model Notes` | Explains the fields | none |
