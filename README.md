# extractio

**extractio** converts AutoCAD DWG layers to GeoJSON — driven entirely by a YAML config file. It connects directly to a running AutoCAD instance, reads geometry and attributes without any manual export step, and writes clean GeoJSON files ready for QGIS, PostGIS, Mapbox, or any GIS pipeline.

No code changes are needed between projects. Everything — geometry type, field schema, field order, spatial joins, calculated values, zone assignment, multi-DWG merging, locking — is controlled by the config file alone.

---

## What It Does

- Connects to AutoCAD over COM — no export, no intermediate file
- Reads entities directly from model space (LWPOLYLINE, INSERT, LINE, POINT, MTEXT, TEXT)
- Supports **polygon**, **point**, and **line** geometry
- Extracts **block attributes** by tag name and maps them to any output field name you choose
- Merges features from multiple DWGs into one GeoJSON with per-source tagging
- Computes area, perimeter, and length from geometry
- Assigns auto-incremented `Connection_ID` keys (globally unique across the project)
- Performs spatial joins — stamps a zone/block identifier on every feature from a reference layer
- Derives dissolved outer-shell boundaries from block/zone polygon layers (no separate DWG source needed)
- Filters phantom reference polygons using text label count as ground truth
- Validates source layer names before extraction and suggests corrections
- Saves a versioned manifest on every run (feature counts + diff vs. previous run)
- LLM-assisted YAML builder (`--build`) via Groq or Gemini

---

## Use Cases

Any project where AutoCAD drawings need to become GeoJSON:

- **Infrastructure mapping** — asset inventories (structures, installations, networks) from site layouts
- **Civil and surveying** — site parcels, boundaries, corridors, and alignments
- **Utilities** — conduit routes, cable paths, service areas from engineering drawings
- **Construction management** — zone boundaries, equipment positions, access routes
- **Urban and land use** — building footprints, plot boundaries, land-use polygons
- **Facilities and asset management** — georeferenced asset databases from as-built drawings
- **Environmental survey** — survey grids, sample areas, observation points from CAD base maps

If your project has AutoCAD drawings with named layers containing geometry and attributes, extractio converts them to GeoJSON with a custom field schema that you define.

---

## How It Works — Overview

```
AutoCAD (open, COM)
        │
        │  reads entities directly from model space
        ▼
  extractio.py ──── config.yaml  (your field schema, layer list, DWG paths)
        │
        │  resolves fields, performs spatial joins, applies filters
        ▼
  output/*.geojson  (one file per layer, CRS defined by you)
```

The executor reads AutoCAD entity geometry directly from the live COM interface. For layers populated with symbol blocks (INSERT entities), it opens each block instance and extracts the attribute tags you specify, mapping them to whatever output field names you choose. It then runs your field resolution rules (constants, attribute lookups, spatial joins, calculations, derivations) to build the full property schema for each feature.

**The config file is the only thing that changes between projects. The extractio.py code is never modified.**

---

## Table of Contents

1. [Requirements](#1-requirements)
2. [Installation](#2-installation)
3. [AutoCAD Setup](#3-autocad-setup)
4. [Quick Start](#4-quick-start)
5. [Execution Model — Two Passes](#5-execution-model--two-passes)
6. [Config File Structure](#6-config-file-structure)
7. [Global Section](#7-global-section)
8. [The Spatial Reference System](#8-the-spatial-reference-system)
9. [Spatial Reference ID from DWG Filenames](#9-spatial-reference-id-from-dwg-filenames)
10. [Layer Types](#10-layer-types)
    - [Type 1 — Spatial Reference Layer](#type-1--spatial-reference-layer)
    - [Type 2 — Derived Boundary Layer](#type-2--derived-boundary-layer)
    - [Type 3 — Standard Layer (point / line / polygon)](#type-3--standard-layer-point--line--polygon)
    - [Type 4 — Derived Child Layer](#type-4--derived-child-layer)
11. [Field Reference](#11-field-reference)
12. [Geometry Params Reference](#12-geometry-params-reference)
13. [Spatial Join Reference](#13-spatial-join-reference)
14. [Derive Transforms Reference](#14-derive-transforms-reference)
15. [The Connection_ID System](#15-the-connection_id-system)
16. [Multi-DWG Projects](#16-multi-dwg-projects)
17. [Working with Block Attributes](#17-working-with-block-attributes)
18. [Advanced Config Features](#18-advanced-config-features)
19. [LLM YAML Builder](#19-llm-yaml-builder)
20. [CLI Reference](#20-cli-reference)
21. [Lock System](#21-lock-system)
22. [Output Format](#22-output-format)
23. [Config and Version Management](#23-config-and-version-management)
24. [Performance Notes](#24-performance-notes)
25. [Troubleshooting](#25-troubleshooting)
26. [Project Setup Checklist](#26-project-setup-checklist)
27. [Files](#27-files)

---

## 1. Requirements

| Item | Minimum | Notes |
|---|---|---|
| Operating System | Windows 10 | The COM bridge to AutoCAD is Windows-only |
| AutoCAD | Any version with COM API | Must be running with DWGs open before extractio starts |
| Python | 3.8+ | Standard CPython |
| pywin32 | any recent | Provides the COM interface to AutoCAD |
| pyyaml | any recent | Reads the config file |
| shapely | 1.8+ | Required only for `derive_from: spatial_reference` (boundary dissolve) layers |

---

## 2. Installation

```
pip install pywin32 pyyaml shapely
```

Optional extras:

```
pip install filelock groq google-generativeai pdfplumber python-docx pandas
```

- `filelock` — safe concurrent version file writes
- `groq` / `google-generativeai` — LLM-assisted YAML builder (`--build`)
- `pdfplumber` / `python-docx` — reference document parsing in `--build` mode
- `pandas` — used in `verify.py`
- `fiona` — File Geodatabase output (`--output-format gdb`); requires GDAL 3.6+

No package installation or virtual environment is required (though using one is good practice). The entire engine is in the single file `extractio.py`.

---

## 3. AutoCAD Setup

extractio reads geometry directly from AutoCAD's model space over COM:

- **AutoCAD must be running** and fully loaded before you run extractio
- **Every DWG listed in your config** must be open as a document in that AutoCAD session
- DWGs must be in **model space** (paper space / layouts are not read)
- Do not close AutoCAD, switch documents, or trigger a DWG regeneration while extractio is running

**Document matching:** extractio finds open documents by comparing the full file path in your config against open document paths in AutoCAD. If the full path does not match (e.g. a mapped drive vs. a UNC path), it falls back to matching by filename only.

**Multiple DWGs:** all DWGs can be open simultaneously in the same AutoCAD session. extractio switches between them programmatically — you do not need to manually activate each one.

**DWG not open:** if a DWG in your config is not open, that layer is skipped with a warning and extraction continues with the remaining layers.

---

## 4. Quick Start

### Step 1 — Set up your config

```
cp samples/sample_config.yaml config.yaml
cp samples/sample_dwg_paths.yaml dwg_paths.yaml
```

Edit `config.yaml` — set `project_name`, `crs`, `output_dir`, and layer definitions.
Edit `dwg_paths.yaml` — fill in the absolute paths to your DWG files.

Or use the interactive wizard to generate a config from the sample template:

```
python yaml_generator.py
```

### Step 2 — Inspect your DWG layer names

Before filling in your config, find the exact AutoCAD layer names:

```
python extractio.py config.yaml --dwg-layers drawing_a
```

AutoCAD layer names are case-sensitive. Copy them exactly.

### Step 3 — Run

```
# Interactive layer selector (toggle layers on/off then press X to run)
python launcher.py

# Run specific layers by name (fuzzy match accepted)
python extractio.py config.yaml --layers "Zone Boundaries" "Equipment Points"

# Run all unlocked layers
python extractio.py config.yaml --run all

# Check what layers are defined and their lock status
python extractio.py config.yaml --list
```

**Config auto-discovery:** if no config path is given, extractio looks for `config.yaml` next to the script, then for a single `*.yaml` file in the same folder.

### Recommended first-run sequence

```
# 1. Inspect DWG layer names
python extractio.py config.yaml --dwg-layers drawing_a

# 2. Extract and verify the spatial reference layer first
python extractio.py config.yaml --unlock "Zone Boundaries"
python extractio.py config.yaml --layers "Zone Boundaries"
python extractio.py config.yaml --lock "Zone Boundaries"

# 3. Unlock and run the rest
python extractio.py config.yaml --unlock "Equipment Points"
python extractio.py config.yaml --run all
python extractio.py config.yaml --lock-all
```

---

## 5. Execution Model — Two Passes

Understanding the two-pass model helps you configure layer ordering correctly and avoid spatial join failures.

### Pass 1 — Build the Spatial Cache

Before any output is written, extractio builds its spatial reference data:

**1a. Spatial reference layer**

The layer with `role: spatial_reference` is loaded first. Its polygons define the zone/block grid used by all spatial joins throughout the run.

- If this layer's GeoJSON already exists on disk (from a prior run), it is loaded from the file. AutoCAD is not read. This is the fast path when the reference grid is stable.
- If the GeoJSON does not exist, it is extracted from AutoCAD before pass 2 begins.

**1b. Spatial join sources**

Any layer referenced by a `spatial_join: {from_layer: "..."}` field in any other layer's config is extracted in pass 1 and registered in the spatial cache so it is available for dependent layers in pass 2.

### Pass 2 — Extract Selected Layers

Layers run in YAML config order, with the spatial reference layer promoted to first position. For each layer:

1. Entities are read from the DWG model space
2. Geometry is extracted (polygon ring, polyline vertices, insert point)
3. All `fields:` are resolved — constants and joins first, derivations second
4. Features are written to a GeoJSON file in `output_dir`
5. The layer's features are registered in the spatial cache for subsequent layers

**Implication for layer ordering:** if layer B uses `spatial_join: {from_layer: "Layer A"}`, then Layer A must appear earlier in the YAML config than Layer B (or be a declared join source in `global.block_no`).

---

## 6. Config File Structure

A config file has two top-level keys:

```yaml
global:
  # project-wide settings: CRS, DWG paths, common fields, spatial reference config
  project_name: "My Project"
  crs: "EPSG:32632"
  output_dir: "./outputs"
  ...

layers:
  # list of layers to extract — one entry per output GeoJSON
  - name: "Zone Boundaries"
    role: spatial_reference
    geometry: polygon
    ...

  - name: "Equipment Points"
    geometry: point
    ...

  - name: "Route Lines"
    geometry: line
    ...
```

Each entry under `layers:` produces one `.geojson` output file. The YAML list order controls execution order.

See [`samples/sample_config.yaml`](samples/sample_config.yaml) for a full working example covering all layer types, all field resolver types, and all common patterns.

---

## 7. Global Section

```yaml
global:
  project_name: "MyProject"              # used in version labels
  crs: "EPSG:32632"                      # output coordinate reference system
  output_dir: "./outputs"                # created automatically if it doesn't exist
  # output_format: geojson               # geojson (default) or gdb
  # gdb_name: MyProject.gdb             # GDB filename; defaults to <project_name>.gdb

  # DWG alias file — maps short names to absolute file paths.
  # Keep this file out of git (it contains machine-specific paths).
  dwg_paths_file: dwg_paths.yaml

  # DWG aliases used by layers below. List form: aliases resolved from dwg_paths_file.
  source_dwgs:
    - drawing_a
    - drawing_b
    - drawing_c

  # Optional: canonical field order reference (another config whose field order is master).
  # All layers in THIS config are reordered to match.
  # field_order_ref: "/path/to/reference_config.yaml"

  # Fields written to every feature on every layer automatically.
  # Layer-specific fields: sections override these.
  common_fields:
    Project_Name: "My Project"
    Organisation: "My Organisation"
    Prepared_By:  "GIS Team"
    Status:       "Draft"

  # Fields left null at extraction time — filled later from an asset management system.
  deferred_fields:
    - Verified_By
    - Approved_By

  # Named geometry calculations available to any layer via {calculate: <name>}
  calculated_fields:
    Area_Ha:      {formula: area,      unit: hectares, round: 2}
    Area_Sqm:     {formula: area,      unit: sqm,      round: 2}
    Perimeter_Km: {formula: perimeter, unit: km,       round: 2}
    Length_m:     {formula: length,    unit: meters,   round: 2}
    Length_Km:    {formula: length,    unit: km,       round: 2}

  # Auto-numbering pattern for Connection_ID.
  # Available vars: {plot_name} {code} {seq:02d}
  connection_id:
    pattern: "{plot_name}_{code}_{seq:02d}"

  # Spatial reference config — fixed internal key name "block_no".
  # primary_source:   used by {spatial_join: primary}
  # secondary_source: used by {spatial_join: secondary}
  block_no:
    primary_source:
      from_dwg:   [drawing_a, drawing_b]
      from_layer: "ZONE-LABELS"
      fallbacks:  ["Zone Labels", "LABELS"]
      from_field: "Contents"
      method:     nearest
    secondary_source:
      from_layer: "Zone Boundaries"
      from_field: "Zone_ID"
      method:     nearest
```

### Choosing a CRS

Use the EPSG code for the UTM zone covering your project area. Examples:

| EPSG code | Zone | Coverage |
|---|---|---|
| EPSG:32629 | UTM 29N | Western Europe, Atlantic |
| EPSG:32632 | UTM 32N | Central Europe |
| EPSG:32637 | UTM 37N | East Africa, Middle East |
| EPSG:32643 | UTM 43N | Central/South Asia |
| EPSG:32644 | UTM 44N | South Asia |
| EPSG:32754 | UTM 54S | Eastern Australia |

Find your UTM zone at [spatialreference.org](https://spatialreference.org) or use QGIS / ArcGIS.

---

## 8. The Spatial Reference System

The spatial reference system is the core mechanism that lets extractio assign zone or block identifiers to features across many layers and DWG files without hard-coding zone values.

**What it is:** a polygon layer in your DWG (or multiple DWGs) that divides the project area into named zones or blocks. Each polygon has an identifier (e.g. `BLK01`, `Z-03`, `SECTOR_A`). extractio loads this layer first and uses it as a lookup grid throughout the rest of the run.

**What it does:**

- Any feature on any other layer receives a zone/block ID stamped automatically by finding which reference polygon contains that feature's centroid (`spatial_join: secondary`)
- `Connection_ID` incorporates the zone ID, making it globally unique across the project
- `from_dwg_name` zone lookups are validated spatially against reference polygons, handling DWGs that span multiple zones

**`block_no` is the fixed internal YAML key** that points to the spatial reference config. The field names (`Block_No`, `Zone_ID`, etc.) in your output are entirely user-defined.

**If you have no zone grid:** replace `spatial_join: secondary` with a constant field, remove `from_dwg_name`, and omit `global.block_no`. The spatial reference layer is optional if you have no zone-based joins.

---

## 9. Spatial Reference ID from DWG Filenames

When you use `{from_dwg_name: true}` in a field, extractio extracts a zone/plot ID from the DWG filename using priority-ordered patterns.

**ID format:** one uppercase letter + one or two digits + optional lowercase suffix. Examples: `A1`, `Z12`, `S3a`, `P09b`.

**Detection patterns (highest priority first):**

| Pattern in filename | Example | Detected |
|---|---|---|
| `Plot Z01` or `plot z01` | `Drawing (Plot Z01).dwg` | `Z01` |
| `Routing A2` | `Routes A2 North.dwg` | `A2` |
| `(S3a)` or `(S3a-` | `Boundary S3a.dwg` | `S3a` |
| `Block Z01` / `BlockZ01` | `Block Z01 Layout.dwg` | `Z01` |
| `_Z01_` or `-Z01-` | `Site_Z01_Layout.dwg` | `Z01` |
| `Z01.dwg` | `Z01.dwg` | `Z01` |
| Standalone token | `Survey Z01 Final.dwg` | `Z01` |

If detection fails, a warning is printed and `from_dwg_name` returns `A?`.

**Fix:** rename the DWG to include a recognisable pattern (e.g. `Layout (Plot A1).dwg`), or replace `from_dwg_name: true` with a constant string.

**Multi-zone DWGs:** a single DWG covering multiple zones (e.g. a master routing file for zones A1 through A5) is handled automatically. extractio spatially checks which reference polygon each feature falls inside and uses that result instead of the filename-derived value. This is logged once per layer.

---

## 10. Layer Types

Four layer types, distinguished by the presence of `role:` and `derive_from:` keys.

---

### Type 1 — Spatial Reference Layer

The zone/block grid layer. One per project. Must appear before any layer that uses `spatial_join: primary` or `spatial_join: secondary`. Declare with `role: spatial_reference`.

```yaml
- name:         "Zone Boundaries"
  role:         spatial_reference       # enables spatial join cache + boundary dissolve
  locked:       false
  cli_aliases:  ["zb", "zones"]
  source_dwg:   drawing_a
  source_layer: "ZONE-BOUNDARY"
  merge_sources:
    - source_dwg:   drawing_a
      sub_plot:     "A"
      source_layer: "ZONE-BOUNDARY"
      fallbacks:    ["ZONE_BOUNDARY", "Zone Boundary"]
    - source_dwg:   drawing_b
      sub_plot:     "B"
      source_layer: "ZONE-BOUNDARY"
  match_mode:   exact
  geometry:     polygon
  code:         "ZB"
  output:       "zone_boundaries.geojson"
  expect:
    min_features: 1
  fields:
    Code:           null
    Category:       "Zone Boundaries"
    Sub_Plot:       {from_merge_source: sub_plot}
    Zone_ID:        {spatial_join: primary}
    Area_Ha:        {calculate: Area_Ha}
    Perimeter_Km:   {calculate: Perimeter_Km}
    Owner:          null
    Remarks:        null
```

**Workflow tip:** once extracted and verified, lock this layer. On all subsequent runs extractio loads the GeoJSON from disk instead of re-reading AutoCAD — faster and guarantees consistency with all downstream spatial joins.

---

### Type 2 — Derived Boundary Layer

Computes the dissolved outer shell of the spatial reference polygons. Does not read from AutoCAD — computed from the already-loaded spatial reference features using shapely. Requires `pip install shapely`.

Use this when you want a merged outer boundary for each sub-plot (e.g. one large polygon per plot from many individual zone polygons).

```yaml
- name:         "Site Boundary"
  derive_from:  spatial_reference       # dissolves Zone Boundaries per sub_plot
  locked:       false
  cli_aliases:  ["sb", "site"]
  merge_sources:
    - source_dwg:   drawing_a
      sub_plot:     "A"
      source_layer: "ZONE-BOUNDARY"
    - source_dwg:   drawing_b
      sub_plot:     "B"
      source_layer: "ZONE-BOUNDARY"
  geometry:     polygon
  code:         "SB"
  output:       "site_boundary.geojson"
  fields:
    Code:           null
    Category:       "Site Boundary"
    Area_Ha:        {calculate: Area_Ha}
    Perimeter_Km:   {calculate: Perimeter_Km}
    Owner:          null
    Remarks:        null
```

---

### Type 3 — Standard Layer (point / line / polygon)

The most common type. Reads entities from one or more AutoCAD layers and writes one GeoJSON.

**Declare the geometry type** with `geometry:`:

| Value | Entity types read | GeoJSON output |
|---|---|---|
| `polygon` | LWPOLYLINE (closed), POLYLINE (closed) | Polygon |
| `line` | LWPOLYLINE, LINE, POLYLINE | LineString |
| `point` | INSERT (block reference), POINT | Point |

If `geometry:` is omitted, it defaults to `polygon`.

---

#### Point layer example

```yaml
- name:         "Equipment Points"
  locked:       false
  cli_aliases:  ["pts", "points"]
  source_dwg:   drawing_a
  source_layer: "POINT-LAYER"
  merge_sources:
    - source_dwg:   drawing_a
      sub_plot:     "A"
      source_layer: "POINT-LAYER"
    - source_dwg:   drawing_b
      sub_plot:     "B"
      source_layer: "POINT-LAYER"
  match_mode:   exact
  geometry:     point
  geometry_params:
    only_insert: true
    block_name:  "POINT_BLOCK"    # leave blank to accept any INSERT on the layer
  code:         "PT"
  output:       "equipment_points.geojson"
  fields:
    Code:           null
    Category:       "Equipment Points"
    Sub_Plot:       {from_merge_source: sub_plot}
    Zone_ID:        {spatial_join: primary}
    Connection_ID:  {auto_sequence: true}
    Asset_Type:     {from_attr: "TYPE"}       # AutoCAD attribute tag "TYPE"
    Attribute_1:    {from_attr: "ATTR_1"}
```

---

#### Line layer example

```yaml
- name:         "Route Lines"
  locked:       false
  cli_aliases:  ["ln", "lines"]
  source_dwg:   drawing_c
  source_layer: "LINE-LAYER"
  merge_sources:
    - source_dwg:   drawing_c
      sub_plot:     "A"
      source_layer: "LINE-LAYER"
      fallbacks:    ["LINE_LAYER", "Line Layer"]
    - source_dwg:   drawing_c
      sub_plot:     "B"
      source_layer: "LINE-LAYER"
  match_mode:   exact
  geometry:     line
  code:         "LN"
  output:       "route_lines.geojson"
  fields:
    Code:           null
    Category:       "Route Lines"
    Sub_Plot:       {from_merge_source: sub_plot}
    Zone_ID:        {spatial_join: primary}
    Connection_ID:  {auto_sequence: true}
    Sub_Type:       {from_merge_source: layer_subtype}
    Length_m:       {calculate: Length_m}
    Length_Km:      {calculate: Length_Km}
```

---

#### Polygon layer example

```yaml
- name:         "Site Polygons"
  locked:       false
  source_dwg:   drawing_a
  source_layer: "POLY-LAYER-A"
  match_mode:   exact
  geometry:     polygon
  geometry_params:
    min_area_sqm: 25.0             # ignore slivers and annotation polygons
  code:         "PL"
  output:       "site_polygons.geojson"
  fields:
    Code:         null
    Plot_No:      {from_dwg_name: true}
    Zone_ID:      {spatial_join: secondary}
    Area_Ha:      {calculate: Area_Ha}
    Perimeter_Km: {calculate: Perimeter_Km}
    Land_Use:     " "
    Owner:        " "
```

---

#### Using match_mode: prefix

When multiple AutoCAD layers share a prefix (e.g. `POLY-LAYER-A`, `POLY-LAYER-B`), use `prefix` mode to capture all of them in one pass.

```yaml
- name:         "Polygon Layer"
  locked:       false
  source_dwg:   drawing_a
  source_layer: "POLY-LAYER"        # prefix: captures POLY-LAYER-A, POLY-LAYER-B, etc.
  match_mode:   prefix
  geometry:     polygon
  code:         "PL"
  output:       "polygon_layer.geojson"
  fields:
    Code:           null
    Category:       "Polygon"
    Zone_ID:        {spatial_join: primary}
    Connection_ID:  {auto_sequence: true}
    Sub_Type:       {from_merge_source: layer_subtype}  # auto-detected from layer suffix
    Length_m:       {calculate: Length_m}
```

---

#### Layer fallbacks

If a CAD layer might have slightly different names across DWG revisions, declare fallbacks:

```yaml
  source_layer: "ZONE-BOUNDARY"
  fallbacks:
    - "ZONE_BOUNDARY"
    - "Zone Boundary"
    - "Zone-Boundary"
```

Fallbacks are tried in order. The first match found is used.

---

### Type 4 — Derived Child Layer

Creates N child point features from each parent feature by expanding block attribute tags. One child point is generated per non-blank attribute tag in the declared `id_fields` list. Child points are positioned evenly along the long axis of the parent's bounding box.

**When to use:** when one parent block (cabinet, rack, panel) represents a multi-slot asset and each slot is a separate attribute tag. Instead of one point per cabinet, you get one point per occupied slot.

```yaml
- name:              "Cabinet Ports"
  derive_from:       parent_layer
  parent_layer_name: "Equipment Cabinets"  # must match an earlier layer's name:
  locked:            false
  output:            "cabinet_ports.geojson"
  code:              "CAB_PORT"
  id_fields:
    - [PORT_01, PORT1]
    - [PORT_02, PORT2]
    - [PORT_03, PORT3]
    - [PORT_04, PORT4]
  fields:
    Code:        "CAB_PORT"
    Port_ID:     {from_attr: Child_ID}
    Zone_ID:     {spatial_join: secondary}
    Cabinet_Ref:
      spatial_join:
        method:     nearest
        from_layer: "Equipment Cabinets"
        from_field: "Connection_ID"
```

**Note:** the parent DWG must be open in AutoCAD. Derived child layers always re-extract the parent live from AutoCAD — they cannot use the cached GeoJSON.

---

## 11. Field Reference

All fields are declared under `fields:` in a layer config. The YAML key order is the exact property order in the output GeoJSON — this controls GIS attribute table column order.

Two fields are always present on every feature and do not need to be declared:

| Field | Value | Notes |
|---|---|---|
| `OBJECTID` | Integer, 1-based, resets per layer | Sequence within this layer only |
| `Connection_ID` | Auto-generated composite key | See [Section 15](#15-the-connection_id-system) |

---

### Constant value

Written identically to every feature.

```yaml
Category:   "Zone Boundaries"
Year:       2024
Status:     "Draft"
Empty_Field: " "    # single space — standard blank placeholder
Null_Field:  null   # written as " " in output
```

---

### from_dwg_name

Reads the zone/plot ID from the DWG filename. See [Section 9](#9-spatial-reference-id-from-dwg-filenames).

**The field must be named `Plot_No`.** The executor uses this specific field name when re-deriving `Connection_ID`. If named anything else, `Connection_ID` will not be updated with the correct zone.

```yaml
Plot_No:
  from_dwg_name: true
```

---

### from_attr

Reads a block attribute tag value from the AutoCAD entity. Tag name is case-insensitive (normalised to uppercase before lookup).

```yaml
Asset_Type:
  from_attr: "TYPE"

Attribute_1:
  from_attr: "ATTR_1"

Serial_No:
  from_attr: "SERIAL_NO"
```

---

### block_attr

Same as `from_attr` but with a fallback list. The executor tries each tag name in order, returning the first non-null value found. Useful when tag names vary between block definitions or DWG revisions.

```yaml
Asset_ID:
  block_attr: "ASSET_ID"
  fallbacks:
    - "ASSETID"
    - "ASSET"
    - "ID"
```

---

### calculate

Computes a geometry measurement. The name must match an entry in `global.calculated_fields`.

```yaml
Area_Ha:
  calculate: Area_Ha

Perimeter_Km:
  calculate: Perimeter_Km

Length_m:
  calculate: Length_m
```

In the global section:

```yaml
global:
  calculated_fields:
    Area_Ha:      {formula: area,      unit: hectares, round: 2}
    Perimeter_Km: {formula: perimeter, unit: km,       round: 2}
    Length_m:     {formula: length,    unit: meters,   round: 2}
```

| formula | Use on | Output |
|---|---|---|
| `area` | polygon | Area in declared unit |
| `perimeter` | polygon | Perimeter in declared unit |
| `length` | line | Length in declared unit |

---

### spatial_join

Assigns a value by finding the nearest matching feature in another layer's spatial cache.

**Primary join** — uses `global.block_no.primary_source` (typically reads text labels):

```yaml
Zone_ID:
  spatial_join: primary
```

**Secondary join** — uses `global.block_no.secondary_source` (typically joins from the spatial reference layer):

```yaml
Zone_ID:
  spatial_join: secondary
```

**Custom nearest join** — finds nearest feature from any cached layer:

```yaml
Nearest_Asset:
  spatial_join:
    method:     nearest
    from_layer: "Equipment Points"
    from_field: "Connection_ID"
```

**Nearest endpoint (line layers)** — finds nearest feature to the line's start and end points separately. extractio infers which field is start vs. end from the field name:

```yaml
Start_Node:
  spatial_join:
    method:     nearest_endpoint
    from_layer: "Junction Boxes"
    from_field: "Connection_ID"

End_Node:
  spatial_join:
    method:     nearest_endpoint
    from_layer: "Junction Boxes"
    from_field: "Connection_ID"
```

**With a format transform** — builds a composite string from the join result and other resolved fields:

```yaml
Full_Ref:
  spatial_join:
    method:     nearest
    from_layer: "Zone Boundaries"
    from_field: "Zone_ID"
    transform:
      format: "{Plot_No}-{Zone_ID}"
```

---

### from_merge_source

Used only in layers with `merge_sources`. Stamps per-source metadata onto features.

| Key | Returns |
|---|---|
| `sub_plot` | The `sub_plot:` value from the matching merge_source entry |
| `code` | The `code:` value from the matching merge_source entry |
| `layer_subtype` | Sub-type from detected layer suffix or `sub_type:` in merge_source |
| `sub_classification` | The `sub_classification:` value from the merge_source entry |

```yaml
Sub_Plot:
  from_merge_source: sub_plot

Sub_Type:
  from_merge_source: layer_subtype
```

---

### from_config

Reads a value from a custom key defined anywhere in the `global:` section. Useful for project-wide constants referenced in multiple layers.

```yaml
# In global:
global:
  route_end_point: "End Point"

# In a layer's fields:
End_Connection:
  from_config: route_end_point
```

---

### resolver (composable chain)

Walk a list of resolvers in order; the first one that returns a non-null, non-empty value wins. Use this when a field should try multiple sources before falling back to a default.

```yaml
Zone_ID:
  resolver:
    - type: from_attr        # 1. try reading a block attribute tag first
      tag:  ZONE_TAG
    - type: spatial_join     # 2. fall back to spatial join from primary source
      source: primary
    - type: fallback         # 3. last resort constant
      value: null
```

Available resolver types:

| type | Parameters | Description |
|---|---|---|
| `from_attr` | `tag` | Read a block attribute tag |
| `calculate` | `name` | Named geometry calculation |
| `spatial_join` | `source` (`primary`/`secondary`) | Spatial join |
| `from_config` | `key` | Value from global config |
| `from_merge_source` | `key` | Per-merge-source metadata |
| `fallback` | `value` | Constant fallback value |

---

### derive

Computes a value from other fields already resolved in pass 1. Runs in pass 2. See [Section 14](#14-derive-transforms-reference) for all transforms.

```yaml
Zone_Block:
  derive:
    transform: auto_sequence
    prefix:    "BLK"
    pad:       2
```

---

## 12. Geometry Params Reference

Declared under `geometry_params:` in a layer config. All keys are optional.

### For point / INSERT layers

| Key | Type | Description |
|---|---|---|
| `only_insert` | bool | Accept only INSERT entities. Skips POINT, TEXT, MTEXT on the same layer. Use for all symbol/block layers. |
| `block_name` | string | Accept only INSERTs whose block definition name matches this value. |

### For polygon layers

| Key | Type | Description |
|---|---|---|
| `only_lwpolyline` | bool | Accept only LWPOLYLINE entities; skips INSERT blocks. |
| `min_area_sqm` | float | Skip polygons with area below this threshold. Filters slivers and annotation polygons. |
| `target_area_sqm` | float | Accept only polygons within `tolerance` sqm of this area. |
| `tolerance` | float | Area tolerance for `target_area_sqm` (default: 5.0). |
| `vertex_count` | int | Accept only polygons with exactly this many vertices. |
| `forced_rotation_deg` | float | Override the INSERT entity's Rotation when transforming block-local vertices to world coordinates. |
| `use_polyline_width` | bool | Expand an LWPOLYLINE into a polygon using its ConstantWidth. |
| `half_width_m` | float | Default half-width when `use_polyline_width` is true and the entity has no explicit width (default: 0.025). |
| `rotate_90` | bool | Rotate polygon geometry 90°. Corrects blocks drawn in landscape orientation. |
| `local_pts` | list of [x, y] | Hardcoded block-local polygon vertices for blocks whose geometry cannot be extracted normally. |

### For line layers

| Key | Type | Description |
|---|---|---|
| `linetype` | string | Accept only LWPOLYLINE entities whose AutoCAD Linetype matches this string (case-insensitive). |

---

## 13. Spatial Join Reference

### Cache registration order

1. **Spatial reference layer** — registered before pass 2 begins
2. **Declared join sources** — any layer referenced by `from_layer:` in any field
3. **Each output layer** — registered after extraction in pass 2

This means: Layer B can only join against Layer A if Layer A appears earlier in the YAML config (or is a pass-1 source).

### spatial_join: primary / secondary

Uses point-in-polygon containment to find the reference polygon that contains the feature's centroid. Falls back to nearest centroid if the feature lies outside all reference polygons (handles edge effects and coordinate noise).

### spatial_join: nearest

Finds the feature with the smallest Euclidean distance from this feature's centroid. No radius cutoff. Warns once per (field, from_layer) pair if the named layer is not in the cache.

### spatial_join: nearest_endpoint

For line layers. Independently finds the nearest cached feature to the line's first vertex (start) and last vertex (end). The field whose name contains "start" receives the start value; the field containing "end" receives the end value.

### spatial_join: nearest_exclusive

A batch post-processing mode ensuring one-to-one assignment — each source feature is matched to at most one target feature per zone. Used when guaranteed uniqueness is required (e.g. each cable segment must connect to a distinct junction point). This runs after the layer is fully extracted.

### Miss warnings

When a spatial join returns no result, a warning is printed once per `(field_name, from_layer)` pair. Repeated misses on the same combination are suppressed.

---

## 14. Derive Transforms Reference

`derive` transforms run in pass 2 after all pass-1 fields are resolved. They compute values from other resolved fields.

---

### `auto_sequence`

Zero-padded sequence string from `OBJECTID`.

```yaml
Block_No:
  derive:
    transform: auto_sequence
    prefix:    "BLK"
    pad:       2
# OBJECTID=4 → "BLK04",  OBJECTID=12 → "BLK12"
```

---

### `block_no_to_connection`

Converts `ZONE_BLKnn` to `ZONE-BLnn`.

```yaml
Linked_Node:
  derive:
    transform:  block_no_to_connection
    from_field: Block_No
# "Z1_BLK03" → "Z1-BL03"
```

---

### `block_no_to_prefixed_connection`

Like above but prepends a prefix.

```yaml
Asset_Ref:
  derive:
    transform:  block_no_to_prefixed_connection
    from_field: Block_No
    prefix:     "REF-"
# "Z1_BLK03" → "REF-Z1-BLK03"
```

---

### `extract_last_sequence`

Takes the last `_`-delimited numeric segment and reformats it.

```yaml
Seq_No:
  derive:
    transform:  extract_last_sequence
    from_field: Connection_ID
    prefix:     "SEQ-"
    pad:        3
# "Z1_PT_07" → "SEQ-007"
```

---

### `strip_last_segment`

Removes the last two separator-delimited segments. Useful for deriving parent IDs from child connection IDs.

```yaml
Parent_Ref:
  derive:
    transform:  strip_last_segment
    from_field: Connection_ID
    separator:  "_"
# "Z1_ASSET_TYPE_05" → "Z1_ASSET"
```

---

### `count_filled`

Counts how many fields in a list have non-null values and maps the count to a string.

```yaml
Occupancy:
  derive:
    transform:   count_filled
    from_fields: [SLOT_01, SLOT_02, SLOT_03, SLOT_04]
    value_map:
      0: "Empty"
      2: "Half"
      4: "Full"
```

---

### `extract_suffix`

Extracts the numeric suffix from a `_BLKnn`-formatted value.

```yaml
Short_Ref:
  derive:
    transform:  extract_suffix
    from_field: Block_No
    prefix:     "BL"
# "Z1_BLK05" → "BL05"
```

---

### `format_reference_id`

Formats a `_BLKnn` value into a document reference code.

```yaml
Doc_Ref:
  derive:
    transform:  format_reference_id
    from_field: Block_No
    prefix:     "DOC-"
# Block_No="Z1_BLK03" → "DOC-Z1-BL-03"
```

---

### `prepend_plot`

Prepends the plot/zone ID to a block text value.

```yaml
Label:
  derive:
    transform: prepend_plot
    format:    "{plot_no}-{text}"
# Plot_No="Z1", block_text="NODE-01" → "Z1-NODE-01"
```

---

## 15. The Connection_ID System

Every feature automatically receives a `Connection_ID`. This is a composite string key that uniquely identifies a feature within the project.

**Default format:** `{Plot_No}_{Code}_{seq:02d}`

Override the pattern in `global.connection_id`:

```yaml
global:
  connection_id:
    pattern: "{plot_name}_{code}_{seq:02d}"
```

- `{plot_name}` — resolved value of the `Plot_No` field
- `{code}` — the layer's `code:` value, or per-source code from `merge_sources`
- `{seq:02d}` — `OBJECTID`, zero-padded

**Examples:**

| Scenario | Connection_ID |
|---|---|
| Zone boundary #1, plot A1, code ZB | `A1_ZB_01` |
| Equipment point #17, plot A1a, code PT | `A1a_PT_17` |
| Route line #4, plot A1b, code LN | `A1b_LN_04` |

`Connection_ID` is seeded at the start of field resolution with a placeholder, then **re-derived after pass 1** once `Plot_No` and `Code` have their final values. The final `Connection_ID` always reflects the actual resolved zone and code.

`Connection_ID` is used as the join key in `spatial_join: nearest` configurations — allowing inter-layer linkage (e.g. each line segment referencing the junction box it connects to) to be fully automated.

---

## 16. Multi-DWG Projects

Most real projects have multiple DWG files — one per zone, or one per drawing type (layout, routing, boundary).

### Declaring DWG aliases

Keep absolute file paths in `dwg_paths.yaml` (gitignored). Reference them by alias everywhere else.

```yaml
# dwg_paths.yaml
drawing_a: "D:/Projects/MyProject/CAD/Drawing-A.dwg"
drawing_b: "D:/Projects/MyProject/CAD/Drawing-B.dwg"
drawing_c: "D:/Projects/MyProject/CAD/Drawing-C.dwg"
drawing_d: null    # null = file not yet received
```

### Merging from multiple DWGs into one layer

```yaml
- name:         "All Equipment"
  geometry:     point
  match_mode:   exact
  source_layer: "POINT-LAYER"
  code:         "PT"
  merge_sources:
    - source_dwg:   drawing_a
      sub_plot:     "A"
      source_layer: "POINT-LAYER"
    - source_dwg:   drawing_b
      sub_plot:     "B"
      source_layer: "POINT-LAYER"
    - source_dwg:   drawing_c
      sub_plot:     "C"
      source_layer: "POINT-LAYER"
  fields:
    Sub_Plot:    {from_merge_source: sub_plot}
    Plot_No:     {from_dwg_name: true}
    Zone_ID:     {spatial_join: primary}
    ...
```

All three DWGs must be open in AutoCAD. If one is missing, that sub-plot is skipped with a warning and extraction continues.

### Zone ID per DWG

`{from_dwg_name: true}` returns a different ID for each DWG based on filename. Combined with `{spatial_join: secondary}` for the reference polygon lookup, features from different DWGs automatically receive the correct zone assignment.

---

## 17. Working with Block Attributes

### What block attributes are

A **block** in AutoCAD is a named symbol definition. When inserted into a drawing, each INSERT entity holds its own values for the block's named attribute fields (ATTDEF). For example, an "Equipment" block might have attributes `TYPE`, `SERIAL_NO`, `CAPACITY`.

### Workflow

**Step 1 — Identify the layer and block type.**

```
python extractio.py config.yaml --dwg-layers drawing_a
```

Find the layer containing your block INSERT entities. In AutoCAD, double-click one entity to see its attribute tags in the attribute editor.

**Step 2 — Configure the layer.**

```yaml
- name:         "Equipment Points"
  geometry:     point
  source_layer: "POINT-LAYER"
  geometry_params:
    only_insert: true           # only block INSERT entities — skip text, POINT
    block_name:  "POINT_BLOCK"  # optional: filter by block definition name
  fields:
    Asset_Type:    {from_attr: "TYPE"}
    Attribute_1:   {from_attr: "ATTR_1"}
    Serial_No:     {from_attr: "SERIAL_NO"}
```

**Step 3 — Run and verify.**

The output GeoJSON will have one feature per block instance, with properties shaped exactly as declared.

### Tag name lookup

- Tag names in `from_attr` and `block_attr` are **normalised to uppercase** before lookup. `from_attr: "serial_no"` and `from_attr: "SERIAL_NO"` are identical.
- If a tag is not found, the field receives `" "` (a single space).

### Finding unknown tag names

Run `--dwg-layers` to list layers, then in AutoCAD double-click a representative entity and open the attribute editor. Or use AutoCAD's `LIST` command on the entity.

To diagnose at runtime, temporarily add a field with the suspected tag name — if the output value is `" "`, the tag name is wrong.

### Blocks without attributes

If a block has no attribute definitions, `from_attr` and `block_attr` return `" "`. All per-feature metadata must then come from constants, spatial joins, or `from_dwg_name`.

---

## 18. Advanced Config Features

### 18.1 Config Inheritance

Declare a parent config with `inherits:`. The parent is deep-merged first; the child config overrides.

```yaml
# project_A.yaml — inherits common global settings
inherits: shared_base.yaml

global:
  project_name: "MyProject_A"
  output_dir:   "./outputs/A"
```

```yaml
# shared_base.yaml — common settings for all projects
global:
  crs:         "EPSG:32632"
  common_fields:
    Organisation: "My Organisation"
    Prepared_By:  "GIS Team"
  calculated_fields:
    Area_Ha: {formula: area, unit: hectares, round: 2}
```

Inheritance is recursive — a parent can itself inherit from a grandparent.

---

### 18.2 Environment Variables

Use `$VAR` or `${VAR}` anywhere in your YAML. Values are substituted from a `.env` file in the same folder as the config, then from `os.environ`.

```yaml
# config.yaml
global:
  output_dir: $OUTPUT_DIR
  crs:        "${DEFAULT_CRS}"
```

```bash
# .env (same folder, gitignored automatically)
OUTPUT_DIR=D:/Projects/MyProject/Output
DEFAULT_CRS=EPSG:32632
```

The `.env` file is never committed — it holds machine-specific paths and secrets. Add it to `.gitignore`.

---

### 18.3 Layer Library

Define reusable layer templates in `global.layer_library`. Reference them with `use: template_name`. The layer's own keys override the template.

```yaml
global:
  layer_library:
    line_defaults:
      geometry:    line
      match_mode:  exact
      locked:      false

layers:
  - name:         "Route Lines"
    use:          line_defaults     # inherits geometry, match_mode, locked
    source_dwg:   drawing_c
    source_layer: "LINE-LAYER"
    code:         "LN"
    output:       "route_lines.geojson"
    fields:
      ...

  - name:         "Polygon Layer"
    use:          line_defaults     # same template
    source_dwg:   drawing_a
    source_layer: "POLY-LAYER"
    code:         "PL"
    output:       "polygon_layer.geojson"
    fields:
      ...
```

---

### 18.4 Common Field Injection Control

Three optional control keys inside `common_fields:`:

```yaml
global:
  common_fields:
    inject_position: first   # "first" (default), "last", or "after:FieldName"

    field_overrides:         # always stamp — layer fields cannot override these
      Prepared_By: "GIS Team"
      Data_Source: "AutoCAD"

    exclude_fields:          # never inject these on any layer
      - Status

    # Regular common fields below:
    Project_Name: "My Project"
    Organisation: "My Organisation"
```

| Key | Description |
|---|---|
| `inject_position: first` | Common fields appear before each layer's own fields (default) |
| `inject_position: last` | Common fields appended after each layer's own fields |
| `inject_position: "after:FieldName"` | Common fields inserted after the named anchor field |
| `field_overrides` | Always-stamp values; cannot be overridden by layer fields |
| `exclude_fields` | Fields in this list are never injected on any layer |

---

### 18.5 Composable Resolver Chains

When a field should try multiple sources in sequence, use a `resolver:` list instead of a single directive. The first resolver that returns a non-null, non-empty value wins.

```yaml
Zone_ID:
  resolver:
    - type: from_attr          # 1st: try reading a block attribute
      tag:  ZONE_TAG
    - type: spatial_join       # 2nd: fall back to spatial join
      source: primary
    - type: fallback           # 3rd: hard fallback
      value: null
```

This is useful for layers where some entities have an attribute already set (from CAD annotation) while others need the value derived spatially.

---

### 18.6 Post-Extraction Validation

Add an `expect:` block to any layer to validate the output feature count after extraction. Mismatches are logged as warnings.

```yaml
- name: "Zone Boundaries"
  ...
  expect:
    min_features: 40      # warn if fewer than 40 features extracted
    max_features: 60      # warn if more than 60 features extracted
  fields:
    ...
```

If the count falls outside the declared range, you see:

```
⚠  EXPECT FAIL: 'Zone Boundaries' has 12 features  (min: 40)
```

If within range:

```
✓  expect: OK  (50 features)
```

Use this to catch DWG changes, layer-name mismatches, or filtering that is too aggressive, before the output reaches downstream tools.

---

## 19. LLM YAML Builder

The LLM builder uses Groq or Gemini to generate a config file from reference documents and plain-English descriptions. **AutoCAD does not need to be running.** The builder is completely independent of AutoCAD — it reads documents and calls an LLM, nothing else.

```
# Set your API key (Windows):
set GROQ_API_KEY=your_key_here

# Run the builder — no AutoCAD required:
python extractio.py config.yaml --build
```

### How it works

The builder collects context from up to four sources, then sends them all to the LLM:

| Source | How it is provided | AutoCAD needed? |
|---|---|---|
| Reference documents (spec sheets, field lists, existing YAMLs) | `--ref` flag or interactive prompt | No |
| DWG scan JSON (exact layer names, entity types, attribute tags) | auto-detected `scan_*.json` in script folder, or `--from-scan` | No (scan itself needs AutoCAD, but only once) |
| User notes / free-text prompt | typed interactively | No |

The LLM reads all provided context and generates one config entry per identified layer. You then review each entry and confirm, skip, or correct it before the YAML is written.

### Basic usage — reference documents only

Pass spec sheets, data dictionaries, or an existing YAML as context:

```
python extractio.py config.yaml --build --ref spec_sheet.pdf
python extractio.py config.yaml --build --ref field_list.docx existing_config.yaml
```

If no `--ref` is given, the builder prompts you to enter document paths or type a description interactively.

### Adding DWG layer names for accuracy

For the LLM to use exact AutoCAD layer names (rather than guessing from document text), generate a scan JSON first. This is the only step that requires AutoCAD:

```
# Step 1 — scan DWGs once (AutoCAD must be open with DWGs loaded):
python extractio.py config.yaml --scan-only

# Step 2 — build YAML offline, no AutoCAD needed:
python extractio.py config.yaml --build --from-scan scan_20250101_120000.json
```

The scan JSON is saved to the script folder with a timestamped name. Subsequent `--build` runs auto-detect the most recent scan file and offer to include it.

### Skip the review loop

Auto-accept all generated layers without interactive confirmation:

```
python extractio.py config.yaml --build --accept-all
```

### Clone from an existing config

Remap an existing config's layer definitions to a new set of DWGs (aliases updated, layer names fuzzy-matched):

```
python extractio.py config.yaml --clone-from other_config.yaml
```

---

## 20. CLI Reference

```
python extractio.py [config.yaml] [options]
```

The config path is optional. If omitted, extractio searches the script folder for `config.yaml` or a single `*.yaml` file.

### Commands

| Command | Description |
|---|---|
| *(no options)* | Launch the interactive layer selector |
| `--list` | Print all layers with lock status, source layer, and exit |
| `--run` or `--run all` | Run all unlocked layers without interaction |
| `--layers "Name A" "Name B"` | Run the named layers regardless of lock state |
| `--unlock "Name"` | Set `locked: false` in the YAML for this layer, then exit |
| `--lock "Name"` | Set `locked: true` in the YAML for this layer, then exit |
| `--unlock-all` | Set `locked: false` for all layers, then exit |
| `--lock-all` | Set `locked: true` for all layers, then exit |
| `--status` | Same as `--list` |
| `--validate` | Validate config structure and exit |
| `--dwg-layers DWG_KEY` | Print all CAD layer names and entity counts from the DWG aliased as `DWG_KEY`, then exit |
| `--scan-only` | Connect to AutoCAD, scan all DWGs, save a layer-name snapshot JSON, exit |
| `--build` | LLM YAML builder — **no AutoCAD required**; reads reference docs and optional scan JSON |
| `--from-scan FILE` | Provide a specific scan JSON to `--build` (overrides auto-detection) |
| `--no-autocad` | Explicitly skip AutoCAD connection check when using `--build` |
| `--versions-dir DIR` | Override default version snapshot directory |
| `--output-format FORMAT` | `geojson` (default) or `gdb` — overrides `output_format` in config |

### Layer name matching

Layer names passed to `--layers`, `--unlock`, and `--lock` are fuzzy-matched. You do not need to type the full exact name. CLI aliases declared in a layer's `cli_aliases:` list are also accepted.

### Interactive mode

Running with no `--run` or `--layers` argument launches the interactive layer selector:

```
  [ 1] [ON ] 🔓 Zone Boundaries
  [ 2] [OFF] 🔓 Site Boundary
  [ 3] [ON ] 🔓 Equipment Points
  [ 4] [OFF] 🔒 Route Lines
────────────────────────────────────────────────────────
  [A] All ON   [N] All OFF   [X] Run
```

Type a number to toggle, `A` to select all unlocked, `N` to deselect all, `X` to run.

### Practical sequences

**Full project from scratch:**

```
python extractio.py config.yaml --dwg-layers drawing_a
python extractio.py config.yaml --unlock "Zone Boundaries"
python extractio.py config.yaml --layers "Zone Boundaries"
python extractio.py config.yaml --lock "Zone Boundaries"
python extractio.py config.yaml --unlock-all
python extractio.py config.yaml --run all
python extractio.py config.yaml --lock-all
```

**Re-run a single layer:**

```
python extractio.py config.yaml --unlock "Equipment Points"
python extractio.py config.yaml --layers "Equipment Points"
python extractio.py config.yaml --lock "Equipment Points"
```

**Using aliases (if declared in config):**

```
python extractio.py config.yaml --layers zb pts   # "zb" = Zone Boundaries, "pts" = Equipment Points
```

---

## 21. Lock System

### What locks do

Every layer in the YAML has a `locked:` boolean. When `locked: true`, extractio will not write that layer's output file, even if the layer is selected. This prevents accidental overwrites of verified outputs when re-running for other layers.

### Default state

Start with `locked: true` on all layers. Unlock explicitly when ready to extract. Re-lock after verifying the output.

### How locks are stored

The lock state is stored directly in the YAML config file. `--lock` and `--unlock` rewrite the `locked:` line for that layer in the file. You can also edit it manually in a text editor.

### Exception: derived layers

Layers with `derive_from:` (`spatial_reference` or `parent_layer`) ignore the lock flag. They always run because they perform computation, not DWG extraction. There is no risk of overwriting raw extraction data.

### Recommended workflow

1. Extract and verify the spatial reference layer (`Zone Boundaries`) first → lock it
2. Extract the first batch of layers → verify output in QGIS or another GIS tool
3. Lock verified layers → move to the next batch
4. Repeat until all layers are extracted and locked
5. For a full re-extraction after DWG updates: `--unlock-all` then `--run all`

---

## 22. Output Format

extractio supports two output formats: **GeoJSON** (default) and **File Geodatabase (GDB)**. The format is selected per-run and does not change anything in the extraction logic — only the final write step differs.

### File structure

```json
{
  "type": "FeatureCollection",
  "name": "equipment_points",
  "crs": {
    "type": "name",
    "properties": { "name": "urn:ogc:def:crs:EPSG::32632" }
  },
  "features": [
    {
      "type": "Feature",
      "geometry": { "type": "Point", "coordinates": [453201.4, 2841093.7] },
      "properties": {
        "OBJECTID": 1,
        "Connection_ID": "A1a_PT_01",
        "Code": "PT",
        "Category": "Equipment Points",
        "Sub_Plot": "A",
        "Zone_ID": "ZN_BLK03",
        "Asset_Type": "Type A",
        "Attribute_1": "value_1"
      }
    }
  ]
}
```

### GeoJSON output (default)

Each layer writes one `.geojson` file (RFC 7946) with an additional `crs` member for compatibility with GIS tools that require explicit CRS declarations.

| Item | Behaviour |
|---|---|
| Coordinates | In the CRS declared in `global.crs` |
| Null / blank values | Written as `" "` (a single space string), not JSON null |
| OBJECTID | Resets to 1 for each layer independently |
| File name | Set by `output:` in the layer config; defaults to `LayerName.geojson` |
| Overwrite | Existing file is overwritten silently on each run (for unlocked layers) |
| Large layers | Features > 500: written in compact JSON (no indent) for speed and file size |
| Small layers | Features ≤ 500: written with `indent=2` for readability |

### File Geodatabase output (GDB)

All layers are written as named feature classes inside a single `.gdb` folder instead of individual `.geojson` files.

**Requirement:** `pip install fiona` with GDAL 3.6 or later (the OpenFileGDB write driver is bundled in GDAL 3.6+ — no ArcGIS licence required).

Verify your GDAL version:

```bash
python -c "import fiona; print(fiona.__gdal_version__)"
```

**Select GDB output at runtime:**

```bash
python extractio.py config.yaml --output-format gdb
```

**Or set it permanently in `config.yaml`:**

```yaml
global:
  output_format: gdb          # geojson (default) or gdb
  gdb_name: MyProject.gdb    # optional; defaults to <project_name>.gdb
```

**`--output-format` on the command line always wins over `output_format` in config.**

| Item | Behaviour |
|---|---|
| GDB location | `<output_dir>/<gdb_name>` (or `<project_name>.gdb` if `gdb_name` omitted) |
| Layer naming | Uses the layer `name:` field from the YAML config |
| Field types | Inferred from the first non-null value in each column (`int`, `float`, `str`) |
| Null values | Written as `null` (native GDB null, not a space string) |
| Driver | `OpenFileGDB` via fiona — no ArcGIS installation needed |
| GDAL requirement | GDAL 3.6 or later |

**Note:** the GDB option writes all layers that are extracted in a single run into the same `.gdb` directory. Layers from separate runs append new feature classes — they do not overwrite existing ones unless the layer name matches exactly.

### Version manifest

After each run, extractio saves a JSON manifest in `~/.cad_tract_versions/<project>/`:

```json
{
  "version": 3,
  "timestamp": "2025-04-14T11:32:10",
  "config_file": "config.yaml",
  "layers": [
    {"layer_name": "Zone Boundaries", "feature_count": 50, ...}
  ],
  "diff": {
    "added": [],
    "changed": [{"layer": "Equipment Points", "feature_count": "240 → 242"}],
    "unchanged": ["Zone Boundaries", "Site Boundary"]
  },
  "resolved_config": {
    "project_name": "MyProject",
    "crs": "EPSG:32632",
    "layers": [...]
  }
}
```

The manifest includes `resolved_config` — the fully merged and inherited config state at run time, so you can reproduce any extraction exactly.

---

## 23. Config and Version Management

### File layout

```
project-folder/
├── extractio.py              ← the engine (tracked in git)
├── config.yaml               ← your project config (gitignored)
├── dwg_paths.yaml            ← machine-specific DWG paths (gitignored)
├── .env                      ← environment variables (gitignored)
├── samples/
│   ├── sample_config.yaml    ← generic annotated template (tracked in git)
│   └── sample_dwg_paths.yaml ← DWG paths template (tracked in git)
└── outputs/
    ├── zone_boundaries.geojson
    └── equipment_points.geojson
```

### What to commit

- `extractio.py` — the tool itself
- `samples/sample_config.yaml` — the generic template
- `samples/sample_dwg_paths.yaml` — the paths template
- `README.md`
- `.gitignore`

### What NOT to commit

- Real config files (`config.yaml`, `*_config.yaml`) — contain client paths and data
- `dwg_paths.yaml` — contains machine-specific absolute paths
- `.env` — may contain API keys
- `outputs/` — GeoJSON files (large, derived, not source)
- `__pycache__/`, `*.pyc`

The `.gitignore` in this repository already excludes all of these.

### Sharing configs in a team

- Share configs through internal channels (not public repos)
- Lock layers that one team member has verified to prevent others from accidentally overwriting them
- Use `--status` to communicate which layers are complete

### Config inheritance for multi-project organisations

If you manage many similar projects, define a `shared_base.yaml` with common settings (CRS, calculated_fields, common_fields) and have each project config inherit from it:

```yaml
# project_A.yaml
inherits: shared_base.yaml
global:
  project_name: MyProject_A
  output_dir:   ./outputs/A
```

---

## 24. Performance Notes

### Typical throughput

| Layer type | Features/second |
|---|---|
| Point (INSERT, no complex geometry) | 200–500 |
| Polygon (LWPOLYLINE) | 50–200 |
| Layers with spatial joins | Slightly slower per feature |

### Biggest single optimisation

Lock the spatial reference layer (`Zone Boundaries`) after first extraction. Subsequent runs load it from disk in milliseconds instead of re-reading AutoCAD.

### Progress display

Two live progress bars are shown: one for the current layer's entity count, one for the overall layer count. Both show elapsed time and estimated time remaining.

### Large DWGs

AutoCAD COM slows with tens of thousands of entities:

- Use `only_insert: true` on point layers to skip text and other entity types
- Use `min_area_sqm` on polygon layers to skip small annotation polygons early
- Use `linetype` filtering on line layers when only one linetype is needed

---

## 25. Troubleshooting

---

### `Cannot connect to AutoCAD`

AutoCAD is not running or is not accessible via COM.

**Fix:** start AutoCAD, open your DWG files, then run again. If the error persists with AutoCAD open:
- Ensure Python and AutoCAD are running as the same Windows user (not one elevated)
- Run from a standard Command Prompt, not a terminal inside an IDE
- Restart AutoCAD and try again

---

### `DWG not open in AutoCAD: myfile.dwg`

AutoCAD is running but this document is not open.

**Fix:** open the file in AutoCAD and run again. Path matching is case-insensitive and falls back to filename-only matching.

---

### `Layer 'MY-LAYER' not found in 'drawing_a'`

The `source_layer` value in your config does not match any layer name in that DWG.

**Fix:** run `--dwg-layers drawing_a` to see the exact list. AutoCAD layer names are case-sensitive and include all spaces, hyphens, and punctuation. Copy the name exactly.

---

### `Could not detect plot from myfile.dwg`

The zone/plot ID could not be extracted from the DWG filename.

**Fix:** either rename the DWG file to include a recognisable pattern (e.g. `Layout (Plot A1).dwg`), or replace `{from_dwg_name: true}` with a constant string:
```yaml
Plot_No: "A1"
```

---

### `spatial_join MISS: 'Zone_ID' — layer 'Zone Boundaries' not found in cache`

A spatial join tried to look up a zone but the spatial reference layer is not loaded.

**Fix:**
1. Extract the spatial reference layer first (or verify its GeoJSON exists in `output_dir`)
2. Confirm `global.block_no.primary_source.from_layer` matches the layer's `name:` value exactly
3. If using `secondary_source`, confirm `from_layer` matches the spatial reference layer name

---

### `spatial_join MISS: field found in cache but returned empty`

The layer is cached but the `from_field` value did not match.

**Fix:** check that `from_field` in `global.block_no` matches the exact field name in that layer's `fields:` section. Field names are case-sensitive.

---

### 0 features extracted for a layer

**Check:**
1. Remove all `geometry_params` temporarily — if features appear, a filter is eliminating them
2. Check `only_insert: true` — if the layer has no INSERT entities, this filters everything
3. Check `min_area_sqm` — if too high, all polygons are filtered
4. Add filters back one at a time to find which is eliminating features

---

### All point features at (0, 0) or the same location

**If using INSERT blocks:** add `geometry_params: only_insert: true`. Without this, MTEXT and text entities on the same layer are also read and their coordinates may default to (0, 0) or a text insertion point.

**If using derived child layer:** the parent DWG must be open in AutoCAD. The derived child layer always re-extracts the parent live.

---

### Polygon geometry looks rotated or distorted

**Rotation:** try `geometry_params: rotate_90: true`, or set `forced_rotation_deg` to the correct value.

**Scale:** the executor uses the INSERT entity's scale factors. Verify the block scale in AutoCAD Properties.

---

### Field values are `" "` instead of actual data

The block attribute tag was not found on this entity.

**Causes:**
- Tag name in `from_attr` is wrong — verify in AutoCAD's attribute editor
- This entity is not an INSERT (it's a POINT or MTEXT) — add `only_insert: true`
- The attribute exists but was left blank in the DWG

**Diagnosis:** run `--dwg-layers` and examine a representative entity in AutoCAD with the `LIST` command.

---

### `Connection_ID` contains `A?`

The zone/plot ID was not resolved — the fallback `A?` is in use.

**Causes:**
- `from_dwg_name` could not extract an ID from the DWG filename
- The spatial reference layer is not loaded, so spatial fallback failed

**Fix:** rename the DWG, replace `from_dwg_name: true` with a constant, or ensure the spatial reference layer is loaded.

---

### `expect:` warning fires unexpectedly

```
⚠  EXPECT FAIL: 'Zone Boundaries' has 12 features  (min: 40)
```

**Causes:**
- A DWG is missing or its layer name changed
- `geometry_params` filters are too aggressive
- The DWG was updated with fewer entities

**Fix:** run `--dwg-layers` to verify entity counts, check filter params, or update `expect.min_features` if the count is legitimately lower.

---

### `.env` variables not substituting

**Causes:**
- `.env` file is not in the same folder as `config.yaml`
- Variable name has a typo
- Value contains `=` — use quotes around the value in `.env`

**Diagnosis:** add a test field `Debug_Var: "$MY_VAR"` and check the output.

---

### `Multiple YAML files found` on startup

More than one `*.yaml` file is in the script folder.

**Fix:** specify the config path explicitly: `python extractio.py my_config.yaml --run all`

---

### `shapely import error` on derived boundary layers

A `derive_from: spatial_reference` layer requires shapely.

**Fix:** `pip install shapely`

---

## 26. Project Setup Checklist

### Before writing the config

- [ ] AutoCAD is installed and all DWG files are accessible
- [ ] You know which DWG layer contains your spatial reference grid (zone/block boundaries)
- [ ] You have a list of output layers you need to produce
- [ ] You know the coordinate reference system (UTM zone / EPSG code) for your project area

### Config setup

- [ ] All DWG file paths in `dwg_paths.yaml` are correct absolute paths
- [ ] `crs` is set to the correct EPSG code
- [ ] `output_dir` is set to a writable folder
- [ ] `global.block_no.primary_source.from_layer` matches the spatial reference layer's `name:` exactly
- [ ] Layer names in `source_layer` have been verified with `--dwg-layers` (not guessed)
- [ ] All layers start with `locked: false` (or whichever is your default) and are reviewed before first run
- [ ] The spatial reference layer appears before any layer using spatial joins in the YAML list
- [ ] Any layer used as a join source (`from_layer:`) appears before the layers that join against it
- [ ] `expect: {min_features: N}` is set on critical layers so failures are caught automatically

### Before running

- [ ] AutoCAD is open
- [ ] All DWGs listed in `dwg_paths.yaml` are open in AutoCAD
- [ ] Only the intended layers are unlocked (`--status` to check)
- [ ] Config file and `dwg_paths.yaml` are not in a publicly visible git repository

### After running

- [ ] Output GeoJSON files are verified in QGIS, ArcGIS, or another GIS tool
- [ ] Feature counts look correct (no suspiciously low or zero counts)
- [ ] Zone/block IDs and `Connection_ID` values are correct on a sample of features
- [ ] Verified layers are locked with `--lock "Layer Name"`
- [ ] Config is backed up or stored in the project's internal document management system

---

## 27. Files

| File | Purpose |
|---|---|
| `extractio.py` | Main extraction engine — the only file you run |
| `launcher.py` | Convenience wrapper (`python launcher.py`) with auto-discovery |
| `launcher.bat` | Windows shortcut — add folder to PATH for `extractio` command |
| `yaml_generator.py` | Interactive wizard to generate a config from the sample template |
| `diagnose.py` | Standalone DWG layer inspector — run before configuring a new layer |
| `verify.py` | Compares extracted GeoJSON against live AutoCAD data |
| `mock_test.py` | Replays extraction logic without AutoCAD (for offline debugging) |
| `samples/sample_config.yaml` | Full annotated config example covering all layer types and field resolvers |
| `samples/sample_dwg_paths.yaml` | DWG paths file template |

---
