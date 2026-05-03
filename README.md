# extractio

**extractio** extracts GIS data from live AutoCAD DWGs and writes clean GeoJSON — driven entirely by a YAML config file. No code changes needed when you add a new project or layer.

---

## What it does

- Connects to a running AutoCAD instance via COM
- Reads entities directly from open DWG files (no intermediate export step)
- Supports **polygon**, **point**, and **line** geometry
- Merges features from multiple DWGs into a single layer with per-source sub_plot tagging
- Computes area, perimeter, and length fields from geometry
- Assigns sequential Connection IDs (`PREFIX_PLOT_CODE_01 … N`)
- Performs spatial joins — stamps Block_No on every feature from MMS text labels in the DWG
- Derives **Plot Boundaries** by dissolving Block Boundary polygons (no separate source needed)
- Filters phantom reference polygons using MMS text label count as ground truth
- Validates source layer names before extraction and suggests corrections
- Saves a versioned manifest on every run (feature counts + diff vs. previous run)
- LLM-assisted YAML builder (`--build`) via Groq or Gemini API

---

## Requirements

**AutoCAD** must be open with the DWG files loaded before running.

```
pip install pywin32 pyyaml shapely
```

Optional extras:
```
pip install filelock groq google-generativeai pdfplumber python-docx pandas
```

---

## Quick start

### 1. Set up your config

Copy the sample and fill in your project values:

```
cp samples/sample_config.yaml config.yaml
cp samples/sample_dwg_paths.yaml dwg_paths.yaml
```

Edit `config.yaml` — set `project_name`, `crs`, `output_dir`, and layer definitions.  
Edit `dwg_paths.yaml` — fill in absolute paths to your DWG files.

Or use the wizard:
```
python yaml_generator.py
```

### 2. Run

```
# Interactive layer selector
python launcher.py --config config.yaml

# Run specific layers
python extractio.py config.yaml --layers "Block Boundaries" "Plot Boundaries"

# Run all unlocked layers
python extractio.py config.yaml --run all
```

### 3. See what layers are defined
```
python extractio.py config.yaml --list
```

---

## CLI reference

```
python extractio.py <config.yaml> [options]

  --layers NAME [NAME ...]   Run specific layers (fuzzy names, aliases accepted)
  --run all                  Run all unlocked layers
  --list                     List all layers with lock status and exit
  --unlock LAYER [...]       Unlock layer(s) and exit
  --lock   LAYER [...]       Lock layer(s) and exit
  --unlock-all               Unlock all layers and exit
  --validate                 Validate config and exit
  --dwg-layers DWG_KEY       Print all layer names in a DWG and exit
  --scan-only                Scan DWGs, save snapshot JSON, exit
  --build                    LLM YAML builder mode
  --from-scan FILE           Use saved DWG scan JSON instead of AutoCAD
  --no-autocad               Skip AutoCAD connection
  --versions-dir DIR         Override version snapshot directory
```

---

## Config file structure

```yaml
global:
  project_name: SolarPark_P1
  crs: "EPSG:32643"
  output_dir: "./outputs"
  dwg_paths_file: dwg_paths.yaml   # resolves DWG aliases
  source_dwgs: [array_layout_P1a, array_layout_P1b]

  common_fields:                   # stamped on every feature
    Plant_Name: "My Solar Plant"
    Country:    "India"

  deferred_fields:                 # left null; filled from asset management later
    - Development_Status
    - Owned_By

  calculated_fields:               # available via {calculate: <name>}
    Area_Ha:      {formula: area,      unit: hectares, round: 2}
    Perimeter_Km: {formula: perimeter, unit: km,       round: 2}
    Length_m:     {formula: length,    unit: meters,   round: 2}

  connection_id:
    pattern: "SP_{plot_name}_{code}_{seq:02d}"

  block_no:                        # reads MMS text labels; assigns Block_No to nearest polygon
    primary_source:
      from_dwg:   [array_layout_P1a, array_layout_P1b]
      from_layer: "MMS Block Numbering"
      from_field: "Contents"
      method:     nearest

layers:
  - name:         "Block Boundaries"
    role:         spatial_reference   # enables spatial join + plot boundary dissolve
    locked:       false
    cli_aliases:  ["bb", "block"]
    source_dwg:   array_layout_P1a
    source_layer: "Boundary-12.5 MW Block"
    merge_sources:
      - source_dwg:   array_layout_P1a
        sub_plot:     "P1a"
        source_layer: "Boundary-12.5 MW Block"
      - source_dwg:   array_layout_P1b
        sub_plot:     "P1b"
        source_layer: "Boundary-12.5 MW Block"
    match_mode:   exact
    geometry:     polygon
    code:         "BL"
    output:       "block_boundaries.geojson"
    fields:
      Sub_Plot:   {from_merge_source: sub_plot}
      Block_No:   {spatial_join: primary}
      Area_Ha:    {calculate: Area_Ha}

  - name:         "Plot Boundaries"
    derive_from:  spatial_reference   # dissolved outer shell of Block Boundaries
    locked:       false
    ...
```

See [`samples/sample_config.yaml`](samples/sample_config.yaml) for a full working example covering polygon, point, and line layers.

---

## Locking layers

Locked layers are skipped during `--run all` — useful to protect layers that are already extracted while you work on others.

```
python extractio.py config.yaml --lock "Solar Tracker"
python extractio.py config.yaml --unlock "Solar Tracker"
python extractio.py config.yaml --unlock-all
```

---

## LLM YAML builder

Scans open DWGs and uses Groq or Gemini to generate a config from plain-English descriptions.

```
set GROQ_API_KEY=your_key_here
python extractio.py config.yaml --build
```

---

## DWG paths file

`dwg_paths.yaml` maps short aliases to absolute file paths. This file is gitignored — copy `samples/sample_dwg_paths.yaml` and fill in your paths.

```yaml
array_layout_P1a: "D:/Projects/SolarPark/CAD/P1/Array Layout/P1a.dwg"
array_layout_P1b: "D:/Projects/SolarPark/CAD/P1/Array Layout/P1b.dwg"
```

---

## Files

| File | Purpose |
|------|---------|
| `extractio.py` | Main extraction engine |
| `launcher.py` | Convenience wrapper (`python launcher.py ...`) |
| `launcher.bat` | Windows shortcut — add folder to PATH for `extractio` command |
| `yaml_generator.py` | Interactive wizard to generate a config from the sample template |
| `diagnose.py` | Standalone DWG layer inspector — run before configuring a new layer |
| `verify.py` | Compares extracted GeoJSON against live AutoCAD data |
| `mock_test.py` | Replays extraction logic without AutoCAD (for debugging) |
| `samples/sample_config.yaml` | Full annotated config example |
| `samples/sample_dwg_paths.yaml` | DWG paths file template |

---

## License

MIT
