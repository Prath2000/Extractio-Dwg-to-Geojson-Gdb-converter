# Extractio — Complete Beginner Guide
### From zero to your first GeoJSON in one sitting

**Author:** Prathamesh Athavale  
**Package:** `extractio` on PyPI  
**What it does:** Reads geometry and data from AutoCAD DWG files and converts them to GeoJSON or File Geodatabase (GDB) — no manual export, no clicking, no code.

---

## Table of Contents

1. [What You Need Before You Start](#1-what-you-need-before-you-start)
2. [Install Python](#2-install-python)
3. [Install Extractio](#3-install-extractio)
4. [Understand the Three Files You Will Work With](#4-understand-the-three-files-you-will-work-with)
5. [Set Up Your Project Folder](#5-set-up-your-project-folder)
6. [Get the Sample Config Files](#6-get-the-sample-config-files)
7. [Edit dwg_paths.yaml — Tell Extractio Where Your DWGs Are](#7-edit-dwg_pathsyaml--tell-extractio-where-your-dwgs-are)
8. [Edit config.yaml — Define What to Extract](#8-edit-configyaml--define-what-to-extract)
9. [Open Your DWGs in AutoCAD](#9-open-your-dwgs-in-autocad)
10. [Run Extractio](#10-run-extractio)
11. [Read the Output](#11-read-the-output)
12. [Optional — Output as File Geodatabase (GDB)](#12-optional--output-as-file-geodatabase-gdb)
13. [Common Errors and Fixes](#13-common-errors-and-fixes)
14. [Quick Reference — Commands You Will Use Every Day](#14-quick-reference--commands-you-will-use-every-day)

---

## 1. What You Need Before You Start

Before installing anything, make sure you have:

| Requirement | Details |
|---|---|
| **Windows PC** | Windows 10 or 11 (extractio only works on Windows — it talks to AutoCAD directly) |
| **AutoCAD** | Any version that supports COM (AutoCAD 2010 and later all do) |
| **Your DWG files** | The files you want to convert. They must be saved on your computer |
| **Python 3.9 or later** | Free — instructions below |

You do **not** need QGIS, ArcGIS, or any other GIS software installed.

---

## 2. Install Python

If you already have Python 3.9+, skip this section. To check, open a terminal (press `Win + R`, type `cmd`, press Enter) and run:

```
python --version
```

If it prints something like `Python 3.11.4`, you are good.

If you get an error or the version is below 3.9:

1. Go to **https://www.python.org/downloads/**
2. Click the big yellow **Download Python 3.x.x** button
3. Run the installer
4. **IMPORTANT:** On the first screen of the installer, tick the box that says **"Add Python to PATH"** before clicking Install Now

After installing, close the terminal, open a new one, and run `python --version` again to confirm.

---

## 3. Install Extractio

Open a terminal (or PowerShell — press `Win + X`, choose "Terminal") and run:

```
pip install extractio
```

That's it. This installs the engine plus creates the `extractio` command you will use to run extractions.

To confirm it installed correctly:

```
extractio --version
```

Expected output:
```
extractio 0.1.1 — by Prathamesh Athavale
```

**Optional extras** — install these if you need them:

```
pip install extractio[shapely]    # needed for dissolved boundary layers
pip install extractio[gdb]        # needed for GDB output (requires GDAL 3.6+)
pip install extractio[llm]        # needed for AI-assisted config generation
pip install extractio[all]        # installs everything
```

---

## 4. Understand the Three Files You Will Work With

Extractio does not have a user interface. You control it entirely through two YAML text files that you create once per project. YAML is not code — it is just structured text with colons and indentation.

| File | What It Does | Do You Edit It? |
|---|---|---|
| `config.yaml` | Tells extractio which layers to extract, what fields to include, and how to calculate/join values | YES — once per project |
| `dwg_paths.yaml` | Tells extractio the full file path to each of your DWG files | YES — once per machine |
| `extractio.py` | The engine itself | NEVER — do not touch this |

Both YAML files live in your project folder alongside your outputs.

---

## 5. Set Up Your Project Folder

Create a folder anywhere on your computer for this project. For example:

```
D:\Projects\MySiteExtraction\
```

Inside that folder, create an `outputs` subfolder:

```
D:\Projects\MySiteExtraction\
└── outputs\
```

All your GeoJSON files will appear in `outputs\` after each run.

---

## 6. Get the Sample Config Files

Extractio ships with two sample files you copy and edit. To find where pip installed them, run:

```
pip show -f extractio
```

Look for lines ending in `sample_config.yaml` and `sample_dwg_paths.yaml` — those are your templates.

Alternatively, download them directly from GitHub:

- **sample_config.yaml** — https://github.com/Prath2000/Extractio-Dwg-to-Geojson-Gdb-converter/blob/master/samples/sample_config.yaml
- **sample_dwg_paths.yaml** — https://github.com/Prath2000/Extractio-Dwg-to-Geojson-Gdb-converter/blob/master/samples/sample_dwg_paths.yaml

Copy both files into your project folder and rename them:

```
D:\Projects\MySiteExtraction\
├── config.yaml           ← renamed from sample_config.yaml
├── dwg_paths.yaml        ← renamed from sample_dwg_paths.yaml
└── outputs\
```

---

## 7. Edit dwg_paths.yaml — Tell Extractio Where Your DWGs Are

Open `dwg_paths.yaml` in any text editor (Notepad works, VS Code is better).

The file looks like this:

```yaml
drawing_a: "D:/Projects/MyProject/CAD/Drawing-A.dwg"
drawing_b: "D:/Projects/MyProject/CAD/Drawing-B.dwg"
drawing_c: "D:/Projects/MyProject/CAD/Drawing-C.dwg"
```

**What to do:**

1. Replace `drawing_a`, `drawing_b`, etc. with short nickname names for your DWGs — use lowercase, no spaces (e.g. `site_plan`, `zone_layout`, `services`)
2. Replace the path strings with the actual full path to each DWG on your machine
3. Use forward slashes `/` in paths, not backslashes `\`

**Example for a real project:**

```yaml
site_plan:    "D:/Projects/Highway/CAD/Site-Plan-Rev3.dwg"
zone_layout:  "D:/Projects/Highway/CAD/Zone-Layout-Rev2.dwg"
```

**Important:** The short names you use here (like `site_plan`) must match exactly what you write in `config.yaml`. Think of them as labels.

If you have only one DWG, just have one line:

```yaml
my_drawing: "D:/path/to/MyDrawing.dwg"
```

---

## 8. Edit config.yaml — Define What to Extract

This is the main file. Open it in a text editor. The file has two main sections: `global` and `layers`.

### 8a. The global section

```yaml
global:
  project_name: MyProject        # a name for your project — used in file naming
  crs: "EPSG:32643"              # coordinate system of your DWG (UTM zone for your area)
  output_dir: "./outputs"        # where GeoJSON files are saved (relative to config.yaml)
  dwg_paths_file: dwg_paths.yaml # keep this line exactly as-is
  source_dwgs:
    - site_plan                  # list the DWG nicknames you defined in dwg_paths.yaml
    - zone_layout
```

**Finding your EPSG code:**  
Go to https://epsg.io/ and search for your location. For most of India, UTM Zone 43N is `EPSG:32643`, Zone 44N is `EPSG:32644`, Zone 45N is `EPSG:32645`.

### 8b. The layers section

Each entry under `layers:` produces one output GeoJSON file. Here is the minimum you need for a simple polygon layer:

```yaml
layers:

  - name:         "Site Boundary"       # name for this layer — appears in the output file
    locked:       false                 # false = will run; true = will be skipped
    source_dwg:   site_plan             # which DWG (must match a nickname from dwg_paths.yaml)
    source_layer: "SITE-BOUNDARY"       # exact name of the CAD layer inside the DWG
    match_mode:   exact                 # exact = layer name must match exactly
    geometry:     polygon               # polygon, point, or line
    code:         "SB"                  # short code used in auto-ID generation
    output:       "site_boundary.geojson"
    fields:
      Category:   "Site Boundary"       # a fixed text value stamped on every feature
      Area_Ha:    {calculate: Area_Ha}  # auto-calculated from geometry
      Notes:      null                  # null = empty field, filled later
```

**How to find the exact CAD layer name:**

You need to know what the layer is called inside AutoCAD. Run this command (AutoCAD must be open with the DWG open):

```
extractio config.yaml --dwg-layers site_plan
```

This prints every layer name in that DWG. Copy the exact name (including capitalisation and hyphens) into `source_layer:`.

### 8c. Geometry types

| If your CAD layer contains... | Use geometry: |
|---|---|
| Closed polylines, hatches, polygonal areas | `polygon` |
| Block inserts, points, symbols | `point` |
| Lines, open polylines, routes, cables | `line` |

### 8d. A minimal working config.yaml

```yaml
global:
  project_name: MyProject
  crs: "EPSG:32643"
  output_dir: "./outputs"
  dwg_paths_file: dwg_paths.yaml
  source_dwgs:
    - my_drawing

layers:

  - name:         "Roads"
    locked:       false
    source_dwg:   my_drawing
    source_layer: "ROAD-CENTRE"
    match_mode:   exact
    geometry:     line
    code:         "RD"
    output:       "roads.geojson"
    fields:
      Category:   "Road"
      Length_m:   {calculate: Length_m}
      Notes:      null
```

---

## 9. Open Your DWGs in AutoCAD

Extractio connects to AutoCAD live while it is running. Before you run extractio:

1. **Open AutoCAD**
2. **Open every DWG file** listed in your `dwg_paths.yaml` inside AutoCAD
3. Make sure the drawings are in **Model Space** (not a Paper Space layout)
4. Do not close AutoCAD while extractio is running

If a DWG is not open in AutoCAD when you run extractio, that layer will be skipped with a warning.

---

## 10. Run Extractio

Open a terminal. Navigate to your project folder:

```
cd "D:\Projects\MySiteExtraction"
```

### See all your layers and their status:

```
extractio config.yaml --list
```

This prints every layer defined in your config, whether it is locked or unlocked, and its source layer name.

### Run all unlocked layers:

```
extractio config.yaml --run all
```

### Run just one specific layer:

```
extractio config.yaml --layers "Roads"
```

### Run the interactive selector (choose layers from a menu):

```
extractio config.yaml
```

This shows a numbered menu — type layer numbers to toggle them on/off, then press Enter to run.

---

## 11. Read the Output

After running, extractio prints a report table:

```
  Layer                                    Features  Status
  ──────────────────────────────────────────────────────────
  Roads                                          87  ✓  OK
  Site Boundary                                   1  ✓  OK
```

Your GeoJSON files are in the `outputs\` folder:

```
D:\Projects\MySiteExtraction\
├── outputs\
│   ├── roads.geojson
│   └── site_boundary.geojson
```

You can open these in:
- **QGIS** — drag and drop directly onto the canvas
- **ArcGIS Pro** — Add Data → browse to the .geojson file
- **Any web GIS tool** — most accept GeoJSON natively

---

## 12. Optional — Output as File Geodatabase (GDB)

If you need a File Geodatabase instead of individual GeoJSON files:

**Step 1 — Install fiona:**

```
pip install fiona
```

**Step 2 — Run with GDB output:**

```
extractio config.yaml --run all --output-format gdb
```

This creates a single `MyProject.gdb` folder in your `outputs\` directory with one feature class per layer.

Or set it permanently in `config.yaml`:

```yaml
global:
  output_format: gdb
  gdb_name: MySite.gdb
```

---

## 13. Common Errors and Fixes

### "No AutoCAD application found"
AutoCAD is not running, or no DWG is open. Start AutoCAD, open your DWG files, then run extractio again.

### "Layer 'ROAD-CENTRE' not found in DWG"
The layer name in `source_layer:` does not match what is in the DWG. Run `--dwg-layers` to get the exact name:
```
extractio config.yaml --dwg-layers my_drawing
```

Then copy the correct name into your config.

### "0 features extracted"
The layer exists in AutoCAD but contains no entities, or the entities are on a different layer than expected. Check in AutoCAD that the layer has visible geometry in Model Space.

### "dwg_paths.yaml not found"
Make sure `dwg_paths.yaml` is in the same folder as `config.yaml`, and that `dwg_paths_file: dwg_paths.yaml` is in your config's global section.

### "pyyaml.scanner.ScannerError"
There is a formatting error in your YAML file. Common causes:
- Used a tab instead of spaces for indentation (YAML requires spaces)
- Missing colon after a key
- Quotes not closed

Open the file in VS Code — it will highlight the error line in red.

### "fiona not installed"
You used `--output-format gdb` but fiona is not installed. Run `pip install fiona`.

---

## 14. Quick Reference — Commands You Will Use Every Day

```bash
# Check extractio version
extractio --version

# List all layers in your config
extractio config.yaml --list

# Show all CAD layer names inside a DWG
extractio config.yaml --dwg-layers my_drawing

# Run all unlocked layers
extractio config.yaml --run all

# Run specific layers by name
extractio config.yaml --layers "Roads" "Site Boundary"

# Unlock a layer (so it runs)
extractio config.yaml --unlock "Roads"

# Lock a layer (so it is skipped)
extractio config.yaml --lock "Roads"

# Run with GDB output
extractio config.yaml --run all --output-format gdb

# Validate your config without running
extractio config.yaml --validate
```

---

*Built with extractio by Prathamesh Athavale*
