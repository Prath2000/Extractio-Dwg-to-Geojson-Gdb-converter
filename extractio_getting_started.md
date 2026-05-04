# Extractio — Complete Beginner Guide
### From zero to your first GeoJSON/GDB in one sitting

**Author:** Prathamesh Athavale  
**Package:** `extractio` on PyPI  
**What it does:** Reads geometry and data from AutoCAD DWG files and converts them to GeoJSON or File Geodatabase (GDB).

---

## Table of Contents

1. [What You Need Before You Start](#1-what-you-need-before-you-start)
2. [Install Python](#2-install-python)
3. [Install Extractio](#3-install-extractio)
4. [Two Ways to Set Up Your Config](#4-two-ways-to-set-up-your-config)
5. [Set Up Your Project Folder](#5-set-up-your-project-folder)
6. [Edit dwg_paths.yaml](#6-edit-dwg_pathsyaml)
7. [PATH A — LLM Config Generation (Recommended)](#7-path-a--llm-config-generation-recommended)
8. [PATH B — Write config.yaml Manually](#8-path-b--write-configyaml-manually)
9. [Open Your DWGs in AutoCAD](#9-open-your-dwgs-in-autocad)
10. [Run Extractio](#10-run-extractio)
11. [Read the Output](#11-read-the-output)
12. [Optional — GDB Output](#12-optional--gdb-output)
13. [Common Errors and Fixes](#13-common-errors-and-fixes)
14. [Quick Reference](#14-quick-reference)

---

## 1. What You Need Before You Start

| Requirement | Details |
|---|---|
| **Windows PC** | Windows 10 or 11 — extractio talks directly to AutoCAD via Windows COM |
| **AutoCAD** | Any version with COM support (2010 and later) |
| **Your DWG files** | The files you want to convert, saved on your machine |
| **Python 3.9+** | Free — instructions below |
| **Groq API key** | Free — only needed for Path A (AI config generation) |

---

## 2. Install Python

Open a terminal (`Win + R` type `cmd` press Enter) and run:

```
python --version
```

If you see `Python 3.9` or higher, skip to Section 3.

If not:
1. Go to **https://www.python.org/downloads/**
2. Click **Download Python**
3. Run the installer
4. **Tick "Add Python to PATH"** before clicking Install Now

Close and reopen the terminal, then confirm with `python --version`.

---

## 3. Install Extractio

```
pip install extractio
```

Confirm it worked:

```
extractio --version
```

Expected output:

```
extractio 0.1.1 — by Prathamesh Athavale
```

Install optional extras depending on what you need:

```
pip install extractio[llm]        <- AI config builder (install this for Path A)
pip install extractio[shapely]    <- dissolved boundary layers
pip install extractio[gdb]        <- GDB output
pip install extractio[all]        <- everything at once
```

---

## 4. Two Ways to Set Up Your Config

Extractio needs a `config.yaml` file that tells it which layers to extract and what fields to produce. There are two ways to create it:

| | Path A — LLM Builder (Recommended) | Path B — Manual |
|---|---|---|
| **Best for** | Everyone, especially beginners | Advanced users only |
| **How** | AI reads your spec sheets and writes the YAML for you | You write YAML by hand |
| **AutoCAD needed?** | No — completely independent of AutoCAD | No (but helpful for layer name lookup) |
| **API key needed?** | Yes — free Groq key | No |

**Use Path A.** It requires no YAML knowledge and the AI handles all the field naming, geometry types, and structure automatically. Even if you tweak the output afterwards, starting from AI-generated YAML is far faster than writing from scratch.

---

## 5. Set Up Your Project Folder

Create a folder for this project:

```
D:\Projects\MySiteExtraction\
```

Create an `outputs` subfolder inside it:

```
D:\Projects\MySiteExtraction\
    outputs\
```

---

## 6. Edit dwg_paths.yaml

Download the template from:
https://github.com/Prath2000/Extractio-Dwg-to-Geojson-Gdb-converter/blob/master/samples/sample_dwg_paths.yaml

Save it into your project folder as `dwg_paths.yaml`. Open it in any text editor and replace the example paths with your own:

```yaml
site_plan:   "D:/Projects/Highway/CAD/Site-Plan-Rev3.dwg"
zone_layout: "D:/Projects/Highway/CAD/Zone-Layout.dwg"
```

Rules:
- Short nicknames only — lowercase, no spaces
- Use forward slashes `/` in paths, not backslashes
- These nicknames are referenced later in config.yaml — they must match exactly

---

## 7. PATH A — LLM Config Generation (Recommended)

This is the fast, beginner-friendly path. The AI reads your project documents and builds the entire `config.yaml` for you. **AutoCAD does not need to be open during this step.**

### Step 1 — Get a free Groq API key

1. Go to **https://console.groq.com**
2. Sign up (free)
3. Go to API Keys and create a key — it starts with `gsk_`

### Step 2 — Create a .env file

In your project folder, create a plain text file named `.env` (no filename before the dot). Add this line:

```
GROQ_API_KEY=gsk_your_key_here
```

Save it. Extractio reads this automatically.

### Step 3 — Gather your reference documents

The AI learns from documents you provide. Good sources:

- **PDF spec sheets** — layer schedules, attribute tables, data dictionaries
- **Word documents** — project briefs, GIS data requirements
- **An existing config.yaml** from a similar project

Even a single PDF that lists layer names and what data each layer should carry is enough.

### Step 4 — (Optional but recommended) Scan your DWGs first

If AutoCAD is open with your DWGs loaded, run this before the builder:

```
extractio --scan-only
```

This saves a snapshot of all layer names from your DWGs into a `scan_*.json` file. The LLM builder finds and uses it automatically — so the AI will know your exact CAD layer names instead of guessing.

Skip this step if AutoCAD is not available yet.

### Step 5 — Run the builder

Open a terminal and go to your project folder:

```
cd "D:\Projects\MySiteExtraction"
```

Run:

```
extractio --build
```

Or pass your reference documents directly to skip the file prompt:

```
extractio --build --ref "D:/Docs/LayerSchedule.pdf" "D:/Docs/DataDictionary.docx"
```

The tool reads your documents and then presents each generated layer for your review:

```
  Layer 1/5 -- "Site Boundary"
  ------------------------------------------------
  geometry:     polygon
  source_layer: SITE-BOUNDARY
  fields:       Category, Area_Ha, Perimeter_Km, Notes

  [A]ccept  [E]dit  [S]kip  [R]egenerate  >
```

- **A** — accept this layer as-is
- **E** — open an editor to tweak the layer before accepting
- **S** — skip this layer (it will not be in the config)
- **R** — ask the AI to regenerate this layer

### Step 6 — Config is saved automatically

After reviewing all layers, `config.yaml` is saved to your project folder. Go to Section 9.

Shortcut to skip all review prompts:

```
extractio --build --accept-all
```

---

## 8. PATH B — Write config.yaml Manually

Only use this if you have no reference documents and need full manual control.

Download the template:
https://github.com/Prath2000/Extractio-Dwg-to-Geojson-Gdb-converter/blob/master/samples/sample_config.yaml

Save it as `config.yaml` in your project folder.

### Global section

```yaml
global:
  project_name: MyProject
  crs: "EPSG:32643"
  output_dir: "./outputs"
  dwg_paths_file: dwg_paths.yaml
  source_dwgs:
    - site_plan
    - zone_layout
```

EPSG codes for India: Zone 43N = EPSG:32643, Zone 44N = EPSG:32644, Zone 45N = EPSG:32645.
Find yours at https://epsg.io/

### Find exact CAD layer names

Open AutoCAD with the DWG loaded, then run:

```
extractio config.yaml --dwg-layers site_plan
```

This prints every layer name in that DWG. Copy the exact name (capitalisation matters) into `source_layer:`.

### A layer entry

```yaml
layers:

  - name:         "Roads"
    locked:       false
    source_dwg:   site_plan
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

Geometry types:

| geometry | Use when layer contains |
|---|---|
| polygon | Closed polylines, hatched areas, zones |
| point | Block inserts, symbols, markers |
| line | Lines, open polylines, routes, cables |

---

## 9. Open Your DWGs in AutoCAD

Before running extraction:

1. Open AutoCAD
2. Open every DWG listed in `dwg_paths.yaml`
3. Make sure they are in **Model Space** (not Paper Space)
4. Do not close AutoCAD while extractio runs

---

## 10. Run Extractio

Navigate to your project folder:

```
cd "D:\Projects\MySiteExtraction"
```

List all layers and their lock status:

```
extractio config.yaml --list
```

Run all unlocked layers:

```
extractio config.yaml --run all
```

Run one specific layer:

```
extractio config.yaml --layers "Roads"
```

Interactive menu:

```
extractio config.yaml
```

---

## 11. Read the Output

Extractio prints a summary when done:

```
  Layer                            Features  Status
  Roads                                  87  OK
  Site Boundary                           1  OK
```

Your GeoJSON files are in the `outputs\` folder. Open them in:
- **QGIS** — drag and drop onto the canvas
- **ArcGIS Pro** — Add Data and browse to the .geojson file
- Any web GIS tool — GeoJSON is universally supported

---

## 12. Optional — GDB Output

Install fiona:

```
pip install fiona
```

Run with GDB output:

```
extractio config.yaml --run all --output-format gdb
```

All layers go into one `.gdb` folder in `outputs\`. Or set it permanently in `config.yaml`:

```yaml
global:
  output_format: gdb
  gdb_name: MySite.gdb
```

---

## 13. Common Errors and Fixes

**"No AutoCAD application found"**
AutoCAD is not running. Open it, load your DWGs, try again.

**"Layer not found in DWG"**
Layer name mismatch. Run `--dwg-layers` to get the exact name from the DWG.

**"0 features extracted"**
Layer exists but has no entities in Model Space. Check in AutoCAD.

**"dwg_paths.yaml not found"**
Put it in the same folder as `config.yaml`.

**YAML formatting error**
Open the file in VS Code — errors are highlighted red. Common causes: tabs instead of spaces, missing colon, unclosed quote.

**"GROQ_API_KEY not set"**
Create a `.env` file in your project folder with:
```
GROQ_API_KEY=gsk_your_key_here
```

**"fiona not installed"**
Run `pip install fiona` before using `--output-format gdb`.

---

## 14. Quick Reference

```
# Check version
extractio --version

# GENERATE config.yaml with AI — no AutoCAD needed
extractio --build
extractio --build --ref "Spec.pdf" "DataDict.docx"
extractio --build --accept-all

# Scan DWGs to give AI real layer names (AutoCAD must be open)
extractio --scan-only

# Inspect your config
extractio config.yaml --list
extractio config.yaml --validate
extractio config.yaml --dwg-layers my_drawing

# Run extraction (AutoCAD must be open)
extractio config.yaml --run all
extractio config.yaml --layers "Roads" "Site Boundary"
extractio config.yaml --run all --output-format gdb

# Lock / unlock layers
extractio config.yaml --unlock "Roads"
extractio config.yaml --lock "Roads"
```

---

*Built with extractio by Prathamesh Athavale*
