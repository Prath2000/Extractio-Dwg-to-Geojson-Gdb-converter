#!/usr/bin/env python3
"""
extractio — YAML Config Generator  (template-based)
====================================================
Generates a new project config YAML by substituting your project values
into the sample template.

Usage:
    python yaml_generator.py
    python yaml_generator.py --output my_project.yaml
    python yaml_generator.py --template path/to/other_template.yaml
"""

import sys
import re
import os
import argparse

# ─── ANSI colours ─────────────────────────────────────────────────────────────

if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleMode(
            ctypes.windll.kernel32.GetStdHandle(-11), 7)
    except Exception:
        pass

R  = "\033[0m"
B  = "\033[1m"
D  = "\033[2m"
CY = "\033[36m"
YL = "\033[33m"
GN = "\033[32m"
RD = "\033[31m"

# ─── UI helpers ───────────────────────────────────────────────────────────────

def banner():
    w = 56
    print(f"\n  {B}{CY}╭{'─'*w}╮{R}")
    print(f"  {B}{CY}│{'extractio · YAML Config Generator':^{w}}│{R}")
    print(f"  {B}{CY}╰{'─'*w}╯{R}")
    print(f"\n  {D}Based on samples/sample_config.yaml — fills in your project data.{R}")
    print(f"  {D}Press Enter to accept defaults [in brackets].  Ctrl+C to cancel.{R}\n")

def section(title, step=None, total=None):
    step_str = f"  {D}({step}/{total}){R}" if step else ""
    print(f"\n  {B}{CY}▸  {title}{R}{step_str}")
    print(f"  {CY}{'─'*52}{R}")

def hint(msg):
    print(f"  {D}  {msg}{R}")

def ok(msg):
    print(f"\n  {GN}{B}✓  {msg}{R}")

def err(msg):
    print(f"  {RD}  {msg}{R}")

def q(label, default=None, allow_blank=False):
    dflt   = f" {D}[{default}]{R}" if default is not None else ""
    prompt = f"  {B}{label}{R}{dflt}  "
    while True:
        val = input(prompt).strip()
        if val:
            return val
        if default is not None:
            return str(default)
        if allow_blank:
            return ""
        err("Required")

def qbool(label, default=True):
    yn     = f"{GN}Y{R}/{D}n{R}" if default else f"{D}y{R}/{GN}N{R}"
    prompt = f"  {B}{label}{R}  {yn}  "
    while True:
        val = input(prompt).strip().lower()
        if not val:           return default
        if val in ("y","yes","1"): return True
        if val in ("n","no","0"):  return False
        err("Enter y or n")

def qchoice(label, choices, descriptions=None, default_idx=0):
    print(f"\n  {B}{label}{R}")
    for i, c in enumerate(choices, 1):
        is_def = (i - 1 == default_idx)
        bullet = f"{CY}›{R}" if is_def else " "
        name   = f"{B}{c}{R}" if is_def else c
        desc   = f"  {D}{descriptions[i-1]}{R}" if descriptions else ""
        print(f"    {bullet} {CY}{i:2}{R}  {name}{desc}")
    prompt = f"\n  {CY}→{R} {D}[{default_idx+1}]{R}  "
    while True:
        val = input(prompt).strip()
        if not val:
            return choices[default_idx]
        if val.isdigit() and 1 <= int(val) <= len(choices):
            return choices[int(val) - 1]
        err(f"Enter 1–{len(choices)}")

# ─── Template defaults (edit here or override at the prompts) ─────────────────

_HERE         = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(_HERE, "samples", "sample_config.yaml")

TPL_PREFIX      = "SP"
TPL_PLANT       = "My Solar Plant"
TPL_COUNTRY     = "India"
TPL_STATE       = "Rajasthan"
TPL_DISTRICT    = "Jaisalmer"
TPL_TALUKA      = ""
TPL_VILLAGE     = ""
TPL_JURISDICTION= "Within Project Limits"
TPL_PLOT        = "P1"
TPL_SUB_A       = "P1a"
TPL_SUB_B       = "P1b"
TPL_HT_END      = "GIS Sub-Station"
TPL_CONN_PAT    = "SP_{plot_name}_{code}_{seq:02d}"
TPL_CRS         = "EPSG:32643"

# DWG role aliases in the sample template
TPL_DWG_ROLES = [
    ("array_layout_P1a", "Array Layout  ·  Sub-plot A"),
    ("array_layout_P1b", "Array Layout  ·  Sub-plot B"),
    ("lt_cable_P1",      "LT Cable / Master working sheet"),
    ("ht_fo_P1a",        "HT & FO Cable  ·  Sub-plot A"),
    ("ht_fo_P1b",        "HT & FO Cable  ·  Sub-plot B"),
    ("pile_layout_P1",   "Pile Layout  (null = not yet received)"),
]

# Layer names in the sample template
TPL_LAYERS = [
    "Block Boundaries",
    "Plot Boundaries",
    "Solar Tracker",
    "HT Cable Trench",
    "LT Cable Trench",
]

# ─── Template manipulation ─────────────────────────────────────────────────────

def apply_subs(text, subs):
    for old, new in sorted(subs.items(), key=lambda x: -len(x[0])):
        if old and old != new:
            text = text.replace(old, new)
    return text

def regex_replace_line(text, key, new_value):
    return re.sub(
        rf'^(  {re.escape(key)}:[ ]+).*$',
        rf'\g<1>{new_value}',
        text,
        flags=re.MULTILINE,
    )

def replace_dwg_path(text, alias, new_path):
    val = f'"{new_path}"' if new_path and new_path.lower() != "null" else "null"
    return re.sub(
        rf'^(    {re.escape(alias)}:[ ]+).*$',
        rf'\g<1>{val}',
        text,
        flags=re.MULTILINE,
    )

def remove_layers(text, names_to_remove):
    if not names_to_remove:
        return text
    lines    = text.split('\n')
    result   = []
    skipping = False
    for line in lines:
        if re.match(r'^  - name:', line):
            m    = re.match(r'^  - name:\s+["\']?(.+?)["\']?\s*$', line)
            name = m.group(1).strip(' "\'') if m else ""
            skipping = name in names_to_remove
        if not skipping:
            result.append(line)
    return '\n'.join(result)

# ─── Generator ────────────────────────────────────────────────────────────────

def generate(template_path=None):
    tpl = template_path or TEMPLATE_PATH
    banner()

    # ── 1 · Project ───────────────────────────────────────────────────────────
    section("PROJECT", 1, 4)
    proj_name  = q("Project name",   default="My Project")
    hint(f"replaces  '{TPL_PREFIX}_'  in all Code values")
    prefix     = q("Code prefix",    default=TPL_PREFIX)
    output_dir = q("Output folder",  default="./output")
    hint("EPSG code for your project area")
    crs        = q("CRS",            default=TPL_CRS)

    # ── 2 · Plot ──────────────────────────────────────────────────────────────
    section("PLOT", 2, 4)
    hint(f"replaces  '{TPL_PLOT}'  in alias names and sub-plot labels")
    plot_id = q("Plot ID",          default=TPL_PLOT)
    hint("label stamped on features from each sub-plot")
    sub_a   = q("Sub-plot A name",  default=f"{plot_id}a")
    sub_b   = q("Sub-plot B name",  default=f"{plot_id}b")

    print()
    hint("location fields stamped on every feature")
    plant_name   = q("Plant_Name",   default=TPL_PLANT)
    country      = q("Country",      default=TPL_COUNTRY)
    state        = q("State",        default=TPL_STATE)
    district     = q("District",     default=TPL_DISTRICT)
    taluka       = q("Taluka",       default=TPL_TALUKA, allow_blank=True)
    village      = q("Village",      default=TPL_VILLAGE, allow_blank=True)
    jurisdiction = q("Jurisdiction", default=TPL_JURISDICTION)

    hint(f"vars: {{plot_name}}  {{code}}  {{seq:02d}}")
    default_pattern = f"{prefix}_{plot_id}_{{code}}_{{seq:02d}}"
    conn_pattern = q("connection_id pattern", default=default_pattern)

    ht_end = q("HT cable end connection", default=TPL_HT_END)

    # ── 3 · Source DWGs ───────────────────────────────────────────────────────
    section("SOURCE DWGs", 3, 4)
    hint("alias names are auto-updated  ·  blank path → null (not yet received)")

    alias_map  = {}
    path_map   = {}

    for old_alias, role in TPL_DWG_ROLES:
        new_alias = (old_alias
                     .replace(TPL_SUB_A, sub_a)
                     .replace(TPL_SUB_B, sub_b)
                     .replace(TPL_PLOT,  plot_id))
        alias_map[old_alias] = new_alias
        print(f"\n  {D}{role}{R}")
        print(f"  {D}  alias  →  {CY}{new_alias}{R}")
        path = q("Path", allow_blank=True)
        path_map[new_alias] = path

    # ── 4 · Layer schema ──────────────────────────────────────────────────────
    section("LAYER SCHEMA", 4, 4)
    hint("the sample template has layers — you can remove any you don't need")

    layers_to_remove = set()

    if not qbool("Keep all layers?", default=True):
        print(f"\n  {B}Template layers:{R}")
        for i, name in enumerate(TPL_LAYERS, 1):
            print(f"    {CY}{i:2}{R}  {name}")
        print()
        hint("enter layer numbers to remove, comma-separated  e.g.  3,5")
        raw = q("Remove layers", allow_blank=True)
        if raw:
            for tok in raw.split(","):
                tok = tok.strip()
                if tok.isdigit() and 1 <= int(tok) <= len(TPL_LAYERS):
                    layers_to_remove.add(TPL_LAYERS[int(tok) - 1])
        if layers_to_remove:
            print(f"\n  {RD}Will remove:{R}")
            for name in sorted(layers_to_remove, key=TPL_LAYERS.index):
                print(f"    {RD}✕{R}  {name}")

    # ── Apply everything to template ──────────────────────────────────────────
    if not os.path.exists(tpl):
        err(f"Template not found: {tpl}")
        sys.exit(1)

    with open(tpl, encoding="utf-8") as f:
        text = f.read()

    # Header
    text = re.sub(r'^# extractio.*$', f'# {proj_name} — extractio Config', text, flags=re.MULTILINE)
    text = re.sub(r'^# Copy this.*$', '# Generated by yaml_generator.py', text, flags=re.MULTILINE)

    # Global scalar lines
    text = regex_replace_line(text, "output_dir", f'"{output_dir}"')
    text = regex_replace_line(text, "crs",        f'"{crs}"')
    text = regex_replace_line(text, "project_name", proj_name)

    # connection_id pattern
    text = text.replace(f'"{TPL_CONN_PAT}"', f'"{conn_pattern}"')

    # HT end connection
    if ht_end != TPL_HT_END:
        text = text.replace(f'"{TPL_HT_END}"', f'"{ht_end}"')

    # Location field values
    loc_subs = {}
    if plant_name   != TPL_PLANT:        loc_subs[f'"{TPL_PLANT}"']        = f'"{plant_name}"'
    if country      != TPL_COUNTRY:      loc_subs[f'"{TPL_COUNTRY}"']      = f'"{country}"'
    if state        != TPL_STATE:        loc_subs[f'"{TPL_STATE}"']        = f'"{state}"'
    if district     != TPL_DISTRICT:     loc_subs[f'"{TPL_DISTRICT}"']     = f'"{district}"'
    if taluka       and taluka != TPL_TALUKA:
        loc_subs[f'"{TPL_TALUKA}"']       = f'"{taluka}"'
    if village      and village != TPL_VILLAGE:
        loc_subs[f'"{TPL_VILLAGE}"']      = f'"{village}"'
    if jurisdiction != TPL_JURISDICTION: loc_subs[f'"{TPL_JURISDICTION}"'] = f'"{jurisdiction}"'
    text = apply_subs(text, loc_subs)

    # Code prefix
    if prefix != TPL_PREFIX:
        text = text.replace(f'"{TPL_PREFIX}_', f'"{prefix}_')
        text = re.sub(rf'(code:\s+){TPL_PREFIX}_', rf'\g<1>{prefix}_', text)

    # DWG alias names
    alias_subs = {old: new for old, new in alias_map.items() if old != new}
    text = apply_subs(text, alias_subs)

    # Sub-plot labels
    plot_subs = {}
    if sub_a != TPL_SUB_A: plot_subs[f'"{TPL_SUB_A}"'] = f'"{sub_a}"'
    if sub_b != TPL_SUB_B: plot_subs[f'"{TPL_SUB_B}"'] = f'"{sub_b}"'
    text = apply_subs(text, plot_subs)

    # DWG paths
    for new_alias, path in path_map.items():
        text = replace_dwg_path(text, new_alias, path)

    # Remove unwanted layers
    text = remove_layers(text, layers_to_remove)

    # Clean up extra blank lines
    text = re.sub(r'\n{3,}', '\n\n', text)

    return text

# ─── Entry point ──────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="extractio YAML Config Generator")
    parser.add_argument("--output", "-o", default=None,
                        help="Output YAML file path  (default: prompted)")
    parser.add_argument("--template", "-t", default=None,
                        help="Template YAML to use  (default: samples/sample_config.yaml)")
    args = parser.parse_args()

    try:
        yaml_content = generate(template_path=args.template)
    except KeyboardInterrupt:
        print("\n\n  Cancelled.")
        sys.exit(0)

    print()
    out_path = args.output or q("Save config to", default="config.yaml")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(yaml_content)

    ok(f"Saved  →  {out_path}")
    print(f"  {D}Verify:{R}  python extractio.py {YL}{out_path}{R} --list\n")
    print(f"  {D}Remember to create dwg_paths.yaml with your actual DWG file paths.{R}\n")


if __name__ == "__main__":
    main()
