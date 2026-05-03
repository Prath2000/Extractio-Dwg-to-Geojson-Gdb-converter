"""
GSA Mock Test Framework
=======================
1. Run gsa_diagnose.py on a layer, save output to a .txt file
2. Run: python gsa_mock_test.py diagnose_output.txt
3. It replays the extraction logic and shows exactly what GeoJSON will be produced
4. Fix logic here first, then apply to executor

Usage:
    python gsa_mock_test.py <diagnose_output.txt> <geometry_type> [options]

Examples:
    python gsa_mock_test.py mms_block.txt point_expand_attrs
    python gsa_mock_test.py idt_layer.txt block_explode_enclose
    python gsa_mock_test.py lt_cable.txt line
"""
import sys, json, math, re, ast

# ── MOCK ENTITY ───────────────────────────────────────────────────────────────
class MockAttrib:
    def __init__(self, tag, val, insert_pt, align_pt=None):
        self.TagString         = tag
        self.TextString        = val
        self._insert           = insert_pt
        self._align            = align_pt or (0.0, 0.0)
    @property
    def InsertionPoint(self):  return self._insert
    @property
    def TextAlignmentPoint(self): return self._align

class MockInsert:
    def __init__(self, name, ip, rotation=0.0, sx=1.0, sy=1.0, attrs=None):
        self.Name            = name
        self._ip             = ip
        self.Rotation        = rotation
        self.XScaleFactor    = sx
        self.YScaleFactor    = sy
        self._attrs          = attrs or []
        self.HasAttributes   = len(self._attrs) > 0
    @property
    def InsertionPoint(self): return self._ip
    def GetAttributes(self):  return self._attrs

class MockLWPoly:
    def __init__(self, coords, bulges=None, closed=False):
        flat = []
        for x, y in coords:
            flat += [x, y]
        self._coords  = flat
        self._bulges  = bulges or [0.0]*len(coords)
        self.Closed   = closed
    @property
    def Coordinates(self): return self._coords
    def GetBulges(self):   return self._bulges

# ── PARSE DIAGNOSE OUTPUT ─────────────────────────────────────────────────────
def parse_diagnose_output(path):
    """Parse gsa_diagnose.py output into mock entities."""
    with open(path) as f:
        lines = f.readlines()

    entities = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if line.startswith('--- First INSERT'):
            # Parse INSERT block
            props = {}
            attrs = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('---'):
                l = lines[i].strip()
                if l.startswith('Block Name'):
                    props['name'] = l.split(':',1)[1].strip()
                elif l.startswith('InsertPt'):
                    m = re.search(r'\(([0-9.\-]+),\s*([0-9.\-]+)\)', l)
                    if m: props['ip'] = (float(m.group(1)), float(m.group(2)))
                elif l.startswith('Rotation'):
                    m = re.search(r'([\d.\-]+)\s*rad', l)
                    if m: props['rot'] = float(m.group(1))
                elif l.startswith('Scale'):
                    m = re.findall(r'[\d.]+', l)
                    if len(m) >= 2:
                        props['sx'] = float(m[0])
                        props['sy'] = float(m[1])
                elif re.match(r'^[A-Z0-9_]+ *=', l):
                    # Attribute line: TAG = 'val'  InsertPt=(x,y)  AlignPt=(x,y)
                    tag_m = re.match(r'^([A-Z0-9_]+)\s*=\s*\'([^\']*)', l)
                    if tag_m:
                        tag = tag_m.group(1)
                        val = tag_m.group(2)
                        ipt = re.search(r'InsertPt=\(([0-9.\-]+),([0-9.\-]+)\)', l)
                        apt = re.search(r'AlignPt=\(([0-9.\-]+),([0-9.\-]+)\)', l)
                        ip2 = (float(ipt.group(1)), float(ipt.group(2))) if ipt else (0,0)
                        ap2 = (float(apt.group(1)), float(apt.group(2))) if apt else None
                        attrs.append(MockAttrib(tag, val, ip2, ap2))
                i += 1
            if 'name' in props and 'ip' in props:
                ent = MockInsert(
                    props['name'], props['ip'],
                    props.get('rot', 0.0),
                    props.get('sx', 1.0), props.get('sy', 1.0),
                    attrs)
                entities.append(('INSERT', ent))

        elif line.startswith('--- First LWPOLYLINE'):
            # Minimal — just store placeholder
            i += 1
        else:
            i += 1

    return entities

# ── GEOMETRY HANDLERS ─────────────────────────────────────────────────────────
def run_point_expand_attrs(entities, expand_attrs, string_field,
                           use_align_pt=True):
    """Test point_expand_attrs extraction."""
    results = []
    for et, ent in entities:
        if et != 'INSERT':
            continue
        tag_vals = {}
        tag_wpts = {}
        if ent.HasAttributes:
            for att in ent.GetAttributes():
                tag = att.TagString.upper().strip()
                val = (att.TextString or '').strip()
                tag_vals[tag] = val
                # Try TextAlignmentPoint first
                wpt = None
                if use_align_pt:
                    try:
                        tap = att.TextAlignmentPoint
                        if tap and (tap[0] != 0 or tap[1] != 0):
                            wpt = (tap[0], tap[1])
                    except: pass
                if not wpt:
                    try:
                        ip = att.InsertionPoint
                        wpt = (ip[0], ip[1])
                    except: pass
                if wpt:
                    tag_wpts[tag] = wpt

        for tag in expand_attrs:
            utag = tag.upper()
            val  = tag_vals.get(utag, '').strip()
            if not val:
                continue
            wpt = tag_wpts.get(utag)
            if not wpt:
                print(f"  WARNING: no position found for {utag}")
                continue
            results.append({
                'tag':   utag,
                'val':   val,
                'point': wpt,
                'all_attrs': dict(tag_vals)
            })

    return results

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        print(__doc__)
        # Run with hardcoded test data if no args
        print("\nRunning with hardcoded test data (MMS Table Block example)...")
        # Simulate one INSERT with 4 ATTRIBs at known positions
        attrs = [
            MockAttrib('ID1', 'B17-R03-T27-I19-S21',
                       insert_pt=(539750.0, 2664308.0),
                       align_pt=(539750.5, 2664316.0)),
            MockAttrib('ID2', 'B17-R03-T27-I19-S22',
                       insert_pt=(539750.0, 2664340.0),
                       align_pt=(539750.5, 2664348.0)),
            MockAttrib('ID3', 'B17-R03-T27-I19-S23',
                       insert_pt=(539750.0, 2664372.0),
                       align_pt=(539750.5, 2664380.0)),
            MockAttrib('ID4', 'B17-R03-T27-I19-S24',
                       insert_pt=(539750.0, 2664404.0),
                       align_pt=(539750.5, 2664412.0)),
        ]
        ent = MockInsert('MMS Table', (539749.0, 2664292.0), 0.0, 1.0, 1.0, attrs)
        entities = [('INSERT', ent)]
    else:
        print(f"Parsing diagnose output: {sys.argv[1]}")
        entities = parse_diagnose_output(sys.argv[1])
        print(f"Parsed {len(entities)} entities")

    geom_type = sys.argv[2] if len(sys.argv) > 2 else 'point_expand_attrs'

    if geom_type == 'point_expand_attrs':
        print(f"\nTesting point_expand_attrs extraction...")
        print(f"  expand_attrs: ID1-ID4")
        print(f"  Using TextAlignmentPoint (falls back to InsertionPoint)")

        results = run_point_expand_attrs(
            entities,
            expand_attrs=['ID1','ID2','ID3','ID4'],
            string_field='Connection_String',
            use_align_pt=True
        )

        print(f"\nResults: {len(results)} point features")
        for r in results:
            print(f"  {r['tag']:4s} | {r['val']:30s} | ({r['point'][0]:.2f}, {r['point'][1]:.2f})")

        if results:
            # Check Y spacing
            ys = [r['point'][1] for r in results]
            if len(ys) > 1:
                spacings = [ys[i+1]-ys[i] for i in range(len(ys)-1)]
                print(f"\n  Y spacings between points: {[f'{s:.2f}m' for s in spacings]}")
                all_same_x = len(set(f"{r['point'][0]:.2f}" for r in results)) == 1
                print(f"  All same X: {all_same_x} (expected: False for spread-along-tracker)")

        # Export mini GeoJSON
        features = [{"type":"Feature",
                     "geometry":{"type":"Point","coordinates":[r['point'][0],r['point'][1]]},
                     "properties":{"String_Code":r['val']}}
                    for r in results]
        out = {"type":"FeatureCollection","features":features}
        out_path = "test_output.geojson"
        with open(out_path, 'w') as f:
            json.dump(out, f, indent=2)
        print(f"\n  Test GeoJSON saved to: {out_path}")
        print(f"  Load in QGIS to verify positions before running full extraction.")

if __name__ == '__main__':
    main()
