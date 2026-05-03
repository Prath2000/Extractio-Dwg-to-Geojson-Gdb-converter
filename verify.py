"""
GSA Output Verifier
===================
Reads extracted geojson and compares against live AutoCAD data.
Shows exactly what was extracted vs what's in the DWG.

Usage: python gsa_verify.py <layer_name>
Example: python gsa_verify.py "String Inverters"
"""
import sys, os, json, math

# ── CONFIG ────────────────────────────────────────────────────────────────────
OUTPUT_DIR = None  # auto-detected from config

def connect_autocad():
    try:
        import win32com.client
        acad = win32com.client.Dispatch("AutoCAD.Application")
        return acad, acad.ActiveDocument
    except Exception as e:
        print(f"AutoCAD connection failed: {e}")
        sys.exit(1)

def find_output_dir():
    """Find output dir from tora_config yaml in current directory."""
    import glob, re
    for f in glob.glob("*.yaml") + glob.glob("tora_config*.yaml"):
        with open(f) as fh:
            content = fh.read()
        m = re.search(r'output_dir:\s*(.+)', content)
        if m:
            d = m.group(1).strip().strip('"').strip("'")
            if os.path.exists(d):
                return d
    return "."

def load_geojson(layer_name, output_dir):
    """Find and load the geojson for a layer."""
    import glob
    # Try common name patterns
    candidates = [
        layer_name.lower().replace(" ", "_") + ".geojson",
        layer_name.lower().replace(" ", "") + ".geojson",
    ]
    for c in candidates:
        path = os.path.join(output_dir, c)
        if os.path.exists(path):
            with open(path, encoding='utf-8') as f:
                return json.load(f), path
    # Search all geojsons
    for path in glob.glob(os.path.join(output_dir, "*.geojson")):
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        name = data.get("name", "")
        if layer_name.lower() in name.lower() or name.lower() in layer_name.lower():
            return data, path
    return None, None

def bbox(coords):
    if not coords:
        return None
    if isinstance(coords[0], (int, float)):
        return coords[0], coords[1], coords[0], coords[1]
    xs = [c[0] for c in coords]
    ys = [c[1] for c in coords]
    return min(xs), min(ys), max(xs), max(ys)

def feature_bbox(feat):
    geom = feat.get("geometry", {})
    gtype = geom.get("type", "")
    coords = geom.get("coordinates", [])
    if gtype == "Point":
        return coords[0], coords[1], coords[0], coords[1]
    elif gtype == "Polygon":
        return bbox(coords[0])
    elif gtype == "LineString":
        return bbox(coords)
    return None

def dist2d(a, b):
    return math.sqrt((a[0]-b[0])**2 + (a[1]-b[1])**2)

def verify_layer(layer_name):
    output_dir = find_output_dir()
    print(f"\nOutput dir: {output_dir}")

    fc, path = load_geojson(layer_name, output_dir)
    if not fc:
        print(f"ERROR: No geojson found for '{layer_name}' in {output_dir}")
        print("Files available:")
        for f in os.listdir(output_dir):
            if f.endswith('.geojson'):
                print(f"  {f}")
        return

    features = fc.get("features", [])
    print(f"\nLoaded: {path}")
    print(f"Features: {len(features)}")

    if not features:
        print("No features to verify.")
        return

    # Analyse geometry
    gtypes = {}
    for f in features:
        gt = f.get("geometry", {}).get("type", "?")
        gtypes[gt] = gtypes.get(gt, 0) + 1
    print(f"Geometry types: {gtypes}")

    # Show first 3 features with bbox
    print(f"\nFirst 3 features:")
    for i, feat in enumerate(features[:3]):
        bb = feature_bbox(feat)
        props = feat.get("properties", {})
        conn_id = props.get("Connection_ID", props.get("connection_id", "?"))
        if bb:
            w = bb[2]-bb[0]
            h = bb[3]-bb[1]
            print(f"  [{i}] ConnID={conn_id}")
            print(f"       BBox: ({bb[0]:.2f},{bb[1]:.2f}) → ({bb[2]:.2f},{bb[3]:.2f})")
            print(f"       Size: {w:.4f}m × {h:.4f}m")
            geom = feat.get("geometry", {})
            if geom.get("type") == "Point":
                c = geom["coordinates"]
                print(f"       Point: ({c[0]:.4f}, {c[1]:.4f})")
            elif geom.get("type") == "LineString":
                coords = geom["coordinates"]
                print(f"       Start: ({coords[0][0]:.4f}, {coords[0][1]:.4f})")
                print(f"       End:   ({coords[-1][0]:.4f}, {coords[-1][1]:.4f})")
                print(f"       Vertices: {len(coords)}")

    # Overall extent
    all_bb = [feature_bbox(f) for f in features if feature_bbox(f)]
    if all_bb:
        min_x = min(b[0] for b in all_bb)
        min_y = min(b[1] for b in all_bb)
        max_x = max(b[2] for b in all_bb)
        max_y = max(b[3] for b in all_bb)
        print(f"\nOverall extent:")
        print(f"  X: {min_x:.2f} to {max_x:.2f}  span={max_x-min_x:.2f}m")
        print(f"  Y: {min_y:.2f} to {max_y:.2f}  span={max_y-min_y:.2f}m")

    # For polygon layers: show size distribution
    polys = [f for f in features if f.get("geometry",{}).get("type") == "Polygon"]
    if polys:
        areas = []
        for f in polys:
            bb = feature_bbox(f)
            if bb:
                w, h = bb[2]-bb[0], bb[3]-bb[1]
                areas.append(w*h)
        if areas:
            areas.sort()
            print(f"\nPolygon size distribution (approx bbox area):")
            print(f"  Min:    {areas[0]:.4f} sqm")
            print(f"  Median: {areas[len(areas)//2]:.4f} sqm")
            print(f"  Max:    {areas[-1]:.4f} sqm")
            # Flag suspiciously large polygons
            median = areas[len(areas)//2]
            outliers = [a for a in areas if a > median * 10]
            if outliers:
                print(f"\n  WARNING: {len(outliers)} features are >10x median size!")
                print(f"  These may be wrong polygons (outer boundary instead of footprint)")
                print(f"  Largest: {outliers[-1]:.2f} sqm vs median {median:.4f} sqm")

    # Now cross-check with live AutoCAD if available
    print(f"\nCross-checking with AutoCAD...")
    try:
        acad, doc = connect_autocad()
        msp = doc.ModelSpace
        # Find matching entities by proximity to first feature
        if features:
            bb = feature_bbox(features[0])
            if bb:
                cx = (bb[0]+bb[2])/2
                cy = (bb[1]+bb[3])/2
                print(f"  Searching AutoCAD near first feature centroid ({cx:.2f}, {cy:.2f})...")
                nearest = []
                for ent in msp:
                    try:
                        et = ent.EntityName.upper()
                        if 'BLOCKREFERENCE' in et:
                            ip = ent.InsertionPoint
                            d = dist2d((ip[0],ip[1]),(cx,cy))
                            if d < 50:
                                nearest.append((d, ent.Layer, ent.Name, ip))
                        elif 'POLYLINE' in et or 'LINE' in et:
                            try:
                                cr = list(ent.Coordinates)
                                if cr:
                                    d = dist2d((cr[0],cr[1]),(cx,cy))
                                    if d < 50:
                                        nearest.append((d, ent.Layer, 'LWPOLY', (cr[0],cr[1])))
                            except: pass
                    except: continue
                nearest.sort()
                if nearest:
                    print(f"  Nearest DWG entities within 50m:")
                    for d, lyr, name, ip in nearest[:5]:
                        print(f"    {d:.2f}m away | layer={lyr} | {name} | pos=({ip[0]:.2f},{ip[1]:.2f})")
                else:
                    print(f"  No DWG entities found within 50m of first feature")
    except Exception as e:
        print(f"  AutoCAD check skipped: {e}")

def main():
    if len(sys.argv) < 2:
        print("Usage: python gsa_verify.py <layer_name>")
        print("Example: python gsa_verify.py \"String Inverters\"")
        print("\nAvailable geojsons:")
        output_dir = find_output_dir()
        for f in os.listdir(output_dir):
            if f.endswith('.geojson'):
                path = os.path.join(output_dir, f)
                with open(path) as fh:
                    fc = json.load(fh)
                n = len(fc.get("features", []))
                print(f"  {f} ({n} features)")
        return

    verify_layer(sys.argv[1])

if __name__ == "__main__":
    main()


def check_all(output_dir=None):
    """Quick sanity check on ALL geojsons in output folder."""
    import glob
    if not output_dir:
        output_dir = find_output_dir()
    
    print(f"\n{'='*60}")
    print(f"  REGRESSION CHECK — All Layers")
    print(f"{'='*60}")
    
    issues = []
    ok = []
    
    for path in sorted(glob.glob(os.path.join(output_dir, "*.geojson"))):
        fname = os.path.basename(path)
        try:
            with open(path, encoding='utf-8') as f:
                fc = json.load(f)
            feats = fc.get("features", [])
            if not feats:
                issues.append(f"  ⚠  {fname}: 0 features")
                continue
            
            gtypes = {}
            for ft in feats:
                gt = ft.get("geometry",{}).get("type","?")
                gtypes[gt] = gtypes.get(gt,0)+1
            
            # Check for suspiciously large polygons
            polys = [f for f in feats if f.get("geometry",{}).get("type")=="Polygon"]
            if polys:
                areas = []
                for f in polys:
                    bb = feature_bbox(f)
                    if bb:
                        areas.append((bb[2]-bb[0])*(bb[3]-bb[1]))
                if areas:
                    median = sorted(areas)[len(areas)//2]
                    outliers = [a for a in areas if a > median * 100]
                    if outliers:
                        issues.append(f"  ✗  {fname}: {len(outliers)} outlier polygons "
                                     f"(largest={outliers[-1]:.1f}sqm vs median={median:.4f}sqm)")
                        continue
            
            ok.append(f"  ✓  {fname}: {len(feats)} features {dict(gtypes)}")
        except Exception as e:
            issues.append(f"  ✗  {fname}: ERROR {e}")
    
    for line in ok:
        print(line)
    if issues:
        print()
        print("ISSUES FOUND:")
        for line in issues:
            print(line)
    else:
        print("\nAll layers OK ✓")

if __name__ == "__main__":
    if len(sys.argv) == 1:
        check_all()
    elif sys.argv[1] == "--all":
        check_all()
    else:
        verify_layer(sys.argv[1])
