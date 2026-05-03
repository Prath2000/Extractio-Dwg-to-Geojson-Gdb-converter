"""
GSA Layer Diagnostics — run standalone to inspect any layer before configuring it.
Usage: python gsa_diagnose.py
Then follow prompts.
"""
import sys, math

def connect_autocad():
    try:
        import win32com.client
        acad = win32com.client.Dispatch("AutoCAD.Application")
        return acad
    except Exception as e:
        print(f"Cannot connect to AutoCAD: {e}")
        sys.exit(1)

def choose_document(acad):
    """List all open DWGs and let user pick one."""
    docs = []
    try:
        for i in range(acad.Documents.Count):
            docs.append(acad.Documents.Item(i))
    except Exception as e:
        print(f"Cannot list documents: {e}")
        return acad.ActiveDocument

    if len(docs) == 1:
        print(f"Active DWG: {docs[0].Name}")
        return docs[0]

    print(f"\nOpen DWG files ({len(docs)}):")
    for i, d in enumerate(docs):
        active = " ← active" if d.Name == acad.ActiveDocument.Name else ""
        print(f"  [{i+1}] {d.Name}{active}")
    print(f"  [Enter] Use active document")

    choice = input("\nChoose DWG number: ").strip()
    if choice == "":
        return acad.ActiveDocument
    try:
        idx = int(choice) - 1
        if 0 <= idx < len(docs):
            return docs[idx]
    except ValueError:
        pass
    return acad.ActiveDocument

def etype(ent):
    try:
        n = ent.EntityName.upper()
        return {"ACDBBLOCKREFERENCE":"INSERT","ACDBPOLYLINE":"LWPOLYLINE",
                "ACDB2DPOLYLINE":"POLYLINE","ACDBLINE":"LINE","ACDBPOINT":"POINT",
                "ACDBMTEXT":"MTEXT","ACDBTEXT":"TEXT","ACDBCIRCLE":"CIRCLE",
                "ACDBARC":"ARC"}.get(n, n)
    except:
        return "UNKNOWN"

def inspect_insert(ent, doc, verbose=True):
    """Print everything about an INSERT entity."""
    print(f"\n  Block Name : {ent.Name}")
    try:
        ip = ent.InsertionPoint
        print(f"  InsertPt   : ({ip[0]:.4f}, {ip[1]:.4f})")
    except: pass
    try: print(f"  Rotation   : {ent.Rotation:.4f} rad = {math.degrees(ent.Rotation):.2f} deg")
    except: pass
    try: print(f"  Scale X/Y  : {ent.XScaleFactor:.4f} / {ent.YScaleFactor:.4f}")
    except: pass

    # Attributes
    try:
        if ent.HasAttributes:
            attrs = ent.GetAttributes()
            print(f"  Attributes ({len(attrs)}):")
            for att in attrs:
                try:
                    tag = att.TagString
                    val = att.TextString
                    try:    ip2 = att.InsertionPoint;  ipt = f"({ip2[0]:.2f},{ip2[1]:.2f})"
                    except: ipt = "N/A"
                    try:    tap = att.TextAlignmentPoint; tapt = f"({tap[0]:.2f},{tap[1]:.2f})"
                    except: tapt = "N/A"
                    try:    mp = att.MidPoint; mpt = f"({mp[0]:.2f},{mp[1]:.2f})"
                    except: mpt = "N/A"
                    print(f"    {tag:8s} = {val!r:30s}  InsertPt={ipt}  AlignPt={tapt}  MidPt={mpt}")
                except: pass
    except: pass

    # Block definition contents — show all polylines with areas
    if verbose:
        try:
            bdef = doc.Blocks.Item(ent.Name)
            etypes = {}
            for i in range(bdef.Count):
                sub = bdef.Item(i)
                st = etype(sub)
                etypes[st] = etypes.get(st, 0) + 1
            print(f"  Block def  : {dict(etypes)}")
            # Show each polyline area
            for i in range(bdef.Count):
                sub = bdef.Item(i)
                st = etype(sub)
                if st == "LWPOLYLINE":
                    try:
                        cr   = list(sub.Coordinates)
                        cpts = [(cr[j],cr[j+1]) for j in range(0,len(cr),2)]
                        a    = abs(sum(cpts[k][0]*cpts[k+1][1]-cpts[k+1][0]*cpts[k][1]
                                       for k in range(len(cpts)-1))/2.0)
                        lyr  = getattr(sub,'Layer','?')
                        print(f"    LWPOLY [{i}] layer={lyr} verts={len(cpts)} area={a:.4f}")
                    except: pass
        except: pass

def inspect_lwpoly(ent):
    try:
        cr   = list(ent.Coordinates)
        pts  = [(cr[i], cr[i+1]) for i in range(0, len(cr), 2)]
        try:    bulges = list(ent.GetBulges())
        except: bulges = []
        has_bulge = any(abs(b) > 1e-9 for b in bulges)
        try:    closed = ent.Closed
        except: closed = False
        area = 0
        try:    area = ent.Area
        except: pass
        print(f"  Vertices   : {len(pts)}  Closed={closed}  HasBulge={has_bulge}  Area={area:.4f}")
        print(f"  First pt   : ({pts[0][0]:.4f}, {pts[0][1]:.4f})")
        print(f"  Last pt    : ({pts[-1][0]:.4f}, {pts[-1][1]:.4f})")
    except Exception as e:
        print(f"  Error: {e}")

def diagnose_layer(msp, doc, layer_name):
    print(f"\n{'='*60}")
    print(f"  Layer: {layer_name}")
    print(f"{'='*60}")

    entities = []
    for ent in msp:
        try:
            if ent.Layer == layer_name:
                entities.append(ent)
        except: pass

    if not entities:
        print("  No entities found on this layer.")
        return

    # Count by type
    counts = {}
    for e in entities:
        t = etype(e)
        counts[t] = counts.get(t, 0) + 1
    print(f"  Total entities: {len(entities)}")
    print(f"  By type: {counts}")

    # Show first of each type in detail
    shown = set()
    for e in entities:
        t = etype(e)
        if t in shown:
            continue
        shown.add(t)
        print(f"\n  --- First {t} ---")
        if t == "INSERT":
            inspect_insert(e, doc)
        elif t == "LWPOLYLINE":
            inspect_lwpoly(e)
        elif t in ("TEXT", "MTEXT"):
            try: print(f"  Text: {e.TextString!r}")
            except: pass
            try:
                ip = e.InsertionPoint
                print(f"  Pos: ({ip[0]:.4f}, {ip[1]:.4f})")
            except: pass

    # For LWPOLYLINE layers: show ALL areas to detect outliers
    polys = [e for e in entities if etype(e) == "LWPOLYLINE"]
    if len(polys) > 1:
        print(f"\n  --- All LWPOLYLINE areas ({len(polys)} total) ---")
        areas = []
        for e in polys:
            try:
                a = e.Area
                areas.append(a)
            except:
                areas.append(0)
        areas_sorted = sorted(enumerate(areas), key=lambda x: x[1], reverse=True)
        for rank, (idx, a) in enumerate(areas_sorted[:10]):
            print(f"  [{idx:3d}] area={a:.4f} sqm")
        if len(polys) > 10:
            print(f"  ... ({len(polys)-10} more)")
        median = sorted(areas)[len(areas)//2]
        outliers = [a for a in areas if a > median * 100]
        if outliers:
            print(f"\n  ⚠ WARNING: {len(outliers)} outlier polygons (>{median*100:.1f} sqm)")
            print(f"  Median area: {median:.4f} sqm")

def main():
    print("GSA Layer Diagnostics")
    print("Connecting to AutoCAD...")
    acad = connect_autocad()
    doc  = choose_document(acad)
    msp  = doc.ModelSpace
    print(f"\nUsing: {doc.Name}")

    # List all layers with entities
    print("\nScanning layers...")
    layer_counts = {}
    for ent in msp:
        try:
            l = ent.Layer
            layer_counts[l] = layer_counts.get(l, 0) + 1
        except: pass

    print(f"Found {len(layer_counts)} layers with entities.")
    print("\nEnter layer name to diagnose (or 'list' to see all, 'quit' to exit):")

    while True:
        inp = input("\n> ").strip()
        if inp.lower() == 'quit':
            break
        elif inp.lower() == 'list':
            for l in sorted(layer_counts):
                print(f"  {layer_counts[l]:5d}  {l}")
        elif inp.lower() == 'switch':
            doc  = choose_document(acad)
            msp  = doc.ModelSpace
            print(f"Switched to: {doc.Name}")
            print("Scanning layers...")
            layer_counts = {}
            for ent in msp:
                try:
                    l = ent.Layer
                    layer_counts[l] = layer_counts.get(l, 0) + 1
                except: pass
            print(f"Found {len(layer_counts)} layers with entities.")

        elif inp.lower().startswith('prefix '):
            prefix = inp[7:].strip()
            matched = [l for l in layer_counts if l.lower().startswith(prefix.lower())]
            print(f"  Layers matching '{prefix}':")
            for l in sorted(matched):
                print(f"    {layer_counts[l]:5d}  {l}")
        else:
            if inp in layer_counts:
                diagnose_layer(msp, doc, inp)
            else:
                # Try prefix
                matched = [l for l in layer_counts if l.lower().startswith(inp.lower())]
                if matched:
                    print(f"  Prefix matched {len(matched)} layers. Diagnosing first: {matched[0]}")
                    diagnose_layer(msp, doc, matched[0])
                else:
                    print(f"  Layer '{inp}' not found.")

if __name__ == "__main__":
    main()
