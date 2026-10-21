#!/usr/bin/env python3
"""
generate_report.py - John Thompson Boiler House Audit Report (laptop version)
==============================================================================
Builds the Word report from a ZIP exported by the JT Audit Capture app.
The phone app can build the same report itself (Generate Word Report); this
script is the laptop fallback and produces the identical document.

USAGE
    py generate_report.py "export.zip"
    py generate_report.py "export.zip" -o "My report.docx"
    py generate_report.py "export.zip" -t JT_Report_Template.docx -r 90

    -t  template (default: JT_Report_Template.docx next to this script)
    -o  output file name (default: <zip name>_Report.docx)
    -r  extra clockwise rotation for ALL photos (90/180/270). Not normally
        needed - phone photo orientation is corrected automatically.

NEEDS: Python 3.8+ and Pillow   ->   py -m pip install pillow
Files in the same folder: template_prep.py, boilerplate.json, JT_Report_Template.docx
"""
import argparse, io, json, os, re, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import template_prep  # noqa: E402

try:
    from PIL import Image, ImageOps
except ImportError:
    sys.exit("Pillow is not installed for this Python. Run:  py -m pip install pillow")

FIRING = {
    "stoker":  {"label": "Coal fired (chain grate stoker)", "fuel": "coal", "part": "stoker"},
    "burner":  {"label": "Oil / gas fired (burner)", "fuel": "oil", "part": "burner"},
    "biomass": {"label": "Biomass fired", "fuel": "biomass", "part": "biomass"},
}
FUEL_FIELDS = {
    "coal": [("supplier", "Supplier / mine"), ("cv", "Calorific value (MJ/kg)"), ("moisture", "Total moisture (%)"),
             ("ash", "Ash content (%)"), ("volatile", "Volatile matter (%)"), ("fixedCarbon", "Fixed carbon (%)"),
             ("sulphur", "Sulphur (%)"), ("aft", "Ash fusion temperature (°C)"), ("size", "Size grading")],
    "oil": [("supplier", "Supplier"), ("cv", "Calorific value (MJ/kg)"), ("density", "Density (kg/m³)"),
            ("viscosity", "Viscosity (cSt @ °C)"), ("sulphur", "Sulphur (%)"), ("flash", "Flash point (°C)"),
            ("water", "Water content (%)"), ("supplyTemp", "Supply temperature at burner (°C)")],
    "biomass": [("supplier", "Supplier / source"), ("cv", "Calorific value (MJ/kg)"), ("moisture", "Moisture content (%)"),
                ("ash", "Ash content (%)"), ("size", "Particle size / grading"), ("bulk", "Bulk density (kg/m³)"),
                ("contaminants", "Contaminants (sand, stones, metal)")],
}
FUEL_COMMON_END = [("analysis", "Fuel analysis report available"), ("analysisDate", "Date of last fuel analysis"), ("notes", "Notes")]
FUEL_PHOTOS = [("FuelSample", "Fuel sample"), ("FuelAnalysis", "Fuel analysis report"), ("FuelStorage", "Fuel storage & handling")]
PRIO = ["High", "Medium", "Low"]
PARTS = ["shared", "stoker", "burner", "biomass", "site"]

BODY_W = 9412                      # twips between the 2.2 cm side margins
EMU_CM = 360000
BOX_W, BOX_H = round(7.6 * EMU_CM), round(8.4 * EMU_CM)
ROW_H = round(9.8 * 567)


def firing_of(b):
    return FIRING.get(b.get("firing"), FIRING["burner"])


def fmt_date(iso):
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", iso or "")
    return "%s/%s/%s" % (m.group(3), m.group(2), m.group(1)) if m else (iso or "")


def clean_sec(s):
    return re.sub(r"\s*\((?:report|data)[^)]*\)\s*$", "", s or "", flags=re.I).strip()


def q_text(it):
    return re.sub(r"^(DATA|MEASURED): ", "", it.get("q", ""))


def base_name(q):
    return re.sub(r"\s*(&\s*)?condition(\s*&\s*test)?\s*$", "", q, flags=re.I).strip()


def finding_of(it, a):
    s = a.get("s")
    t = (a.get("f") or (it or {}).get("f") or "") if s == "Defect" else ((a.get("f") or "") if s == "OK" else "")
    if a.get("v"):
        t = re.sub(r"_{3,}", str(a["v"]), t)
    return t.strip()


def rec_of(it, a):
    s = a.get("s")
    return ((a.get("r") or (it or {}).get("r") or "") if s == "Defect" else ((a.get("r") or "") if s == "OK" else "")).strip()


# --------------------------------------------------------------------------- model
def fix_job(d, bp):
    d.setdefault("answers", {}); d.setdefault("boilers", [])
    sc = d.get("scope") or {}
    d["scope"] = {"requested": sc.get("requested") or [], "done": sc.get("done") or []}
    c = d.setdefault("customer", {})
    if not isinstance(c.get("dist"), list):
        c["dist"] = [{"name": e, "email": e} for e in re.split(r"[,;\n]+", c.get("emails") or "") if e.strip()]
    d["prepared"] = d.get("prepared") or {"name": bp["prepared_by_name"], "title": bp["prepared_by_title"]}
    for b in d["boilers"]:
        if not isinstance(b.get("fuel"), dict):
            b["fuel"] = {"type": b.get("fuel") if isinstance(b.get("fuel"), str) else ""}
        b.setdefault("extras", [])
        if b.get("firing") not in FIRING:
            b["firing"] = "burner"
    qs = d.get("questionSets") or {}
    for p in PARTS:
        qs.setdefault(p, [])
    d["questionSets"] = qs
    return d


def report_model(J):
    qs = J["questionSets"]
    set_of = lambda b: qs["shared"] + qs.get(firing_of(b)["part"], [])
    blabel = lambda b: "Boiler " + str(b.get("no", ""))
    order, secs = [], {}
    prio = {p: [] for p in PRIO}

    def add_prio(p, group, text, section):
        if p in PRIO and text:
            prio[p].append({"group": group, "text": text, "section": section})

    for b in J["boilers"]:
        entries = []
        for sec in set_of(b):
            e = {"name": clean_sec(sec["section"]), "data": [], "find": [], "recs": [], "touched": False}
            for it in sec["items"]:
                a = J["answers"].get("%s|%s" % (b["id"], it.get("id")))
                if not a:
                    continue
                if a.get("s") or a.get("v") or a.get("e"):
                    e["touched"] = True
                if it.get("opts") and a.get("e"):
                    e["data"].append("%s: %s" % (base_name(q_text(it)), a["e"]))
                if it.get("d") in ("data", "measured") and a.get("v"):
                    e["data"].append("%s: %s" % (q_text(it), a["v"]))
                if a.get("s") in ("Defect", "OK"):
                    f, r = finding_of(it, a), rec_of(it, a)
                    if f: e["find"].append(f)
                    if r: e["recs"].append(r)
                    add_prio(a.get("p"), blabel(b), r or f, e["name"])
            entries.append(e)
        for x in b.get("extras", []):
            e = {"name": clean_sec(x.get("name")), "data": [], "find": [], "recs": [], "touched": False}
            for i in x.get("items", []):
                if i.get("s"):
                    e["touched"] = True
                if i.get("s") not in ("Defect", "OK"):
                    continue
                f, r = (i.get("f") or "").strip(), (i.get("r") or "").strip()
                if not f and i.get("s") == "Defect":
                    f = i.get("title", "") + ": defect noted."
                if f: e["find"].append(f)
                if r: e["recs"].append(r)
                add_prio(i.get("p"), blabel(b), r or f, e["name"])
            entries.append(e)
        for e in entries:
            if not e["touched"]:
                continue
            if e["name"] not in secs:
                secs[e["name"]] = []; order.append(e["name"])
            e["boiler"] = b
            secs[e["name"]].append(e)

    site = []
    for sec in qs["site"]:
        e = {"name": clean_sec(sec["section"]), "find": [], "recs": [], "touched": False}
        for it in sec["items"]:
            a = J["answers"].get("site|%s" % it.get("id"))
            if a and (a.get("s") or a.get("v")):
                e["touched"] = True
            if not a or a.get("s") not in ("Defect", "OK"):
                continue
            f, r = finding_of(it, a), rec_of(it, a)
            if f: e["find"].append(f)
            if r: e["recs"].append(r)
            add_prio(a.get("p"), "Site", r or f, e["name"])
        if e["touched"]:
            site.append(e)
    groups = [blabel(b) for b in J["boilers"]] + ["Site"]
    return {"order": order, "secs": secs, "site": site, "prio": prio, "groups": groups}


# --------------------------------------------------------------------------- XML writer
def XE(s):
    s = "" if s is None else str(s)
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", s)


def w_runs(text, bold=False, italic=False, sz=None):
    rpr = ("<w:b/>" if bold else "") + ("<w:i/>" if italic else "") + \
          ('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz) if sz else "")
    rpr = "<w:rPr>%s</w:rPr>" % rpr if rpr else ""
    parts = ["<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>" % (rpr, XE(s)) if s else ""
             for s in ("" if text is None else str(text)).split("\t")]
    return ("<w:r>%s<w:tab/></w:r>" % rpr).join(parts)


def w_p(text="", style=None, keep_next=False, tabs=None, before=None, after=None, ind=None, jc=None,
        bold=False, italic=False, sz=None, raw=None):
    ppr = ""
    if style: ppr += '<w:pStyle w:val="%s"/>' % style
    if keep_next: ppr += "<w:keepNext/>"
    if tabs: ppr += "<w:tabs>" + "".join('<w:tab w:val="left" w:pos="%d"/>' % t for t in tabs) + "</w:tabs>"
    if before is not None or after is not None:
        ppr += "<w:spacing%s%s/>" % (' w:before="%d"' % before if before is not None else "",
                                     ' w:after="%d"' % after if after is not None else "")
    if ind: ppr += '<w:ind w:left="%d" w:hanging="%d"/>' % ind
    if jc: ppr += '<w:jc w:val="%s"/>' % jc
    return "<w:p>%s%s</w:p>" % ("<w:pPr>%s</w:pPr>" % ppr if ppr else "", raw if raw is not None else w_runs(text, bold, italic, sz))


def w_pagebreak():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def w_h(level, text, upper=None):
    up = (level <= 2) if upper is None else upper
    return w_p(str(text).upper() if up else text, style="Heading%d" % level, keep_next=True)


def w_bullet(t):
    return w_p("•\t" + t, style="ListParagraph", ind=(720, 360), tabs=[720], after=40)


def w_label(t):
    return w_p(t, bold=True, keep_next=True, before=120, after=40)


def w_table(rows, widths, bold_first=False):
    grid = "".join('<w:gridCol w:w="%d"/>' % w for w in widths)
    trs = "".join("<w:tr><w:trPr><w:cantSplit/></w:trPr>%s</w:tr>" % "".join(
        '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc>' % (widths[i], w_p(c, bold=bold_first and i == 0, after=0))
        for i, c in enumerate(r)) for r in rows)
    return ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="%d" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            '<w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>' % (sum(widths), grid, trs)) + w_p("", after=0)


def w_image(rid, cx, cy, did):
    return ('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="%d" cy="%d"/>'
            '<wp:docPr id="%d" name="Photo %d"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="%d" name="photo%d.jpg"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
            % (cx, cy, did, did, did, did, rid, cx, cy))


def w_photo_grid(cells):
    col = BODY_W // 2
    def cell(c):
        inner = (w_p("", jc="center", before=60, after=60, keep_next=True, raw=c["img"]) +
                 w_p(c["cap"], jc="center", italic=True, sz=18, after=60)) if c else w_p("", after=0)
        return '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>%s</w:tc>' % (col, inner)
    rows = ""
    for i in range(0, len(cells), 2):
        pair = cells[i:i + 2] + [None] * (2 - len(cells[i:i + 2]))
        rows += '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="%d" w:hRule="atLeast"/></w:trPr>%s%s</w:tr>' % (ROW_H, cell(pair[0]), cell(pair[1]))
    nb = "".join('<w:%s w:val="nil"/>' % s for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return ('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>%s</w:tblBorders>'
            '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="%d"/><w:gridCol w:w="%d"/></w:tblGrid>%s</w:tbl>'
            % (BODY_W, nb, col, col, rows))


# --------------------------------------------------------------------------- photos
def doc_photo(raw, rotate_cw=0):
    im = Image.open(io.BytesIO(raw))
    im = ImageOps.exif_transpose(im)
    if rotate_cw:
        im = im.rotate(-rotate_cw, expand=True)
    if im.mode != "RGB":
        im = im.convert("RGB")
    im.thumbnail((1600, 1600))
    out = io.BytesIO(); im.save(out, format="JPEG", quality=85)
    return out.getvalue(), im.size[0], im.size[1]


def photo_label(J, p):
    if p.get("label"):
        return p["label"]
    for part in PARTS:
        for s in J["questionSets"].get(part, []):
            for it in s["items"]:
                if it.get("t") == p.get("tag"):
                    return q_text(it)
    for t, l in FUEL_PHOTOS:
        if t == p.get("tag"):
            return l
    return re.sub(r"([a-z])([A-Z])", r"\1 \2", p.get("tag") or "Photo")


def photo_order(J, photos):
    rank, n = {}, 0
    qs = J["questionSets"]
    for b in J["boilers"]:
        rank[b["id"] + "|Nameplate"] = n; n += 1
        for t, _ in FUEL_PHOTOS:
            rank[b["id"] + "|" + t] = n; n += 1
        for s in qs["shared"] + qs.get(firing_of(b)["part"], []):
            for it in s["items"]:
                k = b["id"] + "|" + str(it.get("t"))
                if k not in rank:
                    rank[k] = n; n += 1
        for x in b.get("extras", []):
            for i in x.get("items", []):
                rank[b["id"] + "|X" + i["id"]] = n; n += 1
    for s in qs["site"]:
        for it in s["items"]:
            k = "site|" + str(it.get("t"))
            if k not in rank:
                rank[k] = n; n += 1
    idx = {id(p): i for i, p in enumerate(photos)}
    return sorted(photos, key=lambda p: (rank.get("%s|%s" % (p.get("scope"), p.get("tag")), 10 ** 6), idx[id(p)]))


# --------------------------------------------------------------------------- report
def boiler_sentence(b):
    ft = (b.get("fuel") or {}).get("type") or ""
    if b["firing"] == "stoker":
        fire = "coal-fired chain grate stoker"
    elif b["firing"] == "biomass":
        fire = "biomass-fired" + (" (%s)" % ft if ft else "")
    else:
        fire = (ft + "-fired" if ft else "oil / gas-fired") + " burner"
    mm = " ".join(x for x in (b.get("make"), b.get("model")) if x)
    s = "Boiler %s is a %s%s boiler" % (b.get("no", ""), (mm + " ") if mm else "", fire)
    parts = []
    if b.get("capacity"): parts.append("a rated capacity of %s kg/h" % b["capacity"])
    if b.get("pressure"): parts.append("a design pressure of %s bar(g)" % b["pressure"])
    if parts: s += " with " + " and ".join(parts)
    if b.get("year"): s += ", manufactured in %s" % b["year"]
    if b.get("serial"): s += " (serial no. %s)" % b["serial"]
    s += "."
    xs = [x.get("name") for x in b.get("extras", []) if x.get("name")]
    if xs:
        s += " The boiler is fitted with the following additional equipment: %s." % ", ".join(xs)
    return s


def fuel_rows(b):
    f = b.get("fuel") or {}
    rows = [("Fuel type", f["type"])] if f.get("type") else []
    for k, l in FUEL_FIELDS[firing_of(b)["fuel"]] + FUEL_COMMON_END:
        if f.get(k):
            rows.append((l, f[k]))
    return rows


def build_body(J, photos, bp, rotate_cw, media):
    M = report_model(J)
    c = J["customer"]
    cust = " - ".join(x for x in (c.get("name"), c.get("site")) if x)
    out = []
    A = out.append
    # executive summary
    A(w_p("EXECUTIVE SUMMARY", style="ExecTitle"))
    A(w_p(bp["exec_intro"].replace("{customer}", cust or "the customer's site").replace("{date}", fmt_date(c.get("date")))))
    themes, seen = [], set()
    for p in ("High", "Medium"):
        for e in M["prio"][p]:
            key = e["section"] + "|" + e["text"].lower()
            if key in seen: continue
            seen.add(key)
            t = next((x for x in themes if x["s"] == e["section"]), None)
            if not t:
                t = {"s": e["section"], "items": []}; themes.append(t)
            t["items"].append(e["text"])
    if themes:
        A(w_p(bp["exec_recommend_lead"]))
        for t in themes:
            A(w_label(t["s"] + ":"))
            for x in t["items"]: A(w_bullet(x))
    else:
        A(w_p(bp["exec_none"]))
    A(w_pagebreak())
    # 1-3
    A(w_h(1, "Objective")); A(w_p(bp["objective"]))
    A(w_h(1, "Scope of Work"))
    req = [x for x in J["scope"]["requested"] if x and x.strip()]
    A("".join(w_bullet(x) for x in req) if req else w_p(bp["scope_default"]))
    A(w_h(1, "Description of Boiler Plant"))
    for b in J["boilers"]: A(w_p(boiler_sentence(b)))
    if not J["boilers"]: A(w_p("No boilers were recorded."))
    # 4 fuel
    if any(fuel_rows(b) for b in J["boilers"]):
        A(w_h(1, "Fuel"))
        for b in J["boilers"]:
            rows = fuel_rows(b)
            ft = (b.get("fuel") or {}).get("type")
            A(w_h(2, "Boiler %s%s" % (b.get("no", ""), " - " + ft if ft else ""), upper=False))
            A(w_table(rows, [3800, BODY_W - 3800], bold_first=True) if rows else w_p("Fuel details were not recorded."))
    # 5 findings
    A(w_h(1, "Boiler Inspection Findings"))
    if not M["order"]: A(w_p("No boiler inspection findings were recorded."))
    for name in M["order"]:
        A(w_h(2, name))
        for e in M["secs"][name]:
            A(w_h(3, "Boiler %s" % e["boiler"].get("no", ""), upper=False))
            for d in e["data"]: A(w_bullet(d))
            if e["find"]:
                A(w_label("Findings")); [A(w_bullet(f)) for f in e["find"]]
            if e["recs"]:
                A(w_label("Recommendations")); [A(w_bullet(r)) for r in e["recs"]]
            if not (e["data"] or e["find"] or e["recs"]): A(w_p(bp["no_issues"]))
    # 6 site
    if M["site"]:
        A(w_h(1, "Site-Wide Findings and Recommendations"))
        for e in M["site"]:
            A(w_h(2, e["name"]))
            if e["find"]:
                A(w_label("Findings")); [A(w_bullet(f)) for f in e["find"]]
            if e["recs"]:
                A(w_label("Recommendations")); [A(w_bullet(r)) for r in e["recs"]]
            if not (e["find"] or e["recs"]): A(w_p(bp["no_issues"]))
    # 7 efficiency
    A(w_h(1, "Efficiency")); A(w_p(bp["efficiency_text"]))
    firings = []
    for b in J["boilers"]:
        if b["firing"] not in firings and b["firing"] in bp["efficiency_targets"]:
            firings.append(b["firing"])
    if firings:
        A(w_table([["Firing", "Target flue gas O2"]] + [bp["efficiency_targets"][f] for f in firings], [5600, BODY_W - 5600]))
    # 8 recommendations
    A(w_h(1, "Recommendations"))
    for p in PRIO:
        A(w_h(2, p + " Priority", upper=False))
        items = M["prio"][p]
        if not items:
            A(w_p("None.")); continue
        for g in M["groups"]:
            its = [e for e in items if e["group"] == g]
            if not its: continue
            A(w_label(g))
            for e in its: A(w_bullet(e["text"]))
    # 9 work done
    A(w_h(1, "Work Done on Site"))
    done = [x for x in J["scope"]["done"] if x and x.strip()]
    A("".join(w_bullet(x) for x in done) if done else w_p(bp["work_done_none"]))
    # 10 appendix
    A(w_pagebreak()); A(w_h(1, "Appendix A - Boiler Details"))
    for b in J["boilers"]:
        A(w_label("Boiler %s" % b.get("no", "")))
        rows = [("Make", b.get("make")), ("Model", b.get("model")), ("Serial number", b.get("serial")),
                ("Capacity (kg/h)", b.get("capacity")), ("Design pressure (bar g)", b.get("pressure")),
                ("Year of manufacture", b.get("year")), ("Firing", firing_of(b)["label"]), ("Fuel", (b.get("fuel") or {}).get("type"))]
        xs = [x.get("name") for x in b.get("extras", []) if x.get("name")]
        if xs: rows.append(("Additional equipment", ", ".join(xs)))
        A(w_table([[r[0], r[1] or "-"] for r in rows], [3800, BODY_W - 3800], bold_first=True))
    # 11 photographs
    if photos:
        ordered = photo_order(J, photos)
        by_id = {b["id"]: b for b in J["boilers"]}
        lab = lambda s: "Site" if s == "site" else "Boiler %s" % by_id.get(s, {}).get("no", "")
        A(w_pagebreak()); A(w_h(1, "Photographs"))
        fig, did, first = 0, 900000000, True
        for s in [b["id"] for b in J["boilers"]] + ["site"]:
            lst = [p for p in ordered if p.get("scope") == s]
            if not lst: continue
            if not first: A(w_pagebreak())
            first = False
            A(w_h(2, lab(s), upper=False))
            for i in range(0, len(lst), 4):
                if i: A(w_pagebreak()); A(w_label(lab(s) + " (continued)"))
                cells = []
                for p in lst[i:i + 4]:
                    try:
                        data, pw, ph = doc_photo(p["raw"], rotate_cw)
                    except Exception as ex:
                        print("  ! could not read photo %s (%s)" % (p.get("file"), ex))
                        cells.append({"img": w_runs("[photo could not be read]"), "cap": p.get("file", "")}); continue
                    fig += 1; did += 1
                    rid = "rIdJTPh%d" % fig
                    media.append((rid, "media/jtphoto%d.jpg" % fig, data))
                    cx, cy = BOX_W, round(BOX_W * ph / pw)
                    if cy > BOX_H: cy, cx = BOX_H, round(BOX_H * pw / ph)
                    cells.append({"img": w_image(rid, cx, cy, did),
                                  "cap": "Figure %d: %s - %s" % (fig, lab(s), p.get("caption") or photo_label(J, p))})
                A(w_photo_grid(cells))
    return "".join(out)


def build(zip_path, template, out_path, rotate_cw=0):
    bp = json.load(open(os.path.join(HERE, "boilerplate.json"), encoding="utf-8"))
    z = zipfile.ZipFile(zip_path)
    J = fix_job(json.loads(z.read("report_data.json").decode("utf-8")), bp)
    photos = []
    for p in J.get("photos", []):
        try:
            p["raw"] = z.read(p["file"]); photos.append(p)
        except KeyError:
            print("  ! photo listed but missing from zip:", p.get("file"))

    prep = template_prep.prepare(template)
    files = dict(prep["files"])
    media = []
    body = build_body(J, photos, bp, rotate_cw, media)

    c, pr = J["customer"], J["prepared"]
    cust = " - ".join(x for x in (c.get("name"), c.get("site")) if x)
    line = lambda nm, how: prep["dist_line"].replace("@@NAME@@", XE(nm)).replace("@@HOW@@", how)
    dist = "".join(line(d.get("name") or d.get("email"), "(email)") for d in c["dist"] if d.get("name") or d.get("email"))
    dist += line("Master File", "")
    doc = files["word/document.xml"].decode("utf-8")
    for tok, val in (("@@CUSTOMER@@", XE(cust)), ("@@REPORTNO@@", XE(c.get("reportNo", ""))),
                     ("@@PREPNAME@@", XE(pr.get("name") or bp["prepared_by_name"])),
                     ("@@PREPTITLE@@", XE(pr.get("title") or bp["prepared_by_title"])), ("@@DIST@@", dist)):
        doc = doc.replace(tok, val)
    doc = doc.replace("@@BODY@@", body, 1)
    files["word/document.xml"] = doc.encode("utf-8")
    for n in list(files):
        if re.match(r"word/header\d+\.xml$", n):
            x = files[n].decode("utf-8").replace("@@HRN@@", XE(c.get("reportNo", ""))) \
                .replace("@@HCUST@@", XE(cust)).replace("@@HDATE@@", XE(fmt_date(c.get("date"))))
            files[n] = x.encode("utf-8")
    rels = files["word/_rels/document.xml.rels"].decode("utf-8")
    rels = rels.replace("</Relationships>", "".join(
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="%s"/>'
        % (rid, name) for rid, name, _ in media) + "</Relationships>")
    files["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    for rid, name, data in media:
        files["word/" + name] = data

    names = sorted(files, key=lambda n: (n != "[Content_Types].xml", 0))
    with zipfile.ZipFile(out_path, "w") as out:
        for n in names:
            out.writestr(n, files[n], compress_type=zipfile.ZIP_STORED if n.endswith(".jpg") else zipfile.ZIP_DEFLATED)
    return len(media)


def main():
    ap = argparse.ArgumentParser(description="Build the JT boiler audit Word report from an app ZIP.")
    ap.add_argument("zip")
    ap.add_argument("-t", "--template", default=os.path.join(HERE, "JT_Report_Template.docx"))
    ap.add_argument("-o", "--output")
    ap.add_argument("-r", "--rotate", type=int, default=0, choices=[0, 90, 180, 270])
    a = ap.parse_args()
    if not os.path.exists(a.zip): sys.exit("ZIP not found: " + a.zip)
    if not os.path.exists(a.template): sys.exit("Template not found: " + a.template)
    out = a.output or re.sub(r"\.zip$", "", a.zip, flags=re.I) + "_Report.docx"
    try:
        n = build(a.zip, a.template, out, a.rotate)
    except PermissionError:
        sys.exit("Cannot write %s - close it in Word and run again." % out)
    print("Report written: %s  (%d photos)" % (os.path.abspath(out), n))


if __name__ == "__main__":
    main()
