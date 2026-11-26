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
BOX_W, BOX_H = round(7.6 * EMU_CM), round(8.0 * EMU_CM)
ROW_H = round(9.8 * 567)
INL_H, INL_ROW, CAP_MAX = round(5.6 * EMU_CM), round(6.6 * 567), 230   # photos placed inside a section (combustion test)
QS_VER = 6
FUEL_CAT_LABEL = {"coal": "Coal", "oil": "Oil / gas", "biomass": "Biomass"}
PRANK = {"High": 3, "Medium": 2, "Low": 1}
SITE_FUEL_SECS = {"b_burner_0": ("Fuel / HFO Handling & Ring Main", "burner"),
                  "b_site_fuelcoal": ("Fuel Storage & Handling — Coal", "stoker"),
                  "b_site_fuelbio": ("Fuel Storage & Handling — Biomass", "biomass")}
CT_SECTION = {'section': 'Combustion Test — Burner',
 'inline': True,
 'items': [{'q': 'MEASURED: Base load O2 measurement',
            'd': 'measured',
            't': 'CombTestBaseload',
            'f': 'Flue gas O2 measured at ___ % at base load, outside the optimal range.',
            'r': 'Adjust the fuel/air ratio to bring the flue gas oxygen content to between 3% and 6% across the load range.',
            'ok': 'Flue gas O2 measured at ___ % at base load.',
            'p': 'N/A',
            'id': 'b_burner_ct_0'},
           {'q': 'MEASURED: 25% load O2 measurement',
            'd': 'measured',
            't': 'CombTest25load',
            'f': 'Flue gas O2 measured at ___ % at 25% load, outside the optimal range.',
            'r': 'Adjust the fuel/air ratio to bring the flue gas oxygen content to between 3% and 6% across the load range.',
            'ok': 'Flue gas O2 measured at ___ % at 25% load.',
            'p': 'N/A',
            'id': 'b_burner_ct_1'},
           {'q': 'MEASURED: 50% load O2 measurement',
            'd': 'measured',
            't': 'CombTest50load',
            'f': 'Flue gas O2 measured at ___ % at 50% load, outside the optimal range.',
            'r': 'Adjust the fuel/air ratio to bring the flue gas oxygen content to between 3% and 6% across the load range.',
            'ok': 'Flue gas O2 measured at ___ % at 50% load.',
            'p': 'N/A',
            'id': 'b_burner_ct_2'},
           {'q': 'MEASURED: 75% load O2 measurement',
            'd': 'measured',
            't': 'CombTest75load',
            'f': 'Flue gas O2 measured at ___ % at 75% load, outside the optimal range.',
            'r': 'Adjust the fuel/air ratio to bring the flue gas oxygen content to between 3% and 6% across the load range.',
            'ok': 'Flue gas O2 measured at ___ % at 75% load.',
            'p': 'N/A',
            'id': 'b_burner_ct_3'},
           {'q': 'MEASURED: 100% load O2 measurement',
            'd': 'measured',
            't': 'CombTest100load',
            'f': 'Flue gas O2 measured at ___ % at 100% load, outside the optimal range.',
            'r': 'Adjust the fuel/air ratio to bring the flue gas oxygen content to between 3% and 6% across the load range.',
            'ok': 'Flue gas O2 measured at ___ % at 100% load.',
            'p': 'N/A',
            'id': 'b_burner_ct_4'}],
 'id': 'b_burner_ct'}


def firing_of(b):
    return FIRING.get(b.get("firing"), FIRING["burner"])


def fmt_date(iso):
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", iso or "")
    return "%s/%s/%s" % (m.group(3), m.group(2), m.group(1)) if m else (iso or "")


def lines(t):
    return [x.strip() for x in re.split(r"\r?\n+", str(t or "")) if x.strip()]


def set_of(qs, b):
    return qs["shared"] + qs.get(firing_of(b)["part"], [])


def fuels_of(J, b):
    return [f for f in J["fuels"] if f.get("cat") == firing_of(b)["fuel"]]


def boiler_fuel(J, b):
    fs = [f for f in fuels_of(J, b) if f.get("type")]
    pick = next((f for f in fs if f.get("id") == b.get("fuelId")), None)
    return pick["type"] if pick else " / ".join(f["type"] for f in fs)


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
    d["fuels"] = d.get("fuels") or []
    for b in d["boilers"]:
        if isinstance(b.get("fuel"), str):
            b["fuel"] = {"type": b["fuel"]}
        b.setdefault("extras", [])
        if b.get("firing") not in FIRING:
            b["firing"] = "burner"
    qs = d.get("questionSets") or {}
    for p in PARTS:
        qs.setdefault(p, [])
    d["questionSets"] = qs
    if (d.get("schema") or 5) < QS_VER:
        upgrade_job(d)
    return d


# --------------------------------------------------------------------------- older captures -> current layout
# (same logic as upgradeQS / upgradeJob in the app)
def upgrade_qs(qs):
    info = {"ct": None}
    for p in PARTS:
        qs.setdefault(p, [])

    def take(part, pat):
        out = []
        for sec in qs[part]:
            keep = []
            for it in sec["items"]:
                (out if re.search(pat, it.get("q", ""), re.I) else keep).append(it)
            sec["items"] = keep
        return out

    for old, new, f, r in (
            ("Minimum air pressure switch", "Secondary air pressure switch",
             "Secondary air pressure switch faulty / airline blocked.",
             "Replace the secondary air pressure switch and clean/reroute the airline. Burner ignition must be prevented when air pressure is not present."),
            ("Primary air differential pressure switch", "Primary air pressure switch",
             "Primary air pressure switch faulty / airlines blocked or damaged.",
             "Replace the primary air pressure switch; the burner must lock out if the primary air chamber pressure is not correct at ignition.")):
        for sec in qs["burner"]:
            for it in sec["items"]:
                if it.get("q") == old:
                    it["q"] = new
                    if not it.get("fs") and not it.get("rs"):
                        it["f"], it["r"] = f, r
    front = []

    def site_sec(sid, pat):
        x = next((z for z in qs["site"] if z.get("id") == sid or re.search(pat, z.get("section", ""), re.I)), None)
        if x is None:
            x = {"id": sid, "section": SITE_FUEL_SECS[sid][0], "fir": SITE_FUEL_SECS[sid][1], "items": []}
        if not any(x is f for f in front):
            front.append(x)
        return x

    fi = next((i for i, x in enumerate(qs["burner"]) if re.match(r"Fuel / HFO Handling", x.get("section", ""), re.I)), -1)
    ring = take("burner", r"^(Ring ?main|Heat tracing on oil lines)")
    if fi >= 0 or ring:
        if fi >= 0:
            fh = qs["burner"].pop(fi); fh.setdefault("id", "b_burner_0"); fh["section"] = SITE_FUEL_SECS["b_burner_0"][0]; front.append(fh)
        else:
            fh = site_sec("b_burner_0", r"^Fuel / HFO Handling")
        fh["fir"] = "burner"; fh["items"].extend(ring)
        for it in fh["items"]:
            if re.match(r"(DATA|MEASURED): ", it.get("q", "")):
                it["q"] = re.sub(r" at burner$", "", it["q"].replace(" at burner (spec)", " (spec)"))
    coal = take("stoker", r"^Fuel analysis records available")
    if coal:
        x = site_sec("b_site_fuelcoal", r"^Fuel Storage & Handling — Coal"); x["fir"] = "stoker"; x["items"].extend(coal)
    bio = take("biomass", r"^(Fuel storage|Fuel moisture)")
    if bio:
        x = site_sec("b_site_fuelbio", r"^Fuel Storage & Handling — Biomass"); x["fir"] = "biomass"; x["items"].extend(bio)
    for sec in qs["biomass"]:
        if re.match(r"Fuel Handling & Feeding", sec.get("section", ""), re.I):
            sec["section"] = "Fuel Feeding — Biomass"
    qs["site"] = front + [x for x in qs["site"] if not any(x is f for f in front)]
    at = -1
    for i, sec in enumerate(qs["burner"]):
        k = next((n for n, it in enumerate(sec["items"]) if re.match(r"Combustion test", it.get("q", ""), re.I)), -1)
        if k >= 0:
            info["ct"] = {"id": sec["items"][k].get("id"), "t": sec["items"][k].get("t")}
            del sec["items"][k]; at = i
    if not any(x.get("id") == "b_burner_ct" or re.match(r"Combustion Test", x.get("section", "")) for x in qs["burner"]):
        if at < 0:
            at = next((i for i, x in enumerate(qs["burner"]) if re.match(r"Combustion", x.get("section", ""), re.I)), -1)
        qs["burner"].insert(len(qs["burner"]) if at < 0 else at + 1, json.loads(json.dumps(CT_SECTION)))
    qs["_v"] = QS_VER
    return info


def merge_ans(answers, key, a):
    has = lambda x: bool(x) and any(x.get(k) for k in ("s", "v", "f", "r", "e"))
    t = answers.get(key)
    if not has(a):
        return
    if not has(t):
        answers[key] = a; return
    if a.get("s") == "Defect" or not t.get("s"):
        t["s"] = a.get("s") or t.get("s")
    for k in ("f", "r"):
        if a.get(k) and a[k] not in str(t.get(k) or ""):
            t[k] = (t[k] + "\n" if t.get(k) else "") + a[k]
    if a.get("f"): t["fAuto"] = False
    if not t.get("v"): t["v"] = a.get("v")
    if not t.get("e"): t["e"] = a.get("e")
    if PRANK.get(a.get("p"), 0) > PRANK.get(t.get("p"), 0):
        t["p"] = a.get("p")


def upgrade_job(J):
    photos = J.get("photos") or []
    qs = J["questionSets"]
    ct_old, ct_old_tag = ["b_burner_1_16"], ["CombustionTest"]
    if (qs.get("_v") or 0) < QS_VER:
        info = upgrade_qs(qs)
        if info["ct"]:
            ct_old.append(info["ct"]["id"]); ct_old_tag.append(info["ct"]["t"])
    site_it, site_tag = {}, {}
    for sec in qs["site"]:
        for it in sec["items"]:
            site_it[it.get("id")] = it; site_tag[it.get("t")] = it
    fn = len(J["fuels"])
    for b in J["boilers"]:
        st = set_of(qs, b)
        own = {it.get("id") for sec in st for it in sec["items"]}
        own_tag = {it.get("t") for sec in st for it in sec["items"]}
        ct_sec = next((x for x in st if x.get("inline") and re.search(r"combustion test", x.get("section", ""), re.I)), None)
        ct_new = ct_sec["items"][0] if ct_sec and ct_sec["items"] else None
        for k in list(J["answers"]):
            if not k.startswith(b["id"] + "|"):
                continue
            qid = k[len(b["id"]) + 1:]
            if qid in own:
                continue
            if qid in site_it:
                merge_ans(J["answers"], "site|" + qid, J["answers"].pop(k))
            elif ct_new and qid in ct_old:
                merge_ans(J["answers"], b["id"] + "|" + ct_new["id"], J["answers"].pop(k))
        for p in photos:
            if p.get("scope") != b["id"] or p.get("tag") in own_tag:
                continue
            if ct_new and p.get("tag") in ct_old_tag:
                p["tag"] = ct_new["t"]; p["qid"] = ct_new["id"]
            elif p.get("tag") in site_tag:
                p["scope"] = "site"; p["qid"] = site_tag[p["tag"]].get("id")
        bf = b.get("fuel")
        if isinstance(bf, dict):
            cat = firing_of(b)["fuel"]
            f = None
            if any(v for k, v in bf.items() if k != "typeOther"):
                f = next((x for x in J["fuels"] if x.get("cat") == cat and str(x.get("type") or "").lower() == str(bf.get("type") or "").lower()), None)
                if f is None:
                    fn += 1; f = {"id": "f%d" % fn, "cat": cat, "type": ""}; J["fuels"].append(f)
                for k, v in bf.items():
                    if k not in ("id", "cat") and not f.get(k):
                        f[k] = v
                b["fuelId"] = f["id"]
            for p in photos:
                if p.get("scope") == b["id"] and any(t == p.get("tag") for t, _ in FUEL_PHOTOS):
                    if f is None:
                        fn += 1; f = {"id": "f%d" % fn, "cat": cat, "type": ""}; J["fuels"].append(f)
                    p["scope"] = "site"; p["tag"] = "%s_%s" % (p["tag"], f["id"]); p["label"] = ""
            b.pop("fuel", None)
    J["schema"] = QS_VER


def pct_val(v):
    v = str(v)
    return v.strip() + " %" if re.match(r"^\s*\d+([.,]\d+)?\s*$", v) else v.strip()


def report_model(J, ph_keys=None):
    """Every question becomes its own block (title, data, finding, recommendation, its photos);
    plain "good condition" OKs without photos are listed together. ph_keys = set of "scope|tag" with photos."""
    qs = J["questionSets"]
    ph_keys = ph_keys or set()
    blabel = lambda b: "Boiler " + str(b.get("no", ""))
    order, secs = [], {}
    prio = {p: [] for p in PRIO}
    cnt = {"defects": 0}

    def add_prio(p, group, text, section):
        if p in PRIO:
            for t in lines(text):
                prio[p].append({"group": group, "text": t, "section": section})

    def new_entry(name, inline=False, ct=False):
        return {"name": name, "data": [], "blocks": [], "good": [], "find": [], "recs": [], "table": [],
                "thead": ["Load point", "Flue gas O2"] if ct else ["Measurement", "Value"],
                "inline": bool(inline), "tags": [], "touched": False}

    def add_item(e, o):
        has_ph = o["key"] in ph_keys
        if o.get("s") or o.get("v") or o.get("eq") or has_ph:
            e["touched"] = True
        else:
            return
        if o.get("s") == "Defect":
            cnt["defects"] += 1
        ans = o.get("s") in ("Defect", "OK")
        fl = lines(o.get("f")) if ans else []
        rl = lines(o.get("r")) if ans else []
        if ans:
            add_prio(o.get("p"), e["group"], o.get("r") or o.get("f"), e["name"])
        if o.get("d") == "data" and not has_ph and not fl:
            if o.get("v"):
                e["data"].append("%s: %s" % (o["title"], o["v"]))
            return
        if o.get("s") == "OK" and o.get("auto") and not rl and not has_ph and not o.get("v") and not o.get("eq"):
            e["good"].extend(fl); return
        data = []
        if o.get("eq"): data.append("Equipment: %s" % o["eq"])
        if o.get("v"): data.append("%s: %s" % ("Value" if o.get("d") == "data" else "Measured", o["v"]))
        e["blocks"].append({"title": o["title"], "data": data, "find": fl, "recs": rl, "key": o["key"]})

    def entry(sec, scope, group):
        ct = bool(re.search(r"combustion test", sec.get("section", ""), re.I))
        e = new_entry(clean_sec(sec["section"]), sec.get("inline"), ct)
        e["group"] = group
        for it in sec["items"]:
            a = J["answers"].get("%s|%s" % (scope, it.get("id"))) or {}
            if e["inline"]:
                e["tags"].append(it.get("t"))
                if it.get("d") == "measured":
                    e["table"].append([re.sub(r"\s*O2 measurement$", "", q_text(it), flags=re.I),
                                       (pct_val(a["v"]) if ct else str(a["v"])) if a.get("v") else "Not measured"])
                if a.get("s") or a.get("v") or a.get("e") or ("%s|%s" % (scope, it.get("t"))) in ph_keys:
                    e["touched"] = True
                if a.get("s") == "Defect":
                    cnt["defects"] += 1
                if a.get("s") in ("Defect", "OK"):
                    f, r = finding_of(it, a), rec_of(it, a)
                    if not (a.get("s") == "OK" and a.get("fAuto")):
                        e["find"].extend(lines(f))
                    e["recs"].extend(lines(r))
                    add_prio(a.get("p"), group, r or f, e["name"])
                continue
            add_item(e, {"title": q_text(it), "key": "%s|%s" % (scope, it.get("t")), "d": it.get("d"), "s": a.get("s"),
                         "f": finding_of(it, a), "r": rec_of(it, a), "p": a.get("p"),
                         "v": "" if it.get("d") == "check" else a.get("v"),
                         "eq": a.get("e") if it.get("opts") else "", "auto": bool(a.get("fAuto"))})
        return e

    for b in J["boilers"]:
        entries = [entry(sec, b["id"], blabel(b)) for sec in set_of(qs, b)]
        for x in b.get("extras", []):
            e = new_entry(clean_sec(x.get("name"))); e["group"] = blabel(b); e["extra"] = True
            for i in x.get("items", []):
                f = (i.get("f") or "").strip()
                if not f and i.get("s") == "Defect":
                    f = i.get("title", "") + ": defect noted."
                add_item(e, {"title": i.get("title", ""), "key": "%s|X%s" % (b["id"], i.get("id")), "d": "check",
                             "s": i.get("s"), "f": f, "r": (i.get("r") or "").strip(), "p": i.get("p"), "auto": bool(i.get("fAuto"))})
            entries.append(e)
        for e in entries:
            if not e["touched"]:
                continue
            if e["name"] not in secs:
                secs[e["name"]] = []; order.append(e["name"])
            e["boiler"] = b
            secs[e["name"]].append(e)

    site = [e for e in (entry(sec, "site", "Site") for sec in qs["site"]) if e["touched"]]
    groups = [blabel(b) for b in J["boilers"]] + ["Site"]
    return {"order": order, "secs": secs, "site": site, "prio": prio, "groups": groups, "defects": cnt["defects"]}


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


def w_bullet(t, keep=False):
    return w_p("•\t" + t, style="ListParagraph", ind=(720, 360), tabs=[720], after=40, keep_next=keep)


def w_bullet_l(label, t, keep=False):
    return w_p("", style="ListParagraph", ind=(720, 360), tabs=[720], after=40, keep_next=keep,
               raw=w_runs("•\t") + w_runs(label, bold=True) + w_runs(t))


def w_label(t):
    return w_p(t, bold=True, keep_next=True, before=120, after=40)


def w_table(rows, widths, bold_first=False, head=False):
    grid = "".join('<w:gridCol w:w="%d"/>' % w for w in widths)
    trs = "".join("<w:tr><w:trPr><w:cantSplit/></w:trPr>%s</w:tr>" % "".join(
        '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc>' % (widths[i], w_p(c, bold=(bold_first and i == 0) or (head and ri == 0), after=0))
        for i, c in enumerate(r)) for ri, r in enumerate(rows))
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


def w_photo_grid(cells, row_h=ROW_H):
    col = BODY_W // 2
    def cell(c):
        inner = (w_p("", jc="center", before=60, after=60, keep_next=True, raw=c["img"]) +
                 w_p(c["cap"], jc="center", italic=True, sz=18, after=60)) if c else w_p("", after=0)
        return '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>%s</w:tc>' % (col, inner)
    rows = ""
    for i in range(0, len(cells), 2):
        pair = cells[i:i + 2] + [None] * (2 - len(cells[i:i + 2]))
        rows += '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="%d" w:hRule="atLeast"/></w:trPr>%s%s</w:tr>' % (row_h, cell(pair[0]), cell(pair[1]))
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


def photo_ctx(J, p):
    """category = section name, text = the finding written for that question (or the question itself)"""
    qs = J["questionSets"]
    b = None if p.get("scope") == "site" else next((x for x in J["boilers"] if x["id"] == p.get("scope")), None)
    tag = p.get("tag") or ""
    if tag == "Nameplate":
        return {"cat": "Nameplate", "text": "Boiler %s nameplate" % ((b or {}).get("no") or "")}
    fm = re.match(r"^(FuelSample|FuelAnalysis|FuelStorage)(?:_(.+))?$", tag)
    if fm:
        f = next((x for x in J["fuels"] if x.get("id") == fm.group(2)), None)
        return {"cat": "Fuel", "text": dict(FUEL_PHOTOS)[fm.group(1)] + (" - " + f["type"] if f and f.get("type") else "")}
    secs = qs["site"] if p.get("scope") == "site" else (set_of(qs, b) if b else [])
    hit = None
    if p.get("qid"):
        hit = next(((sc, it) for sc in secs for it in sc["items"] if it.get("id") == p["qid"]), None)
    if not hit:
        hit = next(((sc, it) for sc in secs for it in sc["items"] if it.get("t") == tag), None)
    if hit:
        f = lines(finding_of(hit[1], J["answers"].get("%s|%s" % (p.get("scope"), hit[1].get("id"))) or {}))
        return {"cat": clean_sec(hit[0]["section"]), "text": (f[0] if f else "") or p.get("label") or q_text(hit[1])}
    if b:
        for x in b.get("extras", []):
            for i in x.get("items", []):
                if "X" + i["id"] == tag:
                    f = lines(i.get("f")) if i.get("s") in ("Defect", "OK") else []
                    return {"cat": clean_sec(x.get("name")), "text": (f[0] if f else "") or "%s - %s" % (x.get("name"), i.get("title"))}
    return {"cat": "Photo", "text": p.get("label") or re.sub(r"([a-z])([A-Z])", r"\1 \2", tag or "Photo")}


def photo_caption(J, p):
    t = str(p.get("caption") or "").strip() or photo_ctx(J, p)["text"]
    if len(t) > CAP_MAX:
        t = t[:CAP_MAX]; t = t[:max(t.rfind(" "), CAP_MAX - 40)] + "..."
    return t


def photo_order(J, photos):
    rank, n = {}, 0
    qs = J["questionSets"]
    for b in J["boilers"]:
        rank[b["id"] + "|Nameplate"] = n; n += 1
        for sc in set_of(qs, b):
            for it in sc["items"]:
                k = b["id"] + "|" + str(it.get("t"))
                if k not in rank:
                    rank[k] = n; n += 1
        for x in b.get("extras", []):
            for i in x.get("items", []):
                rank[b["id"] + "|X" + i["id"]] = n; n += 1
    for f in J["fuels"]:
        for t, _ in FUEL_PHOTOS:
            rank["site|%s_%s" % (t, f.get("id"))] = n; n += 1
    for sc in qs["site"]:
        for it in sc["items"]:
            k = "site|" + str(it.get("t"))
            if k not in rank:
                rank[k] = n; n += 1
    idx = {id(p): i for i, p in enumerate(photos)}
    return sorted(photos, key=lambda p: (rank.get("%s|%s" % (p.get("scope"), p.get("tag")), 10 ** 6), idx[id(p)]))


# --------------------------------------------------------------------------- report
def boiler_sentence(J, b):
    ft = boiler_fuel(J, b)
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


def fuel_rows(f):
    rows = [("Fuel type", f["type"])] if f.get("type") else []
    for k, l in FUEL_FIELDS.get(f.get("cat"), []) + FUEL_COMMON_END:
        if f.get(k):
            rows.append((l, f[k]))
    return rows


def build_body(J, photos, bp, rotate_cw, media):
    c = J["customer"]
    cust = " - ".join(x for x in (c.get("name"), c.get("site")) if x)
    out = []
    A = out.append
    # photos: each one is printed where it belongs (question, fuel, nameplate); leftovers go to a short chapter at the end
    st = {"fig": 0, "did": 900000000}
    all_ph = photo_order(J, photos)
    ph_by = {}
    for p in all_ph:
        ph_by.setdefault("%s|%s" % (p.get("scope"), p.get("tag")), []).append(p)
    M = report_model(J, set(ph_by))
    placed = set()

    def mk_cell(p, bw, bh):
        placed.add(id(p))
        try:
            data, pw, ph = doc_photo(p["raw"], rotate_cw)
        except Exception as ex:
            print("  ! could not read photo %s (%s)" % (p.get("file"), ex))
            return {"img": w_runs("[photo could not be read]"), "cap": p.get("file", "")}
        st["fig"] += 1; st["did"] += 1
        rid = "rIdJTPh%d" % st["fig"]
        media.append((rid, "media/jtphoto%d.jpg" % st["fig"], data))
        cx, cy = bw, round(bw * ph / pw)
        if cy > bh: cy, cx = bh, round(bh * pw / ph)
        return {"img": w_image(rid, cx, cy, st["did"]), "cap": "Figure %d: %s" % (st["fig"], photo_caption(J, p))}

    def photos_at(keys):
        lst = []
        for k in keys:
            for p in ph_by.get(k, []):
                if id(p) not in placed and not any(p is q for q in lst):
                    lst.append(p)
        if not lst:
            return ""
        return w_photo_grid([mk_cell(p, BOX_W, INL_H) for p in lst], INL_ROW) + w_p("", after=60)

    def write_entry(e, scope):
        h = []
        for d in e["data"]: h.append(w_bullet(d))
        if e["inline"]:
            if e["table"]: h.append(w_table([e["thead"]] + e["table"], [4700, BODY_W - 4700], head=True))
            if e["find"]:
                h.append(w_label("Findings")); h.extend(w_bullet(f) for f in e["find"])
            if e["recs"]:
                h.append(w_label("Recommendations")); h.extend(w_bullet(r) for r in e["recs"])
            h.append(photos_at(["%s|%s" % (scope, t) for t in e["tags"]]))
        else:
            for k in e["blocks"]:
                ph = photos_at([k["key"]])
                rows = [("", d) for d in k["data"]] + [("Finding: ", f) for f in k["find"]] + [("Recommendation: ", r) for r in k["recs"]]
                h.append(w_label(k["title"]))
                for i, (l, t) in enumerate(rows):
                    keep = bool(ph)   # keep the whole block with its first row of photos
                    h.append(w_bullet_l(l, t, keep) if l else w_bullet(t, keep))
                h.append(ph)
            if e["good"]:
                h.append(w_label("Found in good condition")); h.extend(w_bullet(g) for g in e["good"])
        x = "".join(h)
        return x or w_p(bp["no_issues"])

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
        A(w_p(bp["exec_listed"] if M["defects"] else bp["exec_none"]))
    A(w_pagebreak())
    # 1-3
    A(w_h(1, "Objective")); A(w_p(bp["objective"]))
    A(w_h(1, "Scope of Work"))
    req = [x for x in J["scope"]["requested"] if x and x.strip()]
    A("".join(w_bullet(x) for x in req) if req else w_p(bp["scope_default"]))
    A(w_h(1, "Description of Boiler Plant"))
    for b in J["boilers"]: A(w_p(boiler_sentence(J, b)))
    if not J["boilers"]: A(w_p("No boilers were recorded."))
    # 4 fuel
    fuel_keys = lambda f: ["site|%s_%s" % (t, f.get("id")) for t, _ in FUEL_PHOTOS]
    fuels = [f for f in J["fuels"] if fuel_rows(f) or any(k in ph_by for k in fuel_keys(f))]
    if fuels:
        A(w_h(1, "Fuel"))
        for f in fuels:
            A(w_h(2, f.get("type") or (FUEL_CAT_LABEL.get(f.get("cat"), "") + " fuel"), upper=False))
            users = ["Boiler %s" % b.get("no", "") for b in J["boilers"]
                     if any(x is f for x in fuels_of(J, b)) and (not b.get("fuelId") or b.get("fuelId") == f.get("id")
                                                                or not any(x.get("id") == b.get("fuelId") for x in fuels_of(J, b)))]
            if users: A(w_p("Used by: " + ", ".join(users) + "."))
            if fuel_rows(f): A(w_table(fuel_rows(f), [3800, BODY_W - 3800], bold_first=True))
            A(photos_at(fuel_keys(f)))
    # 5 findings
    A(w_h(1, "Boiler Inspection Findings"))
    if not M["order"]: A(w_p("No boiler inspection findings were recorded."))
    for name in M["order"]:
        A(w_h(2, name))
        for e in M["secs"][name]:
            A(w_h(3, "Boiler %s" % e["boiler"].get("no", ""), upper=False))
            A(write_entry(e, e["boiler"]["id"]))
    # 6 site
    if M["site"]:
        A(w_h(1, "Site-Wide Findings and Recommendations"))
        for e in M["site"]:
            A(w_h(2, e["name"]))
            A(write_entry(e, "site"))
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
    if not any(M["prio"][p] for p in PRIO):
        A(w_p(bp["prio_none"]))
    for p in (PRIO if any(M["prio"][p] for p in PRIO) else []):
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
                ("Year of manufacture", b.get("year")), ("Firing", firing_of(b)["label"]), ("Fuel", boiler_fuel(J, b))]
        xs = [x.get("name") for x in b.get("extras", []) if x.get("name")]
        if xs: rows.append(("Additional equipment", ", ".join(xs)))
        A(w_table([[r[0], r[1] or "-"] for r in rows], [3800, BODY_W - 3800], bold_first=True))
        A(photos_at([b["id"] + "|Nameplate"]))
    # 11 other photographs (only photos that have no question / place in the report)
    rest_ph = [p for p in all_ph if id(p) not in placed]
    if rest_ph:
        by_id = {b["id"]: b for b in J["boilers"]}
        lab = lambda s: "Site" if s == "site" else "Boiler %s" % by_id.get(s, {}).get("no", "")
        A(w_pagebreak()); A(w_h(1, "Other Photographs"))
        first = True
        for s in [b["id"] for b in J["boilers"]] + ["site"]:
            lst = [p for p in rest_ph if p.get("scope") == s]
            if not lst: continue
            if not first: A(w_pagebreak())
            first = False
            A(w_h(2, lab(s), upper=False))
            for i in range(0, len(lst), 4):
                if i: A(w_pagebreak()); A(w_label(lab(s) + " (continued)"))
                A(w_photo_grid([mk_cell(p, BOX_W, BOX_H) for p in lst[i:i + 4]]))
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
