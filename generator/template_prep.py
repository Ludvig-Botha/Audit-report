"""
template_prep.py - turns JT_Report_Template.docx into a fill-in skeleton.

Shared by:
  * generator/generate_report.py  (laptop generator, runs it on every report)
  * tools/build_app.py            (bakes the skeleton into the phone app)

so both produce identical Word documents.

What it does to the template (in memory only - the .docx on disk is untouched):
  * keeps the real cover page (JT logo + ACTOM banner) and the disclaimer page,
    replacing the sample customer / report number / distribution / prepared-by
    text in the cover text boxes with tokens
  * removes the sample report body; the cover + disclaimer become section 1,
    the generated report body becomes section 2 (default header/footer,
    3.5 cm top and bottom margins)
  * header1: sample report no / customer / date -> tokens
  * footer1: the static "of 36" -> a live NUMPAGES field
  * Normal style -> 10 pt
  * registers the jpg content type for photos

Tokens left in the output (all XML-escaped by the caller before insertion):
  document.xml : @@CUSTOMER@@ @@REPORTNO@@ @@PREPNAME@@ @@PREPTITLE@@
                 @@DIST@@ (paragraph-level, replaced with distribution lines)
                 @@BODY@@ (paragraph-level, replaced with the report body)
  header*.xml  : @@HRN@@ @@HCUST@@ @@HDATE@@
  dist_line    : one distribution paragraph containing @@NAME@@ and @@HOW@@
"""
import re
import zipfile

W_T = re.compile(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>")


def para_text(p):
    return "".join(W_T.findall(p))


def _set_para_text(p, new_text):
    """Replace the visible text of paragraph XML `p` with new_text.
    Keeps pPr, drawing runs and field runs; the first text run keeps its
    formatting and receives the text; all other plain text/tab runs go."""
    runs = list(re.finditer(r"<w:r\b[^>]*>.*?</w:r>", p, re.S))
    first_done = False
    out, last = [], 0
    for m in runs:
        r = m.group(0)
        out.append(p[last:m.start()])
        last = m.end()
        keep = ("<w:drawing" in r or "<mc:AlternateContent" in r or "<w:fldChar" in r
                or "<w:instrText" in r or "<w:pict" in r)
        if keep:
            out.append(r)
            continue
        if ("<w:t" in r or "<w:tab/>" in r) and not first_done and "<w:t" in r:
            rpr = re.search(r"<w:rPr>.*?</w:rPr>", r, re.S)
            out.append("<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>"
                       % (rpr.group(0) if rpr else "", new_text))
            first_done = True
            continue
        if "<w:t" in r or "<w:tab/>" in r or "<w:br" in r:
            continue          # drop extra text / tab runs
        out.append(r)
    out.append(p[last:])
    res = "".join(out)
    res = re.sub(r"<w:proofErr [^>]*/>", "", res)
    if not first_done:   # paragraph had no text run - append one
        res = res.replace("</w:p>", "<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>" % new_text)
    return res


def _top_children(body):
    """Split body XML into top-level <w:p>/<w:tbl>/<w:sdt> elements, stopping
    at the body's own sectPr. Returns (children, rest)."""
    out, i = [], 0
    starter = re.compile(r"<w:(p|tbl|sdt|sectPr)[ >/]")
    while True:
        m = starter.search(body, i)
        if not m or m.group(1) == "sectPr":
            return out, body[i:]
        tag = m.group(1)
        depth = 0
        end = None
        for mm in re.finditer(r"<(/?)w:%s(?=[ >/])[^>]*?(/?)>" % tag, body[m.start():]):
            if mm.group(2) == "/":
                if depth == 0:
                    end = m.start() + mm.end()
                    break
                continue
            depth += 1 if mm.group(1) == "" else -1
            if depth == 0:
                end = m.start() + mm.end()
                break
        out.append(body[m.start():end])
        i = end


def _numpages_field(rpr):
    return ("<w:r>%s<w:fldChar w:fldCharType=\"begin\"/></w:r>"
            "<w:r>%s<w:instrText xml:space=\"preserve\"> NUMPAGES </w:instrText></w:r>"
            "<w:r>%s<w:fldChar w:fldCharType=\"separate\"/></w:r>"
            "<w:r>%s<w:t>1</w:t></w:r>"
            "<w:r>%s<w:fldChar w:fldCharType=\"end\"/></w:r>") % ((rpr,) * 5)


def prepare(template_path):
    """Return {"files": {path: bytes}, "dist_line": str}.
    files["word/document.xml"] contains the tokens listed in the module doc."""
    z = zipfile.ZipFile(template_path)
    files = {n: z.read(n) for n in z.namelist() if not n.endswith("/")}
    doc = files["word/document.xml"].decode("utf-8")

    b0 = doc.index("<w:body>") + len("<w:body>")
    head = doc[:b0]
    body = doc[b0:doc.rindex("</w:body>")]
    children, _rest = _top_children(body)

    # --- cover + disclaimer = everything before the Executive Summary ---
    cut = None
    for k, c in enumerate(children):
        if 'w:val="ExecTitle"' in c:
            cut = k
            break
    if cut is None:
        raise ValueError("Template has no 'Exec Title' paragraph - cannot find end of cover pages")
    prefix = "".join(children[:cut])

    # --- section 1 properties (the sectPr that carries the header/footer refs) ---
    sects = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)
    sect1 = next((s for s in sects if "w:headerReference" in s and 'w:type="default"' in s), sects[0])
    sect1 = re.sub(r'\sw:rsid\w*="[^"]*"', "", sect1)

    # section 2 (report body): default header/footer only, no title page,
    # continue page numbering, 3.5 cm top AND bottom margin
    sect2 = re.sub(r'<w:(headerReference|footerReference) w:type="(first|even)"[^>]*/>', "", sect1)
    sect2 = sect2.replace("<w:titlePg/>", "")
    sect2 = re.sub(r"<w:pgNumType[^>]*/>", "", sect2)
    sect2 = re.sub(r'(<w:pgMar\b[^>]*w:bottom=")\d+(")', r"\g<1>1985\g<2>", sect2)
    sect2 = re.sub(r'(<w:pgMar\b[^>]*w:top=")\d+(")', r"\g<1>1985\g<2>", sect2)

    # --- cover text boxes ---
    dist_holder = {}
    sample_overflow = ("REPORT ON", "BOILER HOUSE AUDIT AT", "LACTALIS PORT ELIZABETH")

    def cover_para(p):
        t = para_text(p).strip()
        if not t:
            return p
        if t.upper().startswith("REPORT NO. :") or t in sample_overflow:
            return None                                  # hidden sample overflow lines
        if t.upper().startswith("REPORT NO."):
            return _set_para_text(p, "Report No.: @@REPORTNO@@")
        if t.startswith("Lactalis") or t.lower().startswith("customer"):
            return _set_para_text(p, "@@CUSTOMER@@")
        if t == "Ludvig Botha" or t == "Prepared by name":
            return _set_para_text(p, "@@PREPNAME@@")
        if t == "Service Engineer":
            return _set_para_text(p, "@@PREPTITLE@@")
        return p

    # distribution lines appear once per text-box copy; first line of each copy
    # becomes the marker, the rest are removed
    def box_fn(m):
        inner = m.group(1)
        state = {"marked": False}
        def do_p(pm):
            p = pm.group(0)
            t = para_text(p).strip()
            if p.count("<w:tab/>") >= 2 and t.endswith("-1"):
                if "dist_line" not in dist_holder:
                    dist_holder["dist_line"] = _make_dist_line(p)
                if not state["marked"]:
                    state["marked"] = True
                    return "@@DIST@@"
                return ""
            r = cover_para(p)
            return "" if r is None else r
        inner = re.sub(r"<w:p\b.*?</w:p>", do_p, inner, flags=re.S)
        return "<w:txbxContent>" + inner + "</w:txbxContent>"

    prefix = re.sub(r"<w:txbxContent>(.*?)</w:txbxContent>", box_fn, prefix, flags=re.S)

    section_break = "<w:p><w:pPr>%s</w:pPr></w:p>" % sect1
    new_doc = head + prefix + section_break + "@@BODY@@" + sect2 + "</w:body></w:document>"
    files["word/document.xml"] = new_doc.encode("utf-8")
    if "dist_line" not in dist_holder:
        raise ValueError("Could not find the distribution list in the template cover")

    # --- headers: sample report no / customer / date -> tokens ---
    for n in list(files):
        if re.match(r"word/header\d+\.xml$", n):
            x = files[n].decode("utf-8")
            def hdr_p(m):
                p = m.group(0)
                t = para_text(p).strip()
                if not t or "PAGE" in p:
                    return p
                low = t.lower()
                if "report" in low and "no" in low:
                    return _set_para_text(p, "Report No.: @@HRN@@")
                if low.startswith("date"):
                    return _set_para_text(p, "Date: @@HDATE@@")
                return _set_para_text(p, "@@HCUST@@")
            x = re.sub(r"<w:p\b.*?</w:p>", hdr_p, x, flags=re.S)
            files[n] = x.encode("utf-8")

    # --- footers: static "of 36" -> NUMPAGES field ---
    for n in list(files):
        if re.match(r"word/footer\d+\.xml$", n):
            x = files[n].decode("utf-8")
            RUN = r'<w:r\b[^>]*>(?:(?!</w:r>).)*?'
            m = re.search('(' + RUN + r'<w:t xml:space="preserve"> of </w:t></w:r>)'
                          r'((?:<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:t>\d+</w:t></w:r>)+)', x, re.S)
            if m:
                rpr_m = re.search(r"<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>", m.group(1), re.S)
                rpr = rpr_m.group(0) if rpr_m else ""
                x = x[:m.start(2)] + _numpages_field(rpr) + x[m.end(2):]
                files[n] = x.encode("utf-8")

    # --- Normal style 12 pt -> 10 pt ---
    st = files["word/styles.xml"].decode("utf-8")
    def normal_fix(m):
        s = m.group(0)
        s = re.sub(r'<w:sz w:val="\d+"/>', '<w:sz w:val="20"/>', s)
        s = re.sub(r'<w:szCs w:val="\d+"/>', '<w:szCs w:val="20"/>', s)
        return s
    st = re.sub(r'<w:style [^>]*w:styleId="Normal".*?</w:style>', normal_fix, st, count=1, flags=re.S)
    files["word/styles.xml"] = st.encode("utf-8")

    # --- jpg content type ---
    ct = files["[Content_Types].xml"].decode("utf-8")
    for ext in ("jpg", "jpeg"):
        if 'Extension="%s"' % ext not in ct:
            ct = ct.replace("<Default ", '<Default Extension="%s" ContentType="image/jpeg"/><Default ' % ext, 1)
    files["[Content_Types].xml"] = ct.encode("utf-8")

    return {"files": files, "dist_line": dist_holder["dist_line"]}


def _make_dist_line(p):
    """Build a clean distribution paragraph template from the sample line:
    NAME <tab> (HOW) <tab> - <tab> 1, with explicit tab stops so columns align."""
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S)
    ppr = ppr.group(0) if ppr else "<w:pPr></w:pPr>"
    ppr = re.sub(r"<w:tabs>.*?</w:tabs>", "", ppr, flags=re.S)
    tabs = ('<w:tabs><w:tab w:val="left" w:pos="4111"/><w:tab w:val="left" w:pos="5245"/>'
            '<w:tab w:val="left" w:pos="5812"/></w:tabs>')
    ppr = ppr.replace("<w:pPr>", "<w:pPr>" + tabs, 1)
    rpr = re.search(r"<w:r\b[^>]*><w:rPr>(.*?)</w:rPr>", p, re.S)
    rpr = "<w:rPr>%s</w:rPr>" % rpr.group(1) if rpr else ""
    rpr = re.sub(r"<w:lang [^>]*/>", "", rpr)
    def run(t):
        return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, t)
    tab = "<w:r>%s<w:tab/></w:r>" % rpr
    return "<w:p>" + ppr + run("@@NAME@@") + tab + run("@@HOW@@") + tab + run("-") + tab + run("1") + "</w:p>"
