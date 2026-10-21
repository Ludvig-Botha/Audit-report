#!/usr/bin/env python3
"""
build_app.py - builds the single-file phone app.

    python tools/build_app.py [version]

Reads
    app/src/app_template.html        app source with /*QDB*/ /*BP*/ /*TPL*/ /*VER*/ slots
    app/src/questions.json           built-in question library
    generator/boilerplate.json       standard report wording (shared with the laptop generator)
    generator/JT_Report_Template.docx  the JT Word template (prepared by generator/template_prep.py)
Writes
    app/JT_Audit_Capture.html        the file engineers install on their phones
"""
import base64, datetime, json, os, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "generator"))
import template_prep  # noqa: E402


def main():
    ver = sys.argv[1] if len(sys.argv) > 1 else datetime.date.today().strftime("5.%Y%m%d")
    src = open(os.path.join(ROOT, "app", "src", "app_template.html"), encoding="utf-8").read()
    qdb = json.load(open(os.path.join(ROOT, "app", "src", "questions.json"), encoding="utf-8"))
    bp = json.load(open(os.path.join(ROOT, "generator", "boilerplate.json"), encoding="utf-8"))
    prep = template_prep.prepare(os.path.join(ROOT, "generator", "JT_Report_Template.docx"))

    files = {}
    for path, data in prep["files"].items():
        if path.endswith((".xml", ".rels")):
            files[path] = ["x", data.decode("utf-8")]
        else:
            files[path] = ["b", base64.b64encode(data).decode("ascii")]
    tpl = {"files": files, "dist": prep["dist_line"]}

    def js(obj):  # JSON safe to embed inside <script>
        return json.dumps(obj, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

    out = (src.replace("/*VER*/", js(ver)).replace("/*QDB*/", js(qdb))
              .replace("/*BP*/", js(bp)).replace("/*TPL*/", js(tpl)))
    for slot in ("/*VER*/", "/*QDB*/", "/*BP*/", "/*TPL*/"):
        assert slot not in out, slot
    dest = os.path.join(ROOT, "app", "JT_Audit_Capture.html")
    open(dest, "w", encoding="utf-8").write(out)
    print("built %s  (version %s, %d KB)" % (dest, ver, len(out.encode("utf-8")) // 1024))


if __name__ == "__main__":
    main()
