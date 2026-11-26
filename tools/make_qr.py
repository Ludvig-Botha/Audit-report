#!/usr/bin/env python3
"""
make_qr.py - makes the QR code shown under My details -> Share the app.

    py -m pip install segno
    python tools/make_qr.py [url]

Writes app/src/share_qr.svg. Run it only if the app's web address changes,
then rebuild the app with tools/build_app.py.
"""
import os, re, sys
import segno

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
URL = sys.argv[1] if len(sys.argv) > 1 else "https://ludvig-botha.github.io/Audit-report/"

qr = segno.make(URL, error="m", micro=False)
dest = os.path.join(ROOT, "app", "src", "share_qr.svg")
qr.save(dest, kind="svg", scale=1, border=4, dark="#000", light="#fff", xmldecl=False, svgns=True, nl=False)
svg = open(dest, encoding="utf-8").read()
svg = re.sub(r'width="(\d+)" height="(\d+)"', r'viewBox="0 0 \1 \2" shape-rendering="crispEdges"', svg, count=1)
open(dest, "w", encoding="utf-8").write(svg)
open(os.path.join(ROOT, "app", "src", "share_url.txt"), "w").write(URL)
print("QR for %s -> %s (version %s)" % (URL, dest, qr.version))
