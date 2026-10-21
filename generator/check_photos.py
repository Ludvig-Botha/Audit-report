#!/usr/bin/env python3
"""
check_photos.py - inspect the photos inside an audit export zip.
Usage:  py check_photos.py "yourfile.zip"

Prints each photo's stored pixel size and its EXIF orientation tag, so we can
see why images come out rotated. No files are changed.
"""
import sys, zipfile, io, json

try:
    from PIL import Image
    from PIL.ExifTags import TAGS
except ImportError:
    print("Pillow not available to this Python. Run: py -m pip install pillow")
    sys.exit(1)

ORIENT = {
    1: "normal (upright)",
    2: "mirrored",
    3: "rotated 180",
    4: "mirrored + 180",
    5: "mirrored + 90 CCW",
    6: "rotated 90 CW (needs 90 CW to view)",
    7: "mirrored + 90 CW",
    8: "rotated 90 CCW (needs 90 CCW to view)",
}

def main():
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    z = zipfile.ZipFile(sys.argv[1])
    names = [n for n in z.namelist() if n.lower().startswith("photos/")
             and n.lower().endswith((".jpg", ".jpeg"))]
    if not names:
        print("No photos/ found in zip."); return
    print(f"{len(names)} photo(s). Showing first 12:\n")
    counts = {}
    for n in names:
        raw = z.read(n)
        try:
            im = Image.open(io.BytesIO(raw))
            w, h = im.size
            o = None
            exif = im.getexif()
            if exif:
                o = exif.get(274)  # 274 = Orientation
            counts[o] = counts.get(o, 0) + 1
            if names.index(n) < 12:
                shape = "portrait" if h > w else "landscape"
                print(f"  {n.split('/')[-1][:45]:45} {w}x{h} {shape:9} "
                      f"orientation={o} ({ORIENT.get(o,'no tag') if o else 'NO TAG'})")
        except Exception as e:
            print(f"  {n}: could not read ({e})")
    print("\nOrientation tag summary across all photos:")
    for o, c in sorted(counts.items(), key=lambda x: (x[0] is None, x[0])):
        print(f"  orientation {o if o is not None else 'NONE'}: {c} photo(s)"
              f"  - {ORIENT.get(o,'no tag') if o else 'no EXIF tag'}")

if __name__ == "__main__":
    main()
