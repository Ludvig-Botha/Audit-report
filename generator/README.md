# Laptop Report Generator

Builds the same Word report as the phone app, from the app's ZIP export.

```
py -m pip install pillow
py generate_report.py "yourfile.zip"            # -o name.docx  -r 90 (rotate photos, rarely needed)
```

- `template_prep.py` — prepares `JT_Report_Template.docx` (shared with the app build)
- `boilerplate.json` — standard report wording and O2 targets
- `check_photos.py` — shows the orientation tags of the photos in a ZIP
