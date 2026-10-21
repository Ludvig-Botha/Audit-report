# Setup Guide

## Phone app (every engineer)

1. Copy `app/JT_Audit_Capture.html` to the phone.
2. Open it in **Chrome** (Files app → tap the file → *Open with* → Chrome).
3. Chrome menu (⋮) → **Add to Home screen**.
4. In the app: **My details** → your name and job title.

That's all. Reports are generated on the phone with **Generate Report → Generate Word Report**.

**Updating to a new version:** before replacing the app file, open each unfinished
report → *Generate Report* → **Export ZIP**, and *Questions & lists* → **Export**.
Install the new file, then **Import job (ZIP)** and **Import** the questions.
(If the new file opens with your reports already there, nothing needs importing.)

## Laptop generator (optional)

1. Install Python 3 from <https://www.python.org/downloads/> — tick **"Add python.exe to PATH"**.
2. Open a new PowerShell window and run:
   ```
   py -m pip install pillow
   ```
   (Only Pillow is needed now. python-docx and lxml are no longer used.)
3. Put these in one folder, e.g. `C:\BoilerReports`:
   `generate_report.py`, `template_prep.py`, `boilerplate.json`, `JT_Report_Template.docx`, and the ZIP from the phone.
4. Run:
   ```
   cd C:\BoilerReports
   py generate_report.py "yourfile.zip"
   ```

| Problem | Fix |
|---------|-----|
| `pip is not recognized` | use `py -m pip install pillow` |
| `Pillow is not installed for this Python` | run `py -m pip install pillow` with the same `py` you run the script with |
| `Cannot write ... close it in Word` | close the old report in Word, run again |
| Photos tilted | run `py check_photos.py "yourfile.zip"`; if needed add `-r 90` |

## Rebuilding the app (maintainer)

```
python tools/build_app.py
```
Rebuild after editing `app/src/questions.json`, `generator/boilerplate.json`,
`generator/JT_Report_Template.docx` or `app/src/app_template.html`.
