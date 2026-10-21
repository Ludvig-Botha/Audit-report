# Setup Guide

## Publish the app with GitHub Pages (once, by the maintainer)

1. Push this repository to GitHub.
2. Repo **Settings → Pages** → *Build and deployment* → Source **Deploy from a branch** → Branch **main**, folder **/ (root)** → **Save**.
3. After a minute the page shows the address: `https://<github-username>.github.io/<repo-name>/`
   That link opens the app directly. Send it to the engineers.

> GitHub Pages is free for **public** repositories; a **private** repo needs a paid
> GitHub plan (Pro/Team). A public repo means anyone with the link can see the
> app, the question list and the JT report template — decide which suits the company.

**Updates:** push a rebuilt `app/JT_Audit_Capture.html`. Phones pick up the new
version automatically the next time the app is opened with signal (open it twice to see it).

## Phone app (every engineer)

1. Open the GitHub Pages link in **Chrome**.
2. Chrome menu (⋮) → **Add to Home screen** / **Install app**.
3. Open it once with signal — from then on it works offline.
4. In the app: **My details** → your name and job title.

Never open the app from the GitHub repository page itself — GitHub shows the code, it does not run it.

Reports made in a copy opened as a file (WhatsApp/Downloads) are stored separately
from the hosted app. To move them: old copy → *Generate Report* → **Export ZIP**;
hosted app → **Import job (ZIP)**. Same for questions: *Questions & lists* → **Export** / **Import**.

That's all. Reports are generated on the phone with **Generate Report → Generate Word Report**.

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
