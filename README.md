# JT Boiler House Audit — Capture App & Report Generator

Offline Android app for John Thompson boiler-house audits. Engineers walk the
site, answer the checklist per boiler, take photos, and tap **Generate Word
Report** — the finished report, on the real JT template, is built on the phone.
No laptop, no Python, no internet.

| Part | Who uses it | What it does |
|------|-------------|--------------|
| **`app/JT_Audit_Capture.html`** | Every engineer, on an Android phone | Capture the audit and generate the Word report on the phone |
| **`generator/generate_report.py`** | Optional, laptop | Builds the identical report from the app's ZIP export (fallback / batch use) |
| **`tools/build_app.py`** | Whoever maintains the app | Rebuilds the app after changing questions, wording or the template |

---

## For engineers — using the app

**Install (once)**
1. Open the app link in **Chrome**: `https://<github-username>.github.io/<repo-name>/`
   (the repo is published with GitHub Pages — see docs/SETUP.md).
2. Chrome menu (⋮) → **Add to Home screen** / **Install app**.
3. Open it once while you have signal — after that it opens **without signal** too.
4. **My details** → enter your name and job title (printed on the cover as "Prepared by").

**Getting a colleague set up:** *My details & share app* shows a QR code. They scan it with the phone camera, open the link in Chrome and add it to the home screen.

Do **not** open the file from the GitHub repo page — GitHub only shows the code, it doesn't run it.
(Without hosting, the `.html` file can also be sent by WhatsApp/email and opened in Chrome.)

**Doing an audit**
1. **+ New Report** → customer, report number, date, distribution list.
2. **+ Add Boiler** → nameplate photo, details, and firing type: **Coal (stoker)**, **Oil / gas (burner)** or **Biomass**.
3. **Questionnaire** for each boiler - answer each question **OK / Defect / N/A / Skip**. Green ✔ / red ✘ shows what's done.
   - **OK** writes the comment for you ("Gauge glass is in a good condition."). Tap *Edit comment* to change it or add a recommendation.
   - **Defect** opens the finding, recommendation and priority. Pick from the **suggestion drop-downs** or type your own - what you type is added to the drop-down for next time.
   - Priority starts at **N/A**. Pick **High / Medium / Low** to put the item in the priority list at the end of the report.
   - **Combustion Test** (burners): O2 at base load, 25 %, 50 %, 75 % and 100 % load. Photos taken here are printed with the results, not at the back.
   - **+ Add additional equipment** - for things not on every boiler (economiser, bag filter, …).
4. **Site-Wide Questions & Fuel** - the **fuel** (type, analysis, photos), fuel handling and ring main, water treatment, steam traps, insulation, operators. Then **Scope of Work & Work Done**.
5. **Generate Report** → **Generate Word Report**. The `.docx` lands in Downloads; **Share report** sends it straight to email/WhatsApp. **Export ZIP** is still there for the laptop or for moving a job to another phone (**Import job (ZIP)** on the home screen).

**Phone and laptop (OneDrive)**
- Once: Home → **OneDrive** → sign in with the same Microsoft account as the FSO app (see docs/SETUP.md for the one-off Azure step).
- On the device you worked on: open the report → **Save to OneDrive**. On mobile data the photos wait for Wi-Fi unless you say "upload now".
- On the other device: Home → **Open from OneDrive** → tap the report.
- Questions & lists: **Questions & lists → OneDrive → Save** on one device, **Load** on the other.

In the report each photo is printed directly under its question, captioned with the finding. Answer the question first, then take the photo:
every photo is also saved to Downloads as `Category_Finding_B1_Customer_yymmdd_01.jpg` as a backup.

---

## What the report contains

Real JT cover (logo + ACTOM banner, customer, report number, distribution,
prepared by) → template disclaimer page → **Executive Summary** → 1 Objective →
2 Scope of Work → 3 Description of Boiler Plant → 4 Fuel → 5 Boiler Inspection
Findings (each section, each inspected boiler under it, incl. additional
equipment; each question with its finding, recommendation and **its own photos**) → 6 Site-Wide (incl. fuel handling) → 7 Efficiency (O2 targets) → 8 Recommendations
(High / Medium / Low, grouped by boiler) → 9 Work Done → 10 Appendix A (nameplate photos) → 11 Other Photographs (only photos without a question).

Header on every body page: logo, report number, customer, date. Footer: JOHN
THOMPSON and **Page X of Y**.

---

## Laptop generator (optional)

```
py -m pip install pillow
py generate_report.py "yourfile.zip"
```
Keep `generate_report.py`, `template_prep.py`, `boilerplate.json` and
`JT_Report_Template.docx` together in one folder. See [docs/SETUP.md](docs/SETUP.md).

---

## Maintaining the app

| To change… | Edit | Then |
|------------|------|------|
| Built-in questions | `app/src/questions.json` | `python tools/build_app.py` |
| Standard wording, O2 targets | `generator/boilerplate.json` | `python tools/build_app.py` |
| Cover, disclaimer, header/footer, styles | `generator/JT_Report_Template.docx` | `python tools/build_app.py` |
| App behaviour / screens | `app/src/app_template.html` | `python tools/build_app.py` |

The build writes `app/JT_Audit_Capture.html`; send that file to the engineers.
Engineers can also edit questions on their own phone (**Questions & lists**) and
**Export** them to share one library across the team.

---

## Repository layout

```
index.html                     sends the GitHub Pages link straight to the app
app/JT_Audit_Capture.html      the app (built - this is what goes on phones)
app/manifest.json, sw.js, icon-*.png   home-screen install + offline support
app/src/app_template.html      app source
app/src/questions.json         built-in question library
generator/generate_report.py   laptop generator
generator/template_prep.py     template handling shared by app build + generator
generator/boilerplate.json     standard report wording
generator/JT_Report_Template.docx  JT Word template
generator/check_photos.py      photo orientation diagnostic
tools/build_app.py             builds the app
docs/                          setup guide, changelog, question reference
samples/sample_report.docx     example report generated by the app
```

*Internal tooling for John Thompson (a division of ACTOM Pty Ltd) boiler-house
audits. Not an official ACTOM product.*
