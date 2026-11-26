# Changelog

## 5.4
### App
- **OneDrive: Save / Open on any device.** Sign in once (Home > **OneDrive**) with
  the same Microsoft account and Azure Client ID as the FSO app. On this device
  the Client ID is copied from the FSO app automatically, and an existing FSO app
  sign-in is reused.
  - Report menu > **Save to OneDrive** uploads the report to
    `OneDrive > JT Audit > reports > <id>` (report.json + one file per photo)
    and lists it in `reports/index.json`.
  - Home > **Open from OneDrive** lists every saved report with its state:
    *not on this device*, *newer on OneDrive*, *changed here*, *up to date*.
    Tap to open; photos you already have are not downloaded again.
  - **Photos upload on Wi-Fi only.** On mobile data the app asks: upload now,
    or save the report text and let the photos go up by themselves the next
    time the device is on Wi-Fi with the app open. "Upload photos now" on the
    report forces it. Downloads on mobile data ask first too.
  - **No silent overwrites.** Saving over a OneDrive copy that another device
    saved after this one last synced asks first; opening over local changes
    that are not on OneDrive asks first.
  - Photos deleted on a device are removed from OneDrive on its next save.
    Deleting a report in the app keeps the OneDrive copy; delete that from the
    Open from OneDrive list.
  - Photos over 4 MB use a resumable upload session.
- **Questions & lists > OneDrive**: Save / Load the question library and
  suggestion lists (`JT Audit/library.json`), same checks as Import.
- The Microsoft sign-in library is loaded only when needed and pinned with a
  Subresource Integrity hash. Scopes: Files.ReadWrite, User.Read.

### Setup (once)
- Azure portal > App registrations > the FSO app's registration >
  Authentication > Single-page application > add the audit app's address as a
  Redirect URI (shown in the app under Home > OneDrive).

## 5.3
### Report
- **Photos sit with their question.** Every question is its own block: question title, data (equipment / measured value), Finding, Recommendation, then that question's photos, 2 side by side (~6 cm). The block is kept together with its first row of photos.
- OK answers with only the automatic "good condition" comment and no photo are listed together under **Found in good condition** at the end of each section.
- Nameplate photos print in **Appendix A** under the boiler's details; fuel photos in the **Fuel** chapter; additional-equipment photos under their check; combustion test photos stay with the results table.
- The Photographs chapter is gone. Only photos that belong to no question (rare) go to a short **Other Photographs** chapter at the end.
- App and laptop generator produce identical reports.

## 5.2
### App
- **Fuel is site-wide.** The fuel form (type, analysis, photos) and all fuel-handling questions moved out of the boilers into *Site-Wide Questions & Fuel*: HFO handling, **ring main**, oil-line heat tracing, coal fuel analysis records, biomass fuel storage and moisture. Fuel sections only show when the site has that kind of boiler. A site can have more than one fuel.
- **All priorities start at N/A.** A defect only goes into the High / Medium / Low list when the engineer picks a priority. The Generate screen warns how many defects are still N/A.
- **OK writes the comment.** Tapping OK fills in "<item> is in a good condition." (e.g. "Gauge glass is in a good condition.", "Crown valve is in a good condition."). Editable; removed again if the answer is changed to Defect.
- **Suggestion drop-downs.** Every question has a drop-down of suggested findings and one of suggested recommendations. Anything typed by hand is added to that question's list automatically (also for additional equipment). Lists are edited under *Questions & lists*.
- **Photo captions = the finding** of the question the photo belongs to (a typed caption still wins).
- **Photo file names** start with the category, then the finding: `BoilerWaterLevel_GaugeGlassIsInAGoodCondition_B1_Customer_yymmdd_01.jpg`. Names are refreshed from the current answers in the ZIP export.
- **Combustion test** is its own burner section with five readings: base load, 25 %, 50 %, 75 %, 100 % load.
- Renamed: *Minimum air pressure switch* -> **Secondary air pressure switch**; *Primary air differential pressure switch* -> **Primary air pressure switch**.
- Several findings / recommendations on separate lines become separate bullets.

### Report
- Combustion test: results table per boiler with **its photos directly under it**, not in the Photographs chapter.
- Fuel chapter is per fuel ("Used by: Boiler 1, 2").
- Site-wide sections print their data values.
- Executive summary and Recommendations no longer say "no deficiencies / None" when defects exist but none were prioritised.

### Upgrading
- Phones update themselves. Edited question libraries, reports in progress and ZIPs from older versions are converted automatically (answers and photos of moved questions go to Site-Wide; a boiler's fuel becomes a site fuel).
- Laptop generator mirrors all of the above and still reads older ZIPs.

## 5.1
- Hosted on GitHub Pages: the repo link opens the app (root `index.html`), installs to the home screen and **works offline** after the first open (`app/manifest.json`, `app/sw.js`, icons).
- New versions are picked up automatically when opened with signal.

## 5.0
### App
- **Generate Word Report on the phone** — real JT template (cover, disclaimer, header/footer), no laptop or Python needed. Share straight to email/WhatsApp. Export ZIP kept alongside.
- **Import job (ZIP)** — reopen an exported job on any phone, including jobs captured with earlier versions.
- **OK answers can carry a comment and recommendation**, with priority High / Medium / Low / **N/A** (N/A stays out of the priority list).
- **Additional equipment sections** per boiler (economiser, bag filter, …). Names and their checklists are remembered.
- **Biomass** firing type with its own question set (fuel feeding, grate/combustion, ash & emissions).
- **Fuel form** per boiler (coal / oil-gas / biomass): fuel type from a saved list (new types are remembered), analysis values, photos of fuel, analysis report and storage.
- **My details** — each engineer's name/title goes on the cover as "Prepared by". Distribution list with names.
- Questions & lists export/import to share one library across the team.
- Photo captions use the question text instead of tags.

### Report (app and laptop produce identical documents)
- Uses the template's real cover page and disclaimer page.
- Footer now shows the true page total (template had a fixed "of 36").
- Body text 10 pt (template Normal style was 12 pt).
- New Fuel chapter; Boiler Inspection Findings chapter; efficiency O2 target table.
- A boiler is listed under a section only if it was inspected there.
- Photos: 4 per page (2×2), grouped per boiler, evenly spaced, never split across boilers; orientation corrected; resized to 1600 px so reports stay emailable.

### Laptop generator
- Rewritten to share the app's logic; only needs **Pillow** (no python-docx / lxml).
- Reads ZIPs from all app versions.
