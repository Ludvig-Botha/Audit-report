# Changelog

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
