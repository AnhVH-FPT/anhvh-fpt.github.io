# Academic Personal Website: Ha Anh Vu

Static site (plain HTML, CSS, JavaScript). No build step, no dependencies. Works on GitHub Pages as is.

## Project structure

```
.
├── index.html, research.html   # The two pages
├── data/database.xlsx          # The site database (edit with the database manager or Excel)
├── assets/                     # css/, js/, img/
├── admin/                      # Database manager (local tool, NOT published on the website)
│   ├── Open database manager.bat   # double-click to start it on Windows
│   ├── server.py, index.html, app.js, app.css, requirements.txt
│   └── predatory/              # predatory journal / publisher lists used for warnings
├── tools/build_guide.py        # regenerates the "Hướng dẫn" sheet in the workbook (not published)
├── _config.yml                 # GitHub Pages: publish the site, leave out admin/, tools/, README.md, CLAUDE.md
├── README.md, CLAUDE.md        # documentation (not published)
```

Everything needed to edit the site is in this repository. Backups made by the database manager (`admin/backups/`) stay on each computer and are not committed.

## Work on a new computer

1. Install **GitHub Desktop** (sign in with the GitHub account) and **Python 3** from python.org (tick "Add python.exe to PATH" during setup).
2. In GitHub Desktop: **File → Clone repository → `AnhVH-FPT/anhvh-fpt.github.io`**, choose a folder (for example `E:\Research\anhvh-fpt.github.io`).
3. Double-click `admin\Open database manager.bat` in that folder. The first time it installs `openpyxl` automatically, then opens `http://localhost:8765/`.
4. Edit, then in GitHub Desktop: **Commit to main → Push origin**.
5. When you switch between computers, press **Fetch origin / Pull origin** in GitHub Desktop first, so you edit the latest database.

## Run locally

The site reads `data/database.xlsx`. If you double-click `index.html`, browsers block reading that file, so a bar appears at the bottom of the page: click **Choose database.xlsx** and pick the file (it is remembered; choose it again after editing the Excel file). For automatic loading, serve the folder with a local web server instead:

```bash
python3 -m http.server 8000   # then visit http://localhost:8000
```

## Deploy to GitHub Pages

1. Create a GitHub repository. For a personal site, name it `<your-username>.github.io`.
2. Upload every file in this folder to the repository root (keep `index.html` at the top level):
   ```bash
   git init && git add . && git commit -m "Add academic website"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-username>.github.io.git
   git push -u origin main
   ```
3. On GitHub, open **Settings, then Pages**. Under **Build and deployment**, choose **Deploy from a branch**, select branch `main` and folder `/ (root)`, then save.
4. After 1 to 2 minutes the site is live at `https://<your-username>.github.io/`. If the repository has another name, the address is `https://<your-username>.github.io/<repo-name>/` (all paths in this project are relative, so it works either way).

## Customize

- **Avatar**: save a square photo (at least 480x480 px) as `assets/img/avatar.png`. The initials placeholder disappears automatically.
- **Publications and projects**: edit `data/database.xlsx` (see the next section). No HTML or JavaScript changes are needed.
- **Colors and fonts**: change the CSS variables in `:root` and `:root[data-theme=dark]` at the top of `assets/css/style.css`.
- **Prospective students section**: the text in `#join` is a draft. Edit the topics, benefits and requirements to match your group.
- **Custom domain** (optional): add a `CNAME` file containing your domain, then set the domain in Settings, then Pages.

## Edit the home page texts

The fixed texts of both pages are in the sheet **Page Text** (column Page = Home or Research): avatar, name, headline, About, the one-line introductions under each heading, the research page header, steps and boxes. The home-page lists (Education, Experience, courses, honors, academic service) are in **Home Items** (column Section). Edit them in the database manager (group "Chữ trên trang web"); for the avatar choose a picture already in the repository or upload one (it is copied into `assets/img/`). Commit the workbook and any new picture together.

## Update publications and projects (Excel)

**Easiest way: the database manager.** Double-click `admin\Open database manager.bat` (it is in the repository but not published on the website, see `_config.yml`). It starts a small local server (`python admin/server.py`, needs `openpyxl`) and opens `http://localhost:8765/`: every sheet as a searchable, filterable table, a form to add/edit/duplicate/delete rows, data checks (missing fields, topic template, unknown direction codes, predatory journals), and each Save writes straight into `data/database.xlsx` (Excel must be closed). Backups of the previous file are kept in `admin/backups/` and "Undo" restores them. "Xem website" previews the site with the current workbook. Then commit and push with GitHub Desktop as usual.

To edit in Excel instead:

1. Open `data/database.xlsx`. The first sheet, `Hướng dẫn`, explains every column (in Vietnamese).
2. **Projects sheet** (undergraduate thesis projects): `Semester` (for example `Fall 2026`; `FA26` also works), `Title`, `Members` (separated by `;`), `Thesis Status`, `Paper Status`, `Paper Venue`, `Link ID`.
   - `Thesis Status`: **Preparing** (topic just chosen), **In progress** (work ongoing), **Defended** (finished).
   - `Paper Status`: **None** (no badge), **Submitted** (badge "Paper submitted"), **Accepted** (badge "Accepted, see C1").
   - Semesters are sorted automatically: newest year first, and inside a year Spring, then Summer, then Fall. The 2 most recent semesters are expanded (the latest has a "Current" badge); others are collapsed until clicked. In the database manager, Semester is a season dropdown plus a year box.
3. **Publications sheet**: `Type` (`Journal`, `Conference`, `Thesis` or `Preprint`; a paper under review keeps its target kind), `Status` (`Published`, `Accepted`, `In Submission`, `Preprint`, `Thesis`; In Submission papers and preprints get their own groups, S1… and P1…), `Date` (any format works: `dd-mm-yyyy`, `mm-yyyy` or just `yyyy`, or a real Excel date; the year is required and a paper without a year is hidden), `Title`, `Authors` (comma separated), `Venue`, `DOI` (a DOI or full URL), `PDF`, `Code`, `Link ID`.
   - Order inside a group: papers of the same conference or journal are kept together (series such as RIVF 2025 and RIVF 2026 count as one). In **By year**, Published (and Thesis) come first, then Accepted, then In Submission; inside each, venues are ordered by their newest paper and papers by date. In **By type** the same venue grouping applies, without the status ordering. You can enter rows in any order and write a venue in different ways ("RIVF 2026", "IEEE RIVF", or the full name with "(RIVF 2026)"): the site matches them and sorts everything itself. Papers with exactly the same date and venue are ordered alphabetically by title.
   - Papers are always sorted newest first. **By year** view expands the 3 most recent years and collapses the rest; **By type** view expands every group (each can still be collapsed).
4. **Linking a thesis to its paper:** put the same `Link ID` (any short code, for example `AC1`) on the project row and on the paper row. When the project's `Paper Status` is `Accepted`, its badge reads "Accepted, see C1" (the paper's number in Publications) and clicking it opens and scrolls to that paper. Titles do not need to match. Add the accepted paper itself as a normal Publications row.
5. Save, then replace `data/database.xlsx` in your repository (commit and push). The site shows the new data after GitHub Pages redeploys, usually within a minute or two.

The Excel file is public once deployed, so do not put private information in it. The library that reads Excel (SheetJS 0.18.5) is loaded from cdnjs.

## Research topics page (research.html)

"Join my research group" (hero button, the "Join us" menu item and the Join section) opens `research.html`. It reads the same workbook as the main page, `data/database.xlsx`, so **the whole site has one database**. Edit the workbook in Excel, save, commit and push (GitHub Desktop): the site shows the change after GitHub Pages redeploys (1 to 2 minutes). No code changes are needed. The workbook's first sheet, **Hướng dẫn**, explains every sheet in detail (in Vietnamese), including how the site sorts each list.

| Sheet | Used by | One row is |
|---|---|---|
| Publications | index.html | a paper |
| Projects | index.html | a supervised thesis |
| Directions | research.html | a research direction: `Code` (CV, MI, XAI, EDGE, VLM), `Direction`, `Focus` |
| Topics | research.html | a topic (its code, e.g. `MI-07`, is numbered by the site from the direction and the row order: add new topics at the end, mark taken ones `Taken` instead of deleting): `Direction` (a code), `Status` (`Open` or `Taken`), titles, dataset, `Research Description` and `Research Expectation` written in a fixed template ("Context: ...", "Gap: ...", "Expected results:" then "(1) ..." lines; see section 6a of the Hướng dẫn sheet). Only template fields are shown, in a fixed order; lines outside the template are listed in the browser console |
| Conferences | research.html | a conference: `Directions` (codes or `All`), `CORE Rank`, `Submission Deadline`, `Held In` (English month names), `Hybrid`, `Website`, `Registration Fee` (`Regular: ≈USD 900` and `Student: ≈USD 600` on separate lines) |
| Journals | research.html | a no-APC journal: `Rank` (Q1 to Q4), `Directions` (codes or `All`), `Link`, `Time to Publish` (e.g. `4–8 months`; estimates, check them on each journal's page), `Publication Fee` (`Free` or an amount) |
| Venue Notes | research.html | a note shown under the Conferences or Journals table |

Rules: keep sheet names and the header row (row 1) as they are; column order does not matter. Conferences are sorted by deadline on the site, so their row order does not matter either. The Journals `Directions` column was first filled automatically from journal names, so review it.

On the page, topics can be filtered by direction and searched; venues can be filtered by rank, direction, deadline month (including "Next 3 months") and the month the conference is held. Topics and venues are paged (5 to 25 per page). `research.html#MI-01` opens a topic directly. To regenerate the Hướng dẫn sheet after changing its texts, run `python tools/build_guide.py` (needs openpyxl). Project rules for AI assistants are in `CLAUDE.md`.

After changing `assets/js/research.js` or `assets/css/research.css`, bump the `?v=` number in `research.html` so browsers fetch the new file.

**Earlier versions:** every past version is in the git history (GitHub Desktop, History tab). To undo a change, right-click its commit, choose "Revert changes in commit", then push.

## Troubleshooting: "Could not load data/database.xlsx"

- Opened `index.html` by double-click? Browsers block this. Use `python3 -m http.server` (or VS Code Live Server), or view the deployed GitHub Pages site.
- `data/database.xlsx` must be committed to the repository inside the `data` folder with exactly this name (case-sensitive).
- The page needs internet access to load the Excel reader (cdnjs); ad blockers may block it.
