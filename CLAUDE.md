# CLAUDE.md: academic website of Ha Anh Vu

Read this before changing anything. It records how the site works and the decisions the owner has made, so every session keeps the site consistent.

## Working with the owner

- The owner writes in Vietnamese: answer in Vietnamese. Website text stays in **English**; the Excel guide sheet is in **Vietnamese**.
- The owner publishes with **GitHub Desktop** and pushes personally. Never push, never force-push, never open PRs unless asked.
- The owner's local checkout is `E:\Research\anhvh-fpt.github.io` on branch `main` (this is what GitHub Desktop shows; it moved there from `E:\Coding\Resume`, which no longer exists). Commit there, or, if you work in a worktree or another branch, commit there and bring the work into that folder with a fast-forward:
  `git -C "E:/Research/anhvh-fpt.github.io" merge --ff-only <your-branch>`, and tell the owner to press **Push origin**.
- If that fails with `unable to unlink ... data/database.xlsx`, the workbook is open in Excel (Windows locks it). Ask the owner to close Excel. The failed merge leaves the other files half-applied: check they equal your commit, `git checkout --` them, then merge again.
- Keep changes minimal and preview them before saying they are done (see Verify).

## What the site is

Static site on GitHub Pages (repo `AnhVH-FPT/anhvh-fpt.github.io`, branch `main`, root folder). Plain HTML, CSS and JavaScript: **no build step, no npm**. All paths are relative. `_config.yml` keeps `admin/`, `tools/` and the docs off the published site.

| File | Role |
|---|---|
| `index.html` | Profile page: About, Research interests, Supervision (thesis projects), Publications, Experience, Join section, Contact |
| `research.html` | "Join my research group" page: how to join, topic list, where to publish, contact |
| `assets/js/main.js` | index.html: reads Publications, Projects, Directions, Topics; theme toggle; hero animation |
| `assets/js/research.js` | research.html: topics, venue tables, filters, pagination, apply/copy buttons |
| `assets/css/style.css` | Shared design tokens (light/dark) and index layout |
| `assets/css/research.css` | research.html only |
| `assets/img/favicon.svg` | Neural-network icon used by both pages |
| `data/database.xlsx` | **The single database of the whole site** |
| `tools/build_guide.py` | Regenerates the "Hướng dẫn" guide sheet inside the workbook |

Excel is read in the browser with SheetJS **0.18.5 from cdnjs** (`fetch('data/database.xlsx?v=' + Date.now())`, so new data shows without cache problems).

## The database: data/database.xlsx

Excel is the single source of truth: content that lives in the workbook must never be hard-coded in HTML/JS. (index.html keeps static Research-interest cards only as a fallback if the workbook fails to load.) Header in row 1, one record per row, row order does not matter except where noted. Sheets, in this order:

1. **Hướng dẫn**: the only guide sheet, Vietnamese, always first. Never edit it by hand: edit the texts in `tools/build_guide.py` and run `python tools/build_guide.py` (it deletes old guide sheets so there is exactly one). Whenever you change a sheet's columns or how the site reads them, update the guide in the same change.
2. **Page Text** and **Home Items**: the texts of the two pages (owner asked to keep them in few sheets).
   - **Page Text**: `Page` (Home / Research, dropdown), `Field`, `Value`. Fixed fields, one row each. Home (main.js `renderHome`): Avatar (a picture path inside the repo, e.g. `assets/img/avatar.png`), Kicker, Name, Headline, Affiliation, About (blank line = new paragraph), Intro: Research interests / Thesis projects / Publications / Experience / Contact, Courses heading, Join text. Research (research.js `renderPage`): Kicker, Title, Lead, Step 1–3 title/text, Gain heading/list, Requirements heading/list, Own idea heading/text, Intro: Research topics / Where to publish / Contact, Contact note; Research values accept `**bold**` and `[text](link)`, "… list" fields are one item per line. Each field maps to an element with `data-db="<field in lower case>"` on its page.
   - **Home Items**: `Section` (Education / Experience / Courses / Honors / Service, dropdown), `Title`, `Organization`, `Dates`, `Description`; row order inside a Section = order on the page. Education uses Title (degree), Organization (school), Dates (dates, GPA); Experience all four; Courses Title; Honors and Service Title + Description.
   - Both pages keep their current texts in the `data-db` elements as a fallback: an element is replaced only when the workbook has content for it. A tiny script in `<head>` adds `dbwait` (hides `[data-db]`, max 3 s) so old text does not flash; the page JS removes it after rendering or on a load error. The HTML fallback may go stale after edits; that is fine.
3. **Publications**: Type (Journal / Conference / Thesis / Preprint; an in-submission paper keeps its target kind), Status (Published / Accepted / In Submission / Preprint / Thesis; "By type" puts every In Submission paper in its own "In submission" group S1… and preprints in "Preprints" P1…; empty groups are hidden; for Type = Preprint the site shows the venue as "arXiv preprint arXiv:ID" linked to arxiv.org/abs, taking the ID from Venue, PDF or DOI; the database manager has an arXiv box that fills Venue/PDF/DOI and fetches title/authors/date via `GET /api/arxiv?id=`), Date, Title, Authors, Venue, DOI, PDF, Code, Link ID (sorting and venue grouping in main.js; see README).
4. **Projects**: Semester, Title, Members, Thesis Status, Paper Status, Paper Venue, Link ID.
5. **Directions**: Code (CV, MI, XAI, EDGE, VLM), Direction (display name), Focus (shown as "Current trends"), Summary (one line on the home page cards). Row order = order of cards, filter buttons and topic groups.
6. **Topics**: Direction (a code), Status (`Open` / `Taken`), Vietnamese Title, English Title, Dataset, Research Description, Research Expectation.
   - **No Code column.** The site numbers topics: direction code + position among that direction's rows, top to bottom (`MI-01` is the first MI row). Deleting or moving a row renumbers the topics below it, so the guide tells the owner to add rows at the end and mark taken topics `Taken` instead of deleting.
   - **Research Description / Research Expectation use a strict template.** A field starts at the beginning of a line with its label and a colon. Description: `Context`, `Problem formulation`, `Gap`, `Proposed approach`, `Implementation steps`, `Key knowledge to learn`. Expectation: `Deliverables`, `Expected results`, `Baselines`, `Metrics`, `Experimental protocol`, `Minimum success criteria`. Labels are case-insensitive; lines `(1) ...` become numbered lists; the site shows only these fields, always in this order, skips empty ones, and logs lines outside the template to the browser console. `Context` is 3+ sentences (setting, gap, what the student will do; rewritten for all 249 topics) and is the teaser on the topic card, clamped to 3 lines (4 on phones). Closed topic cards all have the same height (title 2 lines, Vietnamese title 1, teaser 3, tags 1 row, spare space above the tags); opened cards show everything. Do not split these into separate columns: the owner rejected that (too many columns).
   - Dataset: `• Name` line followed by its URL line, repeated.
7. **Conferences**: Conference, Full Name, Directions (codes, comma separated, or `All`; blank = All), CORE Rank (A*, A, B, C, Not ranked), Proceedings, Submission Deadline, Held In, Hybrid (Yes/No, shown on the site as a "Hybrid" column and filter), Website, Registration Fee (one line per rate: `Regular: ≈USD 900` / `Student: ≈USD 600`; a single Regular line if there is no student rate; ≈ = approximate, from recent editions). Deadline/Held In are English month words ("Mid December – January"); the site parses months for filters and sorting.
8. **Journals**: Journal, Publisher, Rank (Q1–Q4), Directions, Link, Time to Publish ("4–8 months": estimated months from submission to online publication; the site filters by the upper number and sorts by it with "Fastest first"), Publication Fee (one line per option: `Free (subscription)` for a hybrid no-fee route, `Free` alone for diamond OA, `Open access: USD 3,270` for the APC; hybrid journals have both lines, paid-only journals only the Open access line; APCs from the 2026 Elsevier/Springer Nature/Wiley/IEEE price lists, ≈ = estimate; the owner prefers free journals). Journal fee filter: Free option available / Paid only / Open access up to ≈ USD 2,000 or 3,000; sorts: "Free first, then lowest APC" and "Lowest open-access fee". (Directions and Time to Publish were first filled as estimates by Claude; the owner may correct them.)
9. **Venue Notes**: Applies To (Conferences / Journals / All), Note.

research.js reads each sheet **by name** and never reads the guide sheet (its example tables reuse the same column names).

Editing the workbook from code: use openpyxl, keep data validations (Direction, Status, Rank, Hybrid lists), the grey conditional format for `Taken` rows, header comments (template reminder on the two Topics text columns), freeze panes and filters. `delete_cols`/`insert_cols` do not move validations, conditional formats or comments: set them again. After any change, verify nothing was lost (compare cell text with the previous commit).

## Behaviour the owner asked for (keep it)

- **Sorting is done by the site**, whatever the row order: Journals Q1→Q4 then name (or, with Sort = Fastest first, by Time to Publish upper then lower bound, then rank, then name); Conferences by deadline month (early/mid/late), then CORE rank, then name; both tables also offer Sort = Lowest fee (Regular fee converted to USD with rough rates in research.js, then Student fee; journals Free first); Topics by Directions order then row order; Publications/Projects as in main.js.
- **research.html**: direction chips (counts and lists show Open topics only) + a "Taken" chip, shown when some topic is Taken, that is the only place taken topics appear (codes unchanged) + search; 5/10/15/20/25 per page with numbered pages (remembered in localStorage); `research.html#MI-07` opens a topic on the right page; `research.html?dir=MI` preselects a direction. Venues ("venue finder"): Conferences/Journals tabs with counts, a search box, a panel of filter selects whose options show live counts (conferences: Rank, Direction, Deadline incl. "Next 3 months", Held in, Hybrid (Yes / No), Regular fee cap; journals: Rank, Direction, Time to publish, Publication fee, Publisher), removable chips for active filters + "Clear all", "Showing x–y of n", multi-level sort (Sort select or clickable column headers; the newest choice is level 1, up to 3 levels shown as ↑¹ ↑² and in a "Sorted by … Reset sort" line; clicking the level-1 header again toggles ↑/↓), per-page select, paging, notes under the table. Each tab keeps its own filters and sort. On phones (≤760px) table rows become cards. No "Start preparing" column: the owner removed it.
- Removed columns (owner request, not used by the site): Conferences "Presentation Format", Directions "Topics" count, Conferences "Start Preparing".
- Venue lists: about 200 conferences (Scopus-indexed proceedings only) and 200 journals, none on the owner's predatory journal/publisher lists (copies in `admin/predatory/`; originals in `E:\Research\Topic list\`): check new venues against both lists. The journal tab is called "Journals" (not "Free journals"). CORE Rank follows ICORE2026 (export at portal.core.edu.au; anything without an A*–C rank, including national ranks and multiconferences, is `Not ranked`). Website: the official site, or the conference's dblp page when the site is year-specific or dead (dblp blocks scripts; check it in a browser). Fees without `≈` were read from the registration page (early, member rate); `≈` means estimated. Beware WASET look-alike conferences.
- Topic intro text: individual or team of up to 4, at least 3 months, students prepare their own training hardware.
- Apply button opens a prefilled email; "Copy application text" and a note with email + Zalo cover students whose email app does not open.
- **Header**: on index.html "Join us →" is the last item, styled as the blue `.nav-cta` button; on research.html "← Home" uses the same `.nav-cta` style. Join links open research.html in the **same tab**.
- **Home page Research interests**: compact cards built from the Directions sheet (code, name, Summary, open-topic count), each linking to `research.html?dir=CODE#topics`.
- **Contact**: the same block on index.html and research.html (`.clist` in style.css, copy logic in `assets/js/contact.js`): Email (tiny copy icon), Phone/Zalo `0366765835` (number links to Zalo, tiny copy icon, small Zalo logo button), Research community (owner's Zalo group `https://zalo.me/g/epqegx612`, people icon), Office (AI Department, FPT University, HCMC), Profiles (Google Scholar, ORCID, GitHub). research.html keeps the extra note about the email button below it. The "Your own idea?" panel has two buttons: Email me and Contact (jumps to #contact); no plain-text email line.
- Links: GitHub `https://github.com/AnhVH-FPT`. **No LinkedIn anywhere.**
- Favicon: `assets/img/favicon.svg` (no emoji icon).

## Database manager (admin/, local only)

The owner edits the workbook through a local web app instead of Excel. It is **in the repository** under `admin/` (so a fresh clone on another computer works) but **not published**: `_config.yml` excludes `admin/`, `tools/`, `README.md`, `CLAUDE.md` from GitHub Pages (there is no `.nojekyll`; Jekyll only copies files since no page has front matter, so never add front matter to the HTML). `admin/backups/` is git-ignored. Predatory lists live in `admin/predatory/` (predatory-journals.xlsx, predatory-publishers.xlsx). The .bat checks for Python and installs `openpyxl` from `admin/requirements.txt` on first run.

- `server.py` (stdlib http.server + openpyxl, localhost:8765, `--db` / `--port` / `--no-browser` for testing on a copy): `GET /api/db` returns every sheet (headers, rows as display text, dropdown lists read from the data validations); `POST /api/op` does `update` / `insert` / `delete` / `move` / `set` (bulk) / `rename_code` (Directions code renamed in Topics, Conferences, Journals) with an optimistic version check (file mtime+size), rewrites the record rows in place (new rows copy row-2 styles; hyperlinks re-set on link columns), restores the Topics Direction dropdown, backs up to `admin/backups/` (60 kept) and saves atomically. `POST /api/undo` restores the newest backup. Refuses to write while Excel has the file open. `/site/...` serves the website with the current workbook. Reads the predatory lists from `admin/predatory/` for warnings.
- `index.html`, `app.js`, `app.css`: Vietnamese UI. Tables: quick filter selects, "Lọc nâng cao" (AND conditions: chứa / không chứa / là / không là / trống / không trống / ≥ / ≤), "Sắp xếp" up to 3 levels (or header click, Shift+click adds a level), column picker; values sort by meaning via `sortKey` (dropdown order, dates, semesters, months, fees in USD), blanks last; per-sheet view settings in localStorage. Topic Dataset is edited as rows of name (required) + link (optional), stored as `• Name` + URL lines. Per-sheet settings in `CFG` (table columns, filters, required fields, hints, quick-insert texts, topic template labels); `checkRow` holds the data checks shown in the form, the table dots and the Overview. **When a sheet's columns or rules change, update `CFG`/`checkRow` too** (and the guide).
- The sidebar (`NAV` in app.js) is grouped by web page (index.html, research.html), each with "Nội dung trang" and "Dữ liệu"; entries with `lock` show one sheet pre-filtered (Page Text by Page, Home Items by Section: Education, Experience, Courses, Honors, Service), number rows 1..n inside the view, default new rows to the locked value and move rows among the same Section (`orderWithin`); addresses look like `#Home Items@Education/2`; Page Text is `fixed` (no add/delete, Page and Field read-only, a hint per field); Home Items shows only the fields its Section uses (`ITEM_FIELDS`); the Home/Avatar row has a picture picker: `GET /api/images` lists pictures in the repo (not admin/, tools/, data/), `POST /api/upload-image` copies an uploaded picture into `assets/img/` (ASCII file name, never overwrites). `Open database manager.bat` starts it. The guide sheet tells the owner about it.

## Conventions

- After changing a CSS or JS file, bump its `?v=N` in every HTML file that loads it, so browsers fetch the new file.
- Colors are tokens in `:root` / `:root[data-theme=dark]` in style.css; both themes must work; layouts must work at phone width.
- External scripts only from cdnjs with an exact version; no new dependencies without a reason. Exception (owner's request): the GoatCounter visit counter `https://gc.zgo.at/count.js` at the end of index.html and research.html (account `haanhvu010799`, no cookies, localhost not counted). research.js also sends GoatCounter events `open/<code>`, `apply/<code>` and `copy-application/<code>`.
- Do not put private data in the workbook: it is public once pushed.
- Commit messages: imperative summary line, short body explaining why.

## Verify before reporting done

1. Serve the folder: `python -m http.server 8770` (or the `site` entry in `.claude/launch.json`; port 8765 belongs to the owner's database manager) and open `http://localhost:8770/index.html` and `/research.html`. Opening the HTML by double-click cannot read the workbook.
   Serve the folder that holds the file you want to see. The owner edits `E:\Research\anhvh-fpt.github.io\data\database.xlsx`; a server started inside a worktree shows the worktree's copy, not those edits. To preview the owner's edits: `python -m http.server 8766 --directory E:/Research/anhvh-fpt.github.io` (the `site-main` entry).
2. Check in the page: 249-ish topics load, direction chips and counts, a topic opens with all template fields, venue filters and paging, Publications and Projects still render on index.html, no console errors.
3. If you changed the workbook, also check it opens with the expected sheet list and only one guide sheet.

## Safety nets

- `main` is the only branch; the old backup tag and branch were deleted at the owner's request. Earlier versions remain in the git history.
- On index.html the phone number links to `https://zalo.me/0366765835` (opens a Zalo chat), not `tel:`.
- Rolling back is a decision for the owner; explain the commands (README) rather than running a hard reset or force push.
