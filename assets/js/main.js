/* Academic personal website: Ha Anh Vu */
(() => {
  'use strict';
  const root = document.documentElement;
  const $ = (s, c = document) => c.querySelector(s);
  const $$ = (s, c = document) => [...c.querySelectorAll(s)];
  const store = {
    get: k => { try { return localStorage.getItem(k); } catch { return null; } },
    set: (k, v) => { try { localStorage.setItem(k, v); } catch { /* storage unavailable */ } }
  };

  /* ---------- 1. CONTENT FROM EXCEL (data/database.xlsx) ----------
     Sheets "Publications" and "Projects" are read in the browser with SheetJS; nothing here needs editing
     except ME, the name shown in bold in author lists. */
  const ME = 'Ha Anh Vu';
  const DB_FILE = 'data/database.xlsx';
  // Cells typed or pasted in Excel may use Windows line breaks (CR LF): turn every line break into LF
  const txt = v => String(v == null ? '' : v).replace(/\r\n?/g, '\n').trim();
  const row = r => {                                                   // a missing column simply reads as an empty cell
    const o = Object.fromEntries(Object.entries(r).map(([k, v]) => [k.trim().toLowerCase(), txt(v)]));
    return new Proxy(o, { get: (t, k) => (k in t ? t[k] : '') });
  };
  const esc = s => txt(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const chip = (href, label) => href ? `<a class="chip" href="${esc(href)}" target="_blank" rel="noopener">${label}</a>` : '';
  const doiUrl = d => d && (/^https?:/i.test(d) ? d : 'https://doi.org/' + d);

  const GROUPS = [
    { key: 'journal', label: 'Journal articles', prefix: 'J' },
    { key: 'conference', label: 'Conference proceedings', prefix: 'C' },
    { key: 'thesis', label: 'Thesis', prefix: 'T' },
    { key: 'submission', label: 'In submission', prefix: 'S' },
    { key: 'preprint', label: 'Preprints', prefix: 'P' }
  ];
  const STATUS = { published: 'Published', accepted: 'Accepted', submission: 'In Submission', preprint: 'Preprint', thesis: 'Thesis' };
  // Type = kind of paper (an in-submission paper keeps its target kind: Journal or Conference); Status = where it stands
  const typeOf = s => { s = txt(s).toLowerCase(); return s[0] === 'j' ? 'journal' : s[0] === 'c' ? 'conference' : s[0] === 't' ? 'thesis' : s[0] === 'p' ? 'preprint' : 'other'; };
  const statusOf = s => { s = txt(s).toLowerCase(); return s.includes('accept') ? 'accepted' : s.includes('publish') ? 'published' : s.includes('thesis') ? 'thesis' : s.includes('preprint') ? 'preprint' : 'submission'; };
  // "By type" group: anything In Submission goes to "In submission"; preprints (and the old "Submission / Preprint" type) to "Preprints"
  const groupOf = p => p.status === 'submission' ? 'submission' : p.type === 'preprint' || p.type === 'other' || p.status === 'preprint' ? 'preprint' : p.type;

  const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const TYPE_LABEL = { journal: 'Journal', conference: 'Conference', thesis: 'Thesis', preprint: 'Preprint', other: 'Paper' };

  // <semester-helpers>
  const SEASONS = { sp: 'Spring', spring: 'Spring', su: 'Summer', summer: 'Summer', fa: 'Fall', fall: 'Fall', autumn: 'Fall' };
  const RANK = { Spring: 1, Summer: 2, Fall: 3 };
  const monthOf = v => { v = txt(v).toLowerCase(); const n = parseInt(v, 10); return n >= 1 && n <= 12 ? n : MONTHS.indexOf(v.slice(0, 3)) + 1; };
  // "Fall 2026", "2026 fall", "FA26" ... -> label + sort key (year first; inside a year Fall > Summer > Spring)
  const semOf = s => {
    s = txt(s);
    const w = s.match(/spring|summer|fall|autumn/i) || s.match(/\b(sp|su|fa)(?=\s*-?\d{2})/i);
    const y = s.match(/\d{4}/) || s.match(/\d{2}/);
    if (!w || !y) return { label: s || 'Unspecified semester', key: -1 };
    const season = SEASONS[w[0].toLowerCase()], year = y[0].length === 2 ? 2000 + Number(y[0]) : Number(y[0]);
    return { label: season + ' ' + year, key: year * 10 + RANK[season] };
  };
  // </semester-helpers>

  /* ----- Publications: newest first (Year, Month, Day). By year: 3 latest years open; by type: all open ----- */
  // Preprints: the venue is always the arXiv record, found in Venue, PDF or DOI (e.g. 2501.01234, arxiv.org/abs/2501.01234v2, 10.48550/arXiv.2501.01234)
  const arxivId = p => { const m = [p.venue, p.pdf, p.doi].map(txt).join(' ').match(/arxiv[^0-9]{0,12}(\d{4}\.\d{4,5}(?:v\d+)?)/i) || (p.type === 'preprint' && txt(p.venue).match(/^(\d{4}\.\d{4,5}(?:v\d+)?)$/)); return m ? m[1] : ''; };
  const venueHtml = p => { const id = p.type === 'preprint' && arxivId(p);
    return id ? `<em>arXiv preprint</em> <a href="https://arxiv.org/abs/${esc(id)}" target="_blank" rel="noopener">arXiv:${esc(id)}</a>` : `<em>${esc(p.venue)}</em>`; };
  const pubCard = p => `<article class="pub" id="pub-${p.id}">
    <span class="pid">[${p.id}]</span>
    <div>
      <h4>${esc(p.title)}</h4>
      <p class="au">${esc(p.authors).split(ME).join(`<strong>${ME}</strong>`)}</p>
      <p class="ve">${venueHtml(p)}, ${fmtDate(p)}</p>
      <div class="pmeta"><span class="ptype">${TYPE_LABEL[p.type]}</span><span class="badge ${p.status}">${STATUS[p.status]}</span>${chip(doiUrl(p.doi), 'DOI')}${chip(p.pdf, 'PDF')}${chip(p.code, 'Code')}</div>
    </div></article>`;

  // Venue matching that tolerates messy input: "RIVF 2026", "IEEE RIVF" and the full conference name count as one venue
  const VSTOP = new Set('international intl conference conf proceedings proceeding proc of on the and for in journal symposium workshop transactions ieee acm springer nature elsevier press annual'.split(' '));
  const words = s => s.toLowerCase().replace(/\b(19|20)\d{2}\b/g, ' ').replace(/\b\d+(st|nd|rd|th)\b/g, ' ').split(/[^a-z0-9]+/).filter(w => w.length > 1 && !VSTOP.has(w));
  const venueSets = v => {
    v = txt(v);
    const m = v.match(/\(([^)]*\d{4}[^)]*)\)/);                                   // "(RIVF 2026)" is the short name
    const full = new Set(words(v.replace(/\([^)]*\)/g, x => /\d{4}/.test(x) ? x : ' ')));   // drops "(under review)" etc.
    return { full, short: new Set(m ? words(m[1]) : full.size === 1 ? [...full] : []) };
  };
  const sub = (a, b) => a.size > 0 && [...a].every(w => b.has(w));
  const jac = (a, b) => { const n = [...a].filter(w => b.has(w)).length; return n / (a.size + b.size - n || 1); };
  const sameVenue = (a, b) => jac(a.full, b.full) >= 0.6 || sub(a.short, b.full) || sub(b.short, a.full);
  const clusterVenues = list => {
    const sets = list.map(p => venueSets(p.venue)), root = list.map((_, i) => i);
    const find = i => root[i] === i ? i : (root[i] = find(root[i]));
    for (let i = 0; i < list.length; i++) for (let j = i + 1; j < list.length; j++)
      if (sets[i].full.size && sets[j].full.size && sameVenue(sets[i], sets[j])) root[find(j)] = find(i);
    const label = new Map();                                                       // cluster name does not depend on row order
    list.forEach((p, i) => { const r = find(i), k = [...sets[i].full].sort().join(' '); if (!label.has(r) || k < label.get(r)) label.set(r, k); });
    list.forEach((p, i) => { p.vkey = sets[i].full.size ? label.get(find(i)) : 'zz ' + p.title.toLowerCase(); });
  };
  const STATUS_RANK = { published: 0, thesis: 0, accepted: 1, preprint: 2, submission: 3 };   // published first, in submission last
  // Order a list of papers: (optionally) by status, then keep the same venue together
  // (venues ordered by their newest paper), then newest first inside a venue.
  const arrange = (items, byStatus) => {
    const rank = p => byStatus ? STATUS_RANK[p.status] : 0, best = new Map();
    items.forEach(p => { const k = rank(p) + '|' + p.vkey; best.set(k, Math.max(best.get(k) || 0, p.dkey)); });
    return [...items].sort((a, b) => rank(a) - rank(b)
      || best.get(rank(b) + '|' + b.vkey) - best.get(rank(a) + '|' + a.vkey)
      || (a.vkey < b.vkey ? -1 : a.vkey > b.vkey ? 1 : 0)
      || b.dkey - a.dkey || a.title.toLowerCase().localeCompare(b.title.toLowerCase()) || a.i - b.i);
  };

  // <date-helpers>
  // Date column format: dd-mm-yyyy. The year is required (papers without a year are hidden); day and month are optional.
  // Also understands 03-2026, 2026, 2026-03-15, "15 Mar 2026" and real Excel date cells (serial numbers).
  const parseDate = v => {
    v = txt(v);
    if (/^\d+(\.\d+)?$/.test(v) && parseFloat(v) > 2100) {                          // Excel date cell
      const dt = new Date(Math.floor(parseFloat(v) - 25569) * 864e5);
      return { year: dt.getUTCFullYear(), month: dt.getUTCMonth() + 1, day: dt.getUTCDate() };
    }
    const t = v.replace(/(\d)(st|nd|rd|th)\b/gi, '$1').split(/[^0-9a-z]+/i).filter(Boolean);
    const yi = t.findIndex(x => /^(19|20)\d{2}$/.test(x));
    if (yi < 0) return { year: 0, month: 0, day: 0 };
    const rest = t.filter((_, i) => i !== yi), mi = rest.findIndex(x => /^[a-z]{3,}$/i.test(x) && MONTHS.includes(x.slice(0, 3).toLowerCase()));
    const nums = rest.filter(x => /^\d{1,2}$/.test(x)).map(Number);
    let month = 0, day = 0;
    if (mi >= 0) { month = MONTHS.indexOf(rest[mi].slice(0, 3).toLowerCase()) + 1; day = nums[0] || 0; }
    else if (yi === 0) { month = nums[0] || 0; day = nums[1] || 0; }                 // 2026-03-15
    else if (nums.length > 1) {                                                       // 15-03-2026 (03-15-2026 is also understood)
      day = nums[0]; month = nums[1];
      if (month > 12 && day <= 12) [day, month] = [month, day];
    }
    else if (nums.length === 1) month = nums[0];                                      // 03-2026
    if (month < 1 || month > 12) month = day = 0;
    if (day < 1 || day > 31) day = 0;
    return { year: +t[yi], month, day };
  };
  const pad = n => String(n).padStart(2, '0');
  const fmtDate = p => [p.day && pad(p.day), p.month && pad(p.month), p.year].filter(Boolean).join('-');
  // A real Excel date cell: show only the parts its number format shows ("mmm-yy" -> month and year, "yyyy" -> year)
  const dateCellText = c => {
    const dt = new Date(Math.floor(c.v - 25569) * 864e5), z = String(c.z || ''), full = !z || /general/i.test(z);
    return [full || /d/i.test(z) ? pad(dt.getUTCDate()) : '', full || /m/i.test(z) ? pad(dt.getUTCMonth() + 1) : '', dt.getUTCFullYear()].filter(Boolean).join('-');
  };
  // </date-helpers>

  let PUBS = [], pubView = 'year';
  const drawPubs = () => {
    const sec = (title, items, open) => `<details class="sem"${open ? ' open' : ''}><summary><span>${title}</span><em>${items.length}</em></summary><div class="plist">${
      items.map(pubCard).join('') || '<p class="empty">Nothing in this group yet.</p>'}</div></details>`;
    const html = pubView === 'type'
      ? GROUPS.map(g => [g, PUBS.filter(p => p.group === g.key)]).filter(([, l]) => l.length).map(([g, l]) => sec(g.label, arrange(l, false), true)).join('')   // empty groups hidden
      : [...new Set(PUBS.map(p => p.year))].map((y, i) => sec(y || 'Undated', arrange(PUBS.filter(p => p.year === y), true), i < 3)).join('');
    $('#pubs').innerHTML = html || '<p class="empty">No publications yet.</p>';
    $$('.seg button').forEach(b => b.classList.toggle('on', b.dataset.view === pubView));
  };
  const renderPubs = rows => {
    const all = rows.map(row).filter(r => r.title).map((r, i) => ({ ...r, i, type: typeOf(r.type), status: statusOf(r.status),
      ...parseDate(r.date || [r.day, r.month, r.year].filter(Boolean).join('-')),    // the old Year/Month/Day columns still work
      link: (r['link id'] || '').toLowerCase() }));
    all.filter(p => !p.year).forEach(p => console.warn('Hidden (no year in the Date column): ' + p.title));
    PUBS = all.filter(p => p.year);
    PUBS.sort((a, b) => b.year - a.year || b.month - a.month || b.day - a.day || a.i - b.i);
    PUBS.forEach(p => { p.dkey = p.year * 10000 + p.month * 100 + p.day; p.group = groupOf(p); });
    clusterVenues(PUBS);
    GROUPS.forEach(g => arrange(PUBS.filter(p => p.group === g.key), false).forEach((p, n) => { p.id = g.prefix + (n + 1); }));
    drawPubs();
  };
  $$('.seg button').forEach(b => b.addEventListener('click', () => { pubView = b.dataset.view; drawPubs(); }));

  /* ----- Undergraduate thesis projects: by semester, 2 newest expanded, older collapsed ----- */
  const projCard = r => {
    const ts = r['thesis status'].toLowerCase(), ps = r['paper status'].toLowerCase(), key = (r['link id'] || '').toLowerCase();
    const tcls = /defend|complete/.test(ts) ? 'published' : ts.includes('progress') ? 'defense' : 'thesis';   // green / blue / gray
    const members = r.members.split(/[;\n,]+/).map(s => s.trim()).filter(Boolean).map(esc).join(', ');
    let paper = '';
    if (ps.includes('submit')) paper = `<span class="badge submitted">Paper submitted${r['paper venue'] ? ' to ' + esc(r['paper venue']) : ''}</span>`;
    else if (/accept|publish/.test(ps)) {
      const pub = key && PUBS.find(p => p.link === key);                       // matched by Link ID, not by title
      if (!pub) console.warn('No publication with Link ID "' + r['link id'] + '" for project: ' + r.title);
      paper = pub ? `<a class="badge accepted" href="#pub-${pub.id}" data-go="${pub.id}">Accepted, see ${pub.id}</a>` : '<span class="badge accepted">Accepted</span>';
    }                                                                           // "None" or blank: no paper badge
    return `<article class="proj">
      <h3>${esc(r.title)}</h3>
      ${members ? `<p class="team"><strong>Team:</strong> ${members}</p>` : ''}
      <div class="pmeta">${r['thesis status'] ? `<span class="badge ${tcls}">${esc(r['thesis status'])}</span>` : ''}${paper}</div>
    </article>`;
  };
  const renderProjects = rows => {
    const groups = new Map();
    rows.map(row).filter(r => r.title).forEach(r => {
      const s = semOf(r.semester), g = groups.get(s.label) || { ...s, items: [] };
      g.items.push(r); groups.set(s.label, g);
    });
    // Newest year first; inside a year Spring, then Summer, then Fall (owner's choice). Unreadable semesters go last.
    const list = [...groups.values()].sort((a, b) => (a.key < 0) - (b.key < 0) || Math.floor(b.key / 10) - Math.floor(a.key / 10) || a.key - b.key);
    const newest = [...groups.values()].map(g => g.key).filter(k => k >= 0).sort((a, b) => b - a);   // the two latest semesters start open
    $('#projects').innerHTML = list.map(g => `<details class="sem"${newest.slice(0, 2).includes(g.key) ? ' open' : ''}>
      <summary><span>${esc(g.label)}</span>${g.key >= 0 && g.key === newest[0] ? '<span class="now">Current</span>' : ''}<em>${g.items.length}</em></summary>
      <div class="projects">${g.items.map(projCard).join('')}</div></details>`).join('') || '<p class="empty">No projects yet.</p>';
  };
  // Clicking "Accepted, see C1" opens the paper's group in Publications and scrolls to the paper
  $('#projects').addEventListener('click', e => {
    const a = e.target.closest('[data-go]'); if (!a) return;
    const card = document.getElementById('pub-' + a.dataset.go); if (!card) return;
    e.preventDefault();
    card.closest('details').open = true;
    card.scrollIntoView({ behavior: 'smooth', block: 'center' });
    card.classList.remove('flash'); void card.offsetWidth; card.classList.add('flash');
    setTimeout(() => card.classList.remove('flash'), 2600);
  });

  /* ----- Research interests: one card per row of the Directions sheet, linking to its topics on research.html ----- */
  const renderDirections = (dirRows, topicRows) => {
    const dirs = dirRows.map(row).filter(r => r.code && r.direction);
    if (!dirs.length) return;                                                  // keep the cards written in index.html
    const topics = topicRows.map(row).filter(r => r['english title']);
    const open = c => topics.filter(t => t.direction.toLowerCase() === c.toLowerCase() && !/taken|assigned|closed|full/i.test(t.status)).length;
    $('#dirgrid').innerHTML = dirs.map(d => {
      const n = open(d.code);
      return `<a class="ditem" href="research.html?dir=${encodeURIComponent(d.code)}#topics"><span class="dcode">${esc(d.code)}</span>
        <h3>${esc(d.direction)}</h3>${d.summary ? `<p>${esc(d.summary)}</p>` : ''}${n ? `<span class="dcount">${n} open topics &rarr;</span>` : ''}</a>`;
    }).join('');
  };

  /* ----- Home-page text: sheets "Page Text" (rows with Page = Home) and "Home Items" -----
     Each element marked data-db="..." in index.html keeps its HTML text as a fallback and is replaced
     only when the workbook has content for it, so a missing sheet never empties the page. */
  const paras = t => txt(t).split(/\n\s*\n/).map(p => `<p>${esc(p).replace(/\n/g, '<br>')}</p>`).join('');
  const lines = t => esc(t).replace(/\n/g, '<br>');
  const fill = (key, html) => { const el = $(`[data-db="${key}"]`); if (el && html) el.innerHTML = html; };
  const renderHome = get => {
    const prof = {};
    get('Page Text').map(row).forEach(r => { if (r.field && (!r.page || r.page.toLowerCase() === 'home')) prof[r.field.toLowerCase()] = r.value; });
    for (const [k, v] of Object.entries(prof)) {
      if (!v || k === 'avatar' || k === 'about') continue;
      const el = $(`[data-db="${k}"]`); if (el) el.textContent = v;
    }
    if (prof.about) fill('about', paras(prof.about));
    if (prof.name) { const img = $('[data-db="avatar"]'); if (img) img.alt = 'Portrait of ' + prof.name; }
    if (prof.avatar) {
      const img = $('[data-db="avatar"]');
      if (img && img.getAttribute('src') !== prof.avatar) { img.hidden = false; img.src = prof.avatar; }
    }
    // Home Items: Section decides where a row goes; rows keep their order inside each section
    const items = get('Home Items').map(row).filter(r => r.title), sec = n => items.filter(r => r.section.toLowerCase() === n);
    fill('education', sec('education').map(r =>
      `<p><strong>${esc(r.title)}</strong>${r.organization ? '<br>' + lines(r.organization) : ''}${r.dates ? `<br><span class="date">${lines(r.dates)}</span>` : ''}</p>`).join(''));
    fill('experience', sec('experience').map(r =>
      `<div class="tl-item"><div class="tl-top"><h4>${esc(r.title)}</h4>${r.dates ? `<span class="date">${esc(r.dates)}</span>` : ''}</div>
        ${r.organization ? `<p class="org">${esc(r.organization)}</p>` : ''}${r.description ? paras(r.description) : ''}</div>`).join(''));
    fill('courses', sec('courses').map(r => `<li>${esc(r.title)}</li>`).join(''));
    fill('honors', sec('honors').map(r =>
      `<li><strong>${esc(r.title)}</strong>${r.description ? `<br><span>${lines(r.description)}</span>` : ''}</li>`).join(''));
    fill('service', sec('service').map(r =>
      `<li><strong>${esc(r.title)}</strong>${r.description ? '<br>' + lines(r.description) : ''}</li>`).join(''));
  };
  const homeReady = () => root.classList.remove('dbwait');

  const showWb = wb => {
    const rows = name => {
      const n = wb.SheetNames.find(s => s.trim().toLowerCase() === name.toLowerCase());
      if (!n) return [];
      const ws = wb.Sheets[n], out = XLSX.utils.sheet_to_json(ws, { defval: '', blankrows: true });
      if (!ws['!ref']) return out;
      // Real Excel dates: rewrite them as text that follows the cell's own format (dd-mm-yyyy, mm-yyyy or yyyy)
      const rg = XLSX.utils.decode_range(ws['!ref']);
      for (let c = rg.s.c; c <= rg.e.c; c++) {
        const hd = ws[XLSX.utils.encode_cell({ r: rg.s.r, c })];
        if (!hd || !/^date$/i.test(txt(hd.v))) continue;
        out.forEach((row, i) => {
          const cell = ws[XLSX.utils.encode_cell({ r: rg.s.r + 1 + i, c })];
          if (cell && cell.t === 'n' && cell.v > 2100) row[txt(hd.v)] = dateCellText(cell);
        });
      }
      return out;
    };
    try { renderHome(rows); } catch (e) { console.error(e); }
    homeReady();
    renderPubs(rows('Publications'));
    renderProjects(rows('Projects'));
    renderDirections(rows('Directions'), rows('Topics'));
  };

  // Opened from disk (file://): browsers block fetch, so offer a file picker for a local preview instead
  const localBar = () => {
    if ($('#localbar')) return;
    const bar = document.createElement('div');
    bar.id = 'localbar'; bar.className = 'localbar';
    bar.innerHTML = '<label>Local preview: choose your database.xlsx <input type="file" accept=".xlsx"></label>';
    bar.querySelector('input').addEventListener('change', async e => {
      const f = e.target.files[0];
      if (f) showWb(XLSX.read(await f.arrayBuffer(), { type: 'array', cellNF: true }));
    });
    document.body.appendChild(bar);
  };

  (async () => {
    $('#pubs').innerHTML = $('#projects').innerHTML = '<p class="empty">Loading…</p>';
    try {
      if (!window.XLSX) throw new Error('SheetJS library did not load');
      const res = await fetch(DB_FILE + '?v=' + Date.now());
      if (!res.ok) throw new Error('HTTP ' + res.status);
      showWb(XLSX.read(await res.arrayBuffer(), { type: 'array', cellNF: true }));
    } catch (err) {
      console.error(err);
      homeReady();                                                      // keep the fallback texts
      if (location.protocol === 'file:' && window.XLSX) {
        localBar();
        $('#pubs').innerHTML = $('#projects').innerHTML = '<p class="empty">Opened from disk: use the file picker at the bottom of the page to choose data/database.xlsx. On GitHub Pages the data loads automatically.</p>';
        return;
      }
      const msg = !window.XLSX ? 'The Excel reader library (SheetJS from cdnjs) did not load. Check your internet connection or ad blocker.'
        : /404/.test(err.message) ? 'data/database.xlsx was not found. Check that it is inside the data folder of your repository with exactly this name (GitHub Pages is case-sensitive).'
        : 'Could not read data/database.xlsx (' + err.message + ').';
      $('#pubs').innerHTML = $('#projects').innerHTML = `<p class="empty">${msg}</p>`;
    }
  })();

  /* ---------- 2. THEME (light / dark) ---------- */
  const applyTheme = t => { root.dataset.theme = t; };
  applyTheme(store.get('theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'));
  $('#theme').addEventListener('click', () => {
    const t = root.dataset.theme === 'dark' ? 'light' : 'dark';
    applyTheme(t); store.set('theme', t);
  });

  /* ---------- 3. HERO NEURAL-NETWORK BACKGROUND ---------- */
  const cv = $('#net'), g = cv.getContext('2d'), hero = cv.parentElement;
  const still = matchMedia('(prefers-reduced-motion: reduce)').matches;
  let W = 0, H = 0, nodes = [], mx = -999, my = -999;
  const fit = () => {
    const r = hero.getBoundingClientRect(); W = cv.width = r.width; H = cv.height = r.height;
    nodes = Array.from({ length: Math.round(W * H / 15000) }, () => ({
      x: Math.random() * W, y: Math.random() * H, vx: (Math.random() - .5) * .3, vy: (Math.random() - .5) * .3 }));
    if (still) tick();
  };
  const line = (x1, y1, x2, y2, a) => { g.globalAlpha = a; g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke(); };
  function tick() {
    g.clearRect(0, 0, W, H);
    g.strokeStyle = g.fillStyle = getComputedStyle(root).getPropertyValue('--net').trim();
    nodes.forEach((a, i) => {
      if (!still) {
        a.x += a.vx; a.y += a.vy;
        if (a.x < 0 || a.x > W) a.vx *= -1;
        if (a.y < 0 || a.y > H) a.vy *= -1;
      }
      g.globalAlpha = .55; g.beginPath(); g.arc(a.x, a.y, 2.2, 0, 6.3); g.fill();
      for (let j = i + 1; j < nodes.length; j++) {
        const b = nodes[j], d = Math.hypot(a.x - b.x, a.y - b.y);
        if (d < 130) line(a.x, a.y, b.x, b.y, (1 - d / 130) * .35);
      }
      const dm = Math.hypot(a.x - mx, a.y - my);
      if (dm < 180) line(a.x, a.y, mx, my, (1 - dm / 180) * .8);
    });
    g.globalAlpha = 1;
    if (!still) requestAnimationFrame(tick);
  }
  hero.addEventListener('pointermove', e => { const r = cv.getBoundingClientRect(); mx = e.clientX - r.left; my = e.clientY - r.top; });
  hero.addEventListener('pointerleave', () => { mx = my = -999; });
  addEventListener('resize', fit);
  fit();
  if (!still) tick();
})();
