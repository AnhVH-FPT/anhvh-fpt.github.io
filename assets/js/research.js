/* Research topics page. Everything comes from data/database.xlsx, the same workbook the main page uses:
   sheets Directions, Topics, Conferences, Journals and Venue Notes (see its Instructions sheet). */
(() => {
  'use strict';
  const root = document.documentElement;
  const $ = (s, c = document) => c.querySelector(s);
  const $$ = (s, c = document) => [...c.querySelectorAll(s)];
  const store = {
    get: k => { try { return localStorage.getItem(k); } catch { return null; } },
    set: (k, v) => { try { localStorage.setItem(k, v); } catch { /* storage unavailable */ } }
  };
  const DB_FILE = 'data/database.xlsx';
  const EMAIL = 'AnhVH54@fe.edu.vn';
  const txt = v => String(v == null ? '' : v).trim();
  const esc = s => txt(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const linkify = s => esc(s).replace(/https?:\/\/[^\s<]+[^\s<.,;)]/g, u => `<a href="${u}" target="_blank" rel="noopener">${u.replace(/^https?:\/\/(www\.)?/, '').replace(/\/$/, '')}</a>`);

  /* ---------- theme (shared with the main page) ---------- */
  const applyTheme = t => { root.dataset.theme = t; };
  applyTheme(store.get('theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'));
  $('#theme').addEventListener('click', () => {
    const t = root.dataset.theme === 'dark' ? 'light' : 'dark';
    applyTheme(t); store.set('theme', t);
  });

  /* ---------- reading sheets ---------- */
  // Rows of a sheet as objects keyed by the lower-case header. The header is found by its column names,
  // so title rows above it do no harm. A missing column reads as ''.
  const table = (ws, test) => {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });
    const h = rows.findIndex(r => test(r.map(c => txt(c).toLowerCase())));
    if (h < 0) return null;
    const head = rows[h].map(c => txt(c).toLowerCase());
    return rows.slice(h + 1).map(r => {
      const o = {}; head.forEach((k, i) => { if (k) o[k] = txt(r[i]); });
      return new Proxy(o, { get: (t, k) => (k in t ? t[k] : '') });
    });
  };
  // The sheet with this name if it exists, otherwise the first sheet whose header passes the test.
  // The instructions sheet is never read as data (its example tables use the same column names).
  const findTable = (wb, name, test) => {
    const named = wb.SheetNames.find(n => n.trim().toLowerCase() === name.toLowerCase());
    const others = wb.SheetNames.filter(n => n !== named && !/instruction|hướng dẫn|huong dan/i.test(n));
    for (const n of named ? [named, ...others] : others) { const t = table(wb.Sheets[n], test); if (t) return t; }
    return [];
  };
  const pick = (r, ...keys) => { for (const k of keys) if (r[k]) return r[k]; return ''; };

  // "• Name\nhttps://..." -> [{name, url}]
  const datasets = s => {
    const out = [];
    txt(s).split('\n').map(l => l.trim()).filter(Boolean).forEach(l => {
      if (/^https?:\/\//.test(l)) { const last = out[out.length - 1]; if (last && !last.url) last.url = l; else out.push({ name: l.replace(/^https?:\/\/(www\.)?/, ''), url: l }); }
      else out.push({ name: l.replace(/^[•\-*]\s*/, ''), url: '' });
    });
    return out;
  };


  // Research Description and Research Expectation follow a fixed template (see the Hướng dẫn sheet):
  //   Context: text            <- "Label:" at the start of a line opens a field
  //   Implementation steps:
  //   (1) first step           <- "(1) ..." lines become a numbered list
  // Only these labels are shown, always in this order; a missing or empty field is skipped,
  // and text that is not under one of these labels is not shown.
  const DESC_FIELDS = ['Context', 'Problem formulation', 'Gap', 'Proposed approach', 'Implementation steps', 'Key knowledge to learn'];
  const EXP_FIELDS = ['Deliverables', 'Expected results', 'Baselines', 'Metrics', 'Experimental protocol', 'Minimum success criteria'];
  const parseFields = (text, fields) => {
    const keys = new Map(fields.map(f => [f.toLowerCase(), f])), out = {}, skipped = [];
    let cur = null;
    txt(text).split('\n').forEach(line => {
      const m = line.match(/^\s*([A-Za-z][A-Za-z /&-]*?)\s*:\s*(.*)$/), key = m && keys.get(m[1].toLowerCase().replace(/\s+/g, ' '));
      if (key) { cur = key; out[cur] = m[2].trim(); return; }
      if (!line.trim()) return;
      if (cur) out[cur] += '\n' + line.trim(); else skipped.push(line.trim());
    });
    return { out, skipped };
  };
  // A field: plain lines become a paragraph, "(1) ..." lines a numbered list
  const fieldHtml = (label, value) => {
    const lines = txt(value).split('\n').map(l => l.trim()).filter(Boolean);
    if (!lines.length) return '';
    const items = lines.filter(l => /^\(?\d+[).]\s/.test(l)), text = lines.filter(l => !/^\(?\d+[).]\s/.test(l));
    return `<div class="blk"><h5>${esc(label)}</h5>${text.length ? `<p>${linkify(text.join(' '))}</p>` : ''}${
      items.length ? `<ol>${items.map(l => `<li>${linkify(l.replace(/^\(?\d+[).]\s*/, ''))}</li>`).join('')}</ol>` : ''}</div>`;
  };
  const fieldsHtml = (t, fields) => fields.map(f => fieldHtml(f, t.fields[f])).join('') || '<p class="muted">Details coming soon.</p>';

  /* ---------- pagination ---------- */
  // Page buttons: 1 … 4 5 [6] 7 8 … 25
  const pagerHtml = (page, pages) => {
    if (pages <= 1) return '';
    const nums = [...new Set([1, pages, page - 2, page - 1, page, page + 1, page + 2])].filter(n => n >= 1 && n <= pages).sort((a, b) => a - b);
    const out = [`<button type="button" data-p="${page - 1}" ${page === 1 ? 'disabled' : ''} aria-label="Previous page">&lsaquo;</button>`];
    nums.forEach((n, i) => {
      if (i && n - nums[i - 1] > 1) out.push('<span class="gap">…</span>');
      out.push(`<button type="button" data-p="${n}" class="${n === page ? 'on' : ''}" ${n === page ? 'aria-current="page"' : ''}>${n}</button>`);
    });
    out.push(`<button type="button" data-p="${page + 1}" ${page === pages ? 'disabled' : ''} aria-label="Next page">&rsaquo;</button>`);
    return out.join('');
  };
  // Remembered "per page" choice (5/10/15/20/25)
  const savedPer = id => {
    const sel = $('#' + id), v = store.get(id);
    if (v && [...sel.options].some(o => o.text === v)) sel.value = v;
    return +sel.value || 10;
  };

  /* ---------- directions (sheet "Directions": Code, Direction, Focus) ---------- */
  let DIRS = [];                                        // [{code, name, focus, count}] in sheet order
  const dirByCode = code => DIRS.find(d => d.code.toLowerCase() === txt(code).toLowerCase());
  // "CV, MI" / "All" / "" -> list of direction codes ('All' and blank match every direction)
  const dirCodes = s => /^\s*(all\b.*)?$/i.test(s) ? DIRS.map(d => d.code)
    : txt(s).split(/[,;/]+/).map(x => (dirByCode(x) || {}).code).filter(Boolean);

  /* ---------- topics (sheet "Topics") ---------- */
  let TOPICS = [], dir = 'all', page = 1, per = savedPer('per');
  const wantDir = new URLSearchParams(location.search).get('dir');            // research.html?dir=MI preselects a direction
  const isClosed = t => /taken|assigned|closed|full/i.test(t.status);

  const mailParts = t => {
    const subject = `[Research Topic ${t.code}] ${t.en}`;
    const body = `Dear Mr. Ha Anh Vu,\n\nI would like to apply for the research topic below.\n\nTopic: ${t.code} - ${t.en}\n(${t.vi})\n\nFull name:\nStudent ID:\nMajor and semester:\nTeam members (if any):\nWhy I am interested in this topic:\nRelevant experience (Python, PyTorch, courses, projects):\n\nI have attached my CV and transcript.\n\nBest regards,\n`;
    return { subject, body };
  };
  const mailto = t => {
    const { subject, body } = mailParts(t);
    return `mailto:${EMAIL}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  };

  const card = t => {
    const closed = isClosed(t);
    const ds = datasets(t.dataset);
    return `<details class="topic${closed ? ' closed' : ''}" id="${esc(t.code)}">
      <summary>
        <span class="tcode">${esc(t.code)}</span>
        <div class="thead">
          <h3>${esc(t.en)}</h3>
          ${t.vi ? `<p class="vi">${esc(t.vi)}</p>` : ''}
          <p class="teaser">${esc((t.fields.Context || '').split('\n')[0])}</p>
          <div class="tmeta">
            <span class="tdir">${esc(t.dirName)}</span>
            ${closed ? `<span class="badge thesis">${esc(t.status)}</span>` : ''}
            ${ds.slice(0, 3).map(d => `<span class="dtag">${esc(d.name.replace(/\s*\(.*?\)\s*/g, ' ').trim())}</span>`).join('')}
          </div>
        </div>
      </summary>
      <div class="tbody">
        <div class="tcols">
          <div><h4>Research description</h4>${fieldsHtml(t, DESC_FIELDS)}</div>
          <div><h4>Research expectation</h4>${fieldsHtml(t, EXP_FIELDS)}
            ${ds.length ? `<h4>Datasets</h4><ul class="dlist">${ds.map(d => `<li>${d.url ? `<a href="${esc(d.url)}" target="_blank" rel="noopener">${esc(d.name)}</a>` : esc(d.name)}</li>`).join('')}</ul>` : ''}
          </div>
        </div>
        <div class="tact">
          ${closed ? '<span class="muted">This topic is already taken.</span>'
            : `<a class="btn pri" href="${mailto(t)}">Apply for this topic</a>`}
          ${closed ? '' : `<button type="button" class="btn copy-mail" data-code="${esc(t.code)}">Copy application text</button>`}
          <button type="button" class="btn copy" data-code="${esc(t.code)}">Copy link</button>
          ${closed ? '' : `<p class="tnote">Email app did not open? Copy the application text and send it to <strong>${EMAIL}</strong> from Outlook or Gmail, or message me on Zalo: <strong>0366765835</strong>.</p>`}
        </div>
      </div>
    </details>`;
  };

  const norm = s => s.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd');
  const filtered = () => {
    const q = norm($('#q').value.trim()), only = $('#only').checked;
    return TOPICS.filter(t => (dir === 'all' || t.dir === dir) && (!only || !isClosed(t)) && (!q || t.hay.includes(q)));
  };
  const drawTopics = () => {
    const list = filtered(), pages = Math.max(1, Math.ceil(list.length / per));
    page = Math.min(Math.max(1, page), pages);
    const from = (page - 1) * per, shown = list.slice(from, from + per);
    $('#list').innerHTML = shown.map(card).join('') || '<p class="empty">No topic matches your search.</p>';
    $('#pager').innerHTML = pagerHtml(page, pages);
    $('#count').textContent = list.length ? `${from + 1}–${from + shown.length} of ${list.length} topics` : `0 of ${TOPICS.length} topics`;
    $$('#dirs button').forEach(b => b.classList.toggle('on', b.dataset.dir === dir));
    const d = dirByCode(dir);
    $('#dirfocus').innerHTML = d && d.focus ? `<strong>Current trends:</strong> ${esc(d.focus)}` : '';
  };

  const renderTopics = wb => {
    DIRS = findTable(wb, 'Directions', h => h.includes('code') && h.includes('direction') && !h.includes('english title'))
      .filter(r => r.code && r.direction)
      .map(r => ({ code: r.code, name: r.direction, focus: pick(r, 'focus (current research trends)', 'focus'), count: 0 }));
    const rows = findTable(wb, 'Topics', h => h.includes('english title')).filter(r => r['english title']);
    rows.forEach((r, i) => {
      let d = dirByCode(r.direction) || DIRS.find(x => x.name.toLowerCase() === r.direction.toLowerCase());
      if (!d) {                                                     // a direction code missing from the Directions sheet
        d = { code: r.direction || 'Other', name: r.direction || 'Other', focus: '', count: 0 };
        DIRS.push(d);
      }
      d.count++;
      TOPICS.push({ i, code: r.code || `${d.code}-${String(i + 1).padStart(2, '0')}`, dir: d.code, dirName: d.name,
        en: r['english title'], vi: r['vietnamese title'], dataset: r.dataset, status: r.status,
        desc: r['research description'], expect: r['research expectation'] });
    });
    DIRS = DIRS.filter(d => d.count);
    if (wantDir && dirByCode(wantDir)) dir = dirByCode(wantDir).code;
    // Order: directions as in the Directions sheet, then by topic code number (MI-2 before MI-10), whatever the row order
    const rank = new Map(DIRS.map((d, i) => [d.code, i]));
    TOPICS.sort((a, b) => rank.get(a.dir) - rank.get(b.dir) || a.code.localeCompare(b.code, undefined, { numeric: true }) || a.i - b.i);
    TOPICS.forEach(x => {
      const d = parseFields(x.desc, DESC_FIELDS), e = parseFields(x.expect, EXP_FIELDS);
      x.fields = { ...d.out, ...e.out };
      // Help the editor: lines that do not follow the template are listed in the browser console (F12)
      if (d.skipped.length || e.skipped.length) console.warn(`Topic ${x.code}: text outside the template is not shown:`, [...d.skipped, ...e.skipped]);
      x.hay = norm([x.code, x.en, x.vi, x.dataset, x.desc, x.expect].join(' '));
    });
    $('#onlyWrap').hidden = !TOPICS.some(isClosed);
    $('#dirs').innerHTML = `<button type="button" data-dir="all">All directions <em>${TOPICS.length}</em></button>` +
      DIRS.map(d => `<button type="button" data-dir="${esc(d.code)}">${esc(d.name)} <em>${d.count}</em></button>`).join('');
    drawTopics();
    openFromHash();
  };

  const openFromHash = () => {
    const id = decodeURIComponent(location.hash.slice(1)); if (!id) return;
    let el = document.getElementById(id);
    if (!el && TOPICS.some(t => t.code === id)) {                              // go to the page that holds this topic
      if (!filtered().some(t => t.code === id)) { dir = 'all'; $('#q').value = ''; $('#only').checked = false; }
      page = Math.floor(filtered().findIndex(t => t.code === id) / per) + 1;
      drawTopics(); el = document.getElementById(id);
    }
    if (el && el.classList.contains('topic')) { el.open = true; el.scrollIntoView({ block: 'start' }); }
  };
  addEventListener('hashchange', openFromHash);

  $('#dirs').addEventListener('click', e => { const b = e.target.closest('button'); if (b) { dir = b.dataset.dir; page = 1; drawTopics(); } });
  let qt; $('#q').addEventListener('input', () => { clearTimeout(qt); qt = setTimeout(() => { page = 1; drawTopics(); }, 120); });
  $('#only').addEventListener('change', () => { page = 1; drawTopics(); });
  $('#per').addEventListener('change', () => { per = +$('#per').value; store.set('per', per); page = 1; drawTopics(); });
  $('#pager').addEventListener('click', e => {
    const b = e.target.closest('button[data-p]'); if (!b || b.disabled) return;
    page = +b.dataset.p; drawTopics();
    $('#dirs').scrollIntoView({ block: 'start' });
  });

  // Copy to clipboard, with a manual fallback for browsers that block it
  const copy = async (b, text, done) => {
    const label = b.textContent;
    try { await navigator.clipboard.writeText(text); b.textContent = done; }
    catch {
      const ta = Object.assign(document.createElement('textarea'), { value: text });
      document.body.appendChild(ta); ta.select();
      const ok = document.execCommand && document.execCommand('copy');
      document.body.removeChild(ta);
      if (ok) b.textContent = done; else prompt('Copy this text:', text);
    }
    setTimeout(() => { b.textContent = label; }, 1800);
  };
  $('#list').addEventListener('click', e => {
    const b = e.target.closest('.copy, .copy-mail'); if (!b) return;
    if (b.classList.contains('copy')) return copy(b, location.href.split('#')[0] + '#' + b.dataset.code, 'Link copied');
    const { subject, body } = mailParts(TOPICS.find(t => t.code === b.dataset.code));
    copy(b, `To: ${EMAIL}\nSubject: ${subject}\n\n${body}`, 'Copied');
  });
  $$('.copy-val').forEach(b => b.addEventListener('click', () => copy(b, b.dataset.copy, 'Copied')));

  /* ---------- venues (sheets "Conferences", "Journals", "Venue Notes") ---------- */
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  // "Mid December – January" -> [12, 1]; "February–March" -> [2, 3]; "July (round 1) / September (round 2)" -> [7, 9]
  const monthsOf = s => {
    const out = new Set();
    txt(s).replace(/\(.*?\)/g, ' ').split('/').forEach(part => {
      const found = [...part.matchAll(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/gi)]
        .map(m => MONTHS.findIndex(x => x.slice(0, 3).toLowerCase() === m[0].toLowerCase()) + 1);
      if (found.length === 2 && /[–—-]|\bto\b/.test(part)) {                      // a range, possibly across the new year
        for (let m = found[0], n = 0; n < 12; m = m % 12 + 1, n++) { out.add(m); if (m === found[1]) break; }
      } else found.forEach(m => out.add(m));
    });
    return [...out];
  };
  const thisMonth = new Date().getMonth() + 1;
  const soon = new Set([0, 1, 2].map(k => (thisMonth - 1 + k) % 12 + 1));       // this month and the next two

  let CONF = [], JOUR = [], NOTES = [], tab = 'conf', vf = 'all', vpage = 1, vper = savedPer('vper');
  const vsel = { dir: 'all', dl: 'any', held: 'any' };
  const RANKS = { conf: r => r.rank || 'Not ranked', jour: r => r.rank };
  const venueRows = (data, rk, ignore) => data.filter(x =>
    (ignore === 'rank' || vf === 'all' || rk(x) === vf)
    && (vsel.dir === 'all' || x.dirs.includes(vsel.dir))
    && (tab !== 'conf' || vsel.dl === 'any' || (vsel.dl === 'soon' ? x.dl.some(m => soon.has(m)) : x.dl.includes(+vsel.dl)))
    && (tab !== 'conf' || vsel.held === 'any' || x.held.includes(+vsel.held)));

  const drawVenues = () => {
    const data = tab === 'conf' ? CONF : JOUR, rk = RANKS[tab];
    const order = r => ({ 'A*': '0', A: '1', B: '2', C: '3' }[r] || (/^Q\d/.test(r) ? r : 'z' + r));   // A* first; "Not ranked" last
    const ranks = [...new Set(data.map(rk))].filter(Boolean).sort((a, b) => order(a).localeCompare(order(b)));
    if (!['all', ...ranks].includes(vf)) vf = 'all';
    const base = venueRows(data, rk, 'rank');                                     // rank counts follow the other filters
    $('#venueFilter').innerHTML = [`<button type="button" data-r="all">All <em>${base.length}</em></button>`,
      ...ranks.map(r => `<button type="button" data-r="${esc(r)}">${tab === 'conf' && /^[A-C]/.test(r) ? 'CORE ' : ''}${esc(r)} <em>${base.filter(x => rk(x) === r).length}</em></button>`)].join('');
    $$('#venueFilter button').forEach(b => b.classList.toggle('on', b.dataset.r === vf));
    $$('#venues .seg button').forEach(b => b.classList.toggle('on', b.dataset.tab === tab));
    $$('.conf-only').forEach(el => { el.hidden = tab !== 'conf'; });

    const all = venueRows(data, rk), pages = Math.max(1, Math.ceil(all.length / vper));
    vpage = Math.min(Math.max(1, vpage), pages);
    const rows = all.slice((vpage - 1) * vper, vpage * vper);
    $('#vpager').innerHTML = pagerHtml(vpage, pages);
    $('#vcount').textContent = all.length ? `${(vpage - 1) * vper + 1}–${(vpage - 1) * vper + rows.length} of ${all.length}` : '';
    const a = (u, l) => u ? `<a href="${esc(u)}" target="_blank" rel="noopener">${esc(l)}</a>` : esc(l);
    const dirTags = x => x.all ? '<span class="dtag">All directions</span>' : x.dirs.map(c => `<span class="dtag" title="${esc((dirByCode(c) || {}).name)}">${esc(c)}</span>`).join(' ');
    $('#venueBox').innerHTML = !rows.length ? '<p class="empty">No venue matches these filters.</p>' : tab === 'conf'
      ? `<table><thead><tr><th>Conference</th><th>Rank</th><th>Directions</th><th>Deadline</th><th>Start preparing</th><th>Held in</th><th>Hybrid</th></tr></thead><tbody>${
        rows.map(r => `<tr><td><strong>${a(r.website, r.name)}</strong><small>${esc(r.full)}</small></td><td>${esc(rk(r))}</td><td>${dirTags(r)}</td><td>${esc(r.deadline)}</td><td>${esc(r.prep)}</td><td>${esc(r.heldText)}</td><td>${/yes/i.test(r.hybrid) ? '<span class="badge published">Yes</span>' : esc(r.hybrid)}</td></tr>`).join('')}</tbody></table>`
      : `<table><thead><tr><th>Journal</th><th>Publisher</th><th>Directions</th><th>Rank</th></tr></thead><tbody>${
        rows.map(r => `<tr><td><strong>${a(r.link, r.name)}</strong></td><td>${esc(r.publisher)}</td><td>${dirTags(r)}</td><td><span class="q q${esc(r.rank).slice(1)}">${esc(r.rank)}</span></td></tr>`).join('')}</tbody></table>`;
    const notes = NOTES.filter(n => /^all$/i.test(n.to) || n.to.toLowerCase().startsWith(tab === 'conf' ? 'conf' : 'journ'));
    $('#vnotes').innerHTML = notes.length ? `<details><summary>Notes (${notes.length})</summary><ul>${notes.map(n => `<li>${linkify(n.note)}</li>`).join('')}</ul></details>` : '';
  };

  const renderVenues = wb => {
    const dirsOf = s => ({ dirs: dirCodes(s), all: /^\s*(all\b.*)?$/i.test(s) });
    CONF = findTable(wb, 'Conferences', h => h.includes('conference') && h.includes('website')).filter(r => r.conference).map((r, i) => {
      const deadline = pick(r, 'submission deadline', 'typical submission deadline'), heldText = pick(r, 'held in', 'typical conference month');
      return { i, name: r.conference, full: r['full name'], rank: r['core rank'], deadline, heldText, hybrid: r.hybrid, website: r.website,
        prep: pick(r, 'start preparing', 'suggested start of preparation'), dl: monthsOf(deadline), held: monthsOf(heldText),
        ...dirsOf(pick(r, 'directions', 'research directions')) };
    });
    // Sorting is done here, so rows can be typed into the sheets in any order.
    // Conferences: deadline month (January first; "Mid December – January" counts as December),
    // then early / mid / late in that month, then rank (A* first), then name.
    const RANK_ORDER = { 'A*': 0, A: 1, B: 2, C: 3, Q1: 0, Q2: 1, Q3: 2, Q4: 3 };
    const rankKey = r => r in RANK_ORDER ? RANK_ORDER[r] : 9;
    const dlKey = x => x.dl.length ? x.dl[0] * 10 + (/^\s*early/i.test(x.deadline) ? 0 : /^\s*late/i.test(x.deadline) ? 2 : 1) : 999;
    const byName = (a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
    CONF.sort((a, b) => dlKey(a) - dlKey(b) || rankKey(a.rank) - rankKey(b.rank) || byName(a, b));
    // Journals: Q1, Q2, Q3, Q4 (unranked last), then A to Z
    JOUR = findTable(wb, 'Journals', h => h.some(x => /^journal( name)?$/.test(x)) && h.includes('rank'))
      .map(r => ({ name: pick(r, 'journal', 'journal name'), publisher: r.publisher, rank: r.rank.toUpperCase(), link: r.link, ...dirsOf(r.directions) }))
      .filter(r => r.name && !/^notes?:?$/i.test(r.name) && !r.name.startsWith('•'))
      .sort((a, b) => rankKey(a.rank) - rankKey(b.rank) || byName(a, b));
    NOTES = findTable(wb, 'Venue Notes', h => h.includes('note') && h.includes('applies to')).filter(r => r.note).map(r => ({ to: r['applies to'] || 'All', note: r.note }));

    if (!CONF.length && !JOUR.length) { $('#venues').hidden = true; $('a[href="#venues"]').hidden = true; return; }
    if (!CONF.length) tab = 'jour';
    const opt = (v, l) => `<option value="${v}">${esc(l)}</option>`;
    $('#vdir').innerHTML = opt('all', 'All directions') + DIRS.map(d => opt(esc(d.code), d.name)).join('');
    $('#vdl').innerHTML = opt('any', 'Any month') + opt('soon', 'Next 3 months') + MONTHS.map((m, i) => opt(i + 1, m)).join('');
    $('#vheld').innerHTML = opt('any', 'Any month') + MONTHS.map((m, i) => opt(i + 1, m)).join('');
    drawVenues();
  };
  $('#venues .seg').addEventListener('click', e => { const b = e.target.closest('button'); if (b) { tab = b.dataset.tab; vf = 'all'; vpage = 1; drawVenues(); } });
  $('#venueFilter').addEventListener('click', e => { const b = e.target.closest('button'); if (b) { vf = b.dataset.r; vpage = 1; drawVenues(); } });
  [['#vdir', 'dir'], ['#vdl', 'dl'], ['#vheld', 'held']].forEach(([id, k]) => $(id).addEventListener('change', e => { vsel[k] = e.target.value; vpage = 1; drawVenues(); }));
  $('#vreset').addEventListener('click', () => {
    vf = 'all'; vsel.dir = 'all'; vsel.dl = vsel.held = 'any'; vpage = 1;
    $('#vdir').value = 'all'; $('#vdl').value = $('#vheld').value = 'any'; drawVenues();
  });
  $('#vper').addEventListener('change', () => { vper = +$('#vper').value; store.set('vper', vper); vpage = 1; drawVenues(); });
  $('#vpager').addEventListener('click', e => {
    const b = e.target.closest('button[data-p]'); if (!b || b.disabled) return;
    vpage = +b.dataset.p; drawVenues();
    $('#venues .seg').scrollIntoView({ block: 'start' });
  });

  /* ---------- load ---------- */
  const show = wb => { renderTopics(wb); renderVenues(wb); };
  (async () => {
    try {
      if (!window.XLSX) throw new Error('The Excel reader library (SheetJS from cdnjs) did not load. Check your internet connection or ad blocker.');
      const res = await fetch(DB_FILE + '?v=' + Date.now());
      if (!res.ok) throw new Error(res.status === 404 ? DB_FILE + ' was not found in the repository.' : 'HTTP ' + res.status);
      show(XLSX.read(await res.arrayBuffer(), { type: 'array' }));
      if (!TOPICS.length) $('#list').innerHTML = '<p class="empty">No topics yet: add rows to the Topics sheet of data/database.xlsx.</p>';
    } catch (err) {
      console.error(err);
      const msg = location.protocol === 'file:'
        ? 'Opened from disk: browsers block reading ' + DB_FILE + '. Run a local server (python -m http.server) or view the GitHub Pages site.'
        : 'Could not load the topic list: ' + err.message;
      $('#list').innerHTML = $('#venueBox').innerHTML = `<p class="empty">${esc(msg)}</p>`;
    }
  })();
})();
