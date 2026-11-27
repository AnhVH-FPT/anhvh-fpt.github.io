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
  // Cells typed or pasted in Excel may use Windows line breaks (CR LF): turn every line break into LF
  const txt = v => String(v == null ? '' : v).replace(/\r\n?/g, '\n').trim();
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
    // Taken topics appear only under the "Taken" chip; every other view lists open topics only
    const q = norm($('#q').value.trim());
    return TOPICS.filter(t => (dir === 'taken' ? isClosed(t) : !isClosed(t) && (dir === 'all' || t.dir === dir)) && (!q || t.hay.includes(q)));
  };
  const drawTopics = () => {
    const list = filtered(), pages = Math.max(1, Math.ceil(list.length / per));
    page = Math.min(Math.max(1, page), pages);
    const from = (page - 1) * per, shown = list.slice(from, from + per);
    $('#list').innerHTML = shown.map(card).join('') || `<div class="empty tempty">No topic matches your search${dir !== 'all' ? ' in this group' : ''}.
      ${dir !== 'all' ? '<button type="button" class="linkbtn" data-tfix="all">Search all directions</button>' : ''}<button type="button" class="linkbtn" data-tfix="clear">Clear search</button></div>`;
    $('#pager').innerHTML = pagerHtml(page, pages);
    $('#count').innerHTML = list.length ? `Showing <b>${from + 1}–${from + shown.length}</b> of <b>${list.length}</b> topic${list.length === 1 ? '' : 's'}` : '';
    $$('#dirs button').forEach(b => b.classList.toggle('on', b.dataset.dir === dir));
    const d = dirByCode(dir);
    $('#dirfocus').innerHTML = dir === 'taken' ? 'These topics are already taken by students and are no longer open for applications. They are listed for reference with their original codes.'
      : d && d.focus ? `<strong>Current trends:</strong> ${esc(d.focus)}` : '';
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
      // The code is not stored in Excel: it is the direction code + the topic's position among that direction's rows
      // (first row of MI = MI-01). Deleting or moving a row therefore renumbers the topics after it.
      TOPICS.push({ i, code: `${d.code}-${String(d.count).padStart(2, '0')}`, dir: d.code, dirName: d.name,
        en: r['english title'], vi: r['vietnamese title'], dataset: r.dataset, status: r.status,
        desc: r['research description'], expect: r['research expectation'] });
    });
    DIRS = DIRS.filter(d => d.count);
    if (wantDir && dirByCode(wantDir)) dir = dirByCode(wantDir).code;
    if (wantDir && wantDir.toLowerCase() === 'taken') dir = 'taken';
    // Order: directions as in the Directions sheet, then row order in the Topics sheet (= code number)
    const rank = new Map(DIRS.map((d, i) => [d.code, i]));
    TOPICS.sort((a, b) => rank.get(a.dir) - rank.get(b.dir) || a.i - b.i);
    TOPICS.forEach(x => {
      const d = parseFields(x.desc, DESC_FIELDS), e = parseFields(x.expect, EXP_FIELDS);
      x.fields = { ...d.out, ...e.out };
      // Help the editor: lines that do not follow the template are listed in the browser console (F12)
      if (d.skipped.length || e.skipped.length) console.warn(`Topic ${x.code}: text outside the template is not shown:`, [...d.skipped, ...e.skipped]);
      x.hay = norm([x.code, x.en, x.vi, x.dataset, x.desc, x.expect].join(' '));
    });
    // Chip counts are open topics; taken ones get their own chip (shown only when there are some)
    const open = TOPICS.filter(t => !isClosed(t)), taken = TOPICS.length - open.length;
    $('#dirs').innerHTML = `<button type="button" data-dir="all">All directions <em>${open.length}</em></button>` +
      DIRS.map(d => `<button type="button" data-dir="${esc(d.code)}">${esc(d.name)} <em>${open.filter(t => t.dir === d.code).length}</em></button>`).join('') +
      (taken ? `<button type="button" data-dir="taken" class="taken-chip">Taken <em>${taken}</em></button>` : '');
    if (dir === 'taken' && !taken) dir = 'all';
    drawTopics();
    openFromHash();
  };

  const openFromHash = () => {
    const id = decodeURIComponent(location.hash.slice(1)); if (!id) return;
    let el = document.getElementById(id);
    if (!el && TOPICS.some(t => t.code === id)) {                              // go to the page that holds this topic
      if (!filtered().some(t => t.code === id)) { dir = isClosed(TOPICS.find(t => t.code === id)) ? 'taken' : 'all'; $('#q').value = ''; }
      page = Math.floor(filtered().findIndex(t => t.code === id) / per) + 1;
      drawTopics(); el = document.getElementById(id);
    }
    if (el && el.classList.contains('topic')) { el.open = true; el.scrollIntoView({ block: 'start' }); }
  };
  addEventListener('hashchange', openFromHash);

  $('#dirs').addEventListener('click', e => { const b = e.target.closest('button'); if (b) { dir = b.dataset.dir; page = 1; drawTopics(); } });
  let qt; $('#q').addEventListener('input', () => { clearTimeout(qt); qt = setTimeout(() => { page = 1; drawTopics(); }, 120); });
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

  // Buttons in the "no topic matches" message
  $('#list').addEventListener('click', e => {
    const b = e.target.closest('[data-tfix]'); if (!b) return;
    if (b.dataset.tfix === 'clear') $('#q').value = ''; else dir = 'all';
    page = 1; drawTopics();
  });

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

  /* Venue table state. Each tab keeps its own filters and sort, so switching tabs does not lose them. */
  let CONF = [], JOUR = [], NOTES = [], tab = 'conf', vpage = 1, vper = savedPer('vper');
  const RANK_ORDER = { 'A*': 0, A: 1, B: 2, C: 3, Q1: 0, Q2: 1, Q3: 2, Q4: 3 };
  const rankKey = r => r in RANK_ORDER ? RANK_ORDER[r] : 9;
  const rankOf = x => x.rank || (tab === 'conf' ? 'Not ranked' : '');
  const dlKey = x => x.dl.length ? x.dl[0] * 10 + (/^\s*early/i.test(x.deadline) ? 0 : /^\s*late/i.test(x.deadline) ? 2 : 1) : Infinity;
  // Months until the next deadline, counted from this month (this month = 0)
  const upKey = x => x.dl.length ? Math.min(...x.dl.map(m => (m - thisMonth + 12) % 12)) * 10 + dlKey(x) % 10 : Infinity;
  const DEF = {
    conf: { q: '', rank: 'all', dir: 'all', dl: 'any', held: 'any', fmt: 'any', fee: 'any' },
    jour: { q: '', rank: 'all', dir: 'all', ttp: 'any', fee: 'any', pub: 'all' },
  };
  const F = { conf: { ...DEF.conf }, jour: { ...DEF.jour } };
  const SORT = { conf: { key: 'deadline', dir: 1 }, jour: { key: 'rank', dir: 1 } };

  // Sort keys: a value per row (number or text); Infinity = unknown, always listed last
  const SORTS = {
    conf: { deadline: ['Deadline (Jan → Dec)', dlKey], upcoming: ['Upcoming deadline first', upKey],
      rank: ['Rank (best first)', x => rankKey(x.rank)], fee: ['Lowest fee first', x => x.feeReg],
      held: ['Held in (Jan → Dec)', x => x.held.length ? x.held[0] : Infinity], name: ['Name (A → Z)', x => x.name.toLowerCase()] },
    jour: { rank: ['Rank (Q1 first)', x => rankKey(x.rank)], ttp: ['Fastest to publish', x => x.ttpMax * 100 + x.ttpMin],
      fee: ['Free first, then lowest fee', x => x.feeReg], name: ['Name (A → Z)', x => x.name.toLowerCase()],
      publisher: ['Publisher (A → Z)', x => (x.publisher || '').toLowerCase() || Infinity] },
  };
  const cmpVal = (x, y, d) => x === y ? 0 : x === Infinity ? 1 : y === Infinity ? -1 : (typeof x === 'string' ? x.localeCompare(y) : x - y) * d;
  const sorted = rows => {
    const { key, dir } = SORT[tab], k = SORTS[tab][key][1];
    const tie = tab === 'conf' ? [dlKey, x => rankKey(x.rank)] : [x => rankKey(x.rank)];
    return rows.slice().sort((a, b) => cmpVal(k(a), k(b), dir) || tie.reduce((r, f) => r || cmpVal(f(a), f(b), 1), 0)
      || a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
  };

  // Filter fields per tab: [key, label, options(rows) -> [[value, text]], test(row, value)]
  const monthOpts = (first) => [...first, ...MONTHS.map((m, i) => [String(i + 1), m])];
  const FIELDS = {
    conf: [
      ['rank', 'Rank', null, (x, v) => rankOf(x) === v],
      ['dir', 'Direction', null, (x, v) => x.dirs.includes(v)],
      ['dl', 'Deadline', () => monthOpts([['any', 'Any month'], ['soon', 'Next 3 months']]), (x, v) => v === 'soon' ? x.dl.some(m => soon.has(m)) : x.dl.includes(+v)],
      ['held', 'Held in', () => monthOpts([['any', 'Any month']]), (x, v) => x.held.includes(+v)],
      ['fmt', 'Presentation', () => [['any', 'Any'], ['online', 'Online possible'], ['onsite', 'In person only']], (x, v) => /yes/i.test(x.hybrid) === (v === 'online')],
      ['fee', 'Regular fee', () => [['any', 'Any fee'], ...[400, 600, 800, 1000].map(n => [String(n), `Up to ≈ USD ${n.toLocaleString('en-US')}`])], (x, v) => x.feeReg <= +v],
    ],
    jour: [
      ['rank', 'Rank', null, (x, v) => rankOf(x) === v],
      ['dir', 'Direction', null, (x, v) => x.dirs.includes(v)],
      ['ttp', 'Time to publish', () => [['any', 'Any time'], ['9', 'Up to 9 months'], ['12', 'Up to 12 months'], ['15', 'Up to 15 months'], ['over', 'Over 15 months']],
        (x, v) => v === 'over' ? x.ttpMax > 15 && x.ttpMax < 999 : x.ttpMax <= +v],
      ['fee', 'Publication fee', () => [['any', 'Any'], ['free', 'Free only'], ['paid', 'With a fee']], (x, v) => (x.feeReg === 0) === (v === 'free')],
      ['pub', 'Publisher', null, (x, v) => x.pub === v],
    ],
  };
  const isDef = (k, v) => v === DEF[tab][k];
  const passes = (x, skip) => {
    const f = F[tab];
    if (f.q && !x.hay.includes(norm(f.q))) return false;
    return FIELDS[tab].every(([k, , , test]) => k === skip || isDef(k, f[k]) || test(x, f[k]));
  };
  // Options whose list depends on the data (rank, direction, publisher) show how many rows each would give
  const dynOpts = (k, data) => {
    const base = data.filter(x => passes(x, k)), n = test => base.filter(test).length;
    const anyText = { rank: 'All ranks', dir: 'All directions', pub: 'All publishers' }[k];
    let vals;
    if (k === 'rank') {
      const order = r => ({ 'A*': '0', A: '1', B: '2', C: '3' }[r] || (/^Q\d/.test(r) ? r : 'z' + r));
      vals = [...new Set(data.map(rankOf))].filter(Boolean).sort((a, b) => order(a).localeCompare(order(b)))
        .map(r => [r, (tab === 'conf' && /^[A-C]/.test(r) ? 'CORE ' : '') + r, n(x => rankOf(x) === r)]);
    } else if (k === 'dir') vals = DIRS.map(d => [d.code, d.name, n(x => x.dirs.includes(d.code))]);
    else vals = [...new Set(data.map(x => x.pub).filter(Boolean))].sort((a, b) => a.localeCompare(b))
      .map(p => [p, p, n(x => x.pub === p)]);
    return [['all', `${anyText} (${base.length})`], ...vals.map(([v, l, c]) => [v, `${l} (${c})`])];
  };
  const optText = (k, v) => {
    const fld = FIELDS[tab].find(f => f[0] === k);
    const list = fld[2] ? fld[2]() : dynOpts(k, tab === 'conf' ? CONF : JOUR);
    return ((list.find(o => o[0] === v) || [])[1] || v).replace(/ \(\d+\)$/, '');
  };

  const buildFields = () => {
    $('#vfields').innerHTML = FIELDS[tab].map(([k, label]) =>
      `<label class="fld"><span>${label}</span><select data-k="${k}" aria-label="${label}"></select></label>`).join('');
    $('#vq').value = F[tab].q;
    $('#vq').placeholder = tab === 'conf' ? 'Search name, full name or proceedings…' : 'Search journal or publisher…';
  };

  const drawVenues = () => {
    const data = tab === 'conf' ? CONF : JOUR, f = F[tab];
    $$('#venues .seg button').forEach(b => b.classList.toggle('on', b.dataset.tab === tab));
    $('#vtabConf').textContent = CONF.length; $('#vtabJour').textContent = JOUR.length;
    // filter selects (counts follow the other filters)
    $$('#vfields select').forEach(sel => {
      const k = sel.dataset.k, fld = FIELDS[tab].find(x => x[0] === k);
      const list = fld[2] ? fld[2]() : dynOpts(k, data);
      sel.innerHTML = list.map(([v, l]) => `<option value="${esc(v)}">${esc(l)}</option>`).join('');
      sel.value = f[k];
      sel.closest('.fld').classList.toggle('set', !isDef(k, f[k]));
    });
    // active filters as removable chips
    const act = FIELDS[tab].filter(([k]) => !isDef(k, f[k])).map(([k, label]) =>
      `<button type="button" class="pill" data-k="${k}" title="Remove this filter">${label}: <b>${esc(optText(k, f[k]))}</b> <span aria-hidden="true">×</span></button>`);
    if (f.q) act.unshift(`<button type="button" class="pill" data-k="q" title="Clear search">Search: <b>${esc(f.q)}</b> <span aria-hidden="true">×</span></button>`);
    $('#vactive').innerHTML = act.length ? act.join('') + '<button type="button" class="linkbtn" id="vclear">Clear all</button>' : '';
    // sort select
    $('#vsort').innerHTML = Object.entries(SORTS[tab]).map(([k, [l]]) => `<option value="${k}">${l}</option>`).join('');
    $('#vsort').value = SORT[tab].key;

    const all = sorted(data.filter(x => passes(x)));
    const pages = Math.max(1, Math.ceil(all.length / vper));
    vpage = Math.min(Math.max(1, vpage), pages);
    const rows = all.slice((vpage - 1) * vper, vpage * vper);
    $('#vpager').innerHTML = pagerHtml(vpage, pages);
    const noun = tab === 'conf' ? 'conference' : 'journal';
    $('#vcount').innerHTML = all.length ? `Showing <b>${(vpage - 1) * vper + 1}–${(vpage - 1) * vper + rows.length}</b> of <b>${all.length}</b> ${noun}${all.length === 1 ? '' : 's'}`
      + (all.length < data.length ? ` <span class="muted">(${data.length} in total)</span>` : '') : `No ${noun}s match`;

    const a = (u, l) => u ? `<a href="${esc(u)}" target="_blank" rel="noopener">${esc(l)}</a>` : esc(l);
    const dirTags = x => x.all ? '<span class="dtag">All directions</span>' : x.dirs.map(c => `<span class="dtag" title="${esc((dirByCode(c) || {}).name)}">${esc(c)}</span>`).join(' ');
    const th = (label, key) => {
      if (!key) return `<th>${label}</th>`;
      const on = SORT[tab].key === key, d = SORT[tab].dir;
      return `<th aria-sort="${on ? (d > 0 ? 'ascending' : 'descending') : 'none'}"><button type="button" class="sortbtn${on ? ' on' : ''}" data-sort="${key}">${label}<i aria-hidden="true">${on ? (d > 0 ? '↑' : '↓') : '↕'}</i></button></th>`;
    };
    const rankBadge = x => tab === 'jour' ? `<span class="q q${esc(x.rank).slice(1)}">${esc(x.rank) || '–'}</span>`
      : `<span class="rk${/^[A-C]/.test(rankOf(x)) ? '' : ' nr'}">${esc(rankOf(x))}</span>`;
    if (!rows.length) {
      $('#venueBox').innerHTML = `<div class="empty vempty">No ${noun} matches these filters. <button type="button" class="linkbtn" id="vclear2">Clear all filters</button></div>`;
    } else if (tab === 'conf') {
      $('#venueBox').innerHTML = `<table class="vt"><thead><tr>${th('Conference', 'name')}${th('Rank', 'rank')}${th('Deadline', 'deadline')}${th('Held in', 'held')}<th>Presentation</th>${th('Registration fee', 'fee')}</tr></thead><tbody>${
        rows.map(r => `<tr>
          <td class="vname"><strong>${a(r.website, r.name)}</strong><small>${esc(r.full)}</small><div class="dtags">${dirTags(r)}</div></td>
          <td data-label="Rank">${rankBadge(r)}</td>
          <td data-label="Deadline">${esc(r.deadline) || '–'}</td>
          <td data-label="Held in">${esc(r.heldText) || '–'}</td>
          <td data-label="Presentation">${/yes/i.test(r.hybrid) ? '<span class="badge published">Online possible</span>' : '<span class="badge thesis">In person</span>'}</td>
          <td data-label="Fee" class="fee">${feeHtml(r.feeText)}</td></tr>`).join('')}</tbody></table>`;
    } else {
      $('#venueBox').innerHTML = `<table class="vt"><thead><tr>${th('Journal', 'name')}${th('Publisher', 'publisher')}${th('Rank', 'rank')}${th('Time to publish', 'ttp')}${th('Publication fee', 'fee')}</tr></thead><tbody>${
        rows.map(r => `<tr>
          <td class="vname"><strong>${a(r.link, r.name)}</strong><div class="dtags">${dirTags(r)}</div></td>
          <td data-label="Publisher">${esc(r.publisher) || '–'}</td>
          <td data-label="Rank">${rankBadge(r)}</td>
          <td data-label="Time to publish" class="nowrap">${esc(r.ttp) || '–'}</td>
          <td data-label="Fee" class="fee">${feeHtml(r.feeText)}</td></tr>`).join('')}</tbody></table>`;
    }
    const notes = NOTES.filter(n => /^all$/i.test(n.to) || n.to.toLowerCase().startsWith(tab === 'conf' ? 'conf' : 'journ'));
    $('#vnotes').innerHTML = notes.length ? `<details><summary>Notes on these ${noun}s (${notes.length})</summary><ul>${notes.map(n => `<li>${linkify(n.note)}</li>`).join('')}</ul></details>` : '';
  };

  // "4–8 months" -> [4, 8]; "6 months" -> [6, 6]; blank or no number -> unknown (sorted last, kept only by "Any time")
  const ttpOf = s => { const n = (txt(s).match(/\d+(\.\d+)?/g) || []).map(Number); return n.length ? [Math.min(...n), Math.max(...n)] : [999, 999]; };
  // Fees: "Regular: ≈USD 900" / "Student: ≈USD 600" (one per line) or "Free". Amounts are turned into USD with rough rates, only for sorting and filtering.
  const RATE = { USD: 1, EUR: 1.1, GBP: 1.3, AUD: 0.66, NZD: 0.6, CAD: 0.73, SGD: 0.75, CNY: 0.14, JPY: 0.0067, KRW: 0.00073, INR: 0.012, VND: 0.00004 };
  const usd = line => {
    if (/\bfree\b/i.test(line)) return 0;
    const m = line.replace(/\$/, 'USD ').match(/\b([A-Z]{3})\s*([\d.,]+)/);
    return m && RATE[m[1]] ? parseFloat(m[2].replace(/,/g, '')) * RATE[m[1]] : Infinity;
  };
  const feeOf = s => {
    const lines = txt(s).split('\n').map(l => l.trim()).filter(Boolean);
    const reg = lines.find(l => !/student/i.test(l)) || lines[0] || '', stu = lines.find(l => /student/i.test(l)) || '';
    return { feeText: lines.join('\n'), feeReg: reg ? usd(reg) : Infinity, feeStu: stu ? usd(stu) : Infinity };
  };
  const feeHtml = t => !t ? '–' : /^free$/i.test(t) ? '<span class="badge published">Free</span>'
    : t.split('\n').map(l => { const m = l.match(/^(regular|student)\s*:\s*(.*)$/i); return m ? `<span class="fl"><em>${esc(m[1])}</em> ${esc(m[2])}</span>` : `<span class="fl">${esc(l)}</span>`; }).join('');

  const renderVenues = wb => {
    const dirsOf = s => ({ dirs: dirCodes(s), all: /^\s*(all\b.*)?$/i.test(s) });
    CONF = findTable(wb, 'Conferences', h => h.includes('conference') && h.includes('website')).filter(r => r.conference).map((r, i) => {
      const deadline = pick(r, 'submission deadline', 'typical submission deadline'), heldText = pick(r, 'held in', 'typical conference month');
      return { i, name: r.conference, full: r['full name'], rank: r['core rank'], deadline, heldText, hybrid: r.hybrid, website: r.website,
        dl: monthsOf(deadline), held: monthsOf(heldText),
        hay: norm([r.conference, r['full name'], r.proceedings].join(' ')),
        ...dirsOf(pick(r, 'directions', 'research directions')), ...feeOf(pick(r, 'registration fee', 'fee')) };
    });
    JOUR = findTable(wb, 'Journals', h => h.some(x => /^journal( name)?$/.test(x)) && h.includes('rank'))
      .map(r => ({ name: pick(r, 'journal', 'journal name'), publisher: r.publisher, rank: r.rank.toUpperCase(), link: r.link, ...dirsOf(r.directions),
        hay: norm([pick(r, 'journal', 'journal name'), r.publisher].join(' ')),
        pub: txt(r.publisher).split('/')[0].trim(),   // main publisher, for the Publisher filter ("Springer / KOSOMBE" -> Springer)
        ttp: pick(r, 'time to publish'), ...(([a, b]) => ({ ttpMin: a, ttpMax: b }))(ttpOf(pick(r, 'time to publish'))),
        ...feeOf(pick(r, 'publication fee', 'apc', 'fee')) }))
      .filter(r => r.name && !/^notes?:?$/i.test(r.name) && !r.name.startsWith('•'));
    NOTES = findTable(wb, 'Venue Notes', h => h.includes('note') && h.includes('applies to')).filter(r => r.note).map(r => ({ to: r['applies to'] || 'All', note: r.note }));

    if (!CONF.length && !JOUR.length) { $('#venues').hidden = true; $('a[href="#venues"]').hidden = true; return; }
    if (!CONF.length) tab = 'jour';
    buildFields(); drawVenues();
  };

  const resetAll = () => { F[tab] = { ...DEF[tab] }; vpage = 1; buildFields(); drawVenues(); };
  $('#venues .seg').addEventListener('click', e => { const b = e.target.closest('button'); if (b && b.dataset.tab !== tab) { tab = b.dataset.tab; vpage = 1; buildFields(); drawVenues(); } });
  $('#vfields').addEventListener('change', e => { const k = e.target.dataset.k; if (k) { F[tab][k] = e.target.value; vpage = 1; drawVenues(); } });
  let qTimer;
  $('#vq').addEventListener('input', e => { clearTimeout(qTimer); qTimer = setTimeout(() => { F[tab].q = e.target.value.trim(); vpage = 1; drawVenues(); }, 150); });
  $('#venues').addEventListener('click', e => {
    const pill = e.target.closest('.pill[data-k]');
    if (pill) { const k = pill.dataset.k; F[tab][k] = DEF[tab][k]; if (k === 'q') $('#vq').value = ''; vpage = 1; drawVenues(); return; }
    if (e.target.closest('#vclear, #vclear2')) { resetAll(); return; }
    const sb = e.target.closest('.sortbtn');
    if (sb) { const s = SORT[tab]; s.dir = s.key === sb.dataset.sort ? -s.dir : 1; s.key = sb.dataset.sort; vpage = 1; drawVenues(); }
  });
  $('#vsort').addEventListener('change', e => { SORT[tab] = { key: e.target.value, dir: 1 }; vpage = 1; drawVenues(); });
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
