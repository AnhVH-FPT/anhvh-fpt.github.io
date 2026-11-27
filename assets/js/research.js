/* Research topics page: reads data/research.xlsx (one sheet per research direction, plus journals and conferences). */
(() => {
  'use strict';
  const root = document.documentElement;
  const $ = (s, c = document) => c.querySelector(s);
  const $$ = (s, c = document) => [...c.querySelectorAll(s)];
  const store = {
    get: k => { try { return localStorage.getItem(k); } catch { return null; } },
    set: (k, v) => { try { localStorage.setItem(k, v); } catch { /* storage unavailable */ } }
  };
  const DB_FILE = 'data/research.xlsx';
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
  // Each sheet may have title rows above the header: find the header row by its column names.
  const table = (ws, test) => {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });
    const h = rows.findIndex(r => test(r.map(c => txt(c).toLowerCase())));
    if (h < 0) return null;
    const head = rows[h].map(c => txt(c).toLowerCase());
    const items = rows.slice(h + 1).map(r => {
      const o = {}; head.forEach((k, i) => { if (k) o[k] = txt(r[i]); }); return o;
    });
    return { title: txt((rows[0] || [])[0]), subtitle: txt((rows[1] || [])[0]), items };
  };

  const CODES = [[/medical/i, 'MI'], [/explain|trust|xai/i, 'XAI'], [/iot|edge/i, 'EDGE'], [/language|vlm/i, 'VLM'], [/vision|image/i, 'CV']];
  const codeOf = name => (CODES.find(([re]) => re.test(name)) || [, name.replace(/[^A-Z]/g, '').slice(0, 3) || 'T'])[1];

  // "• Name\nhttps://..." -> [{name, url}]
  const datasets = s => {
    const out = [];
    txt(s).split('\n').map(l => l.trim()).filter(Boolean).forEach(l => {
      if (/^https?:\/\//.test(l)) { const last = out[out.length - 1]; if (last && !last.url) last.url = l; else out.push({ name: l.replace(/^https?:\/\/(www\.)?/, ''), url: l }); }
      else out.push({ name: l.replace(/^[•\-*]\s*/, ''), url: '' });
    });
    return out;
  };

  // "Label: text" lines become headed blocks, "(1) ..." lines become list items.
  const richText = s => {
    const blocks = []; let cur = null;
    txt(s).split('\n').map(l => l.trim()).filter(Boolean).forEach(l => {
      const m = l.match(/^([A-Z][A-Za-z /&-]{2,40}):\s*(.*)$/);
      if (m) { cur = { label: m[1], text: m[2], items: [] }; blocks.push(cur); return; }
      const it = l.match(/^\(?(\d+)[).]\s*(.*)$/);
      if (!cur) { cur = { label: '', text: '', items: [] }; blocks.push(cur); }
      if (it) cur.items.push(it[2]); else cur.text += (cur.text ? ' ' : '') + l;
    });
    return blocks.map(b => `<div class="blk">${b.label ? `<h5>${esc(b.label)}</h5>` : ''}${b.text ? `<p>${linkify(b.text)}</p>` : ''}${
      b.items.length ? `<ol>${b.items.map(i => `<li>${linkify(i)}</li>`).join('')}</ol>` : ''}</div>`).join('');
  };
  const firstSentence = s => { const m = txt(s).match(/Context:\s*([^\n]*)/i); return m ? m[1] : txt(s).split('\n')[0]; };

  /* ---------- topics ---------- */
  let TOPICS = [], DIRS = [], dir = 'all', hasStatus = false;
  const isClosed = t => /taken|assigned|closed|full|ongoing|in progress|đã/i.test(t.status);

  const mailto = t => {
    const subject = `[Research Topic ${t.code}] ${t.en}`;
    const body = `Dear Mr. Ha Anh Vu,\n\nI would like to apply for the research topic below.\n\nTopic: ${t.code} - ${t.en}\n(${t.vi})\n\nFull name:\nStudent ID:\nMajor and semester:\nTeam members (if any):\nWhy I am interested in this topic:\nRelevant experience (Python, PyTorch, courses, projects):\n\nI have attached my CV and transcript.\n\nBest regards,\n`;
    return `mailto:${EMAIL}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  };

  const card = t => {
    const closed = isClosed(t);
    const ds = datasets(t.dataset);
    return `<details class="topic${closed ? ' closed' : ''}" id="${t.code}">
      <summary>
        <span class="tcode">${t.code}</span>
        <div class="thead">
          <h3>${esc(t.en)}</h3>
          ${t.vi ? `<p class="vi">${esc(t.vi)}</p>` : ''}
          <p class="teaser">${esc(firstSentence(t.desc))}</p>
          <div class="tmeta">
            <span class="tdir">${esc(t.dir)}</span>
            ${hasStatus && t.status ? `<span class="badge ${closed ? 'thesis' : 'published'}">${esc(t.status)}</span>` : ''}
            ${ds.filter(d => d.url || d.name).slice(0, 3).map(d => `<span class="dtag">${esc(d.name.replace(/\s*\(.*?\)\s*/g, ' ').trim())}</span>`).join('')}
          </div>
        </div>
      </summary>
      <div class="tbody">
        <div class="tcols">
          <div><h4>Research description</h4>${richText(t.desc)}</div>
          <div><h4>Research expectation</h4>${richText(t.expect)}
            ${ds.length ? `<h4>Datasets</h4><ul class="dlist">${ds.map(d => `<li>${d.url ? `<a href="${esc(d.url)}" target="_blank" rel="noopener">${esc(d.name)}</a>` : esc(d.name)}</li>`).join('')}</ul>` : ''}
          </div>
        </div>
        <div class="tact">
          ${closed ? '<span class="muted">This topic is already taken.</span>'
            : `<a class="btn pri" href="${mailto(t)}">Apply for this topic</a>`}
          <button type="button" class="btn copy" data-code="${t.code}">Copy link</button>
        </div>
      </div>
    </details>`;
  };

  const norm = s => s.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd');
  const drawTopics = () => {
    const q = norm($('#q').value.trim()), only = $('#only').checked;
    const list = TOPICS.filter(t => (dir === 'all' || t.dir === dir) && (!only || !isClosed(t)) && (!q || t.hay.includes(q)));
    $('#list').innerHTML = list.map(card).join('') || '<p class="empty">No topic matches your search.</p>';
    $('#count').textContent = `${list.length} of ${TOPICS.length} topics`;
    $$('#dirs button').forEach(b => b.classList.toggle('on', b.dataset.dir === dir));
    const d = DIRS.find(x => x.name === dir);
    $('#dirfocus').innerHTML = d && d.focus ? `<strong>Current trends:</strong> ${esc(d.focus)}` : '';
  };

  const renderTopics = wb => {
    // Optional "Focus" text per direction from the Overview sheet
    const focus = {};
    const ov = wb.SheetNames.find(n => /overview/i.test(n));
    if (ov) XLSX.utils.sheet_to_json(wb.Sheets[ov], { header: 1, defval: '' }).forEach(r => {
      const cells = r.map(txt).filter(Boolean), last = cells[cells.length - 1] || '';
      if (cells.length >= 2 && last.length > 30) focus[cells[0]] = last;     // the count column may be an uncached formula
    });
    wb.SheetNames.forEach(n => {
      const t = table(wb.Sheets[n], h => h.includes('english title'));
      if (!t) return;
      const name = (t.title.split(/\s+[—–-]\s+/)[0] || n).trim();
      const code = codeOf(name);
      const short = name.replace(/\s*\(.*?\)/g, '');
      const fkey = Object.keys(focus).find(k => k.replace(/\s*\(.*?\)/g, '').toLowerCase() === short.toLowerCase());
      const items = t.items.filter(r => r['english title']).map((r, i) => ({
        code: code + '-' + String(r['no.'] || i + 1).padStart(2, '0'),
        dir: short, en: r['english title'], vi: r['vietnamese title'] || '', dataset: r.dataset || '',
        desc: r['research description'] || '', expect: r['research expectation'] || '', status: r.status || ''
      }));
      items.forEach(x => { x.hay = norm([x.code, x.en, x.vi, x.dataset, x.desc, x.expect].join(' ')); });
      DIRS.push({ name: short, count: items.length, focus: fkey ? focus[fkey] : '' });
      TOPICS.push(...items);
    });
    hasStatus = TOPICS.some(t => t.status);
    $('#onlyWrap').hidden = !hasStatus;
    $('#dirs').innerHTML = `<button type="button" data-dir="all">All directions <em>${TOPICS.length}</em></button>` +
      DIRS.map(d => `<button type="button" data-dir="${esc(d.name)}">${esc(d.name)} <em>${d.count}</em></button>`).join('');
    drawTopics();
    openFromHash();
  };

  const openFromHash = () => {
    const id = decodeURIComponent(location.hash.slice(1)); if (!id) return;
    let el = document.getElementById(id);
    if (!el && TOPICS.some(t => t.code === id)) { dir = 'all'; $('#q').value = ''; drawTopics(); el = document.getElementById(id); }
    if (el && el.classList.contains('topic')) { el.open = true; el.scrollIntoView({ block: 'start' }); }
  };
  addEventListener('hashchange', openFromHash);

  $('#dirs').addEventListener('click', e => { const b = e.target.closest('button'); if (b) { dir = b.dataset.dir; drawTopics(); } });
  let qt; $('#q').addEventListener('input', () => { clearTimeout(qt); qt = setTimeout(drawTopics, 120); });
  $('#only').addEventListener('change', drawTopics);
  $('#list').addEventListener('click', async e => {
    const b = e.target.closest('.copy'); if (!b) return;
    const url = location.href.split('#')[0] + '#' + b.dataset.code;
    try { await navigator.clipboard.writeText(url); b.textContent = 'Link copied'; } catch { prompt('Copy this link:', url); }
    setTimeout(() => { b.textContent = 'Copy link'; }, 1800);
  });

  /* ---------- venues ---------- */
  let CONF = [], JOUR = [], tab = 'conf', vf = 'all';
  const RANKS = { conf: r => r['core rank'] || 'Not ranked', jour: r => r.rank || '' };
  const drawVenues = () => {
    const data = tab === 'conf' ? CONF : JOUR, rk = RANKS[tab];
    const order = r => ({ 'A*': '0', A: '1', B: '2', C: '3' }[r] || (/^Q\d/.test(r) ? r : 'z' + r));   // A* first; "Not ranked" last
    const ranks = [...new Set(data.map(rk))].filter(Boolean).sort((a, b) => order(a).localeCompare(order(b)));
    if (!['all', ...ranks].includes(vf)) vf = 'all';
    $('#venueFilter').innerHTML = [`<button type="button" data-r="all">All <em>${data.length}</em></button>`,
      ...ranks.map(r => `<button type="button" data-r="${esc(r)}">${tab === 'conf' && /^[A-C]/.test(r) ? 'CORE ' : ''}${esc(r)} <em>${data.filter(x => rk(x) === r).length}</em></button>`)].join('');
    $$('#venueFilter button').forEach(b => b.classList.toggle('on', b.dataset.r === vf));
    $$('#venues .seg button').forEach(b => b.classList.toggle('on', b.dataset.tab === tab));
    const rows = data.filter(x => vf === 'all' || rk(x) === vf);
    const a = (u, l) => u ? `<a href="${esc(u)}" target="_blank" rel="noopener">${esc(l)}</a>` : esc(l);
    $('#venueBox').innerHTML = tab === 'conf'
      ? `<table><thead><tr><th>Conference</th><th>Rank</th><th>Directions</th><th>Deadline</th><th>Start preparing</th><th>Held in</th><th>Hybrid</th></tr></thead><tbody>${
        rows.map(r => `<tr><td><strong>${a(r.website, r.conference)}</strong><small>${esc(r['full name'])}</small></td><td>${esc(rk(r))}</td><td>${esc(r['research directions'])}</td><td>${esc(r['typical submission deadline'])}</td><td>${esc(r['suggested start of preparation'])}</td><td>${esc(r['typical conference month'])}</td><td>${/yes/i.test(r.hybrid) ? '<span class="badge published">Yes</span>' : esc(r.hybrid)}</td></tr>`).join('')}</tbody></table>`
      : `<table><thead><tr><th>Journal</th><th>Publisher</th><th>Rank</th></tr></thead><tbody>${
        rows.map(r => `<tr><td><strong>${a(r.link, r['journal name'])}</strong></td><td>${esc(r.publisher)}</td><td><span class="q q${esc(r.rank).slice(1)}">${esc(r.rank)}</span></td></tr>`).join('')}</tbody></table>`;
    if (!rows.length) $('#venueBox').innerHTML = '<p class="empty">No venues listed.</p>';
  };
  const renderVenues = wb => {
    wb.SheetNames.forEach(n => {
      const c = table(wb.Sheets[n], h => h.includes('conference') && h.includes('website'));
      if (c) CONF = c.items.filter(r => r.conference);
      const j = table(wb.Sheets[n], h => h.includes('journal name'));
      if (j) JOUR = j.items.filter(r => r['journal name']);
    });
    if (!CONF.length && !JOUR.length) { $('#venues').hidden = true; $('a[href="#venues"]').hidden = true; return; }
    if (!CONF.length) tab = 'jour';
    drawVenues();
  };
  $('#venues .seg').addEventListener('click', e => { const b = e.target.closest('button'); if (b) { tab = b.dataset.tab; vf = 'all'; drawVenues(); } });
  $('#venueFilter').addEventListener('click', e => { const b = e.target.closest('button'); if (b) { vf = b.dataset.r; drawVenues(); } });

  /* ---------- load ---------- */
  const show = wb => { renderTopics(wb); renderVenues(wb); };
  (async () => {
    try {
      if (!window.XLSX) throw new Error('The Excel reader library (SheetJS from cdnjs) did not load. Check your internet connection or ad blocker.');
      if (window.RESEARCH_BASE64) return show(XLSX.read(window.RESEARCH_BASE64, { type: 'base64' }));
      const res = await fetch(DB_FILE + '?v=' + Date.now());
      if (!res.ok) throw new Error(res.status === 404 ? 'data/research.xlsx was not found in the repository.' : 'HTTP ' + res.status);
      show(XLSX.read(await res.arrayBuffer(), { type: 'array' }));
    } catch (err) {
      console.error(err);
      const msg = location.protocol === 'file:'
        ? 'Opened from disk: browsers block reading data/research.xlsx. Run a local server (python -m http.server) or view the GitHub Pages site.'
        : 'Could not load the topic list: ' + err.message;
      $('#list').innerHTML = $('#venueBox').innerHTML = `<p class="empty">${esc(msg)}</p>`;
    }
  })();
})();
