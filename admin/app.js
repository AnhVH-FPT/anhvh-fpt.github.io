/* Database manager: edits data/database.xlsx through admin/server.py */
(() => {
'use strict';
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = s => String(s ?? '').normalize('NFKD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/&/g, ' and ').replace(/\(.*?\)/g, ' ').replace(/[^a-z0-9]+/g, ' ').trim();
const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const hasMonth = s => new RegExp('\\b(' + MONTHS.join('|') + ')', 'i').test(s || '');
const isUrl = s => /^https?:\/\/\S+$/i.test(s || '');
const DESC = ['Context', 'Problem formulation', 'Gap', 'Proposed approach', 'Implementation steps', 'Key knowledge to learn'];
const EXPC = ['Deliverables', 'Expected results', 'Baselines', 'Metrics', 'Experimental protocol', 'Minimum success criteria'];

/* ------------------------------------------------------------ per-sheet settings */
const linkIds = sheet => [...new Set(rowsOf(sheet).map(r => r['Link ID']).filter(Boolean))];
const CFG = {
  Publications: {
    icon: '📄', desc: 'Bài báo trong mục Publications của trang chủ.', key: 'Title', title: r => r.Title,
    cols: ['Type', 'Status', 'Date', 'Title', 'Venue'], filters: ['Type', 'Status'],
    required: ['Type', 'Status', 'Date', 'Title', 'Authors', 'Venue'],
    fields: {
      Date: { hint: 'dd-mm-yyyy, mm-yyyy hoặc chỉ yyyy. Không có năm thì bài bị ẩn.', ph: '15-03-2026' },
      Title: { long: 2, wide: 1 }, Authors: { wide: 1, hint: 'Cách nhau bằng dấu phẩy. "Ha Anh Vu" tự in đậm.' }, Venue: { wide: 1 },
      DOI: { hint: 'Mã DOI hoặc link đầy đủ.' }, PDF: { url: 1 }, Code: { url: 1 },
      'Link ID': { list: () => linkIds('Projects'), hint: 'Mã nối với đề tài trong Projects (Paper Status = Accepted).' },
    },
    site: () => 'index.html#publications',
  },
  Projects: {
    icon: '🎓', desc: 'Đề tài luận văn đang hướng dẫn, mục Undergraduate thesis projects.', key: 'Title', title: r => r.Title,
    cols: ['Semester', 'Title', 'Members', 'Thesis Status', 'Paper Status'], filters: ['Semester', 'Thesis Status', 'Paper Status'],
    required: ['Semester', 'Title', 'Thesis Status'],
    fields: {
      Semester: { list: () => [...new Set(rowsOf('Projects').map(r => r.Semester).filter(Boolean))], hint: 'Fall/Summer/Spring + năm, hoặc FA26, SU26, SP26.' },
      Title: { long: 2, wide: 1 }, Members: { long: 2, wide: 1, hint: 'Tên sinh viên, cách nhau bằng dấu chấm phẩy (;).' },
      'Paper Venue': { hint: 'Hiện sau "Paper submitted to".' },
      'Link ID': { list: () => linkIds('Publications'), hint: 'Khi Paper Status = Accepted: cùng mã với dòng bài báo trong Publications.' },
    },
    site: () => 'index.html#supervision',
  },
  Directions: {
    icon: '🧭', desc: 'Các hướng nghiên cứu. Thứ tự dòng = thứ tự thẻ, nút lọc và nhóm đề tài.', key: 'Code', title: r => r.Code + ' · ' + r.Direction, ordered: true,
    cols: ['Code', 'Direction', 'Summary'], required: ['Code', 'Direction'],
    fields: {
      Code: { hint: 'Chữ in hoa, không dấu, không khoảng trắng. Đổi mã sẽ cập nhật mọi tham chiếu.' },
      'Focus (current research trends)': { long: 3, wide: 1, hint: 'Hiện thành dòng "Current trends" khi chọn hướng.' },
      Summary: { long: 2, wide: 1, hint: 'Khoảng 10 từ tiếng Anh, hiện trên thẻ Research interests.' },
    },
    site: () => 'index.html#research',
  },
  Topics: {
    icon: '🧪', desc: 'Đề tài cho sinh viên, mục Research topics của research.html. Mã đề tài = hướng + vị trí dòng trong hướng.', key: 'English Title',
    title: r => r['English Title'], codeCol: true, filters: ['Direction', 'Status'],
    cols: ['Direction', 'Status', 'English Title', 'Dataset'], required: ['Direction', 'Status', 'English Title', 'Research Description', 'Research Expectation'],
    fields: {
      'Vietnamese Title': { long: 2, wide: 1 }, 'English Title': { long: 2, wide: 1 },
      Dataset: { datasets: 1, wide: 1, hint: 'Tên bắt buộc, link không bắt buộc. Trên web tên hiện thành link tới trang dataset.' },
      'Research Description': { big: 1, wide: 1, template: DESC },
      'Research Expectation': { big: 1, wide: 1, template: EXPC },
    },
    site: r => 'research.html#' + r._code,
  },
  Conferences: {
    icon: '🎤', desc: 'Hội nghị có kỷ yếu Scopus, tab Conferences của mục Where to publish.', key: 'Conference', title: r => r.Conference + (r['Full Name'] ? ' · ' + r['Full Name'] : ''),
    cols: ['Conference', 'Full Name', 'CORE Rank', 'Submission Deadline', 'Hybrid', 'Registration Fee'], filters: ['CORE Rank', 'Directions', 'Hybrid'],
    required: ['Conference'],
    fields: {
      'Full Name': { wide: 1 }, Directions: { multi: 1, wide: 1 },
      Proceedings: { hint: 'Phải được Scopus index. Dùng cho ô tìm kiếm.', list: () => [...new Set(rowsOf('Conferences').map(r => r.Proceedings).filter(Boolean))] },
      'Submission Deadline': { hint: 'Tên tháng tiếng Anh, ví dụ "Mid December – January".' },
      'Held In': { hint: 'Tháng tổ chức, ví dụ "October–November".' },
      Website: { url: 1, wide: 1, hint: 'Trang chính thức, hoặc trang dblp nếu web chỉ có trang từng năm.' },
      'Registration Fee': { long: 2, wide: 1, hint: 'Mỗi mức một dòng. ≈ = ước lượng.', quick: ['Regular: ≈USD ', 'Student: ≈USD ', 'Regular: USD ', 'Student: USD '] },
    },
    site: () => 'research.html#venues',
  },
  Journals: {
    icon: '📚', desc: 'Tạp chí, tab Journals của mục Where to publish.', key: 'Journal', title: r => r.Journal,
    cols: ['Journal', 'Publisher', 'Rank', 'Time to Publish', 'Publication Fee'], filters: ['Rank', 'Publisher', 'Directions'],
    required: ['Journal'],
    fields: {
      Journal: { wide: 1 }, Publisher: { list: () => [...new Set(rowsOf('Journals').map(r => r.Publisher).filter(Boolean))].sort() },
      Directions: { multi: 1, wide: 1 }, Link: { url: 1, wide: 1 },
      'Time to Publish': { hint: 'Số tháng từ lúc nộp đến khi đăng online, ví dụ "4–8 months".', quick: ['3–6 months', '4–8 months', '6–10 months', '8–14 months'] },
      'Publication Fee': { long: 2, hint: 'Mỗi cách đăng một dòng.', quick: ['Free (subscription)', 'Free', 'Open access: USD ', 'Open access: ≈USD '] },
    },
    site: () => 'research.html#venues',
  },
  'Venue Notes': {
    icon: '📝', desc: 'Ghi chú hiện dưới bảng hội nghị/tạp chí. Thứ tự dòng = thứ tự hiện.', key: null, title: r => (r.Note || '').slice(0, 80), ordered: true,
    cols: ['Applies To', 'Note'], filters: ['Applies To'], required: ['Applies To', 'Note'],
    fields: { Note: { long: 4, wide: 1 } },
    site: () => 'research.html#venues',
  },
};
const cfg = s => CFG[s] || { icon: '▦', desc: '', cols: DB.sheets[s].headers.slice(0, 5), required: [], fields: {}, title: r => Object.values(r)[0] };

/* ------------------------------------------------------------ state */
let DB = null, PRED = null, view = 'overview', editing = null, busy = false, lastVersion = '';
const UI = {};                              // per sheet: {q, filters, sort, sel:Set}
const ui = s => UI[s] || (UI[s] = { q: '', filters: {}, conds: [], sort: [], cols: null, panel: null, sel: new Set() });

const rowsOf = sheet => {
  const S = DB && DB.sheets[sheet]; if (!S) return [];
  if (S._objs) return S._objs;
  const counters = {};
  S._objs = S.rows.map((vals, i) => {
    const o = { _i: i };
    S.headers.forEach((h, c) => o[h] = vals[c] ?? '');
    if (sheet === 'Topics') { const d = o.Direction || '?'; counters[d] = (counters[d] || 0) + 1; o._code = d + '-' + String(counters[d]).padStart(2, '0'); }
    return o;
  });
  return S._objs;
};
const codes = () => rowsOf('Directions').map(r => r.Code).filter(Boolean);

/* ------------------------------------------------------------ server */
async function api(path, body) {
  const r = await fetch(path, body ? { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) } : {});
  let j = {}; try { j = await r.json(); } catch { }
  if (!r.ok) { const e = new Error(j.error || ('Lỗi ' + r.status)); e.status = r.status; throw e; }
  return j;
}
async function reload(quiet) {
  DB = await api('/api/db');
  lastVersion = DB.version;
  ISSUES = null;
  $('#file').textContent = DB.file;
  setSync('ok', 'Đã đồng bộ với database.xlsx');
  renderNav(); render();
  if (!quiet) hideBanner();
}
async function op(body, okMsg) {
  if (busy) return null;
  busy = true; setSync('', 'Đang lưu vào Excel…');
  try {
    const res = await api('/api/op', { ...body, version: DB.version });
    await reload(true);
    toast(okMsg || 'Đã lưu vào database.xlsx', 'ok');
    return res;
  } catch (e) {
    if (e.status === 409) { await reload(true); banner(e.message, 'warn'); }
    else if (e.status === 423) { setSync('warn', 'Excel đang mở file'); dialog('Không ghi được file', esc(e.message), [{ t: 'Đã hiểu', primary: 1 }]); }
    else { setSync('err', 'Lỗi khi lưu'); toast(e.message, 'err'); }
    return null;
  } finally { busy = false; }
}

/* ------------------------------------------------------------ checks */
let ISSUES = null;
function checkRow(sheet, r, all) {
  const C = cfg(sheet), S = DB.sheets[sheet], out = [];
  const add = (level, field, msg) => out.push({ level, field, msg });
  for (const h of C.required || []) if (!String(r[h] || '').trim()) add('err', h, 'Bắt buộc phải nhập.');
  for (const [h, opts] of Object.entries(S.options || {})) {
    if (!Array.isArray(opts) || !r[h]) continue;
    if (!opts.includes(r[h])) add('err', h, `Giá trị "${r[h]}" không có trong danh sách: ${opts.join(', ')}.`);
  }
  for (const [h, f] of Object.entries(C.fields || {})) if (f.url && r[h] && !isUrl(r[h])) add('warn', h, 'Link nên bắt đầu bằng https://');
  if (C.key && r[C.key]) {
    const dup = (all || rowsOf(sheet)).find(o => o._i !== r._i && norm(o[C.key]) === norm(r[C.key]));
    if (dup) add('warn', C.key, `Trùng với dòng ${dup._i + 2}.`);
  }
  const multiCheck = h => {
    if (!r[h]) return;
    const bad = r[h].split(',').map(s => s.trim()).filter(s => s && s !== 'All' && !codes().includes(s));
    if (bad.length) add('err', h, 'Mã hướng không tồn tại: ' + bad.join(', '));
  };
  if (sheet === 'Publications') {
    if (r.Date && !/\d{4}/.test(r.Date)) add('err', 'Date', 'Chưa có năm (4 chữ số): bài sẽ bị ẩn trên web.');
    if (r.Status === 'Preprint' && r.Type !== 'Preprint') add('warn', 'Type', 'Status = Preprint thì Type nên là Preprint.');
    if (r.Type === 'Preprint' && !['Preprint', 'In Submission'].includes(r.Status)) add('warn', 'Status', 'Bài Preprint nên có Status = Preprint (hoặc In Submission nếu đang nộp).');
    if (r.Type === 'Preprint' && !/arxiv[^0-9]{0,12}\d{4}\.\d{4,5}/i.test([r.Venue, r.PDF, r.DOI].join(' '))) add('warn', 'Venue', 'Preprint chưa có arXiv ID: dán ID vào ô arXiv ở đầu biểu mẫu.');
    if (r.Type === 'Thesis' && r.Status !== 'Thesis') add('warn', 'Status', 'Luận văn nên có Status = Thesis.');
    if (r['Link ID']) { const d = rowsOf(sheet).find(o => o._i !== r._i && o['Link ID'] && o['Link ID'].toLowerCase() === r['Link ID'].toLowerCase()); if (d) add('err', 'Link ID', `Link ID trùng với dòng ${d._i + 2}.`); }
  }
  if (sheet === 'Projects') {
    if (r.Semester && !(/spring|summer|fall|autumn/i.test(r.Semester) || /\b(sp|su|fa)\s*-?\d{2}/i.test(r.Semester))) add('warn', 'Semester', 'Không nhận ra học kỳ (Fall 2026, FA26…).');
    if (r['Paper Status'] === 'Accepted') {
      if (!r['Link ID']) add('warn', 'Link ID', 'Paper Status = Accepted nên cần Link ID để nối với bài báo.');
      else if (!linkIds('Publications').some(x => x.toLowerCase() === r['Link ID'].toLowerCase())) add('err', 'Link ID', 'Không có bài báo nào trong Publications có Link ID này.');
    }
    if (r['Paper Status'] === 'Submitted' && !r['Paper Venue']) add('warn', 'Paper Venue', 'Nên ghi nơi đã nộp bài.');
  }
  if (sheet === 'Directions' && r.Code && !/^[A-Z0-9]+$/.test(r.Code)) add('err', 'Code', 'Chỉ dùng chữ in hoa và số, không dấu, không khoảng trắng.');
  if (sheet === 'Topics') {
    if (r.Direction && !codes().includes(r.Direction)) add('err', 'Direction', 'Mã hướng không có trong sheet Directions.');
    for (const [h, labels] of [['Research Description', DESC], ['Research Expectation', EXPC]]) {
      if (!r[h]) continue;
      const t = tplCheck(r[h], labels);
      if (t.missing.length) add('err', h, 'Thiếu mục: ' + t.missing.join(', ') + '.');
      if (t.stray.length) add('warn', h, `${t.stray.length} dòng nằm ngoài khung mẫu (web sẽ bỏ qua): "${t.stray[0].slice(0, 60)}"`);
      if (h === 'Research Description' && t.fields.Context !== undefined) {
        const n = sentences(t.fields.Context);
        if (n < 3) add('warn', h, `Context mới có ${n} câu, cần ít nhất 3 câu.`);
      }
    }
    if (r.Dataset) {
      const ds = parseDatasets(r.Dataset);
      if (ds.some(d => !d.name)) add('err', 'Dataset', 'Có dataset chưa nhập tên.');
      const badUrl = ds.find(d => d.url && !isUrl(d.url));
      if (badUrl) add('warn', 'Dataset', `Link của "${badUrl.name}" nên bắt đầu bằng https://`);
    }
  }
  if (sheet === 'Conferences') {
    multiCheck('Directions');
    for (const h of ['Submission Deadline', 'Held In']) if (r[h] && !hasMonth(r[h])) add('warn', h, 'Không có tên tháng tiếng Anh: bộ lọc và sắp xếp sẽ bỏ qua.');
    if (r['Registration Fee']) {
      const bad = r['Registration Fee'].split('\n').filter(l => l.trim() && !/^(regular|student)\s*:/i.test(l.trim()));
      if (bad.length) add('warn', 'Registration Fee', 'Mỗi dòng nên bắt đầu bằng "Regular:" hoặc "Student:".');
      else if (!/\d/.test(r['Registration Fee'])) add('warn', 'Registration Fee', 'Chưa có số tiền.');
    }
    if (/waset|world academy of science/i.test((r['Full Name'] || '') + ' ' + (r.Website || ''))) add('err', 'Website', 'Có dấu hiệu hội nghị WASET (săn mồi).');
  }
  if (sheet === 'Journals') {
    multiCheck('Directions');
    if (r['Time to Publish'] && !/\d/.test(r['Time to Publish'])) add('warn', 'Time to Publish', 'Nên ghi số tháng, ví dụ "4–8 months".');
    if (r['Publication Fee']) {
      const bad = r['Publication Fee'].split('\n').filter(l => l.trim() && !/^(free|open access)/i.test(l.trim()));
      if (bad.length) add('warn', 'Publication Fee', 'Mỗi dòng nên là "Free (subscription)", "Free" hoặc "Open access: USD …".');
    }
    if (PRED) {
      if (r.Journal && PRED.journals.includes(norm(r.Journal))) add('err', 'Journal', 'Tạp chí này có trong danh sách tạp chí săn mồi.');
      if (r.Publisher && PRED.publishers.includes(norm(r.Publisher))) add('err', 'Publisher', 'Nhà xuất bản này có trong danh sách nhà xuất bản săn mồi.');
      const d = host(r.Link); if (d && PRED.domains.some(x => d === x || d.endsWith('.' + x))) add('err', 'Link', 'Tên miền này thuộc danh sách săn mồi (' + d + ').');
    }
  }
  return out;
}
const host = u => { try { return new URL(u).hostname.replace(/^www\./, ''); } catch { return ''; } };
const sentences = s => (String(s).replace(/\b(e\.g|i\.e|etc|vs|et al)\./gi, '$1').match(/[^.!?]+[.!?]+(\s|$)/g) || []).length;
function tplCheck(text, labels) {
  const fields = {}, stray = []; let cur = null;
  const re = new RegExp('^\\s*(' + labels.map(l => l.replace(/ /g, '\\s+')).join('|') + ')\\s*:\\s*(.*)$', 'i');
  for (const line of String(text).split('\n')) {
    const m = line.match(re);
    if (m) { cur = labels.find(l => l.toLowerCase() === m[1].replace(/\s+/g, ' ').toLowerCase()); fields[cur] = m[2]; }
    else if (cur) fields[cur] += '\n' + line;
    else if (line.trim()) stray.push(line.trim());
  }
  return { fields, stray, missing: labels.filter(l => !(l in fields) || !fields[l].trim()) };
}
// Dataset cell = "• Name" line, then (optional) its URL line, repeated. Same reading as research.js.
function parseDatasets(text) {
  const out = [];
  String(text || '').split('\n').map(l => l.trim()).filter(Boolean).forEach(l => {
    if (/^https?:\/\//i.test(l)) { const last = out[out.length - 1]; if (last && !last.url) last.url = l; else out.push({ name: '', url: l }); }
    else out.push({ name: l.replace(/^[•\-*]\s*/, ''), url: '' });
  });
  return out;
}
const joinDatasets = list => list.filter(d => d.name.trim() || d.url.trim())
  .map(d => [d.name.trim() ? '• ' + d.name.trim() : '', d.url.trim()].filter(Boolean).join('\n')).join('\n');

function allIssues() {
  if (ISSUES) return ISSUES;
  ISSUES = {};
  for (const s of DB.order) { const rows = rowsOf(s); ISSUES[s] = rows.map(r => checkRow(s, r, rows)); }
  return ISSUES;
}

/* ------------------------------------------------------------ chrome */
function setSync(cls, text) { const el = $('#sync'); el.className = 'sync ' + cls; el.querySelector('span').textContent = text; }
function banner(text, cls) { const b = $('#banner'); b.className = 'banner ' + (cls || ''); b.innerHTML = esc(text) + '<button class="btn sm" type="button">Đóng</button>'; b.hidden = false; b.querySelector('button').onclick = hideBanner; }
function hideBanner() { $('#banner').hidden = true; }
function toast(text, cls) {
  const t = document.createElement('div'); t.className = 'toast ' + (cls || ''); t.textContent = text;
  $('#toasts').append(t); setTimeout(() => t.remove(), cls === 'err' ? 6000 : 2600);
}
function dialog(title, html, buttons) {
  return new Promise(res => {
    $('#mtitle').textContent = title; $('#mtext').innerHTML = html;
    const act = $('#mact'); act.innerHTML = '';
    for (const b of buttons) {
      const el = document.createElement('button'); el.type = 'button';
      el.className = 'btn ' + (b.primary ? 'primary' : b.danger ? 'dangerfill' : '');
      el.textContent = b.t; el.onclick = () => { $('#modal').hidden = true; res(b.v); };
      act.append(el);
    }
    $('#modal').hidden = false; act.lastChild.focus();
  });
}
const confirmBox = (title, html, yes, danger) => dialog(title, html, [{ t: 'Hủy', v: false }, { t: yes, v: true, primary: !danger, danger }]);

function renderNav() {
  const iss = allIssues();
  const count = s => iss[s].reduce((n, a) => n + (a.some(x => x.level === 'err') ? 1 : 0), 0);
  const total = DB.order.reduce((n, s) => n + count(s), 0);
  $('#nav').innerHTML = `<a href="#overview" class="${view === 'overview' ? 'on' : ''}"><span>◈</span> Tổng quan ${total ? `<span class="bad">${total}</span>` : ''}</a>
    <div class="grp">Sheets</div>` + DB.order.map(s => {
    const c = count(s);
    return `<a href="#${encodeURIComponent(s)}" class="${view === s ? 'on' : ''}"><span>${cfg(s).icon}</span> ${esc(s)}${c ? `<span class="bad" title="${c} dòng có lỗi">${c}</span>` : ''}<span class="n">${rowsOf(s).length}</span></a>`;
  }).join('');
}

function render() {
  const top = $('#topact');
  if (view === 'overview' || !DB.sheets[view]) {
    view = 'overview';
    $('#title').textContent = 'Tổng quan'; $('#subtitle').textContent = 'Mọi thay đổi ở đây được ghi thẳng vào data/database.xlsx.';
    top.innerHTML = ''; renderOverview(); return;
  }
  const C = cfg(view);
  $('#title').textContent = view; $('#subtitle').textContent = C.desc;
  top.innerHTML = `<a class="btn" target="_blank" rel="noopener" href="/site/${esc(C.site ? C.site({}).replace(/#.*/, '') : 'index.html')}">Xem trên web ↗</a>
    <button class="btn primary" type="button" id="addbtn">＋ Thêm dòng</button>`;
  $('#addbtn').onclick = () => openEditor(view, null);
  renderSheet();
}

/* ------------------------------------------------------------ user guide (Overview page) */
const GUIDE_HTML = `
<div class="panel guide"><h3>Hướng dẫn quản lý database</h3>
<p>Bấm vào từng mục để mở. Mọi thao tác Lưu ở đây ghi thẳng vào <code>data/database.xlsx</code>, đúng file mà website đọc.</p>

<details open><summary>1. Quy trình cập nhật website (5 bước)</summary>
<ol class="steps">
  <li><b>Mở công cụ:</b> nhấp đúp <code>admin\\Open database manager.bat</code> trong thư mục website (thư mục clone từ GitHub). Trình duyệt mở <code>http://localhost:8765/</code>. Giữ cửa sổ đen (server) mở trong lúc làm việc.</li>
  <li><b>Đóng Excel</b> nếu đang mở file database.xlsx (Windows khóa file đang mở, công cụ sẽ không lưu được). Chấm màu ở góc dưới bên trái báo trạng thái: xanh = sẵn sàng, vàng = Excel đang mở file.</li>
  <li><b>Sửa dữ liệu:</b> chọn sheet ở cột bên trái → bấm một dòng để sửa, hoặc “＋ Thêm dòng” → bấm <b>Lưu</b> (Ctrl + S).</li>
  <li><b>Xem trước:</b> bấm “Xem website (dữ liệu hiện tại)” ở góc dưới bên trái, hoặc “Xem trên web” trong biểu mẫu, để thấy trang web với dữ liệu vừa sửa.</li>
  <li><b>Đưa lên mạng:</b> mở GitHub Desktop → thấy file <code>data/database.xlsx</code> thay đổi → ghi tóm tắt (ví dụ “Thêm đề tài MI”) → <b>Commit to main</b> → <b>Push origin</b>. Website cập nhật sau khoảng 1–2 phút.</li>
</ol></details>

<details><summary>2. Bố cục màn hình</summary>
<ul class="steps">
  <li><b>Cột trái:</b> Tổng quan và 7 sheet, kèm số dòng. Số trong ô đỏ = số dòng đang có lỗi.</li>
  <li><b>Tổng quan:</b> số liệu nhanh, danh sách <b>Kiểm tra dữ liệu</b> (bấm vào một mục để mở đúng dòng cần sửa) và hướng dẫn này.</li>
  <li><b>Góc dưới trái:</b> trạng thái kết nối, nút <b>Hoàn tác thao tác cuối</b>, nút xem website, đường dẫn file đang sửa.</li>
  <li><b>Biểu mẫu sửa</b> mở từ bên phải. Ô có dấu <span style="color:var(--err)">*</span> là bắt buộc. Viền đỏ = lỗi, viền cam = cảnh báo; khung tóm tắt ở đầu biểu mẫu liệt kê tất cả.</li>
</ul></details>

<details><summary>3. Tìm kiếm, lọc, sắp xếp, chọn cột</summary>
<ul class="steps">
  <li><b>Tìm kiếm</b> (phím <kbd>/</kbd>): tìm trong mọi cột, không phân biệt dấu và hoa thường. Ở Topics tìm được cả mã đề tài (MI-07).</li>
  <li><b>Lọc nhanh:</b> các ô chọn trên thanh công cụ (ví dụ Direction, Status, Rank). Ô “Kiểm tra” chỉ hiện các dòng có lỗi hoặc cảnh báo.</li>
  <li><b>Lọc nâng cao:</b> thêm nhiều điều kiện, dòng phải thỏa <b>tất cả</b>. Phép so sánh: chứa, không chứa, là, không là, trống, không trống, ≥, ≤. Phép ≥/≤ hiểu đúng nghĩa: tháng (“Submission Deadline ≥ Mar”), ngày (“Date ≥ 2025”), học kỳ, hạng (A* &gt; A &gt; B), phí (đổi ra USD).</li>
  <li><b>Sắp xếp:</b> bấm tiêu đề cột (bấm lại để đảo ↑/↓, lần ba để bỏ). Giữ <kbd>Shift</kbd> khi bấm để thêm tầng 2, 3. Hoặc dùng nút “Sắp xếp” để chọn tối đa 3 tầng. Ô trống luôn nằm cuối. “Về thứ tự trong Excel” bỏ sắp xếp.</li>
  <li><b>Cột:</b> chọn cột nào hiện trong bảng (chỉ đổi cách xem, không đổi file).</li>
  <li>Các lọc và sắp xếp đang bật hiện thành nhãn có dấu ×; “Xóa hết lọc &amp; sắp xếp” để về mặc định. Mỗi sheet nhớ cách xem riêng.</li>
  <li><b>Sửa hàng loạt:</b> tick ô vuông đầu dòng (hoặc ô ở tiêu đề để chọn hết các dòng đang hiện) → thanh “Đã chọn” cho phép đặt cùng một giá trị (ví dụ Status = Taken) hoặc xóa nhiều dòng.</li>
</ul></details>

<details><summary>4. Thêm, sửa, nhân bản, xóa, hoàn tác</summary>
<ul class="steps">
  <li><b>Thêm:</b> “＋ Thêm dòng” → điền → “Thêm vào Excel”. Dòng mới luôn nằm cuối sheet.</li>
  <li><b>Sửa:</b> bấm vào dòng → sửa → <b>Lưu</b> (Ctrl + S). <kbd>Esc</kbd> để đóng; nếu còn thay đổi chưa lưu, công cụ sẽ hỏi lại.</li>
  <li><b>Nhân bản:</b> mở dòng → “Nhân bản” tạo biểu mẫu mới với cùng nội dung (tiện khi thêm hội nghị/tạp chí tương tự).</li>
  <li><b>Xóa:</b> mở dòng → “Xóa”. Đề tài thì nên <b>đánh dấu Taken</b> thay vì xóa (xem mục Topics).</li>
  <li><b>Hoàn tác:</b> trước mỗi lần lưu, bản cũ được sao lưu vào <code>admin\\backups</code> (giữ 60 bản gần nhất). Nút “Hoàn tác thao tác cuối” đưa file về trước lần lưu gần nhất; bấm nhiều lần để lùi tiếp.</li>
  <li>Nếu bạn sửa file trong Excel trong lúc công cụ đang mở, công cụ tự tải lại dữ liệu mới. Nếu cùng một dòng vừa bị đổi ở nơi khác, công cụ không ghi đè mà tải lại để bạn kiểm tra.</li>
</ul></details>

<details><summary>5. Publications (bài báo)</summary>
<ul class="steps">
  <li><b>Type</b>: Journal, Conference, Thesis, Preprint. <b>Status</b>: Published, Accepted, In Submission, Preprint, Thesis.</li>
  <li><b>Bài đang nộp/phản biện:</b> Type = nơi nộp (Journal hoặc Conference), Status = In Submission. Trên web nằm ở nhóm “In submission” (S1, S2…).</li>
  <li><b>Preprint (arXiv):</b> chọn Type = Preprint → ô “arXiv ID hoặc link” hiện ra → dán link arXiv: Venue, PDF, DOI tự điền. Bấm “Lấy tên bài, tác giả, ngày từ arXiv” để điền nốt các ô còn trống. Web hiện “arXiv preprint arXiv:…” có link, nhóm “Preprints” (P1, P2…).</li>
  <li><b>Date:</b> dd-mm-yyyy, mm-yyyy hoặc chỉ yyyy. Bài không có năm sẽ bị ẩn khỏi web.</li>
  <li><b>Authors:</b> cách nhau bằng dấu phẩy; “Ha Anh Vu” tự in đậm.</li>
  <li><b>Link ID:</b> mã ngắn (ví dụ AC1) để nối bài báo với đề tài luận văn trong Projects (xem mục 6).</li>
  <li>Web tự sắp xếp (mới nhất trước) và gom các bài cùng hội nghị; bạn không cần sắp thứ tự dòng.</li>
</ul></details>

<details><summary>6. Projects (đề tài luận văn đang hướng dẫn)</summary>
<ul class="steps">
  <li><b>Semester:</b> Fall/Summer/Spring + năm (Fall 2026) hoặc FA26, SU26, SP26. Web tự xếp học kỳ mới nhất lên trước.</li>
  <li><b>Members:</b> tên sinh viên cách nhau bằng dấu chấm phẩy (;).</li>
  <li><b>Thesis Status:</b> Preparing → In progress → Defended.</li>
  <li><b>Paper Status:</b> None, Submitted (ghi thêm Paper Venue), Accepted. Khi Accepted: thêm bài vào Publications, đặt <b>cùng một Link ID</b> ở cả hai dòng; web hiện “Accepted, see C1” và bấm vào sẽ nhảy tới bài báo. Công cụ báo lỗi nếu Link ID không khớp.</li>
</ul></details>

<details><summary>7. Directions (hướng nghiên cứu)</summary>
<ul class="steps">
  <li>Thứ tự dòng = thứ tự thẻ Research interests, nút lọc và nhóm đề tài. Mở một hướng → nút ↑ ↓ để đổi thứ tự.</li>
  <li><b>Code</b> viết hoa, không dấu (CV, MI…). <b>Đổi mã</b> trong công cụ sẽ tự cập nhật mọi đề tài, hội nghị, tạp chí dùng mã đó (mã đề tài trên web đổi theo, ví dụ CV-01 → CVX-01).</li>
  <li><b>Summary</b> (khoảng 10 từ tiếng Anh) hiện trên thẻ ở trang chủ; <b>Focus</b> hiện thành dòng “Current trends”.</li>
  <li>Không nên xóa hướng còn đề tài: các đề tài đó sẽ thành một nhóm lạ trên web.</li>
</ul></details>

<details><summary>8. Topics (đề tài cho sinh viên)</summary>
<ul class="steps">
  <li><b>Mã đề tài</b> (CV-01, MI-07…) do web tự đánh theo vị trí dòng trong hướng. Vì vậy: <b>thêm đề tài mới vào cuối</b>, và khi có người nhận thì <b>đặt Status = Taken</b> (nút “Đánh dấu Taken”) thay vì xóa; xóa sẽ làm đổi mã các đề tài phía dưới. Biểu mẫu hiện trước mã mà đề tài sẽ có.</li>
  <li>Đề tài Taken không hiện trong danh sách chính, chỉ trong mục “Taken” trên web.</li>
  <li><b>Dataset:</b> mỗi dòng gồm Tên (bắt buộc) và Link (không bắt buộc); “＋ Thêm dataset” để thêm. Dán “Tên + link” vào ô tên sẽ tự tách. Web hiện tên thành link.</li>
  <li><b>Research Description</b> phải đủ 6 mục, mỗi mục bắt đầu bằng đúng tên mục và dấu hai chấm ở đầu dòng: Context, Problem formulation, Gap, Proposed approach, Implementation steps, Key knowledge to learn. <b>Research Expectation</b>: Deliverables, Expected results, Baselines, Metrics, Experimental protocol, Minimum success criteria. Dấu ✓/✗ dưới ô cho biết mục nào còn thiếu; nút “Chèn khung mẫu còn thiếu” thêm sẵn tên mục.</li>
  <li><b>Context</b> là đoạn tóm tắt trên thẻ đề tài: ít nhất 3 câu (bối cảnh, khó khăn hiện tại, sinh viên sẽ làm gì). Góc phải ô hiện số từ và số câu.</li>
  <li>Các bước hoặc kết quả đánh số: xuống dòng sau dấu hai chấm, mỗi ý một dòng bắt đầu bằng (1), (2), (3)…</li>
</ul></details>

<details><summary>9. Conferences (hội nghị)</summary>
<ul class="steps">
  <li>Chỉ thêm hội nghị có kỷ yếu được <b>Scopus</b> index; ghi nơi xuất bản ở Proceedings (cột này dùng cho ô tìm kiếm trên web). Công cụ cảnh báo hội nghị có dấu hiệu WASET.</li>
  <li><b>CORE Rank:</b> theo bảng ICORE2026 (portal.core.edu.au); không có hạng A*–C thì chọn Not ranked.</li>
  <li><b>Directions:</b> bấm chọn mã hướng, hoặc All.</li>
  <li><b>Submission Deadline / Held In:</b> ghi tên tháng tiếng Anh (“Mid December – January”, “October–November”) để bộ lọc và sắp xếp trên web hiểu được.</li>
  <li><b>Hybrid:</b> Yes = cho trình bày online; No = phải đến trực tiếp.</li>
  <li><b>Website:</b> trang chính thức; nếu hội nghị chỉ có trang từng năm thì dùng trang dblp.</li>
  <li><b>Registration Fee:</b> mỗi mức một dòng, dùng nút điền nhanh: “Regular: ≈USD 900” và “Student: ≈USD 600”. Có ≈ = ước lượng; không có ≈ = số đọc từ trang đăng ký.</li>
</ul></details>

<details><summary>10. Journals (tạp chí)</summary>
<ul class="steps">
  <li>Công cụ đối chiếu tên tạp chí, nhà xuất bản và tên miền với hai danh sách săn mồi trong <code>admin\\predatory</code> (có danh sách mới thì thay hai file ở đó, giữ nguyên tên file) và báo lỗi nếu trùng. Đừng lưu tạp chí bị báo lỗi này.</li>
  <li><b>Rank:</b> Q1–Q4 theo SCImago.</li>
  <li><b>Time to Publish:</b> số tháng từ lúc nộp đến khi đăng online, ví dụ “4–8 months” (có nút điền nhanh).</li>
  <li><b>Publication Fee:</b> mỗi cách đăng một dòng: “Free (subscription)” (hybrid, đăng miễn phí), “Free” (open access không thu phí), “Open access: USD 3,270” (phí APC). Tạp chí hybrid ghi cả hai dòng.</li>
</ul></details>

<details><summary>11. Venue Notes</summary>
<ul class="steps">
  <li>Ghi chú hiện dưới bảng hội nghị/tạp chí. Applies To = Conferences, Journals hoặc All. Thứ tự dòng = thứ tự hiện; dùng ↑ ↓ để đổi.</li>
</ul></details>

<details><summary>12. Phím tắt</summary>
<ul class="steps">
  <li><kbd>/</kbd> tìm kiếm · <kbd>Ctrl</kbd> + <kbd>S</kbd> lưu biểu mẫu · <kbd>Esc</kbd> đóng biểu mẫu/hộp thoại · <kbd>Shift</kbd> + bấm tiêu đề cột: thêm tầng sắp xếp.</li>
</ul></details>

<details><summary>13. Khi gặp sự cố</summary>
<ul class="steps">
  <li><b>“Không ghi được file … đang mở trong Excel”:</b> đóng Excel (lưu hoặc bỏ thay đổi trong Excel) rồi bấm Lưu lại.</li>
  <li><b>“Mất kết nối”:</b> cửa sổ đen đã bị đóng. Nhấp đúp lại file .bat rồi tải lại trang.</li>
  <li><b>Trang không đổi sau khi cập nhật công cụ:</b> bấm Ctrl + F5.</li>
  <li><b>Lưu nhầm:</b> bấm “Hoàn tác thao tác cuối”. Các bản cũ hơn nằm trong <code>admin\\backups</code> (mở bằng Excel để lấy lại dữ liệu).</li>
  <li><b>Web chưa đổi sau khi Push:</b> đợi 1–2 phút rồi tải lại trang (Ctrl + F5).</li>
  <li>Sheet “Hướng dẫn” trong file Excel giải thích chi tiết từng cột nếu bạn muốn sửa trực tiếp trong Excel.</li>
</ul></details>
</div>`;

/* ------------------------------------------------------------ overview */
function renderOverview() {
  const iss = allIssues();
  const t = rowsOf('Topics'), open = t.filter(r => r.Status !== 'Taken').length;
  const card = (s, v, sub) => `<a class="card" href="#${encodeURIComponent(s)}"><div class="k">${cfg(s).icon} ${esc(s)}</div><div class="v">${v}</div><div class="s">${sub}</div></a>`;
  const pubs = rowsOf('Publications'), proj = rowsOf('Projects');
  const list = [];
  for (const s of DB.order) iss[s].forEach((a, i) => a.forEach(x => list.push({ s, i, ...x })));
  list.sort((a, b) => (a.level === 'err' ? 0 : 1) - (b.level === 'err' ? 0 : 1));
  const errs = list.filter(x => x.level === 'err').length;
  $('#view').innerHTML = `<div class="cards">
      ${card('Topics', t.length, `${open} Open · ${t.length - open} Taken`)}
      ${card('Conferences', rowsOf('Conferences').length, 'Hội nghị Scopus')}
      ${card('Journals', rowsOf('Journals').length, 'Tạp chí')}
      ${card('Publications', pubs.length, pubs.filter(p => p.Status === 'Published').length + ' Published')}
      ${card('Projects', proj.length, [...new Set(proj.map(p => p.Semester))].length + ' học kỳ')}
      ${card('Directions', rowsOf('Directions').length, codes().join(' · '))}
    </div>
    <div class="panel"><h3>Kiểm tra dữ liệu</h3>
      <p>${list.length ? `${errs} lỗi, ${list.length - errs} cảnh báo. Bấm vào một dòng để sửa.` : 'Không có lỗi hay cảnh báo nào. Dữ liệu sẵn sàng để đưa lên web.'}</p>
      <ul class="issues">${list.slice(0, 150).map(x => {
        const r = rowsOf(x.s)[x.i];
        return `<li><span class="tag ${x.level}">${x.level === 'err' ? 'Lỗi' : 'Cảnh báo'}</span>
          <span><a data-go="${esc(x.s)}|${x.i}">${esc(x.s)} · ${esc(r._code || 'dòng ' + (x.i + 2))}</a> <span class="where">(${esc(x.field)})</span><br>${esc(x.msg)}</span></li>`;
      }).join('')}</ul>${list.length > 150 ? `<p>… và ${list.length - 150} mục khác.</p>` : ''}</div>
    ${GUIDE_HTML}`;
  $('#view').querySelectorAll('[data-go]').forEach(a => a.onclick = () => { const [s, i] = a.dataset.go.split('|'); go(s); openEditor(s, +i); });
}

/* ------------------------------------------------------------ sheet table */
// Values sort by meaning, not text: dropdown order (A* > A > B, Q1 > Q4, Open/Taken), dates, semesters, months, fees.
const SEASON = { spring: 1, sp: 1, summer: 2, su: 2, fall: 3, fa: 3, autumn: 3 };
const MONTH_COLS = ['Submission Deadline', 'Held In'];
const NUM_COLS = ['Registration Fee', 'Publication Fee', 'Time to Publish'];
const RATE = { USD: 1, EUR: 1.1, GBP: 1.3, AUD: 0.66, NZD: 0.6, CAD: 0.73, SGD: 0.75, CNY: 0.14, JPY: 0.0067, KRW: 0.00073, INR: 0.012, VND: 0.00004 };
function sortKey(sheet, h, v) {
  if (h === '#') return v;
  v = String(v ?? '').trim();
  if (!v) return null;
  const opts = DB.sheets[sheet].options[h];
  if (Array.isArray(opts) && opts.includes(v)) return opts.indexOf(v);
  if (h === 'Date') {
    const y = v.match(/\d{4}/); if (!y) return null;
    const rest = v.replace(y[0], '').match(/\d{1,2}/g) || [];
    const [a, b] = rest.map(Number);                       // dd-mm or mm
    const m = rest.length >= 2 ? b : rest.length ? a : 0, d = rest.length >= 2 ? a : 0;
    return +y[0] * 10000 + m * 100 + d;
  }
  if (h === 'Semester') {
    const w = v.match(/spring|summer|fall|autumn|\b(sp|su|fa)(?=\s*-?\d{2})/i), y = v.match(/\d{4}/) || v.match(/\d{2}/);
    return y ? (y[0].length === 2 ? 2000 + +y[0] : +y[0]) * 10 + (w ? SEASON[w[0].toLowerCase()] : 0) : null;
  }
  if (MONTH_COLS.includes(h)) {
    const m = v.toLowerCase().match(new RegExp('(early|mid|late)?[\\s–-]*(' + MONTHS.join('|') + ')'));
    return m ? MONTHS.indexOf(m[2]) * 3 + ({ early: 0, mid: 1, late: 2 }[m[1]] ?? 0) : null;
  }
  if (h === 'Publication Fee' && /^free/im.test(v) && !/open access/i.test(v)) return 0;
  if (h === 'Registration Fee' || h === 'Publication Fee') {          // first amount, converted to USD (same rough rates as research.js)
    const m = v.replace(/,/g, '').match(/\b(USD|EUR|GBP|AUD|NZD|CAD|SGD|CNY|JPY|KRW|INR|VND)\s*([\d.]+)/i);
    if (m) return Math.round(+m[2] * RATE[m[1].toUpperCase()]);
  }
  if (NUM_COLS.includes(h)) { const n = v.replace(/,/g, '').match(/\d+(\.\d+)?/); return n ? +n[0] : null; }
  return v.toLowerCase();
}
const cmp = (a, b) => (typeof a === 'number' && typeof b === 'number') ? a - b : String(a).localeCompare(String(b), 'vi', { numeric: true });

const OPS = [['contains', 'chứa'], ['ncontains', 'không chứa'], ['eq', 'là'], ['neq', 'không là'], ['empty', 'trống'], ['nempty', 'không trống'], ['gte', '≥ (từ)'], ['lte', '≤ (đến)']];
const opLabel = o => (OPS.find(x => x[0] === o) || [, o])[1];
function testCond(sheet, r, c) {
  const raw = c.h === '_code' ? r._code : r[c.h], v = String(raw ?? '').trim();
  const hay = norm(v), w = norm(c.v);
  const multi = cfg(sheet).fields?.[c.h]?.multi;
  const parts = multi ? (v || 'All').split(',').map(s => s.trim()) : null;
  switch (c.op) {
    case 'contains': return hay.includes(w);
    case 'ncontains': return !hay.includes(w);
    case 'eq': return multi ? parts.includes(c.v) || (c.v !== 'All' && parts.includes('All')) : v === c.v;
    case 'neq': return multi ? !parts.includes(c.v) : v !== c.v;
    case 'empty': return !v;
    case 'nempty': return !!v;
    case 'gte': case 'lte': {
      const a = sortKey(sheet, c.h, v), b = sortKey(sheet, c.h, c.v);
      if (a == null || b == null) return false;
      return c.op === 'gte' ? cmp(a, b) >= 0 : cmp(a, b) <= 0;
    }
  }
  return true;
}

const PREF = 'dbm:';
function loadPrefs(sheet, st) {
  try { const p = JSON.parse(localStorage.getItem(PREF + sheet) || '{}'); Object.assign(st, { conds: p.conds || [], sort: p.sort || [], cols: p.cols || null, filters: p.filters || {} }); } catch { }
}
function savePrefs(sheet) {
  const st = ui(sheet);
  try { localStorage.setItem(PREF + sheet, JSON.stringify({ conds: st.conds, sort: st.sort, cols: st.cols, filters: st.filters })); } catch { }
}

function filtered(sheet) {
  const st = ui(sheet), H = DB.sheets[sheet].headers;
  const q = norm(st.q).split(' ').filter(Boolean);
  let rows = rowsOf(sheet).filter(r => {
    for (const [h, v] of Object.entries(st.filters)) {
      if (!v || h === '_issues') continue;
      if (!testCond(sheet, r, { h, op: 'eq', v })) return false;
    }
    if (st.filters._issues) { const a = allIssues()[sheet][r._i]; if (!a.some(x => st.filters._issues === 'any' || x.level === st.filters._issues)) return false; }
    for (const c of st.conds) if (c.h && (c.op === 'empty' || c.op === 'nempty' || c.v) && !testCond(sheet, r, c)) return false;
    if (!q.length) return true;
    const hay = norm(H.map(h => r[h]).join(' ') + ' ' + (r._code || ''));
    return q.every(w => hay.includes(w));
  });
  const levels = st.sort.filter(s => s.h);
  if (levels.length) {
    rows = [...rows].sort((a, b) => {
      for (const { h, dir } of levels) {
        const x = sortKey(sheet, h, h === '#' ? a._i : a[h]), y = sortKey(sheet, h, h === '#' ? b._i : b[h]);
        if (x == null && y == null) continue;
        if (x == null) return 1; if (y == null) return -1;      // blanks always last
        const c = cmp(x, y); if (c) return c * dir;
      }
      return a._i - b._i;
    });
  }
  return rows;
}

function cellHtml(sheet, h, v) {
  if (!v) return '<span class="dim">—</span>';
  const opts = DB.sheets[sheet].options[h];
  if (h === 'Status' || h === 'Thesis Status' || h === 'Paper Status' || h === 'Hybrid') {
    const cls = /^(open|published|defended|accepted|yes)$/i.test(v) ? 'ok' : /^(taken|no|none)$/i.test(v) ? 'grey' : 'warn';
    return `<span class="tag ${cls}">${esc(v)}</span>`;
  }
  if (Array.isArray(opts) || h === 'Directions' || h === 'Rank') return v.split(',').map(s => `<span class="tag">${esc(s.trim())}</span>`).join('');
  if (h === 'Dataset') return `<span class="clip">${esc(parseDatasets(v).map(d => d.name).join(', '))}</span>`;
  if (/^https?:/.test(v)) return `<a href="${esc(v)}" target="_blank" rel="noopener" onclick="event.stopPropagation()">${esc(host(v) || v)}</a>`;
  return `<span class="clip">${esc(v).replace(/\n/g, '<br>')}</span>`;
}

function valueChoices(sheet, h) {
  if (cfg(sheet).fields?.[h]?.multi) return ['All', ...codes()];
  const o = DB.sheets[sheet].options[h];
  if (Array.isArray(o)) return o;
  const vals = [...new Set(rowsOf(sheet).map(r => String(r[h] ?? '').trim()).filter(Boolean))];
  return vals.length <= 40 ? vals.sort((a, b) => cmp(sortKey(sheet, h, a), sortKey(sheet, h, b))) : null;
}

function renderSheet() {
  const sheet = view, st = ui(sheet), C = cfg(sheet), S = DB.sheets[sheet];
  if (!st._loaded) { loadPrefs(sheet, st); st._loaded = true; }
  const defCols = (C.cols || S.headers.slice(0, 5)).filter(h => S.headers.includes(h));
  const cols = (st.cols || defCols).filter(h => S.headers.includes(h));
  const rows = filtered(sheet), iss = allIssues()[sheet];
  const total = rowsOf(sheet).length;
  const quick = h => {
    const vals = valueChoices(sheet, h) || [];
    return `<select data-f="${esc(h)}" class="${st.filters[h] ? 'act' : ''}" aria-label="${esc(h)}"><option value="">${esc(h)}: tất cả</option>${vals.map(v => `<option ${st.filters[h] === v ? 'selected' : ''}>${esc(v)}</option>`).join('')}</select>`;
  };
  const nErr = iss.filter(a => a.some(x => x.level === 'err')).length, nAny = iss.filter(a => a.length).length;
  const lvl = h => st.sort.findIndex(s => s.h === h);
  const arrow = h => { const k = lvl(h); if (k < 0) return ''; return `<span class="ar">${st.sort[k].dir > 0 ? '↑' : '↓'}${st.sort.length > 1 ? `<sup>${k + 1}</sup>` : ''}</span>`; };
  const selected = [...st.sel].filter(i => i < total);
  const bulkOpts = Object.entries(S.options).filter(([, o]) => Array.isArray(o));
  const colOpts = (cur, forFilter) => [...(forFilter ? (C.codeCol ? [['_code', 'Mã đề tài']] : []) : [['#', C.codeCol ? 'Thứ tự trong Excel (mã)' : 'Thứ tự trong Excel']]), ...S.headers.map(h => [h, h])]
    .map(([v, t]) => `<option value="${esc(v)}" ${v === cur ? 'selected' : ''}>${esc(t)}</option>`).join('');
  const condVal = (c, k) => {
    if (c.op === 'empty' || c.op === 'nempty') return '<span></span>';
    const ch = c.h && c.h !== '_code' ? valueChoices(sheet, c.h) : null;
    return ch && c.op !== 'contains' && c.op !== 'ncontains'
      ? `<select data-cv="${k}"><option value="">— chọn —</option>${ch.map(v => `<option ${v === c.v ? 'selected' : ''}>${esc(v)}</option>`).join('')}</select>`
      : `<input data-cv="${k}" value="${esc(c.v || '')}" placeholder="Giá trị…">`;
  };
  const isActive = c => c.h && (c.op === 'empty' || c.op === 'nempty' || c.v);
  const nConds = st.conds.filter(isActive).length;
  const chips = [
    ...Object.entries(st.filters).filter(([, v]) => v).map(([h, v]) => `<span class="chip">${h === '_issues' ? 'Kiểm tra' : esc(h)}: <b>${esc(h === '_issues' ? (v === 'err' ? 'có lỗi' : 'có lỗi/cảnh báo') : v)}</b><button type="button" data-xf="${esc(h)}" aria-label="Bỏ">×</button></span>`),
    ...st.conds.map((c, k) => isActive(c) ? `<span class="chip">${esc(c.h === '_code' ? 'Mã' : c.h)} ${esc(opLabel(c.op))} ${c.op === 'empty' || c.op === 'nempty' ? '' : `<b>${esc(c.v)}</b>`}<button type="button" data-xc="${k}" aria-label="Bỏ">×</button></span>` : ''),
    ...st.sort.map((s, k) => s.h ? `<span class="chip sort">⇅ ${st.sort.length > 1 ? (k + 1) + '. ' : ''}${esc(s.h === '#' ? 'Thứ tự Excel' : s.h)} ${s.dir > 0 ? '↑' : '↓'}<button type="button" data-xs="${k}" aria-label="Bỏ">×</button></span>` : ''),
  ].filter(Boolean);

  $('#view').innerHTML = `
    <div class="tools">
      <label class="search"><input type="search" id="q" placeholder="Tìm trong mọi cột…  (phím /)" value="${esc(st.q)}"></label>
      ${(C.filters || []).filter(h => S.headers.includes(h)).map(quick).join('')}
      <select data-f="_issues" class="${st.filters._issues ? 'act' : ''}" aria-label="Kiểm tra"><option value="">Kiểm tra: tất cả</option>
        <option value="err" ${st.filters._issues === 'err' ? 'selected' : ''}>Có lỗi (${nErr})</option>
        <option value="any" ${st.filters._issues === 'any' ? 'selected' : ''}>Có lỗi/cảnh báo (${nAny})</option></select>
      <span class="tgroup">
        <button class="btn sm ${st.panel === 'filter' ? 'on' : ''}" type="button" data-panel="filter">Lọc nâng cao${nConds ? ` <span class="badge">${nConds}</span>` : ''}</button>
        <button class="btn sm ${st.panel === 'sort' ? 'on' : ''}" type="button" data-panel="sort">Sắp xếp${st.sort.length ? ` <span class="badge">${st.sort.length}</span>` : ''}</button>
        <button class="btn sm ${st.panel === 'cols' ? 'on' : ''}" type="button" data-panel="cols">Cột</button>
      </span>
      <span class="count">${rows.length === total ? total : `<b>${rows.length}</b> / ${total}`} dòng</span>
    </div>
    ${st.panel === 'filter' ? `<div class="fpanel"><div class="ph"><b>Lọc nâng cao</b><span>Dòng phải thỏa <b>tất cả</b> điều kiện. ≥ / ≤ hiểu đúng tháng, ngày, học kỳ, hạng và số tiền.</span></div>
      ${st.conds.map((c, k) => `<div class="frow"><select data-ch="${k}"><option value="">— chọn cột —</option>${colOpts(c.h, true)}</select>
        <select data-co="${k}">${OPS.map(([o, t]) => `<option value="${o}" ${o === c.op ? 'selected' : ''}>${t}</option>`).join('')}</select>
        ${condVal(c, k)}<button type="button" class="iconbtn" data-xc="${k}" aria-label="Xóa điều kiện">✕</button></div>`).join('')}
      <div class="pfoot"><button class="btn sm" type="button" id="addcond">＋ Thêm điều kiện</button>${st.conds.length ? '<button class="btn sm" type="button" id="clrcond">Xóa mọi điều kiện</button>' : ''}</div></div>` : ''}
    ${st.panel === 'sort' ? `<div class="fpanel"><div class="ph"><b>Sắp xếp nhiều tầng</b><span>Xét tầng 1 trước; bằng nhau thì xét tầng 2, rồi 3. Ô trống luôn nằm cuối. Mẹo: bấm tiêu đề cột để sắp xếp, giữ <kbd>Shift</kbd> khi bấm để thêm tầng.</span></div>
      ${st.sort.map((s, k) => `<div class="frow"><span class="lv">${k + 1}</span><select data-sh="${k}">${colOpts(s.h, false)}</select>
        <select data-sd="${k}"><option value="1" ${s.dir > 0 ? 'selected' : ''}>Tăng dần (A→Z, cũ→mới, nhỏ→lớn)</option><option value="-1" ${s.dir < 0 ? 'selected' : ''}>Giảm dần (Z→A, mới→cũ, lớn→nhỏ)</option></select>
        <button type="button" class="iconbtn" data-xs="${k}" aria-label="Bỏ tầng">✕</button></div>`).join('')}
      <div class="pfoot">${st.sort.length < 3 ? '<button class="btn sm" type="button" id="addsort">＋ Thêm tầng</button>' : ''}${st.sort.length ? '<button class="btn sm" type="button" id="clrsort">Về thứ tự trong Excel</button>' : ''}</div></div>` : ''}
    ${st.panel === 'cols' ? `<div class="fpanel"><div class="ph"><b>Cột hiển thị trong bảng</b><span>Chỉ đổi cách xem, không đổi file Excel.</span></div>
      <div class="colpick">${S.headers.map(h => `<label><input type="checkbox" data-col="${esc(h)}" ${cols.includes(h) ? 'checked' : ''}> ${esc(h)}</label>`).join('')}</div>
      <div class="pfoot"><button class="btn sm" type="button" id="defcols">Về mặc định</button></div></div>` : ''}
    ${chips.length ? `<div class="chips-active">${chips.join('')}<button class="linkbtn" type="button" id="clrall">Xóa hết lọc &amp; sắp xếp</button></div>` : ''}
    ${selected.length ? `<div class="bulk"><b>Đã chọn ${selected.length}</b>
      ${bulkOpts.map(([h, o]) => `<select data-bulk="${esc(h)}"><option value="">Đặt ${esc(h)}…</option>${o.map(v => `<option>${esc(v)}</option>`).join('')}</select>`).join('')}
      <button class="btn sm danger" type="button" id="bdel">Xóa ${selected.length} dòng</button>
      <button class="btn sm" type="button" id="bnone">Bỏ chọn</button></div>` : ''}
    <div class="tablewrap">${rows.length ? `<table class="grid"><thead><tr>
      <th class="ck nos"><input type="checkbox" id="ckall" aria-label="Chọn tất cả" ${rows.every(r => st.sel.has(r._i)) ? 'checked' : ''}></th>
      <th data-s="#" title="Bấm: sắp xếp. Shift + bấm: thêm tầng">${C.codeCol ? 'Mã' : '#'} ${arrow('#')}</th>
      ${cols.map((h, k) => `<th data-s="${esc(h)}" class="${k > 2 ? 'hide-sm' : ''}" title="Bấm: sắp xếp. Shift + bấm: thêm tầng">${esc(h)} ${arrow(h)}</th>`).join('')}
    </tr></thead><tbody>${rows.map(r => {
      const a = iss[r._i], lv = a.some(x => x.level === 'err') ? 'err' : a.length ? 'warn' : '';
      return `<tr data-i="${r._i}" class="${st.sel.has(r._i) ? 'sel' : ''} ${r.Status === 'Taken' ? 'taken' : ''}">
        <td class="ck"><input type="checkbox" data-ck="${r._i}" ${st.sel.has(r._i) ? 'checked' : ''} aria-label="Chọn dòng"></td>
        <td class="rn">${C.codeCol ? `<span class="code">${esc(r._code)}</span>` : r._i + 2}${lv ? `<span class="issue-dot ${lv}" title="${esc(a.map(x => x.field + ': ' + x.msg).join('\n'))}"></span>` : ''}</td>
        ${cols.map((h, k) => `<td class="${k > 2 ? 'hide-sm' : ''}">${cellHtml(sheet, h, r[h])}</td>`).join('')}</tr>`;
    }).join('')}</tbody></table>` : `<div class="empty">Không có dòng nào khớp. ${total ? '<button class="btn sm" type="button" id="clrall2">Xóa hết bộ lọc</button>' : 'Bấm “Thêm dòng” để bắt đầu.'}</div>`}</div>`;

  const v = $('#view');
  const redraw = () => { savePrefs(sheet); renderSheet(); };
  const on = (id, fn) => { const el = $('#' + id); if (el) el.onclick = fn; };
  const q = $('#q'); let qt;
  q.oninput = () => { clearTimeout(qt); qt = setTimeout(() => { st.q = q.value; const pos = q.selectionStart; renderSheet(); const n = $('#q'); n.focus(); n.setSelectionRange(pos, pos); }, 150); };
  v.querySelectorAll('select[data-f]').forEach(s => s.onchange = () => { st.filters[s.dataset.f] = s.value; redraw(); });
  v.querySelectorAll('[data-panel]').forEach(b => b.onclick = () => {
    st.panel = st.panel === b.dataset.panel ? null : b.dataset.panel;
    if (st.panel === 'filter' && !st.conds.length) st.conds.push({ h: '', op: 'contains', v: '' });
    redraw();
  });
  v.querySelectorAll('[data-ch]').forEach(s => s.onchange = () => {
    const c = st.conds[+s.dataset.ch]; c.h = s.value; c.v = '';
    if (c.h && c.h !== '_code' && valueChoices(sheet, c.h) && c.op === 'contains') c.op = 'eq';
    redraw();
  });
  v.querySelectorAll('[data-co]').forEach(s => s.onchange = () => { st.conds[+s.dataset.co].op = s.value; redraw(); });
  v.querySelectorAll('[data-cv]').forEach(el => {
    if (el.tagName === 'SELECT') el.onchange = () => { st.conds[+el.dataset.cv].v = el.value; redraw(); };
    else {
      let t;
      el.oninput = () => { clearTimeout(t); t = setTimeout(() => {
        const k = el.dataset.cv, pos = el.selectionStart; st.conds[+k].v = el.value; savePrefs(sheet); renderSheet();
        const n = v.querySelector(`[data-cv="${k}"]`); if (n) { n.focus(); n.setSelectionRange(pos, pos); }
      }, 250); };
    }
  });
  v.querySelectorAll('[data-xc]').forEach(b => b.onclick = () => { st.conds.splice(+b.dataset.xc, 1); redraw(); });
  v.querySelectorAll('[data-xf]').forEach(b => b.onclick = () => { st.filters[b.dataset.xf] = ''; redraw(); });
  v.querySelectorAll('[data-xs]').forEach(b => b.onclick = () => { st.sort.splice(+b.dataset.xs, 1); redraw(); });
  v.querySelectorAll('[data-sh]').forEach(s => s.onchange = () => { st.sort[+s.dataset.sh].h = s.value; redraw(); });
  v.querySelectorAll('[data-sd]').forEach(s => s.onchange = () => { st.sort[+s.dataset.sd].dir = +s.value; redraw(); });
  v.querySelectorAll('[data-col]').forEach(c => c.onchange = () => { st.cols = S.headers.filter(h => v.querySelector(`[data-col="${CSS.escape(h)}"]`).checked); redraw(); });
  on('addcond', () => { st.conds.push({ h: '', op: 'contains', v: '' }); redraw(); });
  on('clrcond', () => { st.conds = []; redraw(); });
  on('addsort', () => { const used = st.sort.map(s => s.h); st.sort.push({ h: S.headers.find(h => !used.includes(h)) || '#', dir: 1 }); redraw(); });
  on('clrsort', () => { st.sort = []; redraw(); });
  on('defcols', () => { st.cols = null; redraw(); });
  const clearAll = () => { st.q = ''; st.filters = {}; st.conds = []; st.sort = []; redraw(); };
  on('clrall', clearAll); on('clrall2', clearAll);
  v.querySelectorAll('th[data-s]').forEach(th => th.onclick = e => {
    const h = th.dataset.s, k = lvl(h);
    if (e.shiftKey) {                                           // add a level / flip it / remove it
      if (k < 0) { if (st.sort.length < 3) st.sort.push({ h, dir: 1 }); }
      else if (st.sort[k].dir > 0) st.sort[k].dir = -1; else st.sort.splice(k, 1);
    } else {                                                    // single column: ↑ → ↓ → off
      st.sort = k === 0 && st.sort.length === 1 ? (st.sort[0].dir > 0 ? [{ h, dir: -1 }] : []) : [{ h, dir: 1 }];
    }
    redraw();
  });
  v.querySelectorAll('tbody tr').forEach(tr => tr.onclick = e => { if (!e.target.closest('input,a')) openEditor(sheet, +tr.dataset.i); });
  v.querySelectorAll('[data-ck]').forEach(c => c.onchange = () => { const i = +c.dataset.ck; c.checked ? st.sel.add(i) : st.sel.delete(i); renderSheet(); });
  const ca = $('#ckall'); if (ca) ca.onchange = e => { rows.forEach(r => e.target.checked ? st.sel.add(r._i) : st.sel.delete(r._i)); renderSheet(); };
  on('bnone', () => { st.sel.clear(); renderSheet(); });
  on('bdel', async () => {
    const warn = sheet === 'Topics' ? '<br><br>Xóa đề tài sẽ <b>đổi mã</b> các đề tài nằm dưới trong cùng hướng. Nên đặt Status = Taken thay vì xóa.' : '';
    if (!await confirmBox('Xóa nhiều dòng', `Xóa <b>${selected.length}</b> dòng khỏi sheet ${esc(sheet)}? Có thể hoàn tác.${warn}`, 'Xóa', true)) return;
    if (await op({ sheet, op: 'delete', indexes: selected }, `Đã xóa ${selected.length} dòng`)) { st.sel.clear(); renderSheet(); }
  });
  v.querySelectorAll('[data-bulk]').forEach(s => s.onchange = async () => {
    if (!s.value) return;
    const h = s.dataset.bulk, val = s.value;
    if (!await confirmBox('Sửa nhiều dòng', `Đặt <b>${esc(h)} = ${esc(val)}</b> cho ${selected.length} dòng?`, 'Áp dụng')) { s.value = ''; return; }
    if (await op({ sheet, op: 'set', column: h, value: val, indexes: selected }, `Đã cập nhật ${selected.length} dòng`)) { st.sel.clear(); renderSheet(); }
  });
}

/* ------------------------------------------------------------ editor drawer */
let dirty = false;
function fieldHtml(sheet, h, val, isNew) {
  const C = cfg(sheet), f = (C.fields || {})[h] || {}, S = DB.sheets[sheet];
  const opts = S.options[h], id = 'f_' + h.replace(/\W+/g, '_');
  const req = (C.required || []).includes(h);
  const long = f.big || f.long || (String(val).length > 80 || String(val).includes('\n'));
  let input;
  if (f.datasets) {
    const list = parseDatasets(val); if (!list.length) list.push({ name: '', url: '' });
    input = `<input type="hidden" id="${id}" data-h="${esc(h)}" value="${esc(val)}"><div class="dsrows" data-ds="${id}">${list.map(dsRow).join('')}</div>
      <div><button type="button" class="btn sm" data-dsadd="${id}">＋ Thêm dataset</button></div>`;
  } else if (f.multi) {
    const set = new Set(String(val || '').split(',').map(s => s.trim()).filter(Boolean));
    input = `<input type="hidden" id="${id}" data-h="${esc(h)}" value="${esc(val)}"><div class="chips" data-multi="${id}">
      ${['All', ...codes()].map(c => `<button type="button" data-v="${esc(c)}" class="${set.has(c) ? 'on' : ''}">${esc(c)}</button>`).join('')}</div>`;
  } else if (Array.isArray(opts)) {
    const list = opts.includes(val) || !val ? opts : [val, ...opts];
    input = `<select id="${id}" data-h="${esc(h)}">${req ? '' : '<option value="">—</option>'}${list.map(o => `<option ${o === val ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
  } else if (long) {
    input = `<textarea id="${id}" data-h="${esc(h)}" class="${f.big ? 'big' : ''}" rows="${f.long || 3}" spellcheck="true">${esc(val)}</textarea>`;
  } else {
    const dl = f.list ? `<datalist id="${id}_l">${f.list().map(o => `<option value="${esc(o)}">`).join('')}</datalist>` : '';
    input = `<input id="${id}" data-h="${esc(h)}" value="${esc(val)}" ${f.ph ? `placeholder="${esc(f.ph)}"` : ''} ${f.list ? `list="${id}_l"` : ''} ${f.url ? 'type="url" inputmode="url"' : ''} autocomplete="off">${dl}`;
  }
  const quick = f.quick ? `<div class="quick">${f.quick.map(q => `<button type="button" data-q="${esc(q)}" data-for="${id}">${esc(q)}${q.endsWith(' ') ? '…' : ''}</button>`).join('')}</div>` : '';
  const tpl = f.template ? `<div class="quick"><button type="button" data-tpl="${id}">Chèn khung mẫu còn thiếu</button></div><div class="tpl" data-tplv="${id}"></div>` : '';
  const aux = f.url ? `<a class="aux" data-open="${id}" href="#" tabindex="-1">Mở link ↗</a>` : f.big ? `<span class="aux" data-wc="${id}"></span>` : '';
  return `<div class="f ${f.wide || long ? 'wide' : ''}" data-field="${esc(h)}"><label for="${id}">${esc(h)}${req ? '<span class="req">*</span>' : ''}${aux}</label>
    ${input}${quick}${tpl}${f.hint ? `<span class="hint">${esc(f.hint)}</span>` : ''}<span class="msg"></span></div>`;
}

function dsRow(d) {
  return `<div class="dsrow"><input class="dsn" placeholder="Tên dataset (bắt buộc)" value="${esc(d.name)}" aria-label="Tên dataset">
    <input class="dsu" type="url" inputmode="url" placeholder="https://… (link, không bắt buộc)" value="${esc(d.url)}" aria-label="Link dataset">
    <button type="button" class="iconbtn dsx" title="Xóa dataset này" aria-label="Xóa dataset này">✕</button></div>`;
}

function formValues() {
  const o = {};
  $('#dbody').querySelectorAll('[data-h]').forEach(el => o[el.dataset.h] = el.value);
  return o;
}

function openEditor(sheet, index, preset) {
  const S = DB.sheets[sheet], C = cfg(sheet);
  const isNew = index == null;
  const r = isNew ? { _i: -1, ...(preset || {}) } : rowsOf(sheet)[index];
  if (!r) return;
  editing = { sheet, index, isNew, before: isNew ? null : S.rows[index].slice() };
  if (isNew && !preset) {                              // sensible defaults
    for (const h of S.headers) { const o = S.options[h]; if (Array.isArray(o) && (C.required || []).includes(h)) r[h] = o[0]; }
    if (sheet === 'Topics') { r.Status = 'Open'; const f = ui('Topics').filters.Direction; r.Direction = codes().includes(f) ? f : codes()[0]; }
    if (sheet === 'Conferences') r.Directions = 'All';
  }
  const title = isNew ? 'Thêm dòng mới' : (C.title(r) || 'Dòng ' + (index + 2));
  $('#dkicker').textContent = sheet + (isNew ? '' : ' · ' + (r._code ? r._code + ' · ' : '') + 'dòng ' + (index + 2) + ' trong Excel');
  $('#dtitle').textContent = title;
  const topicNote = sheet === 'Topics' ? `<div class="note">Mã đề tài ${isNew ? 'sẽ là' : 'hiện là'} <b class="code" id="tcode"></b>. Mã tính theo vị trí dòng trong hướng: thêm vào cuối, và khi có người nhận đề tài thì đặt Status = <b>Taken</b> thay vì xóa.</div>` : '';
  $('#dbody').innerHTML = `${topicNote}<div class="sumbox" id="sum" hidden></div><div class="fgrid">${S.headers.map(h => fieldHtml(sheet, h, r[h] ?? '', isNew)).join('')}</div>`;
  const site = C.site && !isNew ? C.site(r) : null;
  $('#dfoot').innerHTML = `<button class="btn primary" type="button" id="dsave">${isNew ? 'Thêm vào Excel' : 'Lưu'} <kbd style="color:#dfe7ff;border-color:#7d97ff">Ctrl S</kbd></button>
    <button class="btn" type="button" id="dcancel">Hủy</button><span class="sp"></span>
    ${!isNew && C.ordered ? `<button class="btn sm" type="button" id="dup" title="Chuyển lên một dòng" ${index === 0 ? 'disabled' : ''}>↑</button><button class="btn sm" type="button" id="ddown" title="Chuyển xuống một dòng" ${index === S.rows.length - 1 ? 'disabled' : ''}>↓</button>` : ''}
    ${site ? `<a class="btn sm" target="_blank" rel="noopener" href="/site/${esc(site)}">Xem trên web ↗</a>` : ''}
    ${!isNew ? `<button class="btn sm" type="button" id="dcopy">Nhân bản</button>` : ''}
    ${!isNew && sheet === 'Topics' ? `<button class="btn sm" type="button" id="dtaken">${r.Status === 'Taken' ? 'Mở lại (Open)' : 'Đánh dấu Taken'}</button>` : ''}
    ${!isNew ? `<button class="btn sm danger" type="button" id="ddel">Xóa</button>` : ''}`;
  wireForm(sheet);
  dirty = false;
  $('#drawer').classList.add('open'); $('#drawer').setAttribute('aria-hidden', 'false'); $('#scrim').hidden = false;
  history.replaceState(null, '', '#' + encodeURIComponent(sheet) + (isNew ? '/new' : '/' + (index + 2)));
  setTimeout(() => { const first = $('#dbody').querySelector('input:not([type=hidden]),select,textarea'); first && isNew && first.focus(); }, 220);
}

const ARXIV_RE = /(\d{4}\.\d{4,5}(?:v\d+)?)/;
// Publications with Type = Preprint: an arXiv box fills Venue/PDF/DOI from the ID, and can fetch title/authors/date
function arxivBox(sheet) {
  if (sheet !== 'Publications') return;
  const body = $('#dbody'), F = h => body.querySelector(`[data-h="${h}"]`);
  const typeEl = F('Type'); if (!typeEl) return;
  const set = (h, v, onlyIfEmpty) => { const el = F(h); if (!el || (onlyIfEmpty && el.value.trim())) return false; el.value = v; return true; };
  const box = document.createElement('div');
  box.className = 'note arx'; box.hidden = true;
  box.innerHTML = `<b>Preprint arXiv</b> Dán arXiv ID hoặc link (ví dụ 2501.01234 hoặc https://arxiv.org/abs/2501.01234): Venue, PDF và DOI được điền tự động.
    <div class="arxrow"><input id="arxid" placeholder="arXiv ID hoặc link" autocomplete="off"><button type="button" class="btn sm" id="arxget">Lấy tên bài, tác giả, ngày từ arXiv</button></div><span class="hint" id="arxmsg"></span>`;
  body.querySelector('.fgrid').before(box);
  const idIn = box.querySelector('#arxid'), msg = box.querySelector('#arxmsg');
  const cur = () => { const m = [F('Venue').value, F('PDF')?.value, F('DOI')?.value].join(' ').match(/arxiv[^0-9]{0,12}(\d{4}\.\d{4,5}(?:v\d+)?)/i); return m ? m[1] : ''; };
  const apply = id => {
    F('Venue').value = 'arXiv preprint arXiv:' + id;
    const base = id.replace(/v\d+$/, '');
    if (!F('PDF').value.trim() || /arxiv\.org/i.test(F('PDF').value)) F('PDF').value = 'https://arxiv.org/pdf/' + id;
    if (!F('DOI').value.trim() || /arxiv/i.test(F('DOI').value)) F('DOI').value = '10.48550/arXiv.' + base;
    msg.textContent = 'Đã điền Venue, PDF, DOI theo arXiv:' + id + '.'; changed();
  };
  const sync = () => { const on = typeEl.value === 'Preprint'; box.hidden = !on; F('Venue').readOnly = on && !!cur(); if (on && !idIn.value) idIn.value = cur(); };
  typeEl.addEventListener('change', () => {
    if (typeEl.value === 'Preprint') { const st = F('Status'); if (st && st.value !== 'In Submission') st.value = 'Preprint'; if (cur()) apply(cur()); }
    sync();
  });
  idIn.addEventListener('input', () => { const m = idIn.value.match(ARXIV_RE); if (m) { apply(m[1]); sync(); } });
  box.querySelector('#arxget').onclick = async () => {
    const m = (idIn.value || cur()).match(ARXIV_RE); if (!m) { msg.textContent = 'Nhập arXiv ID trước.'; return; }
    msg.textContent = 'Đang hỏi arXiv…';
    try {
      const r = await api('/api/arxiv?id=' + encodeURIComponent(m[1]));
      apply(m[1]);
      const filled = [['Title', r.title], ['Authors', r.authors], ['Date', r.date]].filter(([h, v]) => v && set(h, v, true)).map(([h]) => h);
      msg.textContent = filled.length ? `Đã điền ${filled.join(', ')} từ arXiv (ô đã có chữ thì giữ nguyên).` : 'Các ô Title, Authors, Date đã có chữ nên giữ nguyên.';
      changed();
    } catch (e) { msg.textContent = e.message; }
  };
  sync();
}

function wireForm(sheet) {
  const body = $('#dbody');
  arxivBox(sheet);
  body.querySelectorAll('[data-multi]').forEach(box => box.onclick = e => {
    const b = e.target.closest('button'); if (!b) return;
    const inp = $('#' + box.dataset.multi);
    if (b.dataset.v === 'All') box.querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b && !b.classList.contains('on')));
    else { b.classList.toggle('on'); box.querySelector('[data-v="All"]').classList.remove('on'); }
    inp.value = [...box.querySelectorAll('button.on')].map(x => x.dataset.v).join(', ');
    changed();
  });
  body.querySelectorAll('[data-ds]').forEach(box => {
    const inp = $('#' + box.dataset.ds);
    const sync = () => {
      const rows = [...box.querySelectorAll('.dsrow')];
      inp.value = joinDatasets(rows.map(r => ({ name: r.querySelector('.dsn').value, url: r.querySelector('.dsu').value })));
      rows.forEach(r => r.querySelector('.dsn').classList.toggle('need', !r.querySelector('.dsn').value.trim() && !!r.querySelector('.dsu').value.trim()));
    };
    box.addEventListener('input', sync);
    box.addEventListener('click', e => {
      const x = e.target.closest('.dsx'); if (!x) return;
      if (box.querySelectorAll('.dsrow').length > 1) x.closest('.dsrow').remove();
      else x.closest('.dsrow').querySelectorAll('input').forEach(i => i.value = '');
      sync(); changed();
    });
    body.querySelector(`[data-dsadd="${box.dataset.ds}"]`).onclick = () => {
      box.insertAdjacentHTML('beforeend', dsRow({ name: '', url: '' }));
      box.lastElementChild.querySelector('.dsn').focus();
    };
    box.addEventListener('paste', e => {                       // "Name https://..." pasted into the name box: split it
      const t = e.target; if (!t.classList.contains('dsn')) return;
      const text = (e.clipboardData || window.clipboardData).getData('text');
      const m = text.match(/^([\s\S]*?)\s*(https?:\/\/\S+)\s*$/);
      if (m && m[1].trim()) { e.preventDefault(); t.value = m[1].replace(/^[•\-*]\s*/, '').trim(); t.closest('.dsrow').querySelector('.dsu').value = m[2]; sync(); changed(); }
    });
    sync();
  });
  body.querySelectorAll('[data-q]').forEach(b => b.onclick = () => {
    const el = $('#' + b.dataset.for), q = b.dataset.q;
    if (el.tagName === 'TEXTAREA') { el.value = (el.value.trim() ? el.value.replace(/\s*$/, '\n') : '') + q; }
    else el.value = q;
    el.focus(); el.setSelectionRange(el.value.length, el.value.length); changed();
  });
  body.querySelectorAll('[data-tpl]').forEach(b => b.onclick = () => {
    const el = $('#' + b.dataset.tpl), labels = cfg(sheet).fields[el.dataset.h].template;
    const t = tplCheck(el.value, labels);
    const miss = labels.filter(l => !(l in t.fields));
    if (!miss.length) { toast('Đã có đủ các mục của khung mẫu.'); return; }
    el.value = (el.value.trim() ? el.value.replace(/\s*$/, '\n') : '') + miss.map(l => l + ': ').join('\n');
    el.focus(); changed();
  });
  body.querySelectorAll('[data-open]').forEach(a => a.onclick = e => { e.preventDefault(); const v = $('#' + a.dataset.open).value; if (isUrl(v)) window.open(v, '_blank', 'noopener'); });
  body.addEventListener('input', changed);
  body.addEventListener('change', changed);
  $('#dsave').onclick = saveEditor;
  $('#dcancel').onclick = closeEditor;
  const { index } = editing;
  $('#dup') && ($('#dup').onclick = () => moveRow(index, index - 1));
  $('#ddown') && ($('#ddown').onclick = () => moveRow(index, index + 1));
  $('#dcopy') && ($('#dcopy').onclick = async () => {
    if (dirty && !await confirmBox('Bỏ thay đổi?', 'Các thay đổi chưa lưu sẽ mất.', 'Tiếp tục')) return;
    const v = formValues(); const k = cfg(sheet).key; if (k && v[k]) v[k] += ' (copy)';
    openEditor(sheet, null, v);
  });
  $('#dtaken') && ($('#dtaken').onclick = async () => {
    const el = body.querySelector('[data-h="Status"]'); el.value = el.value === 'Taken' ? 'Open' : 'Taken'; changed(); await saveEditor();
  });
  $('#ddel') && ($('#ddel').onclick = deleteRow);
  validateForm();
}
function changed() { dirty = true; validateForm(); }

function validateForm() {
  if (!editing) return [];
  const { sheet, index, isNew } = editing;
  const v = formValues();
  const r = { _i: isNew ? -1 : index, ...v };
  const issues = checkRow(sheet, r);
  const body = $('#dbody');
  body.querySelectorAll('.f').forEach(f => {
    const mine = issues.filter(x => x.field === f.dataset.field);
    f.classList.toggle('bad', mine.some(x => x.level === 'err'));
    f.classList.toggle('warned', !mine.some(x => x.level === 'err') && mine.length > 0);
    const m = f.querySelector('.msg'); m.className = 'msg ' + (mine[0]?.level || ''); m.textContent = mine.map(x => x.msg).join(' ');
  });
  if (sheet === 'Topics') {
    const d = v.Direction || '?';
    const rows = rowsOf('Topics');
    const n = isNew ? rows.filter(o => o.Direction === d).length + 1 : rows.filter(o => o.Direction === d && o._i <= index).length + (rows[index].Direction === d ? 0 : 1);
    $('#tcode').textContent = d + '-' + String(n).padStart(2, '0');
    for (const [h, labels] of [['Research Description', DESC], ['Research Expectation', EXPC]]) {
      const el = body.querySelector(`[data-h="${h}"]`); if (!el) continue;
      const t = tplCheck(el.value, labels);
      body.querySelector(`[data-tplv="${el.id}"]`).innerHTML = labels.map(l => `<span class="${t.missing.includes(l) ? 'miss' : ''}">${t.missing.includes(l) ? '✗' : '✓'} ${esc(l)}</span>`).join('');
      const wc = body.querySelector(`[data-wc="${el.id}"]`);
      if (wc) wc.textContent = h === 'Research Description' && t.fields.Context !== undefined ? `Context: ${t.fields.Context.trim().split(/\s+/).filter(Boolean).length} từ, ${sentences(t.fields.Context)} câu` : '';
    }
  }
  const sum = $('#sum'), errs = issues.filter(x => x.level === 'err');
  sum.hidden = !issues.length;
  sum.innerHTML = issues.length ? `<b>${errs.length ? errs.length + ' lỗi cần sửa' : 'Không có lỗi'}${issues.length - errs.length ? ` · ${issues.length - errs.length} cảnh báo` : ''}</b>` +
    issues.map(x => `<div class="${x.level}">• ${esc(x.field)}: ${esc(x.msg)}</div>`).join('') : '';
  return issues;
}

async function saveEditor() {
  if (!editing || busy) return;
  const { sheet, index, isNew, before } = editing;
  const S = DB.sheets[sheet];
  const issues = validateForm(), errs = issues.filter(x => x.level === 'err');
  if (errs.length) {
    const go = await dialog('Dữ liệu còn lỗi', `${errs.map(x => `• <b>${esc(x.field)}</b>: ${esc(x.msg)}`).join('<br>')}<br><br>Web có thể hiển thị sai hoặc ẩn dòng này.`,
      [{ t: 'Quay lại sửa', v: false, primary: 1 }, { t: 'Vẫn lưu', v: true }]);
    if (!go) return;
  }
  const v = formValues();
  const values = S.headers.map(h => v[h] ?? '');
  // Directions: a changed code is renamed everywhere first
  if (sheet === 'Directions' && !isNew && before[0] && v.Code && v.Code.trim().toUpperCase() !== before[0]) {
    const used = rowsOf('Topics').filter(t => t.Direction === before[0]).length;
    if (!await confirmBox('Đổi mã hướng', `Đổi mã <b>${esc(before[0])}</b> thành <b>${esc(v.Code.trim().toUpperCase())}</b> ở mọi sheet (${used} đề tài, cùng các hội nghị và tạp chí dùng mã này)? Mã đề tài trên web sẽ đổi theo, ví dụ ${esc(before[0])}-01 → ${esc(v.Code.trim().toUpperCase())}-01.`, 'Đổi mã')) return;
    const res = await op({ sheet, op: 'rename_code', old: before[0], new: v.Code }, 'Đã đổi mã hướng ở mọi sheet');
    if (!res) return;
    values[0] = v.Code.trim().toUpperCase();
    const r2 = await op({ sheet, op: 'update', index, values, before: DB.sheets[sheet].rows[index] }, 'Đã lưu');
    if (r2) { dirty = false; closeEditor(true); flash(index); }
    return;
  }
  const res = isNew ? await op({ sheet, op: 'insert', values }, 'Đã thêm vào database.xlsx')
    : await op({ sheet, op: 'update', index, values, before }, 'Đã lưu vào database.xlsx');
  if (res) { dirty = false; closeEditor(true); if (view !== sheet) go(sheet); flash(res.index); }
}

async function moveRow(i, j) {
  if (dirty && !await confirmBox('Bỏ thay đổi?', 'Các thay đổi chưa lưu sẽ mất khi di chuyển dòng.', 'Tiếp tục')) return;
  const sheet = editing.sheet;
  const res = await op({ sheet, op: 'move', index: i, to: j }, 'Đã đổi thứ tự');
  if (res) { dirty = false; openEditor(sheet, j); flash(j); }
}

async function deleteRow() {
  const { sheet, index } = editing;
  const r = rowsOf(sheet)[index], C = cfg(sheet);
  if (sheet === 'Topics') {
    const ans = await dialog('Xóa đề tài?', `Xóa đề tài <b>${esc(r._code)}</b> sẽ đổi mã của các đề tài ${esc(r.Direction)} nằm dưới nó (link cũ như research.html#${esc(r._code)} sẽ trỏ sai).<br><br>Nếu đề tài đã có người nhận, nên <b>đánh dấu Taken</b>.`,
      [{ t: 'Hủy', v: 0 }, { t: 'Đánh dấu Taken', v: 1, primary: 1 }, { t: 'Vẫn xóa', v: 2, danger: 1 }]);
    if (ans === 1) { $('#dbody').querySelector('[data-h="Status"]').value = 'Taken'; changed(); return saveEditor(); }
    if (ans !== 2) return;
  } else if (sheet === 'Directions' && rowsOf('Topics').some(t => t.Direction === r.Code)) {
    const n = rowsOf('Topics').filter(t => t.Direction === r.Code).length;
    if (!await confirmBox('Xóa hướng đang được dùng', `Hướng <b>${esc(r.Code)}</b> còn ${n} đề tài. Xóa hướng thì các đề tài này sẽ thành một nhóm lạ trên web. Vẫn xóa?`, 'Vẫn xóa', true)) return;
  } else if (!await confirmBox('Xóa dòng?', `Xóa <b>${esc(C.title(r) || 'dòng ' + (index + 2))}</b> khỏi sheet ${esc(sheet)}? Có thể hoàn tác bằng nút “Hoàn tác”.`, 'Xóa', true)) return;
  if (await op({ sheet, op: 'delete', index }, 'Đã xóa dòng')) { dirty = false; closeEditor(true); }
}

async function closeEditor(force) {
  if (!editing) return;
  if (force !== true && dirty && !await confirmBox('Bỏ thay đổi?', 'Bạn có thay đổi chưa lưu. Đóng mà không lưu?', 'Đóng, không lưu', true)) return;
  editing = null; dirty = false;
  $('#drawer').classList.remove('open'); $('#drawer').setAttribute('aria-hidden', 'true'); $('#scrim').hidden = true;
  history.replaceState(null, '', '#' + encodeURIComponent(view));
}
function flash(i) {
  requestAnimationFrame(() => { const tr = $('#view').querySelector(`tr[data-i="${i}"]`); if (tr) { tr.classList.add('flash'); tr.scrollIntoView({ block: 'nearest' }); } });
}

/* ------------------------------------------------------------ routing + events */
function go(name) { view = name; $('#side').classList.remove('open'); renderNav(); render(); if (location.hash.split('/')[0] !== '#' + encodeURIComponent(name)) history.replaceState(null, '', '#' + encodeURIComponent(name)); $('#view').scrollTop = 0; }
function route() {
  const [s, row] = decodeURIComponent(location.hash.slice(1)).split('/');
  const name = DB.sheets[s] ? s : 'overview';
  if (editing) { closeEditor(true); }
  view = name; renderNav(); render();
  if (DB.sheets[s] && row) { if (row === 'new') openEditor(s, null); else if (+row >= 2 && rowsOf(s)[+row - 2]) openEditor(s, +row - 2); }
}
addEventListener('hashchange', async () => {
  if (editing && dirty && !await confirmBox('Bỏ thay đổi?', 'Bạn có thay đổi chưa lưu. Rời biểu mẫu mà không lưu?', 'Rời đi', true)) {
    history.replaceState(null, '', '#' + encodeURIComponent(editing.sheet) + (editing.isNew ? '/new' : '/' + (editing.index + 2))); return;
  }
  dirty = false; route();
});
$('#dclose').onclick = closeEditor; $('#scrim').onclick = closeEditor;
$('#menu').onclick = () => $('#side').classList.toggle('open');
$('#undo').onclick = async () => {
  if (!DB.backups) { toast('Chưa có thao tác nào để hoàn tác.'); return; }
  if (!await confirmBox('Hoàn tác', 'Đưa file database.xlsx về trạng thái <b>trước thao tác lưu gần nhất</b>? Có thể bấm nhiều lần để lùi tiếp.', 'Hoàn tác')) return;
  try { const r = await api('/api/undo', {}); closeEditor(true); await reload(); toast('Đã hoàn tác thao tác gần nhất', 'ok'); }
  catch (e) { e.status === 423 ? dialog('Không hoàn tác được', esc(e.message), [{ t: 'Đã hiểu', primary: 1 }]) : toast(e.message, 'err'); }
};
addEventListener('keydown', e => {
  if (!$('#modal').hidden) { if (e.key === 'Escape') $('#mact').firstChild.click(); return; }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); if (editing) saveEditor(); return; }
  if (e.key === 'Escape' && editing) { closeEditor(); return; }
  if (e.key === '/' && !editing && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName) && $('#q')) { e.preventDefault(); $('#q').focus(); }
});
addEventListener('beforeunload', e => { if (dirty) { e.preventDefault(); e.returnValue = ''; } });

// Watch the file: picks up edits made in Excel and warns while Excel has it open
setInterval(async () => {
  if (busy || document.hidden) return;
  try {
    const v = await api('/api/version');
    if (v.locked) setSync('warn', 'Excel đang mở file: đóng Excel trước khi lưu');
    else if (!busy) setSync('ok', 'Đã đồng bộ với database.xlsx');
    if (v.version !== lastVersion) {
      if (editing && dirty) { lastVersion = v.version; banner('File database.xlsx vừa thay đổi bên ngoài (Excel?). Lưu biểu mẫu này sẽ được kiểm tra xung đột; nếu dòng đã đổi, dữ liệu sẽ tải lại.'); }
      else { const ed = editing; await reload(true); if (ed && !ed.isNew) openEditor(ed.sheet, ed.index); toast('Đã tải lại dữ liệu mới từ file Excel'); }
    }
  } catch { setSync('err', 'Mất kết nối: cửa sổ server đã đóng?'); }
}, 3000);

(async () => {
  try {
    await reload();
    route();
    try { PRED = await api('/api/predatory'); if (PRED.journals.length) { ISSUES = null; renderNav(); if (view === 'overview' || view === 'Journals') render(); } } catch { }
  } catch (e) {
    $('#view').innerHTML = `<div class="panel"><h3>Không kết nối được server</h3><p>Hãy mở công cụ bằng file <b>Open database manager.bat</b> (hoặc chạy <code>python admin/server.py</code>), rồi tải lại trang này.</p><p>${esc(e.message)}</p></div>`;
    setSync('err', 'Không kết nối được server');
  }
})();
})();
