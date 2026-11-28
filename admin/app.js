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
    icon: '📄', desc: 'Các bài báo hiện ở mục Publications trên trang chủ.', key: 'Title', title: r => r.Title,
    cols: ['Type', 'Status', 'Date', 'Title', 'Venue'], filters: ['Type', 'Status'],
    required: ['Type', 'Status', 'Date', 'Title', 'Authors', 'Venue'],
    fields: {
      Date: { hint: 'Ghi theo dạng dd-mm-yyyy, mm-yyyy hoặc chỉ yyyy. Bài không có năm sẽ không hiện trên web.', ph: '15-03-2026' },
      Title: { long: 2, wide: 1 }, Authors: { wide: 1, hint: 'Các tác giả cách nhau bằng dấu phẩy. Tên "Ha Anh Vu" được tự động in đậm.' }, Venue: { wide: 1 },
      DOI: { hint: 'Mã DOI hoặc đường link đầy đủ.' }, PDF: { url: 1 }, Code: { url: 1 },
      'Link ID': { list: () => linkIds('Projects'), hint: 'Mã ngắn để nối bài báo với đề tài luận văn tương ứng trong Projects (khi bài được nhận).' },
    },
    site: () => 'index.html#publications',
  },
  Projects: {
    icon: '🎓', desc: 'Các đề tài luận văn đang hướng dẫn, hiện ở mục Undergraduate thesis projects trên trang chủ.', key: 'Title', title: r => r.Title,
    cols: ['Semester', 'Title', 'Members', 'Thesis Status', 'Paper Status'], filters: ['Semester', 'Thesis Status', 'Paper Status'],
    required: ['Semester', 'Title', 'Thesis Status'],
    fields: {
      Semester: { semester: 1, hint: 'Chọn mùa và nhập năm. Trên web, năm mới nhất hiện trước; trong một năm: Spring, rồi Summer, rồi Fall.' },
      Title: { long: 2, wide: 1 }, Members: { long: 2, wide: 1, hint: 'Tên các sinh viên, cách nhau bằng dấu chấm phẩy (;).' },
      'Paper Venue': { hint: 'Nơi đã gửi bài. Trên web hiện sau dòng "Paper submitted to".' },
      'Link ID': { list: () => linkIds('Publications'), hint: 'Ghi giống hệt Link ID của bài báo tương ứng trong Publications.' },
    },
    site: () => 'index.html#supervision',
  },
  Directions: {
    icon: '🧭', desc: 'Các hướng nghiên cứu. Thứ tự các dòng cũng là thứ tự thẻ, nút lọc và nhóm đề tài trên web.', key: 'Code', title: r => r.Code + ' · ' + r.Direction, ordered: true,
    cols: ['Code', 'Direction', 'Summary'], required: ['Code', 'Direction'],
    fields: {
      Code: { hint: 'Viết hoa, không dấu, không có khoảng trắng. Khi đổi mã, mọi nơi đang dùng mã này sẽ được cập nhật theo.' },
      'Focus (current research trends)': { long: 3, wide: 1, hint: 'Hiện thành dòng "Current trends" khi sinh viên chọn hướng này.' },
      Summary: { long: 2, wide: 1, hint: 'Một câu tiếng Anh khoảng 10 từ, hiện trên thẻ Research interests ở trang chủ.' },
    },
    site: () => 'index.html#research',
  },
  Topics: {
    icon: '🧪', desc: 'Các đề tài cho sinh viên, hiện ở mục Research topics trên research.html. Mã đề tài gồm mã hướng và vị trí của dòng trong hướng đó.', key: 'English Title',
    title: r => r['English Title'], codeCol: true, filters: ['Direction', 'Status'],
    cols: ['Direction', 'Status', 'English Title', 'Dataset'], required: ['Direction', 'Status', 'English Title', 'Research Description', 'Research Expectation'],
    fields: {
      'Vietnamese Title': { long: 2, wide: 1 }, 'English Title': { long: 2, wide: 1 },
      Dataset: { datasets: 1, wide: 1, hint: 'Phải có tên, link có thể bỏ trống. Trên web, tên dataset hiện thành link tới trang của dataset.' },
      'Research Description': { big: 1, wide: 1, template: DESC },
      'Research Expectation': { big: 1, wide: 1, template: EXPC },
    },
    site: r => 'research.html#' + r._code,
  },
  Conferences: {
    icon: '🎤', desc: 'Các hội nghị có kỷ yếu Scopus, hiện ở tab Conferences trong mục Where to publish.', key: 'Conference', title: r => r.Conference + (r['Full Name'] ? ' · ' + r['Full Name'] : ''),
    cols: ['Conference', 'Full Name', 'CORE Rank', 'Submission Deadline', 'Hybrid', 'Registration Fee'], filters: ['CORE Rank', 'Directions', 'Hybrid'],
    required: ['Conference'],
    fields: {
      'Full Name': { wide: 1 }, Directions: { multi: 1, wide: 1 },
      Proceedings: { hint: 'Nơi xuất bản kỷ yếu (phải được Scopus index). Cột này cũng dùng cho ô tìm kiếm trên web.', list: () => [...new Set(rowsOf('Conferences').map(r => r.Proceedings).filter(Boolean))] },
      'Submission Deadline': { hint: 'Ghi tên tháng bằng tiếng Anh, ví dụ "Mid December – January".' },
      'Held In': { hint: 'Tháng diễn ra hội nghị, ví dụ "October–November".' },
      Website: { url: 1, wide: 1, hint: 'Trang chính thức của hội nghị; nếu hội nghị chỉ có trang riêng cho từng năm thì dùng trang dblp.' },
      'Registration Fee': { long: 2, wide: 1, hint: 'Mỗi mức phí ghi một dòng. Dấu ≈ nghĩa là số ước lượng.', quick: ['Regular: ≈USD ', 'Student: ≈USD ', 'Regular: USD ', 'Student: USD '] },
    },
    site: () => 'research.html#venues',
  },
  Journals: {
    icon: '📚', desc: 'Các tạp chí, hiện ở tab Journals trong mục Where to publish.', key: 'Journal', title: r => r.Journal,
    cols: ['Journal', 'Publisher', 'Rank', 'Time to Publish', 'Publication Fee'], filters: ['Rank', 'Publisher', 'Directions'],
    required: ['Journal'],
    fields: {
      Journal: { wide: 1 }, Publisher: { list: () => [...new Set(rowsOf('Journals').map(r => r.Publisher).filter(Boolean))].sort() },
      Directions: { multi: 1, wide: 1 }, Link: { url: 1, wide: 1 },
      'Time to Publish': { hint: 'Số tháng từ lúc gửi bài tới khi bài được đăng online, ví dụ "4–8 months".', quick: ['3–6 months', '4–8 months', '6–10 months', '8–14 months'] },
      'Publication Fee': { long: 2, hint: 'Mỗi hình thức đăng ghi một dòng.', quick: ['Free (subscription)', 'Free', 'Open access: USD ', 'Open access: ≈USD '] },
    },
    site: () => 'research.html#venues',
  },
  'Page Text': {
    icon: '🧾', group: 'pages', desc: 'Các dòng chữ cố định trên hai trang web: Home là trang chủ, Research là research.html. Bấm vào một dòng để sửa nội dung.',
    key: null, title: r => r.Page + ' · ' + r.Field, fixed: true, cols: ['Page', 'Field', 'Value'], filters: ['Page'], required: ['Page', 'Field'],
    fields: { Field: {}, Value: { long: 3, wide: 1 } }, site: r => (r.Page === 'Research' ? 'research.html' : 'index.html'),
  },
  'Home Items': {
    icon: '🗂️', group: 'pages', desc: 'Các danh sách trên trang chủ: Education, Experience, Courses, Honors, Service (cột Section). Các dòng cùng Section hiện trên web theo đúng thứ tự trong bảng.',
    key: null, title: r => r.Section + ' · ' + r.Title, ordered: true, orderWithin: 'Section', cols: ['Section', 'Title', 'Organization', 'Dates'], filters: ['Section'], required: ['Section', 'Title'],
    fields: { Title: { wide: 1 }, Organization: { wide: 1 }, Dates: { wide: 1 }, Description: { long: 4, wide: 1 } },
    site: r => (r.Section === 'Education' ? 'index.html#about' : 'index.html#experience'),
  },
  'Venue Notes': {
    icon: '📝', desc: 'Các ghi chú hiện bên dưới bảng hội nghị và tạp chí, theo đúng thứ tự trong bảng.', key: null, title: r => (r.Note || '').slice(0, 80), ordered: true,
    cols: ['Applies To', 'Note'], filters: ['Applies To'], required: ['Applies To', 'Note'],
    fields: { Note: { long: 4, wide: 1 } },
    site: () => 'research.html#venues',
  },
};

// Home page (index.html): texts in the Profile sheet are fixed fields, one per row
const PROFILE_HINT = {
  'Avatar': 'Ảnh đại diện. Chọn một ảnh có sẵn trong thư mục website, hoặc tải ảnh từ máy lên (ảnh sẽ được chép vào assets/img). Nên dùng ảnh vuông, ít nhất 480×480 px.',
  'Kicker': 'Dòng chữ nhỏ màu xanh ngay trên tên, ở đầu trang chủ.',
  'Name': 'Tên hiển thị cỡ lớn ở đầu trang chủ.',
  'Headline': 'Câu giới thiệu ngắn ngay dưới tên.',
  'Affiliation': 'Đơn vị công tác, hiện dưới câu giới thiệu.',
  'About': 'Nội dung mục About. Muốn xuống đoạn thì để một dòng trống giữa hai đoạn.',
  'Intro: Research interests': 'Câu giới thiệu dưới tiêu đề Research interests.',
  'Intro: Thesis projects': 'Câu giới thiệu dưới tiêu đề Undergraduate thesis projects.',
  'Intro: Publications': 'Câu giới thiệu dưới tiêu đề Publications.',
  'Intro: Experience': 'Câu giới thiệu dưới tiêu đề Experience.',
  'Courses heading': 'Tiêu đề của danh sách môn dạy trong mục Experience.',
  'Join text': 'Đoạn chữ trong khung màu xanh đậm "Join my research group".',
  'Intro: Contact': 'Câu giới thiệu dưới tiêu đề Contact.',
};
const PAGE_HINT = {
  'Kicker': 'Dòng chữ nhỏ màu xanh ngay trên tiêu đề, ở đầu trang research.html.',
  'Title': 'Tiêu đề lớn ở đầu trang research.html.',
  'Lead': 'Đoạn giới thiệu ngay dưới tiêu đề.',
  'Step 1 title': 'Tiêu đề của thẻ bước 1.', 'Step 1 text': 'Nội dung của thẻ bước 1.',
  'Step 2 title': 'Tiêu đề của thẻ bước 2.', 'Step 2 text': 'Nội dung của thẻ bước 2.',
  'Step 3 title': 'Tiêu đề của thẻ bước 3.', 'Step 3 text': 'Nội dung của thẻ bước 3.',
  'Gain heading': 'Tiêu đề cột thứ nhất trong khung màu xanh đậm.', 'Gain list': 'Các quyền lợi, mỗi ý ghi một dòng.',
  'Requirements heading': 'Tiêu đề cột thứ hai.', 'Requirements list': 'Các yêu cầu, mỗi ý ghi một dòng.',
  'Own idea heading': 'Tiêu đề cột thứ ba.', 'Own idea text': 'Đoạn chữ phía trên hai nút Email me và Contact.',
  'Intro: Research topics': 'Câu giới thiệu dưới tiêu đề Research topics.',
  'Intro: Where to publish': 'Câu giới thiệu dưới tiêu đề Where to publish.',
  'Intro: Contact': 'Câu giới thiệu dưới tiêu đề Contact của research.html.',
  'Contact note': 'Dòng ghi chú nhỏ ở cuối mục Contact.',
};
const MARKUP_HINT = ' Có thể viết **chữ đậm** và [chữ](link), ví dụ [Contact](#contact).';
const ITEM_FIELDS = {
  Education: { Title: 'Tên bằng cấp (in đậm), ví dụ "M.S. in Computational Science".', Organization: 'Tên trường và quốc gia.', Dates: 'Thời gian học, GPA, xếp loại (hiện bằng chữ nhỏ màu xám).' },
  Experience: { Title: 'Chức danh, ví dụ "AI Lecturer".', Organization: 'Đơn vị công tác (hiện bằng chữ xanh).', Dates: 'Thời gian, ví dụ "Sep 2025 – Present".', Description: 'Mô tả công việc. Để một dòng trống để tách đoạn. Công việc mới nhất nên để trên cùng.' },
  Courses: { Title: 'Tên môn học.' },
  Honors: { Title: 'Tên giải thưởng (in đậm).', Description: 'Mô tả ngắn về giải thưởng.' },
  Service: { Title: 'Vai trò (in đậm), ví dụ "Conference reviewer".', Description: 'Tên hội nghị hoặc tạp chí và năm.' },
};
let IMAGES = [];
async function loadImages() { try { IMAGES = (await api('/api/images')).images; } catch { IMAGES = []; } }

const cfg = s => CFG[s] || { icon: '▦', desc: '', cols: DB.sheets[s].headers.slice(0, 5), required: [], fields: {}, title: r => Object.values(r)[0] };

/* ------------------------------------------------------------ state */
let DB = null, PRED = null, view = 'overview', lock = null, editing = null, busy = false, lastVersion = '';
// Sidebar, grouped by web page. An entry with "lock" shows one sheet already filtered on one column
// (e.g. Home Items where Section = Education), so each part of a page has its own place.
const NAV = [
  { page: 'Trang chủ · index.html' },
  { sub: 'Page content' },
  { s: 'Page Text', lock: ['Page', 'Home'], label: 'Header & text', icon: '🏠', desc: 'Avatar, tên, câu giới thiệu, mục About, các câu giới thiệu dưới tiêu đề và khung Join trên trang chủ. Bấm vào một dòng để sửa.' },
  { s: 'Home Items', lock: ['Section', 'Education'], label: 'Education', icon: '🏫', desc: 'Khung Education trong mục About. Các dòng hiện trên web theo đúng thứ tự trong bảng; dùng nút ↑ ↓ trong khung chỉnh sửa để đổi.' },
  { s: 'Home Items', lock: ['Section', 'Experience'], label: 'Experience', icon: '💼', desc: 'Mốc thời gian Academic trong mục Experience. Nên để công việc gần đây nhất lên đầu.' },
  { s: 'Home Items', lock: ['Section', 'Courses'], label: 'Courses', icon: '📘', desc: 'Danh sách môn dạy trong mục Experience. Tiêu đề của danh sách được sửa ở Header & text, dòng Courses heading.' },
  { s: 'Home Items', lock: ['Section', 'Honors'], label: 'Honors and awards', icon: '🏆', desc: 'Danh sách giải thưởng (Honors and awards) trong mục Experience.' },
  { s: 'Home Items', lock: ['Section', 'Service'], label: 'Academic service', icon: '🤝', desc: 'Các hoạt động học thuật (Academic service) trong mục Experience.' },
  { sub: 'Data' },
  { s: 'Publications' }, { s: 'Projects' }, { s: 'Directions', note: 'Dùng cho cả hai trang' },
  { page: 'Trang research.html' },
  { sub: 'Page content' },
  { s: 'Page Text', lock: ['Page', 'Research'], label: 'Header & text', icon: '🧾', desc: 'Phần đầu trang, ba bước đăng ký, khung quyền lợi và yêu cầu, các câu giới thiệu và ghi chú trên research.html. Có thể viết **chữ đậm** và [chữ](link).' },
  { sub: 'Data' },
  { s: 'Topics' }, { s: 'Conferences' }, { s: 'Journals' }, { s: 'Venue Notes' },
];
const navEntry = () => NAV.find(n => n.s === view && (lock ? n.lock && n.lock[1] === lock.v : !n.lock));
const vkey = () => view + (lock ? '@' + lock.v : '');
const inLock = r => !lock || r[lock.h] === lock.v;
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
  setSync('ok', 'Đã kết nối với database.xlsx');
  renderNav(); render();
  if (!quiet) hideBanner();
}
async function op(body, okMsg) {
  if (busy) return null;
  busy = true; setSync('', 'Đang lưu…');
  try {
    const res = await api('/api/op', { ...body, version: DB.version });
    await reload(true);
    toast(okMsg || 'Đã lưu', 'ok');
    return res;
  } catch (e) {
    if (e.status === 409) { await reload(true); banner(e.message, 'warn'); }
    else if (e.status === 423) { setSync('warn', 'Excel đang mở file'); dialog('Chưa lưu được', esc(e.message), [{ t: 'Đóng', primary: 1 }]); }
    else { setSync('err', 'Lưu không thành công'); toast(e.message, 'err'); }
    return null;
  } finally { busy = false; }
}

/* ------------------------------------------------------------ checks */
let ISSUES = null;
function checkRow(sheet, r, all) {
  const C = cfg(sheet), S = DB.sheets[sheet], out = [];
  const add = (level, field, msg) => out.push({ level, field, msg });
  for (const h of C.required || []) if (!String(r[h] || '').trim()) add('err', h, 'Không được để trống.');
  for (const [h, opts] of Object.entries(S.options || {})) {
    if (!Array.isArray(opts) || !r[h]) continue;
    if (!opts.includes(r[h])) add('err', h, `"${r[h]}" không có trong danh sách cho phép (${opts.join(', ')}).`);
  }
  for (const [h, f] of Object.entries(C.fields || {})) if (f.url && r[h] && !isUrl(r[h])) add('warn', h, 'Link nên bắt đầu bằng https://.');
  if (C.key && r[C.key]) {
    const dup = (all || rowsOf(sheet)).find(o => o._i !== r._i && norm(o[C.key]) === norm(r[C.key]));
    if (dup) add('warn', C.key, `Bị trùng với dòng ${dup._i + 1}.`);
  }
  const multiCheck = h => {
    if (!r[h]) return;
    const bad = r[h].split(',').map(s => s.trim()).filter(s => s && s !== 'All' && !codes().includes(s));
    if (bad.length) add('err', h, 'Không có mã hướng: ' + bad.join(', ') + '.');
  };
  if (sheet === 'Publications') {
    if (r.Date && !/\d{4}/.test(r.Date)) add('err', 'Date', 'Thiếu năm (4 chữ số) nên bài sẽ không hiện trên web.');
    if (r.Status === 'Preprint' && r.Type !== 'Preprint') add('warn', 'Type', 'Status là Preprint thì Type cũng nên chọn Preprint.');
    if (r.Type === 'Preprint' && !['Preprint', 'In Submission'].includes(r.Status)) add('warn', 'Status', 'Bài preprint nên chọn Status là Preprint (hoặc In Submission nếu đang gửi bài).');
    if (r.Type === 'Preprint' && !/arxiv[^0-9]{0,12}\d{4}\.\d{4,5}/i.test([r.Venue, r.PDF, r.DOI].join(' '))) add('warn', 'Venue', 'Chưa có arXiv ID. Hãy dán ID vào ô arXiv ở đầu khung chỉnh sửa.');
    if (r.Type === 'Thesis' && r.Status !== 'Thesis') add('warn', 'Status', 'Luận văn nên chọn Status là Thesis.');
    if (r['Link ID']) { const d = rowsOf(sheet).find(o => o._i !== r._i && o['Link ID'] && o['Link ID'].toLowerCase() === r['Link ID'].toLowerCase()); if (d) add('err', 'Link ID', `Link ID bị trùng với dòng ${d._i + 1}.`); }
  }
  if (sheet === 'Projects') {
    if (r.Semester) { const sv = parseSemester(r.Semester); if (!sv.season || !sv.year) add('err', 'Semester', 'Chưa đúng dạng học kỳ: chọn Spring, Summer hoặc Fall và nhập năm.'); }
    if (r['Paper Status'] === 'Accepted') {
      if (!r['Link ID']) add('warn', 'Link ID', 'Bài đã được nhận, nên điền Link ID để nối với bài báo.');
      else if (!linkIds('Publications').some(x => x.toLowerCase() === r['Link ID'].toLowerCase())) add('err', 'Link ID', 'Trong Publications không có bài báo nào mang Link ID này.');
    }
    if (r['Paper Status'] === 'Submitted' && !r['Paper Venue']) add('warn', 'Paper Venue', 'Nên ghi nơi đã gửi bài.');
  }
  if (sheet === 'Page Text' && r.Page === 'Home') {
    if (r.Field === 'Avatar') {
      if (!r.Value) add('err', 'Value', 'Chưa chọn ảnh, trang chủ sẽ chỉ hiện chữ viết tắt.');
      else if (IMAGES.length && !IMAGES.some(i => i.path === r.Value)) add('err', 'Value', `Không tìm thấy ảnh "${r.Value}" trong thư mục website.`);
    } else if (['Name', 'Headline'].includes(r.Field) && !r.Value) add('warn', 'Value', 'Đang để trống, trang chủ sẽ dùng chữ có sẵn trong index.html.');
  }
  if (sheet === 'Home Items' && r.Section && ITEM_FIELDS[r.Section]) {
    const used = ITEM_FIELDS[r.Section];
    for (const h of ['Organization', 'Dates', 'Description']) if (r[h] && !used[h]) add('warn', h, `Phần ${r.Section} không dùng cột ${h}, nên nội dung này sẽ không hiện trên web.`);
  }
  if (sheet === 'Directions' && r.Code && !/^[A-Z0-9]+$/.test(r.Code)) add('err', 'Code', 'Chỉ dùng chữ hoa và số, không dấu, không có khoảng trắng.');
  if (sheet === 'Topics') {
    if (r.Direction && !codes().includes(r.Direction)) add('err', 'Direction', 'Mã hướng này không có trong Directions.');
    for (const [h, labels] of [['Research Description', DESC], ['Research Expectation', EXPC]]) {
      if (!r[h]) continue;
      const t = tplCheck(r[h], labels);
      if (t.missing.length) add('err', h, 'Còn thiếu: ' + t.missing.join(', ') + '.');
      if (t.stray.length) add('warn', h, `Có ${t.stray.length} dòng không thuộc mục nào nên web sẽ bỏ qua, ví dụ: "${t.stray[0].slice(0, 60)}"`);
      if (h === 'Research Description' && t.fields.Context !== undefined) {
        const n = sentences(t.fields.Context);
        if (n < 3) add('warn', h, `Context mới có ${n} câu, cần ít nhất 3 câu.`);
      }
    }
    if (r.Dataset) {
      const ds = parseDatasets(r.Dataset);
      if (ds.some(d => !d.name)) add('err', 'Dataset', 'Có dataset chưa có tên.');
      const badUrl = ds.find(d => d.url && !isUrl(d.url));
      if (badUrl) add('warn', 'Dataset', `Link của "${badUrl.name}" nên bắt đầu bằng https://.`);
    }
  }
  if (sheet === 'Conferences') {
    multiCheck('Directions');
    for (const h of ['Submission Deadline', 'Held In']) if (r[h] && !hasMonth(r[h])) add('warn', h, 'Chưa có tên tháng bằng tiếng Anh, nên bộ lọc và sắp xếp trên web sẽ bỏ qua dòng này.');
    if (r['Registration Fee']) {
      const bad = r['Registration Fee'].split('\n').filter(l => l.trim() && !/^(regular|student)\s*:/i.test(l.trim()));
      if (bad.length) add('warn', 'Registration Fee', 'Mỗi dòng nên bắt đầu bằng "Regular:" hoặc "Student:".');
      else if (!/\d/.test(r['Registration Fee'])) add('warn', 'Registration Fee', 'Chưa ghi số tiền.');
    }
    if (/waset|world academy of science/i.test((r['Full Name'] || '') + ' ' + (r.Website || ''))) add('err', 'Website', 'Có dấu hiệu là hội nghị WASET (hội nghị săn mồi).');
  }
  if (sheet === 'Journals') {
    multiCheck('Directions');
    if (r['Time to Publish'] && !/\d/.test(r['Time to Publish'])) add('warn', 'Time to Publish', 'Nên ghi số tháng, ví dụ "4–8 months".');
    if (r['Publication Fee']) {
      const bad = r['Publication Fee'].split('\n').filter(l => l.trim() && !/^(free|open access)/i.test(l.trim()));
      if (bad.length) add('warn', 'Publication Fee', 'Mỗi dòng nên có dạng "Free (subscription)", "Free" hoặc "Open access: USD …".');
    }
    if (PRED) {
      if (r.Journal && PRED.journals.includes(norm(r.Journal))) add('err', 'Journal', 'Tạp chí này nằm trong danh sách tạp chí săn mồi.');
      if (r.Publisher && PRED.publishers.includes(norm(r.Publisher))) add('err', 'Publisher', 'Nhà xuất bản này nằm trong danh sách nhà xuất bản săn mồi.');
      const d = host(r.Link); if (d && PRED.domains.some(x => d === x || d.endsWith('.' + x))) add('err', 'Link', 'Tên miền ' + d + ' nằm trong danh sách săn mồi.');
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
  const bad = (sheet, lk) => iss[sheet].reduce((n, a, i) => n + (a.some(x => x.level === 'err') && (!lk || rowsOf(sheet)[i][lk[0]] === lk[1]) ? 1 : 0), 0);
  const total = DB.order.reduce((n, s) => n + bad(s), 0);
  const listed = new Set(NAV.filter(n => n.s && !n.lock).map(n => n.s).concat(NAV.filter(n => n.lock).map(n => n.s)));
  const extra = DB.order.filter(s => !listed.has(s)).map(s => ({ s }));
  const entries = NAV.filter(n => !n.s || DB.sheets[n.s]).concat(extra.length ? [{ page: 'Các sheet khác' }, ...extra] : []);
  $('#nav').innerHTML = `<a href="#overview" class="${view === 'overview' ? 'on' : ''}"><span>◈</span> Tổng quan ${total ? `<span class="bad">${total}</span>` : ''}</a>` +
    entries.map(n => {
      if (n.page) return `<div class="grp">${esc(n.page)}</div>`;
      if (n.sub) return `<div class="sub">${esc(n.sub)}</div>`;
      const key = n.s + (n.lock ? '@' + n.lock[1] : ''), on = vkey() === key && view !== 'overview';
      const c = bad(n.s, n.lock), cnt = n.lock ? rowsOf(n.s).filter(r => r[n.lock[0]] === n.lock[1]).length : rowsOf(n.s).length;
      return `<a href="#${encodeURIComponent(key)}" class="${on ? 'on' : ''}" ${n.note ? `title="${esc(n.note)}"` : ''}><span>${n.icon || cfg(n.s).icon}</span> ${esc(n.label || n.s)}${c ? `<span class="bad" title="${c} dòng có lỗi">${c}</span>` : ''}<span class="n">${cnt}</span></a>`;
    }).join('');
}

function render() {
  const top = $('#topact');
  if (view === 'overview' || !DB.sheets[view]) {
    view = 'overview';
    $('#title').textContent = 'Tổng quan'; $('#subtitle').textContent = 'Mọi thay đổi ở đây được lưu ngay vào data/database.xlsx.';
    top.innerHTML = ''; renderOverview(); return;
  }
  const C = cfg(view), N = navEntry();
  $('#title').textContent = N && N.label ? N.label : view;
  $('#subtitle').textContent = N && N.desc ? N.desc + ' (sheet ' + view + ')' : C.desc;
  top.innerHTML = `<a class="btn" target="_blank" rel="noopener" href="/site/${esc(C.site ? C.site(lock ? { [lock.h]: lock.v } : {}) : 'index.html')}">Xem trên web ↗</a>
    ${C.fixed ? '' : '<button class="btn primary" type="button" id="addbtn">＋ Thêm mới</button>'}`;
  if ($('#addbtn')) $('#addbtn').onclick = () => openEditor(view, null);
  renderSheet();
}

/* ------------------------------------------------------------ user guide (Overview page) */
const GUIDE_HTML = `
<div class="panel guide"><h3>Hướng dẫn sử dụng</h3>
<p>Bấm vào từng mục để xem. Mỗi lần bấm <b>Lưu</b>, thay đổi được ghi ngay vào file <code>data/database.xlsx</code> – chính là file mà website đọc.</p>

<details open><summary>1. Cập nhật website trong 5 bước</summary>
<ol class="steps">
  <li><b>Mở công cụ:</b> trong thư mục website (thư mục đã clone từ GitHub), nhấp đúp <code>admin\\Open database manager.bat</code>. Trình duyệt sẽ mở <code>http://localhost:8765/</code>. Để nguyên cửa sổ màu đen trong lúc làm việc.</li>
  <li><b>Đóng Excel</b> nếu đang mở file database.xlsx, vì Windows khóa file đang mở và công cụ sẽ không lưu được. Chấm tròn ở góc dưới bên trái cho biết trạng thái: xanh là sẵn sàng, vàng là Excel đang mở file.</li>
  <li><b>Chỉnh sửa:</b> chọn một mục ở cột bên trái, bấm vào dòng cần sửa (hoặc bấm “＋ Thêm mới”), sửa xong bấm <b>Lưu</b> (hoặc Ctrl + S).</li>
  <li><b>Xem trước:</b> bấm “Xem website” ở góc dưới bên trái, hoặc nút “Xem trên web” trong khung chỉnh sửa, để xem trang web với dữ liệu vừa sửa.</li>
  <li><b>Đăng lên web:</b> mở GitHub Desktop, ghi vài chữ mô tả thay đổi (ví dụ “Thêm đề tài MI”), bấm <b>Commit to main</b> rồi <b>Push origin</b>. Khoảng 1–2 phút sau website sẽ cập nhật.</li>
</ol></details>

<details><summary>2. Các phần trên màn hình</summary>
<ul class="steps">
  <li><b>Cột bên trái:</b> chia theo hai trang web là <b>index.html</b> (trang chủ) và <b>research.html</b>. Mỗi trang gồm phần <b>Page content</b> (chữ và danh sách hiển thị trên trang) và phần <b>Data</b> (bài báo, đề tài, hội nghị…). Số bên phải là số dòng; số trong ô đỏ là số dòng đang có lỗi.</li>
  <li><b>Tổng quan:</b> số liệu nhanh, danh sách <b>Kiểm tra dữ liệu</b> (bấm vào một lỗi để mở đúng dòng cần sửa) và phần hướng dẫn này.</li>
  <li><b>Góc dưới bên trái:</b> trạng thái kết nối, nút <b>Hoàn tác</b>, nút xem website và đường dẫn tới file đang sửa.</li>
  <li><b>Khung chỉnh sửa</b> mở ra ở bên phải. Ô có dấu <span style="color:var(--err)">*</span> là bắt buộc. Viền đỏ là lỗi, viền cam là cảnh báo; tất cả được tóm tắt ở đầu khung.</li>
</ul></details>

<details><summary>3. Tìm kiếm, lọc, sắp xếp, chọn cột</summary>
<ul class="steps">
  <li><b>Tìm kiếm</b> (phím <kbd>/</kbd>): tìm trong mọi cột, gõ có dấu hay không dấu đều được. Ở Topics có thể tìm theo mã đề tài, ví dụ MI-07.</li>
  <li><b>Lọc nhanh:</b> các ô chọn trên thanh công cụ (Direction, Status, Rank…). Ô “Kiểm tra” giúp chỉ xem những dòng có lỗi hoặc cảnh báo.</li>
  <li><b>Lọc nâng cao:</b> đặt nhiều điều kiện cùng lúc; chỉ những dòng thỏa mãn <b>tất cả</b> điều kiện mới hiện ra. Các phép so sánh: chứa, không chứa, bằng, khác, để trống, có nội dung, từ (≥), đến (≤). Phép “từ/đến” hiểu đúng tháng (“Submission Deadline từ Mar”), năm (“Date từ 2025”), học kỳ, hạng (A* &gt; A &gt; B) và số tiền (quy ra USD).</li>
  <li><b>Sắp xếp:</b> bấm vào tên cột để sắp xếp (bấm lần nữa để đảo chiều, lần thứ ba để bỏ). Giữ <kbd>Shift</kbd> khi bấm để sắp xếp thêm theo cột thứ hai, thứ ba. Nút “Sắp xếp” cho phép chọn tối đa 3 cột ưu tiên. Ô trống luôn nằm cuối.</li>
  <li><b>Chọn cột:</b> chọn những cột muốn hiện trong bảng (chỉ thay đổi cách xem, không ảnh hưởng tới file).</li>
  <li>Các bộ lọc và cách sắp xếp đang dùng hiện thành nhãn có dấu ×; bấm “Bỏ hết lọc và sắp xếp” để trở về mặc định. Mỗi mục tự nhớ cách xem riêng.</li>
  <li><b>Sửa nhiều dòng một lúc:</b> tích ô vuông đầu dòng (hoặc ô ở hàng tiêu đề để chọn tất cả các dòng đang hiện). Thanh “Đã chọn” cho phép gán cùng một giá trị cho tất cả (ví dụ Status là Taken) hoặc xóa các dòng đã chọn.</li>
</ul></details>

<details><summary>4. Thêm, sửa, tạo bản sao, xóa, hoàn tác</summary>
<ul class="steps">
  <li><b>Thêm:</b> bấm “＋ Thêm mới”, điền thông tin rồi bấm “Thêm vào Excel”. Dòng mới luôn được thêm vào cuối.</li>
  <li><b>Sửa:</b> bấm vào dòng, sửa rồi bấm <b>Lưu</b> (Ctrl + S). Bấm <kbd>Esc</kbd> để đóng; nếu còn thay đổi chưa lưu, công cụ sẽ hỏi lại.</li>
  <li><b>Tạo bản sao:</b> mở một dòng rồi bấm “Tạo bản sao” để có khung thêm mới với nội dung giống hệt – tiện khi thêm hội nghị hay tạp chí tương tự.</li>
  <li><b>Xóa:</b> mở dòng rồi bấm “Xóa”. Với đề tài, nên <b>đánh dấu Taken</b> thay vì xóa (xem mục Topics).</li>
  <li><b>Hoàn tác:</b> trước mỗi lần lưu, bản cũ được sao lưu vào <code>admin\\backups</code> (giữ 60 bản gần nhất). Nút “Hoàn tác” đưa file về trạng thái trước lần lưu gần nhất; bấm nhiều lần để quay lại xa hơn.</li>
  <li>Nếu bạn sửa file bằng Excel trong lúc công cụ đang mở, công cụ sẽ tự tải lại dữ liệu mới. Nếu dòng bạn đang sửa vừa bị thay đổi ở nơi khác, công cụ sẽ không ghi đè mà tải lại để bạn kiểm tra.</li>
</ul></details>

<details><summary>5. Page content: avatar, About, Education, Experience và các đoạn giới thiệu</summary>
<ul class="steps">
  <li>Trong phần <b>Page content</b> của mỗi trang, sửa xong bấm Lưu rồi “Xem trên web” để kiểm tra ngay.</li>
  <li><b>Header &amp; text</b> (trong file Excel là sheet <b>Page Text</b>): không thêm hay xóa dòng, chỉ sửa cột Value; mỗi dòng có ghi chú cho biết chữ đó nằm ở đâu trên trang.
    <br>• Trang chủ: <b>Avatar</b> (chọn ảnh có sẵn trong thư mục website, hoặc bấm “Tải ảnh từ máy lên…” để chép ảnh vào <code>assets/img</code>), Kicker / Name / Headline / Affiliation ở đầu trang, <b>About</b> (để một dòng trống giữa hai đoạn), các dòng Intro: … dưới mỗi tiêu đề, Courses heading và Join text.
    <br>• Trang research.html: Kicker / Title / Lead, ba bước đăng ký, khung quyền lợi và yêu cầu (các mục “… list” ghi mỗi ý một dòng), các dòng Intro: … và Contact note. Có thể viết <code>**chữ đậm**</code> và <code>[chữ](link)</code>.</li>
  <li><b>Education, Experience, Courses, Honors and awards, Academic service</b> (trong file Excel là sheet <b>Home Items</b>): khung chỉnh sửa chỉ hiện những ô mà phần đó dùng. Thứ tự các dòng chính là thứ tự hiển thị trên web; dùng nút ↑ ↓ để đổi.</li>
  <li>Khi đăng lên web, nhớ commit cả <code>data/database.xlsx</code> lẫn ảnh mới trong <code>assets/img</code> (GitHub Desktop sẽ tự liệt kê), rồi bấm Push origin.</li>
</ul></details>

<details><summary>6. Publications (bài báo)</summary>
<ul class="steps">
  <li><b>Type</b>: Journal, Conference, Thesis, Preprint. <b>Status</b>: Published, Accepted, In Submission, Preprint, Thesis.</li>
  <li><b>Bài đang gửi đi hoặc đang phản biện:</b> chọn Type theo nơi gửi (Journal hoặc Conference) và Status là In Submission. Trên web, các bài này nằm trong nhóm “In submission” (S1, S2…).</li>
  <li><b>Preprint (arXiv):</b> chọn Type là Preprint, ô “arXiv ID hoặc link” sẽ hiện ra. Dán link arXiv vào, Venue, PDF và DOI sẽ tự điền. Bấm “Lấy thông tin từ arXiv” để điền nốt tên bài, tác giả và ngày (chỉ điền vào các ô còn trống). Trên web, venue hiện là “arXiv preprint arXiv:…” kèm link, trong nhóm “Preprints” (P1, P2…).</li>
  <li><b>Date:</b> dd-mm-yyyy, mm-yyyy hoặc chỉ yyyy. Bài không có năm sẽ không hiện trên web.</li>
  <li><b>Authors:</b> cách nhau bằng dấu phẩy; tên “Ha Anh Vu” được tự động in đậm.</li>
  <li><b>Link ID:</b> một mã ngắn (ví dụ AC1) dùng để nối bài báo với đề tài luận văn trong Projects (xem mục 7).</li>
  <li>Web tự sắp xếp (mới nhất lên trước) và gom các bài cùng hội nghị; bạn không cần quan tâm thứ tự dòng.</li>
</ul></details>

<details><summary>7. Projects (đề tài luận văn đang hướng dẫn)</summary>
<ul class="steps">
  <li><b>Semester:</b> chọn mùa (Spring, Summer hoặc Fall) và nhập năm. Trên web, năm mới nhất hiện trước; trong cùng một năm, Spring đứng trước Summer, Summer trước Fall. Hai học kỳ gần nhất được mở sẵn.</li>
  <li><b>Members:</b> tên các sinh viên, cách nhau bằng dấu chấm phẩy (;).</li>
  <li><b>Thesis Status:</b> Preparing (sắp bắt đầu), In progress (đang làm), Defended (đã bảo vệ).</li>
  <li><b>Paper Status:</b> None (chưa có bài), Submitted (đã gửi bài – khi đó điền thêm Paper Venue), Accepted (đã được nhận). Khi bài được nhận: thêm bài vào Publications và ghi <b>cùng một Link ID</b> ở cả hai dòng; trên web sẽ hiện “Accepted, see C1”, bấm vào sẽ chuyển tới bài báo. Công cụ sẽ báo lỗi nếu Link ID không khớp.</li>
</ul></details>

<details><summary>8. Directions (hướng nghiên cứu)</summary>
<ul class="steps">
  <li>Thứ tự các dòng cũng là thứ tự các thẻ Research interests, các nút lọc và các nhóm đề tài trên web. Mở một hướng rồi dùng nút ↑ ↓ để đổi thứ tự.</li>
  <li><b>Code</b> viết hoa, không dấu (CV, MI…). Khi <b>đổi mã</b> trong công cụ, mọi đề tài, hội nghị và tạp chí đang dùng mã đó sẽ được cập nhật theo (mã đề tài trên web cũng đổi, ví dụ CV-01 thành CVX-01).</li>
  <li><b>Summary</b> (khoảng 10 từ tiếng Anh) hiện trên thẻ ở trang chủ; <b>Focus</b> hiện thành dòng “Current trends”.</li>
  <li>Không nên xóa một hướng khi vẫn còn đề tài thuộc hướng đó, vì các đề tài này sẽ bị tách thành một nhóm lạ trên web.</li>
</ul></details>

<details><summary>9. Topics (đề tài cho sinh viên)</summary>
<ul class="steps">
  <li><b>Mã đề tài</b> (CV-01, MI-07…) được web tự đánh theo vị trí của dòng trong từng hướng. Vì vậy hãy <b>thêm đề tài mới vào cuối</b>, và khi đã có người nhận thì <b>đánh dấu Taken</b> thay vì xóa – xóa sẽ làm thay đổi mã của các đề tài phía sau. Khung chỉnh sửa luôn cho biết trước mã mà đề tài sẽ có.</li>
  <li>Đề tài Taken không hiện trong danh sách chính, chỉ hiện ở mục “Taken” trên web.</li>
  <li><b>Dataset:</b> mỗi dòng gồm tên (bắt buộc) và link (không bắt buộc); bấm “＋ Thêm dataset” để thêm dòng. Nếu dán cả tên lẫn link vào ô tên, công cụ sẽ tự tách ra. Trên web, tên dataset hiện thành link.</li>
  <li><b>Research Description</b> phải có đủ 6 mục, mỗi mục bắt đầu ở đầu dòng bằng đúng tên mục và dấu hai chấm: Context, Problem formulation, Gap, Proposed approach, Implementation steps, Key knowledge to learn. <b>Research Expectation</b> cũng vậy: Deliverables, Expected results, Baselines, Metrics, Experimental protocol, Minimum success criteria. Dấu ✓/✗ bên dưới ô cho biết mục nào còn thiếu; nút “Thêm các mục còn thiếu” sẽ điền sẵn tên mục.</li>
  <li><b>Context</b> là đoạn tóm tắt hiện trên thẻ đề tài, cần ít nhất 3 câu: bối cảnh, khó khăn hiện tại và việc sinh viên sẽ làm. Góc phải của ô có hiện số từ và số câu.</li>
  <li>Để đánh số các bước hoặc kết quả: xuống dòng sau dấu hai chấm, mỗi ý một dòng, bắt đầu bằng (1), (2), (3)…</li>
</ul></details>

<details><summary>10. Conferences (hội nghị)</summary>
<ul class="steps">
  <li>Chỉ thêm hội nghị có kỷ yếu được <b>Scopus</b> index; ghi nơi xuất bản kỷ yếu ở cột Proceedings (cột này cũng được dùng cho ô tìm kiếm trên web). Công cụ sẽ cảnh báo các hội nghị có dấu hiệu WASET.</li>
  <li><b>CORE Rank:</b> theo bảng ICORE2026 (portal.core.edu.au); nếu không có hạng A*–C thì chọn Not ranked.</li>
  <li><b>Directions:</b> bấm chọn các mã hướng phù hợp, hoặc All.</li>
  <li><b>Submission Deadline / Held In:</b> ghi tên tháng bằng tiếng Anh (“Mid December – January”, “October–November”) để bộ lọc và sắp xếp trên web đọc được.</li>
  <li><b>Hybrid:</b> Yes nếu cho phép trình bày online; No nếu phải đến trình bày trực tiếp.</li>
  <li><b>Website:</b> trang chính thức; nếu hội nghị chỉ có trang riêng cho từng năm thì dùng trang dblp.</li>
  <li><b>Registration Fee:</b> mỗi mức phí một dòng, có nút điền nhanh “Regular: ≈USD 900” và “Student: ≈USD 600”. Dấu ≈ nghĩa là số ước lượng; không có ≈ là số lấy từ trang đăng ký.</li>
</ul></details>

<details><summary>11. Journals (tạp chí)</summary>
<ul class="steps">
  <li>Công cụ đối chiếu tên tạp chí, nhà xuất bản và tên miền với hai danh sách săn mồi trong <code>admin\\predatory</code> và báo lỗi nếu trùng. Khi có danh sách mới, chỉ cần thay hai file trong thư mục đó (giữ nguyên tên file). Không nên lưu tạp chí đang bị báo lỗi này.</li>
  <li><b>Rank:</b> Q1–Q4 theo SCImago.</li>
  <li><b>Time to Publish:</b> số tháng từ lúc gửi bài tới khi bài được đăng online, ví dụ “4–8 months” (có nút điền nhanh).</li>
  <li><b>Publication Fee:</b> mỗi hình thức đăng một dòng: “Free (subscription)” (tạp chí hybrid, đăng không mất phí), “Free” (open access miễn phí), “Open access: USD 3,270” (phí APC). Tạp chí hybrid ghi cả hai dòng.</li>
</ul></details>

<details><summary>12. Venue Notes</summary>
<ul class="steps">
  <li>Các ghi chú hiện bên dưới bảng hội nghị và tạp chí. Applies To chọn Conferences, Journals hoặc All. Thứ tự dòng chính là thứ tự hiển thị; dùng ↑ ↓ để đổi.</li>
</ul></details>

<details><summary>13. Phím tắt</summary>
<ul class="steps">
  <li><kbd>/</kbd> tìm kiếm · <kbd>Ctrl</kbd> + <kbd>S</kbd> lưu · <kbd>Esc</kbd> đóng khung chỉnh sửa hoặc hộp thoại · <kbd>Shift</kbd> + bấm tên cột: sắp xếp thêm theo cột đó.</li>
</ul></details>

<details><summary>14. Khi gặp sự cố</summary>
<ul class="steps">
  <li><b>Báo “File database.xlsx đang mở trong Excel”:</b> đóng Excel (lưu hoặc bỏ thay đổi bên Excel) rồi bấm Lưu lại.</li>
  <li><b>Báo “Mất kết nối”:</b> cửa sổ màu đen đã bị tắt. Nhấp đúp lại file .bat rồi tải lại trang.</li>
  <li><b>Giao diện không thay đổi sau khi cập nhật công cụ:</b> bấm Ctrl + F5.</li>
  <li><b>Lỡ lưu nhầm:</b> bấm “Hoàn tác”. Các bản cũ hơn nằm trong <code>admin\\backups</code> (mở bằng Excel để lấy lại dữ liệu).</li>
  <li><b>Đã Push mà web chưa đổi:</b> chờ 1–2 phút rồi tải lại trang (Ctrl + F5).</li>
  <li>Sheet “Hướng dẫn” trong file Excel giải thích chi tiết từng cột, nếu bạn muốn sửa trực tiếp bằng Excel.</li>
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
      ${card('Conferences', rowsOf('Conferences').length, 'Hội nghị có kỷ yếu Scopus')}
      ${card('Journals', rowsOf('Journals').length, 'Tạp chí khoa học')}
      ${card('Publications', pubs.length, pubs.filter(p => p.Status === 'Published').length + ' Published')}
      ${card('Projects', proj.length, [...new Set(proj.map(p => p.Semester))].length + ' học kỳ')}
      ${card('Directions', rowsOf('Directions').length, codes().join(' · '))}
    </div>
    <div class="panel"><h3>Kiểm tra dữ liệu</h3>
      <p>${list.length ? `Có ${errs} lỗi và ${list.length - errs} cảnh báo. Bấm vào từng mục để sửa.` : 'Không có lỗi hay cảnh báo nào, dữ liệu đã sẵn sàng để đăng lên web.'}</p>
      <ul class="issues">${list.slice(0, 150).map(x => {
        const r = rowsOf(x.s)[x.i];
        return `<li><span class="tag ${x.level}">${x.level === 'err' ? 'Lỗi' : 'Cảnh báo'}</span>
          <span><a data-go="${esc(x.s)}|${x.i}">${esc(x.s)} · ${esc(r._code || 'dòng ' + (x.i + 1))}</a> <span class="where">(${esc(x.field)})</span><br>${esc(x.msg)}</span></li>`;
      }).join('')}</ul>${list.length > 150 ? `<p>… và ${list.length - 150} mục khác.</p>` : ''}</div>
    ${GUIDE_HTML}`;
  $('#view').querySelectorAll('[data-go]').forEach(a => a.onclick = () => { const [s, i] = a.dataset.go.split('|'); const n = NAV.find(x => x.s === s && x.lock && rowsOf(s)[+i][x.lock[0]] === x.lock[1]); go(s, n ? n.lock[1] : null); openEditor(s, +i); });
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

const OPS = [['contains', 'chứa'], ['ncontains', 'không chứa'], ['eq', 'bằng'], ['neq', 'khác'], ['empty', 'để trống'], ['nempty', 'có nội dung'], ['gte', 'từ (≥)'], ['lte', 'đến (≤)']];
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
  const st = ui(sheet === view ? vkey() : sheet), H = DB.sheets[sheet].headers;
  const q = norm(st.q).split(' ').filter(Boolean);
  let rows = rowsOf(sheet).filter(r => {
    if (sheet === view && !inLock(r)) return false;
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
  const sheet = view, key = vkey(), st = ui(key), C = cfg(sheet), S = DB.sheets[sheet];
  if (!st._loaded) { loadPrefs(key, st); st._loaded = true; }
  const defCols = (C.cols || S.headers.slice(0, 5)).filter(h => S.headers.includes(h) && !(lock && h === lock.h));
  const cols = (st.cols || defCols).filter(h => S.headers.includes(h) && !(lock && h === lock.h));
  const rows = filtered(sheet), iss = allIssues()[sheet];
  const mine = rowsOf(sheet).filter(inLock), total = mine.length;
  const pos = new Map(mine.map((r, k) => [r._i, k + 1]));          // number rows 1..n inside this view
  const quick = h => {
    const vals = valueChoices(sheet, h) || [];
    return `<select data-f="${esc(h)}" class="${st.filters[h] ? 'act' : ''}" aria-label="${esc(h)}"><option value="">${esc(h)}: tất cả</option>${vals.map(v => `<option ${st.filters[h] === v ? 'selected' : ''}>${esc(v)}</option>`).join('')}</select>`;
  };
  const nErr = iss.filter(a => a.some(x => x.level === 'err')).length, nAny = iss.filter(a => a.length).length;
  const lvl = h => st.sort.findIndex(s => s.h === h);
  const arrow = h => { const k = lvl(h); if (k < 0) return ''; return `<span class="ar">${st.sort[k].dir > 0 ? '↑' : '↓'}${st.sort.length > 1 ? `<sup>${k + 1}</sup>` : ''}</span>`; };
  const selected = [...st.sel].filter(i => i < total);
  const bulkOpts = Object.entries(S.options).filter(([, o]) => Array.isArray(o));
  const colOpts = (cur, forFilter) => [...(forFilter ? (C.codeCol ? [['_code', 'Mã đề tài']] : []) : [['#', C.codeCol ? 'Thứ tự gốc (theo mã đề tài)' : 'Thứ tự gốc trong Excel']]), ...S.headers.map(h => [h, h])]
    .map(([v, t]) => `<option value="${esc(v)}" ${v === cur ? 'selected' : ''}>${esc(t)}</option>`).join('');
  const condVal = (c, k) => {
    if (c.op === 'empty' || c.op === 'nempty') return '<span></span>';
    const ch = c.h && c.h !== '_code' ? valueChoices(sheet, c.h) : null;
    return ch && c.op !== 'contains' && c.op !== 'ncontains'
      ? `<select data-cv="${k}"><option value="">— Chọn —</option>${ch.map(v => `<option ${v === c.v ? 'selected' : ''}>${esc(v)}</option>`).join('')}</select>`
      : `<input data-cv="${k}" value="${esc(c.v || '')}" placeholder="Nhập giá trị…">`;
  };
  const isActive = c => c.h && (c.op === 'empty' || c.op === 'nempty' || c.v);
  const nConds = st.conds.filter(isActive).length;
  const chips = [
    ...Object.entries(st.filters).filter(([, v]) => v).map(([h, v]) => `<span class="chip">${h === '_issues' ? 'Kiểm tra' : esc(h)}: <b>${esc(h === '_issues' ? (v === 'err' ? 'có lỗi' : 'có lỗi hoặc cảnh báo') : v)}</b><button type="button" data-xf="${esc(h)}" aria-label="Bỏ">×</button></span>`),
    ...st.conds.map((c, k) => isActive(c) ? `<span class="chip">${esc(c.h === '_code' ? 'Mã' : c.h)} ${esc(opLabel(c.op))} ${c.op === 'empty' || c.op === 'nempty' ? '' : `<b>${esc(c.v)}</b>`}<button type="button" data-xc="${k}" aria-label="Bỏ">×</button></span>` : ''),
    ...st.sort.map((s, k) => s.h ? `<span class="chip sort">⇅ ${st.sort.length > 1 ? (k + 1) + '. ' : ''}${esc(s.h === '#' ? 'Thứ tự gốc' : s.h)} ${s.dir > 0 ? '↑' : '↓'}<button type="button" data-xs="${k}" aria-label="Bỏ">×</button></span>` : ''),
  ].filter(Boolean);

  $('#view').innerHTML = `
    <div class="tools">
      <label class="search"><input type="search" id="q" placeholder="Tìm kiếm…  (phím /)" value="${esc(st.q)}"></label>
      ${(C.filters || []).filter(h => S.headers.includes(h) && !(lock && h === lock.h)).map(quick).join('')}
      <select data-f="_issues" class="${st.filters._issues ? 'act' : ''}" aria-label="Kiểm tra"><option value="">Kiểm tra: tất cả</option>
        <option value="err" ${st.filters._issues === 'err' ? 'selected' : ''}>Có lỗi (${nErr})</option>
        <option value="any" ${st.filters._issues === 'any' ? 'selected' : ''}>Có lỗi hoặc cảnh báo (${nAny})</option></select>
      <span class="tgroup">
        <button class="btn sm ${st.panel === 'filter' ? 'on' : ''}" type="button" data-panel="filter">Lọc nâng cao${nConds ? ` <span class="badge">${nConds}</span>` : ''}</button>
        <button class="btn sm ${st.panel === 'sort' ? 'on' : ''}" type="button" data-panel="sort">Sắp xếp${st.sort.length ? ` <span class="badge">${st.sort.length}</span>` : ''}</button>
        <button class="btn sm ${st.panel === 'cols' ? 'on' : ''}" type="button" data-panel="cols">Chọn cột</button>
      </span>
      <span class="count">${rows.length === total ? total : `<b>${rows.length}</b> / ${total}`} dòng</span>
    </div>
    ${st.panel === 'filter' ? `<div class="fpanel"><div class="ph"><b>Lọc nâng cao</b><span>Chỉ hiện những dòng thỏa mãn <b>tất cả</b> điều kiện. Phép “từ/đến” hiểu đúng tháng, ngày, học kỳ, hạng và số tiền.</span></div>
      ${st.conds.map((c, k) => `<div class="frow"><select data-ch="${k}"><option value="">— Chọn cột —</option>${colOpts(c.h, true)}</select>
        <select data-co="${k}">${OPS.map(([o, t]) => `<option value="${o}" ${o === c.op ? 'selected' : ''}>${t}</option>`).join('')}</select>
        ${condVal(c, k)}<button type="button" class="iconbtn" data-xc="${k}" aria-label="Xóa điều kiện">✕</button></div>`).join('')}
      <div class="pfoot"><button class="btn sm" type="button" id="addcond">＋ Thêm điều kiện</button>${st.conds.length ? '<button class="btn sm" type="button" id="clrcond">Xóa tất cả điều kiện</button>' : ''}</div></div>` : ''}
    ${st.panel === 'sort' ? `<div class="fpanel"><div class="ph"><b>Sắp xếp theo nhiều cột</b><span>Sắp theo cột ưu tiên 1 trước; nếu bằng nhau thì xét đến cột ưu tiên 2, rồi 3. Ô trống luôn nằm cuối. Mẹo: bấm vào tên cột để sắp xếp, giữ <kbd>Shift</kbd> khi bấm để thêm cột.</span></div>
      ${st.sort.map((s, k) => `<div class="frow"><span class="lv">${k + 1}</span><select data-sh="${k}">${colOpts(s.h, false)}</select>
        <select data-sd="${k}"><option value="1" ${s.dir > 0 ? 'selected' : ''}>Tăng dần (A→Z, cũ → mới, nhỏ → lớn)</option><option value="-1" ${s.dir < 0 ? 'selected' : ''}>Giảm dần (Z→A, mới → cũ, lớn → nhỏ)</option></select>
        <button type="button" class="iconbtn" data-xs="${k}" aria-label="Bỏ cột ưu tiên này">✕</button></div>`).join('')}
      <div class="pfoot">${st.sort.length < 3 ? '<button class="btn sm" type="button" id="addsort">＋ Thêm cột ưu tiên</button>' : ''}${st.sort.length ? '<button class="btn sm" type="button" id="clrsort">Trở về thứ tự gốc</button>' : ''}</div></div>` : ''}
    ${st.panel === 'cols' ? `<div class="fpanel"><div class="ph"><b>Các cột hiện trong bảng</b><span>Chỉ thay đổi cách xem, không ảnh hưởng tới file Excel.</span></div>
      <div class="colpick">${S.headers.map(h => `<label><input type="checkbox" data-col="${esc(h)}" ${cols.includes(h) ? 'checked' : ''}> ${esc(h)}</label>`).join('')}</div>
      <div class="pfoot"><button class="btn sm" type="button" id="defcols">Trở về mặc định</button></div></div>` : ''}
    ${chips.length ? `<div class="chips-active">${chips.join('')}<button class="linkbtn" type="button" id="clrall">Bỏ hết lọc và sắp xếp</button></div>` : ''}
    ${selected.length ? `<div class="bulk"><b>Đã chọn ${selected.length}</b>
      ${bulkOpts.map(([h, o]) => `<select data-bulk="${esc(h)}"><option value="">Gán ${esc(h)}…</option>${o.map(v => `<option>${esc(v)}</option>`).join('')}</select>`).join('')}
      <button class="btn sm danger" type="button" id="bdel">Xóa ${selected.length} dòng</button>
      <button class="btn sm" type="button" id="bnone">Bỏ chọn</button></div>` : ''}
    <div class="tablewrap">${rows.length ? `<table class="grid"><thead><tr>
      <th class="ck nos"><input type="checkbox" id="ckall" aria-label="Chọn tất cả" ${rows.every(r => st.sel.has(r._i)) ? 'checked' : ''}></th>
      <th data-s="#" title="Bấm để sắp xếp. Giữ Shift khi bấm để sắp xếp thêm theo cột này">${C.codeCol ? 'Mã' : '#'} ${arrow('#')}</th>
      ${cols.map((h, k) => `<th data-s="${esc(h)}" class="${k > 2 ? 'hide-sm' : ''}" title="Bấm để sắp xếp. Giữ Shift khi bấm để sắp xếp thêm theo cột này">${esc(h)} ${arrow(h)}</th>`).join('')}
    </tr></thead><tbody>${rows.map(r => {
      const a = iss[r._i], lv = a.some(x => x.level === 'err') ? 'err' : a.length ? 'warn' : '';
      return `<tr data-i="${r._i}" class="${st.sel.has(r._i) ? 'sel' : ''} ${r.Status === 'Taken' ? 'taken' : ''}">
        <td class="ck"><input type="checkbox" data-ck="${r._i}" ${st.sel.has(r._i) ? 'checked' : ''} aria-label="Chọn dòng"></td>
        <td class="rn">${C.codeCol ? `<span class="code">${esc(r._code)}</span>` : pos.get(r._i)}${lv ? `<span class="issue-dot ${lv}" title="${esc(a.map(x => x.field + ': ' + x.msg).join('\n'))}"></span>` : ''}</td>
        ${cols.map((h, k) => `<td class="${k > 2 ? 'hide-sm' : ''}">${cellHtml(sheet, h, r[h])}</td>`).join('')}</tr>`;
    }).join('')}</tbody></table>` : `<div class="empty">Không có dòng nào phù hợp. ${total ? '<button class="btn sm" type="button" id="clrall2">Bỏ hết bộ lọc</button>' : 'Bấm “Thêm mới” để bắt đầu.'}</div>`}</div>`;

  const v = $('#view');
  const redraw = () => { savePrefs(key); renderSheet(); };
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
        const k = el.dataset.cv, pos = el.selectionStart; st.conds[+k].v = el.value; savePrefs(key); renderSheet();
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
    const warn = sheet === 'Topics' ? '<br><br>Xóa đề tài sẽ làm <b>thay đổi mã</b> của các đề tài phía sau trong cùng hướng. Nên đánh dấu Taken thay vì xóa.' : '';
    if (!await confirmBox('Xóa các dòng đã chọn?', `Xóa <b>${selected.length}</b> dòng khỏi ${esc(sheet)}? Bạn có thể hoàn tác sau đó.${warn}`, 'Xóa', true)) return;
    if (await op({ sheet, op: 'delete', indexes: selected }, `Đã xóa ${selected.length} dòng`)) { st.sel.clear(); renderSheet(); }
  });
  v.querySelectorAll('[data-bulk]').forEach(s => s.onchange = async () => {
    if (!s.value) return;
    const h = s.dataset.bulk, val = s.value;
    if (!await confirmBox('Sửa các dòng đã chọn?', `Gán <b>${esc(h)}</b> thành <b>${esc(val)}</b> cho ${selected.length} dòng?`, 'Áp dụng')) { s.value = ''; return; }
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
  if (f.semester) {
    const sv = parseSemester(val), yr = sv.year || new Date().getFullYear();
    input = `<input type="hidden" id="${id}" data-h="${esc(h)}" value="${esc(val)}"><div class="semrow" data-sem="${id}">
      <select class="sems" aria-label="Học kỳ">${SEASONS.map(x => `<option ${x === (sv.season || 'Fall') ? 'selected' : ''}>${x}</option>`).join('')}</select>
      <input class="semy" type="number" min="2000" max="2100" step="1" value="${yr}" aria-label="Năm"></div>`;
  } else if (f.datasets) {
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
  const tpl = f.template ? `<div class="quick"><button type="button" data-tpl="${id}">Thêm các mục còn thiếu</button></div><div class="tpl" data-tplv="${id}"></div>` : '';
  const aux = f.url ? `<a class="aux" data-open="${id}" href="#" tabindex="-1">Mở link ↗</a>` : f.big ? `<span class="aux" data-wc="${id}"></span>` : '';
  return `<div class="f ${f.wide || long ? 'wide' : ''}" data-field="${esc(h)}"><label for="${id}">${esc(h)}${req ? '<span class="req">*</span>' : ''}${aux}</label>
    ${input}${quick}${tpl}${f.hint ? `<span class="hint">${esc(f.hint)}</span>` : ''}<span class="msg"></span></div>`;
}

// Semester = "Spring 2026" / "Summer 2026" / "Fall 2026"; older values like FA26 are read too
const SEASONS = ['Spring', 'Summer', 'Fall'];
function parseSemester(v) {
  v = String(v || '');
  const w = v.match(/spring|summer|fall|autumn|\b(sp|su|fa)(?=\s*-?\d{2})/i), y = v.match(/\d{4}/) || v.match(/\d{2}/);
  const map = { spring: 'Spring', sp: 'Spring', summer: 'Summer', su: 'Summer', fall: 'Fall', fa: 'Fall', autumn: 'Fall' };
  return { season: w ? map[w[0].toLowerCase()] : '', year: y ? (y[0].length === 2 ? 2000 + +y[0] : +y[0]) : 0 };
}

function dsRow(d) {
  return `<div class="dsrow"><input class="dsn" placeholder="Tên dataset (bắt buộc)" value="${esc(d.name)}" aria-label="Tên dataset">
    <input class="dsu" type="url" inputmode="url" placeholder="https://… (không bắt buộc)" value="${esc(d.url)}" aria-label="Link dataset">
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
    if (lock && sheet === view) r[lock.h] = lock.v;                 // a new row belongs to the part being viewed
  }
  // ↑ ↓ move a row among the rows of the same group (Home Items: same Section), which is the order the page shows
  const grp = C.orderWithin, sibs = rowsOf(sheet).filter(o => !grp || o[grp] === r[grp]).map(o => o._i), at = sibs.indexOf(index);
  editing.prev = at > 0 ? sibs[at - 1] : null; editing.next = at >= 0 && at < sibs.length - 1 ? sibs[at + 1] : null;
  const num = at >= 0 ? at + 1 : index + 1;
  const title = isNew ? 'Thêm mới' : (C.title(r) || 'Dòng ' + num);
  $('#dkicker').textContent = sheet + (isNew ? '' : ' · ' + (r._code ? r._code + ' · ' : '') + 'dòng ' + num);
  $('#dtitle').textContent = title;
  const topicNote = sheet === 'Topics' ? `<div class="note">Mã đề tài ${isNew ? 'sẽ là' : 'hiện là'} <b class="code" id="tcode"></b>. Mã được đánh theo vị trí của dòng trong hướng, vì vậy hãy thêm đề tài mới vào cuối; khi đề tài đã có người nhận thì chọn Status là <b>Taken</b> thay vì xóa.</div>` : '';
  $('#dbody').innerHTML = `${topicNote}<div class="sumbox" id="sum" hidden></div><div class="fgrid">${S.headers.map(h => fieldHtml(sheet, h, r[h] ?? '', isNew)).join('')}</div>`;
  const site = C.site && !isNew ? C.site(r) : null;
  $('#dfoot').innerHTML = `<button class="btn primary" type="button" id="dsave">${isNew ? 'Thêm vào Excel' : 'Lưu'} <kbd style="color:#dfe7ff;border-color:#7d97ff">Ctrl S</kbd></button>
    <button class="btn" type="button" id="dcancel">Hủy</button><span class="sp"></span>
    ${!isNew && C.ordered ? `<button class="btn sm" type="button" id="dup" title="Đưa lên trên" ${editing.prev == null ? 'disabled' : ''}>↑</button><button class="btn sm" type="button" id="ddown" title="Đưa xuống dưới" ${editing.next == null ? 'disabled' : ''}>↓</button>` : ''}
    ${site ? `<a class="btn sm" target="_blank" rel="noopener" href="/site/${esc(site)}">Xem trên web ↗</a>` : ''}
    ${!isNew && !C.fixed ? `<button class="btn sm" type="button" id="dcopy">Tạo bản sao</button>` : ''}
    ${!isNew && sheet === 'Topics' ? `<button class="btn sm" type="button" id="dtaken">${r.Status === 'Taken' ? 'Mở lại đề tài' : 'Đánh dấu Taken'}</button>` : ''}
    ${!isNew && !C.fixed ? `<button class="btn sm danger" type="button" id="ddel">Xóa</button>` : ''}`;
  wireForm(sheet);
  dirty = false;
  $('#drawer').classList.add('open'); $('#drawer').setAttribute('aria-hidden', 'false'); $('#scrim').hidden = false;
  history.replaceState(null, '', '#' + encodeURIComponent(sheet === view ? vkey() : sheet) + (isNew ? '/new' : '/' + (index + 1)));
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
  box.innerHTML = `<b>Bài preprint trên arXiv.</b> Dán arXiv ID hoặc link (ví dụ 2501.01234 hoặc https://arxiv.org/abs/2501.01234), các ô Venue, PDF và DOI sẽ tự điền.
    <div class="arxrow"><input id="arxid" placeholder="arXiv ID hoặc link" autocomplete="off"><button type="button" class="btn sm" id="arxget">Lấy thông tin từ arXiv</button></div><span class="hint" id="arxmsg"></span>`;
  body.querySelector('.fgrid').before(box);
  const idIn = box.querySelector('#arxid'), msg = box.querySelector('#arxmsg');
  const cur = () => { const m = [F('Venue').value, F('PDF')?.value, F('DOI')?.value].join(' ').match(/arxiv[^0-9]{0,12}(\d{4}\.\d{4,5}(?:v\d+)?)/i); return m ? m[1] : ''; };
  const apply = id => {
    F('Venue').value = 'arXiv preprint arXiv:' + id;
    const base = id.replace(/v\d+$/, '');
    if (!F('PDF').value.trim() || /arxiv\.org/i.test(F('PDF').value)) F('PDF').value = 'https://arxiv.org/pdf/' + id;
    if (!F('DOI').value.trim() || /arxiv/i.test(F('DOI').value)) F('DOI').value = '10.48550/arXiv.' + base;
    msg.textContent = 'Đã điền Venue, PDF và DOI theo arXiv:' + id + '.'; changed();
  };
  const sync = () => { const on = typeEl.value === 'Preprint'; box.hidden = !on; F('Venue').readOnly = on && !!cur(); if (on && !idIn.value) idIn.value = cur(); };
  typeEl.addEventListener('change', () => {
    if (typeEl.value === 'Preprint') { const st = F('Status'); if (st && st.value !== 'In Submission') st.value = 'Preprint'; if (cur()) apply(cur()); }
    sync();
  });
  idIn.addEventListener('input', () => { const m = idIn.value.match(ARXIV_RE); if (m) { apply(m[1]); sync(); } });
  box.querySelector('#arxget').onclick = async () => {
    const m = (idIn.value || cur()).match(ARXIV_RE); if (!m) { msg.textContent = 'Hãy nhập arXiv ID trước.'; return; }
    msg.textContent = 'Đang lấy thông tin từ arXiv…';
    try {
      const r = await api('/api/arxiv?id=' + encodeURIComponent(m[1]));
      apply(m[1]);
      const filled = [['Title', r.title], ['Authors', r.authors], ['Date', r.date]].filter(([h, v]) => v && set(h, v, true)).map(([h]) => h);
      msg.textContent = filled.length ? `Đã điền ${filled.join(', ')} từ arXiv (các ô đã có nội dung được giữ nguyên).` : 'Title, Authors và Date đã có nội dung nên được giữ nguyên.';
      changed();
    } catch (e) { msg.textContent = e.message; }
  };
  sync();
}

function profileBox(sheet) {
  if (!cfg(sheet).fixed || !editing || editing.isNew) return;
  const body = $('#dbody'), field = body.querySelector('[data-h="Field"]'), value = body.querySelector('[data-h="Value"]');
  const pageEl = body.querySelector('[data-h="Page"]'), page = pageEl ? pageEl.value : '';
  field.readOnly = true; if (pageEl) pageEl.disabled = true;
  const box = value.closest('.f'), hint = page === 'Home' ? PROFILE_HINT[field.value] : PAGE_HINT[field.value] && PAGE_HINT[field.value] + MARKUP_HINT;
  if (hint) box.querySelector('.hint')?.remove(), box.querySelector('.msg').insertAdjacentHTML('beforebegin', `<span class="hint">${esc(hint)}</span>`);
  if (page !== 'Home' || field.value !== 'Avatar') return;
  // Avatar: pick a picture that is inside the repository, or upload one (copied into assets/img)
  const input = document.createElement('input'); input.type = 'hidden'; input.dataset.h = 'Value'; input.id = value.id; input.value = value.value;
  value.replaceWith(input);
  const pick = document.createElement('div'); pick.className = 'imgpick';
  box.querySelector('.hint').before(pick);
  const draw = () => {
    pick.innerHTML = `<div class="imgcur"><img src="/site/${esc(input.value)}?t=${Date.now()}" alt="" onerror="this.style.visibility='hidden'"><div><b>Ảnh đang dùng:</b> <code>${esc(input.value || '(chưa chọn)')}</code>
        <div class="imgbtns"><label class="btn sm">Tải ảnh từ máy lên…<input type="file" accept="image/png,image/jpeg,image/webp,image/gif" hidden></label><button type="button" class="btn sm" data-reload>Tải lại danh sách ảnh</button></div></div></div>
      <div class="imggrid">${IMAGES.filter(i => !/favicon/i.test(i.path)).map(i => `<button type="button" class="${i.path === input.value ? 'on' : ''}" data-img="${esc(i.path)}" title="${esc(i.path)}"><img src="/site/${esc(i.path)}" alt="" loading="lazy"><span>${esc(i.path.split('/').pop())}</span></button>`).join('') || '<p class="hint">Chưa có ảnh nào trong thư mục website.</p>'}</div>`;
    pick.querySelectorAll('[data-img]').forEach(b => b.onclick = () => { input.value = b.dataset.img; changed(); draw(); });
    pick.querySelector('[data-reload]').onclick = async () => { await loadImages(); draw(); };
    pick.querySelector('input[type=file]').onchange = async e => {
      const f = e.target.files[0]; if (!f) return;
      const data = await new Promise(res => { const rd = new FileReader(); rd.onload = () => res(rd.result); rd.readAsDataURL(f); });
      try {
        const r = await api('/api/upload-image', { name: f.name, data });
        await loadImages(); input.value = r.path; changed(); draw();
        toast('Đã chép ảnh vào ' + r.path + '. Bấm Lưu để dùng ảnh này làm avatar.', 'ok');
      } catch (err) { toast(err.message, 'err'); }
    };
  };
  draw();
}

// Home Items: show only the fields the chosen Section uses, with a hint for each
function itemsBox(sheet) {
  if (sheet !== 'Home Items') return;
  const body = $('#dbody'), sec = body.querySelector('[data-h="Section"]'); if (!sec) return;
  const sync = () => {
    const used = ITEM_FIELDS[sec.value] || {};
    for (const h of ['Title', 'Organization', 'Dates', 'Description']) {
      const f = body.querySelector(`.f[data-field="${h}"]`); if (!f) continue;
      const el = f.querySelector('[data-h]');
      f.hidden = !used[h] && !(el && el.value.trim());
      let hint = f.querySelector('.hint.sec'); if (!hint) { hint = document.createElement('span'); hint.className = 'hint sec'; f.querySelector('.msg').before(hint); }
      hint.textContent = used[h] || '';
    }
  };
  sec.addEventListener('change', sync); sync();
}

function semesterBox() {
  $('#dbody').querySelectorAll('[data-sem]').forEach(box => {
    const inp = $('#' + box.dataset.sem), s = box.querySelector('.sems'), y = box.querySelector('.semy');
    const sync = () => { inp.value = y.value ? s.value + ' ' + y.value : ''; };
    box.addEventListener('input', sync); box.addEventListener('change', () => { sync(); changed(); });
    sync();
  });
}
function projectsBox(sheet) {
  if (sheet !== 'Projects') return;
  const body = $('#dbody'), ps = body.querySelector('[data-h="Paper Status"]'); if (!ps) return;
  const sync = () => {
    for (const [h, show] of [['Paper Venue', ps.value === 'Submitted'], ['Link ID', ps.value === 'Accepted']]) {
      const f = body.querySelector(`.f[data-field="${h}"]`), el = f && f.querySelector('[data-h]');
      if (f) f.hidden = !show && !(el && el.value.trim());
    }
  };
  ps.addEventListener('change', sync); sync();
}

function wireForm(sheet) {
  const body = $('#dbody');
  semesterBox();
  projectsBox(sheet);
  arxivBox(sheet);
  profileBox(sheet);
  itemsBox(sheet);
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
    if (!miss.length) { toast('Đã có đủ các mục.'); return; }
    el.value = (el.value.trim() ? el.value.replace(/\s*$/, '\n') : '') + miss.map(l => l + ': ').join('\n');
    el.focus(); changed();
  });
  body.querySelectorAll('[data-open]').forEach(a => a.onclick = e => { e.preventDefault(); const v = $('#' + a.dataset.open).value; if (isUrl(v)) window.open(v, '_blank', 'noopener'); });
  body.addEventListener('input', changed);
  body.addEventListener('change', changed);
  $('#dsave').onclick = saveEditor;
  $('#dcancel').onclick = closeEditor;
  const { index } = editing;
  $('#dup') && ($('#dup').onclick = () => editing.prev != null && moveRow(index, editing.prev));
  $('#ddown') && ($('#ddown').onclick = () => editing.next != null && moveRow(index, editing.next));
  $('#dcopy') && ($('#dcopy').onclick = async () => {
    if (dirty && !await confirmBox('Bỏ các thay đổi?', 'Những thay đổi chưa lưu sẽ bị mất.', 'Tiếp tục')) return;
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
  sum.innerHTML = issues.length ? `<b>${errs.length ? 'Có ' + errs.length + ' lỗi cần sửa' : 'Không có lỗi'}${issues.length - errs.length ? ` · ${issues.length - errs.length} cảnh báo` : ''}</b>` +
    issues.map(x => `<div class="${x.level}">• ${esc(x.field)}: ${esc(x.msg)}</div>`).join('') : '';
  return issues;
}

async function saveEditor() {
  if (!editing || busy) return;
  const { sheet, index, isNew, before } = editing;
  const S = DB.sheets[sheet];
  const issues = validateForm(), errs = issues.filter(x => x.level === 'err');
  if (errs.length) {
    const go = await dialog('Dữ liệu vẫn còn lỗi', `${errs.map(x => `• <b>${esc(x.field)}</b>: ${esc(x.msg)}`).join('<br>')}<br><br>Nếu vẫn lưu, web có thể hiển thị sai hoặc không hiện dòng này.`,
      [{ t: 'Quay lại sửa', v: false, primary: 1 }, { t: 'Vẫn lưu', v: true }]);
    if (!go) return;
  }
  const v = formValues();
  const values = S.headers.map(h => v[h] ?? '');
  // Directions: a changed code is renamed everywhere first
  if (sheet === 'Directions' && !isNew && before[0] && v.Code && v.Code.trim().toUpperCase() !== before[0]) {
    const used = rowsOf('Topics').filter(t => t.Direction === before[0]).length;
    if (!await confirmBox('Đổi mã hướng?', `Đổi mã <b>${esc(before[0])}</b> thành <b>${esc(v.Code.trim().toUpperCase())}</b> ở tất cả các sheet (${used} đề tài, cùng các hội nghị và tạp chí đang dùng mã này)? Mã đề tài trên web cũng sẽ đổi theo, ví dụ ${esc(before[0])}-01 thành ${esc(v.Code.trim().toUpperCase())}-01.`, 'Đổi mã')) return;
    const res = await op({ sheet, op: 'rename_code', old: before[0], new: v.Code }, 'Đã đổi mã hướng ở tất cả các sheet');
    if (!res) return;
    values[0] = v.Code.trim().toUpperCase();
    const r2 = await op({ sheet, op: 'update', index, values, before: DB.sheets[sheet].rows[index] }, 'Đã lưu');
    if (r2) { dirty = false; closeEditor(true); flash(index); }
    return;
  }
  const res = isNew ? await op({ sheet, op: 'insert', values }, 'Đã thêm')
    : await op({ sheet, op: 'update', index, values, before }, 'Đã lưu');
  if (res) { dirty = false; closeEditor(true); if (view !== sheet) go(sheet); flash(res.index); }
}

async function moveRow(i, j) {
  if (dirty && !await confirmBox('Bỏ các thay đổi?', 'Đổi vị trí dòng sẽ làm mất những thay đổi chưa lưu.', 'Tiếp tục')) return;
  const sheet = editing.sheet;
  const res = await op({ sheet, op: 'move', index: i, to: j }, 'Đã đổi thứ tự');
  if (res) { dirty = false; openEditor(sheet, j); flash(j); }
}

async function deleteRow() {
  const { sheet, index } = editing;
  const r = rowsOf(sheet)[index], C = cfg(sheet);
  if (sheet === 'Topics') {
    const ans = await dialog('Xóa đề tài?', `Xóa đề tài <b>${esc(r._code)}</b> sẽ làm thay đổi mã của các đề tài ${esc(r.Direction)} phía sau, nên các link cũ như research.html#${esc(r._code)} sẽ trỏ sai.<br><br>Nếu đề tài đã có người nhận, nên <b>đánh dấu Taken</b> thay vì xóa.`,
      [{ t: 'Hủy', v: 0 }, { t: 'Đánh dấu Taken', v: 1, primary: 1 }, { t: 'Vẫn xóa', v: 2, danger: 1 }]);
    if (ans === 1) { $('#dbody').querySelector('[data-h="Status"]').value = 'Taken'; changed(); return saveEditor(); }
    if (ans !== 2) return;
  } else if (sheet === 'Directions' && rowsOf('Topics').some(t => t.Direction === r.Code)) {
    const n = rowsOf('Topics').filter(t => t.Direction === r.Code).length;
    if (!await confirmBox('Xóa hướng đang có đề tài?', `Hướng <b>${esc(r.Code)}</b> vẫn còn ${n} đề tài. Nếu xóa, các đề tài này sẽ bị tách thành một nhóm lạ trên web. Bạn vẫn muốn xóa?`, 'Vẫn xóa', true)) return;
  } else if (!await confirmBox('Xóa dòng này?', `Xóa <b>${esc(C.title(r) || 'dòng ' + (index + 1))}</b> khỏi ${esc(sheet)}? Bạn có thể khôi phục bằng nút “Hoàn tác”.`, 'Xóa', true)) return;
  if (await op({ sheet, op: 'delete', index }, 'Đã xóa dòng')) { dirty = false; closeEditor(true); }
}

async function closeEditor(force) {
  if (!editing) return;
  if (force !== true && dirty && !await confirmBox('Đóng mà không lưu?', 'Bạn có thay đổi chưa lưu. Đóng khung chỉnh sửa sẽ bỏ các thay đổi đó.', 'Đóng, không lưu', true)) return;
  editing = null; dirty = false;
  $('#drawer').classList.remove('open'); $('#drawer').setAttribute('aria-hidden', 'true'); $('#scrim').hidden = true;
  history.replaceState(null, '', '#' + encodeURIComponent(vkey()));
}
function flash(i) {
  requestAnimationFrame(() => { const tr = $('#view').querySelector(`tr[data-i="${i}"]`); if (tr) { tr.classList.add('flash'); tr.scrollIntoView({ block: 'nearest' }); } });
}

/* ------------------------------------------------------------ routing + events */
function setView(name, lv) {
  const n = lv != null && NAV.find(x => x.s === name && x.lock && x.lock[1] === lv);
  view = name; lock = n ? { h: n.lock[0], v: n.lock[1] } : null;
}
function go(name, lv) { setView(name, lv); $('#side').classList.remove('open'); renderNav(); render(); if (location.hash.split('/')[0] !== '#' + encodeURIComponent(vkey())) history.replaceState(null, '', '#' + encodeURIComponent(vkey())); $('#view').scrollTop = 0; }
function route() {
  const [k, row] = decodeURIComponent(location.hash.slice(1)).split('/');
  const [s, lv] = (k || '').split('@');
  if (editing) { closeEditor(true); }
  setView(DB.sheets[s] ? s : 'overview', lv); renderNav(); render();
  if (DB.sheets[s] && row) { if (row === 'new') openEditor(s, null); else if (+row >= 1 && rowsOf(s)[+row - 1]) openEditor(s, +row - 1); }
}
addEventListener('hashchange', async () => {
  if (editing && dirty && !await confirmBox('Rời đi mà không lưu?', 'Bạn có thay đổi chưa lưu. Chuyển sang mục khác sẽ bỏ các thay đổi đó.', 'Rời đi', true)) {
    history.replaceState(null, '', '#' + encodeURIComponent(editing.sheet === view ? vkey() : editing.sheet) + (editing.isNew ? '/new' : '/' + (editing.index + 1))); return;
  }
  dirty = false; route();
});
$('#dclose').onclick = closeEditor; $('#scrim').onclick = closeEditor;
$('#menu').onclick = () => $('#side').classList.toggle('open');
$('#undo').onclick = async () => {
  if (!DB.backups) { toast('Chưa có gì để hoàn tác.'); return; }
  if (!await confirmBox('Hoàn tác lần lưu gần nhất?', 'File database.xlsx sẽ trở về trạng thái <b>trước lần lưu gần nhất</b>. Bạn có thể bấm nhiều lần để quay lại xa hơn.', 'Hoàn tác')) return;
  try { const r = await api('/api/undo', {}); closeEditor(true); await reload(); toast('Đã hoàn tác lần lưu gần nhất', 'ok'); }
  catch (e) { e.status === 423 ? dialog('Chưa hoàn tác được', esc(e.message), [{ t: 'Đóng', primary: 1 }]) : toast(e.message, 'err'); }
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
    if (v.locked) setSync('warn', 'Excel đang mở file, hãy đóng Excel trước khi lưu');
    else if (!busy) setSync('ok', 'Đã kết nối với database.xlsx');
    if (v.version !== lastVersion) {
      if (editing && dirty) { lastVersion = v.version; banner('File database.xlsx vừa được sửa ở nơi khác (có thể bằng Excel). Khi bạn bấm Lưu, công cụ sẽ kiểm tra; nếu dòng này đã bị thay đổi, dữ liệu sẽ được tải lại để bạn xem trước.'); }
      else { const ed = editing; await reload(true); if (ed && !ed.isNew) openEditor(ed.sheet, ed.index); toast('Đã cập nhật dữ liệu mới từ file Excel'); }
    }
  } catch { setSync('err', 'Mất kết nối – có thể cửa sổ màu đen đã bị tắt'); }
}, 3000);

(async () => {
  try {
    await loadImages();
    await reload();
    route();
    try { PRED = await api('/api/predatory'); if (PRED.journals.length) { ISSUES = null; renderNav(); if (view === 'overview' || view === 'Journals') render(); } } catch { }
  } catch (e) {
    $('#view').innerHTML = `<div class="panel"><h3>Không kết nối được với công cụ</h3><p>Hãy mở công cụ bằng file <b>Open database manager.bat</b> (hoặc chạy lệnh <code>python admin/server.py</code>), rồi tải lại trang này.</p><p>${esc(e.message)}</p></div>`;
    setSync('err', 'Không kết nối được với công cụ');
  }
})();
})();
