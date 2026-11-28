"""Database manager for the academic website (local only, never published).

Run:  python admin/server.py          (or double-click "Open database manager.bat")
Then: http://localhost:8765/          (opens automatically)

Every change made in the page is written straight into data/database.xlsx with openpyxl, keeping the
workbook's dropdowns, formatting, hyperlinks and the guide sheet. Before each write the previous file is
copied to admin/backups/ (the newest 60 are kept), so "Undo" can put it back.
"""
import argparse, datetime as dt, glob, json, os, re, shutil, sys, tempfile, threading, unicodedata, webbrowser
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, unquote

try:
    import openpyxl
    from openpyxl.worksheet.datavalidation import DataValidation
    from copy import copy
except ImportError:
    print('openpyxl is missing. Install it with:  python -m pip install openpyxl')
    input('Press Enter to close'); sys.exit(1)

HERE = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.dirname(HERE)
ap = argparse.ArgumentParser()
ap.add_argument('--db', default=os.path.join(SITE, 'data', 'database.xlsx'))
ap.add_argument('--port', type=int, default=8765)
ap.add_argument('--no-browser', action='store_true')
ap.add_argument('--predatory', default=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'predatory'),
                help='folder with the predatory journal/publisher lists (default: admin/predatory in the repo)')
ARGS = ap.parse_args()
DB = os.path.abspath(ARGS.db)
BACKUPS = os.path.join(HERE, 'backups')
KEEP_BACKUPS = 60
GUIDE = 'Hướng dẫn'
SHEET_ORDER = ['Profile', 'Education', 'Experience', 'Courses', 'Honors', 'Service',
               'Publications', 'Projects', 'Directions', 'Topics', 'Conferences', 'Journals', 'Venue Notes']
LOCK = threading.Lock()


class UserError(Exception):
    def __init__(self, msg, status=400):
        super().__init__(msg); self.status = status


def version():
    st = os.stat(DB)
    return f'{st.st_mtime_ns}-{st.st_size}'


def show(v):
    """Cell value -> the text shown in the manager."""
    if v is None: return ''
    if isinstance(v, dt.datetime):
        return v.strftime('%d-%m-%Y') if (v.hour, v.minute) == (0, 0) else v.strftime('%d-%m-%Y %H:%M')
    if isinstance(v, dt.date): return v.strftime('%d-%m-%Y')
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v)


def headers(ws):
    hs = []
    for c in ws[1]:
        if c.value is None or str(c.value).strip() == '': break
        hs.append(str(c.value).strip())
    return hs


def data_rows(ws, n):
    """All record rows (as cell-value lists), stopping at the last non-empty row."""
    rows = [list(r[:n]) for r in ws.iter_rows(min_row=2, max_col=n, values_only=True)]
    while rows and all(v is None or str(v).strip() == '' for v in rows[-1]): rows.pop()
    return rows


def list_options(ws, hs):
    """Dropdown lists from the sheet's data validations: {header: [options]}."""
    out = {}
    for dv in ws.data_validations.dataValidation:
        if dv.type != 'list' or not dv.formula1: continue
        f = dv.formula1.strip()
        for rng in str(dv.sqref).split():
            col = re.match(r'([A-Z]+)', rng).group(1)
            idx = openpyxl.utils.column_index_from_string(col) - 1
            if idx >= len(hs): continue
            if f.startswith('"'):
                out[hs[idx]] = [x.strip() for x in f.strip('"').split(',')]
            else:
                out[hs[idx]] = {'ref': f.lstrip('=')}
    return out


def link_columns(ws, n, rows):
    cols = set()
    for r in range(2, len(rows) + 2):
        for c in range(1, n + 1):
            if ws.cell(r, c).hyperlink: cols.add(c - 1)
    return cols


def load(path=None):
    return openpyxl.load_workbook(path or DB)


def snapshot():
    wb = load()
    out = {'version': version(), 'file': DB, 'sheets': {}, 'order': []}
    for name in [s for s in SHEET_ORDER if s in wb.sheetnames] + [s for s in wb.sheetnames if s not in SHEET_ORDER and s != GUIDE]:
        ws = wb[name]; hs = headers(ws); rows = data_rows(ws, len(hs))
        out['order'].append(name)
        out['sheets'][name] = {'headers': hs, 'rows': [[show(v) for v in r] for r in rows], 'options': list_options(ws, hs)}
    # resolve dropdowns that point at another sheet's column (Topics Direction -> Directions!A)
    for s in out['sheets'].values():
        for h, o in list(s['options'].items()):
            if isinstance(o, dict):
                m = re.match(r"'?([^'!]+)'?!\$?([A-Z]+)", o['ref'])
                src = out['sheets'].get(m.group(1)) if m else None
                if src:
                    i = openpyxl.utils.column_index_from_string(m.group(2)) - 1
                    s['options'][h] = [r[i] for r in src['rows'] if i < len(r) and r[i]]
                else:
                    del s['options'][h]
    out['backups'] = len(backup_list())
    return out


# ---------------------------------------------------------------- writing
def backup_list():
    return sorted(glob.glob(os.path.join(BACKUPS, 'database-*.xlsx')))


def make_backup(label):
    os.makedirs(BACKUPS, exist_ok=True)
    stamp = dt.datetime.now().strftime('%Y%m%d-%H%M%S-%f')[:-3]
    safe = re.sub(r'[^A-Za-z0-9_-]+', '_', label)[:40]
    shutil.copy2(DB, os.path.join(BACKUPS, f'database-{stamp}-{safe}.xlsx'))
    for old in backup_list()[:-KEEP_BACKUPS]:
        try: os.remove(old)
        except OSError: pass


def excel_locked():
    """Excel keeps an open workbook locked on Windows, so opening it for writing fails."""
    try:
        with open(DB, 'r+b'): return False
    except PermissionError:
        return True


EXCEL_MSG = 'File database.xlsx đang mở trong Excel nên không ghi được. Hãy đóng Excel (lưu hoặc bỏ thay đổi trong Excel) rồi thử lại.'


def restore_topics_dropdown(wb):
    """Excel stores the Topics Direction list (taken from another sheet) where openpyxl cannot read it; put it back."""
    if 'Topics' not in wb.sheetnames: return
    tw = wb['Topics']
    if any('A2' in str(dv.sqref) for dv in tw.data_validations.dataValidation): return
    msg = 'Direction code from the Directions sheet (CV, MI, XAI, EDGE, VLM, ...)'
    dv = DataValidation(type='list', formula1='=Directions!$A$2:$A$50', allow_blank=True, showErrorMessage=True,
                        showInputMessage=True, errorTitle='Invalid value', error=msg, promptTitle='Choose', prompt=msg)
    dv.add('A2:A2000'); tw.add_data_validation(dv)


def save(wb, label):
    if excel_locked(): raise UserError(EXCEL_MSG, 423)
    restore_topics_dropdown(wb)
    make_backup(label)
    fd, tmp = tempfile.mkstemp(suffix='.xlsx', dir=os.path.dirname(DB)); os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, DB)
    except PermissionError:
        raise UserError(EXCEL_MSG, 423)
    finally:
        if os.path.exists(tmp): os.remove(tmp)


def to_cell(text, old):
    """Text from the form -> cell value. Unchanged text keeps the original value (dates, numbers)."""
    if old is not None and show(old) == text: return old
    text = text.replace('\r\n', '\n').strip()
    return text if text != '' else None


def write_sheet(ws, n, rows, old_count):
    """Rewrite all record rows in place, copying style from the row above for new rows."""
    tpl = 2
    links = link_columns(ws, n, [None] * old_count)
    for i, r in enumerate(rows):
        rr = i + 2
        if rr > old_count + 1:                                   # a brand-new row: copy look from row 2
            for c in range(1, n + 1):
                src, dst = ws.cell(tpl, c), ws.cell(rr, c)
                if src.has_style:
                    dst.font, dst.alignment, dst.border, dst.fill = copy(src.font), copy(src.alignment), copy(src.border), copy(src.fill)
                    dst.number_format = src.number_format
            h = ws.row_dimensions[tpl].height
            if h: ws.row_dimensions[rr].height = h
        for c in range(n):
            v = r[c] if c < len(r) else None
            cell = ws.cell(rr, c + 1)
            cell.value = v
            if c in links:
                cell.hyperlink = str(v) if isinstance(v, str) and re.match(r'https?://', v) else None
            elif cell.hyperlink:
                cell.hyperlink = None
    for rr in range(len(rows) + 2, old_count + 2):              # rows left over after a delete
        for c in range(1, n + 1):
            cell = ws.cell(rr, c); cell.value = None
            if cell.hyperlink: cell.hyperlink = None
    if ws.auto_filter.ref:
        ws.auto_filter.ref = f'A1:{openpyxl.utils.get_column_letter(n)}{max(len(rows) + 1, 2)}'


def apply_op(op):
    name = op.get('sheet')
    with LOCK:
        if op.get('version') != version():
            raise UserError('File Excel vừa được thay đổi ở nơi khác (có thể bạn đã sửa trong Excel). Dữ liệu đã được tải lại, hãy kiểm tra rồi thao tác lại.', 409)
        wb = load()
        if name not in wb.sheetnames or name == GUIDE: raise UserError('Không có sheet ' + str(name))
        ws = wb[name]; hs = headers(ws); n = len(hs)
        rows = data_rows(ws, n); old_count = len(rows)
        kind = op.get('op'); i = op.get('index')

        def check_index(k):
            if not isinstance(k, int) or not 0 <= k < len(rows): raise UserError('Dòng không tồn tại (dữ liệu đã đổi?). Hãy tải lại.', 409)

        def record(vals, old=None):
            vals = list(vals) + [''] * (n - len(vals))
            return [to_cell(str(vals[c] if vals[c] is not None else ''), old[c] if old else None) for c in range(n)]

        if kind == 'update':
            check_index(i)
            if op.get('before') is not None and [show(v) for v in rows[i]] != list(op['before']):
                raise UserError('Dòng này vừa bị thay đổi ở nơi khác. Dữ liệu đã được tải lại.', 409)
            rows[i] = record(op['values'], rows[i]); label = f'edit {name} row {i + 2}'
        elif kind == 'insert':
            at = op.get('at', len(rows))
            if not isinstance(at, int) or not 0 <= at <= len(rows): at = len(rows)
            rows.insert(at, record(op['values'])); label = f'add {name} row {at + 2}'; i = at
        elif kind == 'delete':
            idx = sorted(set(op.get('indexes') or [i]), reverse=True)
            for k in idx: check_index(k)
            for k in idx: rows.pop(k)
            label = f'delete {len(idx)} from {name}'
        elif kind == 'move':
            check_index(i); j = op.get('to')
            if not isinstance(j, int) or not 0 <= j < len(rows): raise UserError('Vị trí mới không hợp lệ.')
            rows.insert(j, rows.pop(i)); label = f'move {name} row {i + 2} to {j + 2}'; i = j
        elif kind == 'set':                                      # same value in one column for many rows
            col = hs.index(op['column']) if op.get('column') in hs else None
            if col is None: raise UserError('Không có cột ' + str(op.get('column')))
            for k in op.get('indexes') or []:
                check_index(k); rows[k][col] = to_cell(str(op.get('value', '')), rows[k][col])
            label = f'set {op["column"]} in {name}'
        elif kind == 'rename_code':                              # Directions: change a code everywhere
            old, new = op['old'].strip(), op['new'].strip().upper()
            if not re.fullmatch(r'[A-Z0-9]+', new): raise UserError('Mã mới chỉ gồm chữ in hoa và số, không dấu, không khoảng trắng.')
            ci = hs.index('Code'); codes = [show(r[ci]) for r in rows]
            if new in codes: raise UserError(f'Mã {new} đã tồn tại.')
            if old not in codes: raise UserError(f'Không có mã {old}.')
            rows[codes.index(old)][ci] = new
            for other, col in (('Topics', 'Direction'), ('Conferences', 'Directions'), ('Journals', 'Directions')):
                if other not in wb.sheetnames: continue
                ows = wb[other]; oh = headers(ows)
                if col not in oh: continue
                c = oh.index(col) + 1
                for rr in range(2, ows.max_row + 1):
                    v = ows.cell(rr, c).value
                    if isinstance(v, str) and old in [x.strip() for x in v.split(',')]:
                        ows.cell(rr, c).value = ', '.join(new if x.strip() == old else x.strip() for x in v.split(','))
            label = f'rename code {old} to {new}'
        else:
            raise UserError('Thao tác không hợp lệ: ' + str(kind))
        write_sheet(ws, n, rows, old_count)
        save(wb, label)
        return {'ok': True, 'index': i, 'message': label}


def undo():
    with LOCK:
        b = backup_list()
        if not b: raise UserError('Không còn bản sao lưu nào để hoàn tác.')
        if excel_locked(): raise UserError(EXCEL_MSG, 423)
        try:
            shutil.copy2(b[-1], DB)
        except PermissionError:
            raise UserError(EXCEL_MSG, 423)
        label = os.path.basename(b[-1])
        os.remove(b[-1])
        os.utime(DB, None)                                      # new version, so open pages reload
        return {'ok': True, 'message': 'Đã hoàn tác: ' + label}


# ---------------------------------------------------------------- predatory lists
def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode().lower()
    s = s.replace('&', ' and ')
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9]+', ' ', re.sub(r'\(.*?\)', ' ', s))).strip()


def domain(u):
    try:
        h = urlparse(str(u if '://' in str(u) else 'http://' + str(u))).hostname or ''
        return h[4:] if h.startswith('www.') else h
    except Exception:
        return ''


def predatory_lists():
    out = {'journals': [], 'publishers': [], 'domains': [], 'source': ''}
    folder = ARGS.predatory
    if not os.path.isdir(folder): return out
    for f in glob.glob(os.path.join(folder, '*.xlsx')):                 # predatory-journals.xlsx / predatory-publishers.xlsx
        if os.path.basename(f).startswith('~$'): continue
        key = 'publishers' if ('publisher' in f.lower() or 'nhà xuất bản' in f.lower()) else 'journals'
        try:
            wb = openpyxl.load_workbook(f, read_only=True)
            for r in wb.worksheets[0].iter_rows(min_row=2, values_only=True):
                if len(r) > 1 and r[1]:
                    out[key].append(norm(r[1]))
                    d = domain(r[2]) if len(r) > 2 and r[2] else ''
                    if d and d not in ('hindawi.com', 'mdpi.com', 'frontiersin.org', 'scirp.org', 'sciencedirect.com', 'springer.com', 'tandfonline.com'):
                        out['domains'].append(d)
            out['source'] += os.path.basename(f) + '; '
        except Exception as e:
            print('Cannot read', f, e)
    out['journals'] = sorted(set(out['journals'])); out['publishers'] = sorted(set(out['publishers'])); out['domains'] = sorted(set(out['domains']))
    return out


PRED = None


# ---------------------------------------------------------------- pictures inside the repository (avatar)
IMG_EXT = ('.png', '.jpg', '.jpeg', '.webp', '.gif', '.svg')
SKIP_DIRS = {'.git', 'admin', '.claude', 'node_modules', 'tools', 'data'}


def list_images():
    out = []
    for dirpath, dirnames, files in os.walk(SITE):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS and not d.startswith('.')]
        for f in files:
            if f.lower().endswith(IMG_EXT):
                full = os.path.join(dirpath, f)
                out.append({'path': os.path.relpath(full, SITE).replace(os.sep, '/'), 'size': os.path.getsize(full)})
    return sorted(out, key=lambda x: x['path'])


def upload_image(body):
    import base64
    name = os.path.basename(str(body.get('name') or 'image.png'))
    stem, ext = os.path.splitext(name)
    ext = ext.lower()
    if ext not in IMG_EXT or ext == '.svg': raise UserError('Chỉ nhận ảnh PNG, JPG, WEBP hoặc GIF.')
    data = str(body.get('data') or '')
    if ',' in data: data = data.split(',', 1)[1]                  # data:image/png;base64,....
    raw = base64.b64decode(data)
    if len(raw) > 8 * 1024 * 1024: raise UserError('Ảnh lớn hơn 8 MB. Hãy thu nhỏ ảnh (khoảng 800×800 px là đủ cho avatar).')
    stem = re.sub(r'[^A-Za-z0-9_-]+', '-', unicodedata.normalize('NFKD', stem).encode('ascii', 'ignore').decode()).strip('-').lower() or 'image'
    folder = os.path.join(SITE, 'assets', 'img'); os.makedirs(folder, exist_ok=True)
    target, n = os.path.join(folder, stem + ext), 2
    while os.path.exists(target):                                  # never overwrite a picture already in the repo
        target = os.path.join(folder, f'{stem}-{n}{ext}'); n += 1
    with open(target, 'wb') as f: f.write(raw)
    return {'ok': True, 'path': os.path.relpath(target, SITE).replace(os.sep, '/')}


# ---------------------------------------------------------------- arXiv lookup (for Preprint rows)
def arxiv_lookup(aid):
    import urllib.request, xml.etree.ElementTree as ET
    if not re.fullmatch(r'\d{4}\.\d{4,5}(v\d+)?', aid or ''): raise UserError('arXiv ID không hợp lệ (ví dụ 2501.01234).')
    try:
        with urllib.request.urlopen('https://export.arxiv.org/api/query?id_list=' + aid, timeout=15) as r: xml = r.read()
    except Exception as e:
        raise UserError('Không kết nối được arXiv: ' + str(e), 502)
    ns = {'a': 'http://www.w3.org/2005/Atom'}
    e = ET.fromstring(xml).find('a:entry', ns)
    if e is None or e.find('a:title', ns) is None or not (e.find('a:title', ns).text or '').strip() or (e.find('a:id', ns) is not None and 'api/errors' in (e.find('a:id', ns).text or '')):
        raise UserError('arXiv không có bài ' + aid + '.', 404)
    clean = lambda t: re.sub(r'\s+', ' ', t or '').strip()
    pub = (e.find('a:published', ns).text or '')[:10]                       # yyyy-mm-dd
    return {'id': aid, 'title': clean(e.find('a:title', ns).text),
            'authors': ', '.join(clean(a.find('a:name', ns).text) for a in e.findall('a:author', ns)),
            'date': '-'.join(reversed(pub.split('-'))) if pub else ''}


# ---------------------------------------------------------------- HTTP
TYPES = {'.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
         '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.webp': 'image/webp',
         '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.json': 'application/json', '.ico': 'image/x-icon',
         '.pdf': 'application/pdf', '.woff2': 'font/woff2'}


class H(BaseHTTPRequestHandler):
    def log_message(self, fmt, *a):
        if 'api' in (a[0] if a else ''): sys.stderr.write('%s\n' % (a[0],))

    def send(self, code, body, ctype='application/json; charset=utf-8'):
        if not isinstance(body, bytes): body = json.dumps(body, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype); self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers(); self.wfile.write(body)

    def file(self, root, rel):
        p = os.path.normpath(os.path.join(root, rel.lstrip('/')))
        if not (p + os.sep).startswith(root + os.sep) or not os.path.isfile(p): return self.send(404, b'Not found', 'text/plain')
        if os.sep + 'admin' + os.sep + 'backups' in p: return self.send(404, b'Not found', 'text/plain')
        with open(p, 'rb') as f: self.send(200, f.read(), TYPES.get(os.path.splitext(p)[1].lower(), 'application/octet-stream'))

    def do_GET(self):
        path = unquote(urlparse(self.path).path)
        try:
            if path == '/api/db': return self.send(200, snapshot())
            if path == '/api/version': return self.send(200, {'version': version(), 'locked': excel_locked(), 'backups': len(backup_list())})
            if path == '/api/predatory':
                global PRED
                if PRED is None: PRED = predatory_lists()
                return self.send(200, PRED)
            if path == '/api/images': return self.send(200, {'images': list_images()})
            if path == '/api/arxiv':
                from urllib.parse import parse_qs
                return self.send(200, arxiv_lookup((parse_qs(urlparse(self.path).query).get('id') or [''])[0]))
            if path.startswith('/site/'):                         # preview the website with the current workbook
                rel = path[len('/site/'):] or 'index.html'
                if rel == 'data/database.xlsx':
                    with open(DB, 'rb') as f: return self.send(200, f.read(), TYPES['.xlsx'])
                return self.file(SITE, rel)
            if path in ('/', '/index.html'): return self.file(HERE, 'index.html')
            return self.file(HERE, path)
        except UserError as e:
            self.send(e.status, {'error': str(e)})
        except Exception as e:
            self.send(500, {'error': f'{type(e).__name__}: {e}'})

    def do_POST(self):
        path = urlparse(self.path).path
        if self.headers.get('Origin') not in (None, f'http://127.0.0.1:{ARGS.port}', f'http://localhost:{ARGS.port}'):
            return self.send(403, {'error': 'Forbidden'})
        try:
            body = json.loads(self.rfile.read(int(self.headers.get('Content-Length') or 0)) or b'{}')
            if path == '/api/op': return self.send(200, apply_op(body))
            if path == '/api/undo': return self.send(200, undo())
            if path == '/api/upload-image': return self.send(200, upload_image(body))
            self.send(404, {'error': 'Not found'})
        except UserError as e:
            self.send(e.status, {'error': str(e)})
        except Exception as e:
            self.send(500, {'error': f'{type(e).__name__}: {e}'})


def main():
    if not os.path.isfile(DB):
        print('Cannot find the workbook:', DB); input('Press Enter to close'); sys.exit(1)
    try:
        srv = ThreadingHTTPServer(('127.0.0.1', ARGS.port), H)
    except OSError:
        url = f'http://localhost:{ARGS.port}/'
        print('The manager is already running:', url)
        if not ARGS.no_browser: webbrowser.open(url)
        return
    url = f'http://localhost:{ARGS.port}/'
    print('Database manager running at', url)
    print('Workbook:', DB)
    print('Keep this window open while you work. Close it (or press Ctrl+C) to stop.')
    if not ARGS.no_browser: threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == '__main__':
    main()
