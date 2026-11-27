"""Rebuild the "Hướng dẫn" (guide) sheet of data/database.xlsx: detailed Vietnamese instructions for every sheet.

Run from anywhere:  python tools/build_guide.py
Edit the texts below, never the sheet by hand, so the guide can always be regenerated. Needs openpyxl.
The script replaces every old guide sheet (so there is exactly one) and leaves all data sheets untouched.
Close the workbook in Excel first: Windows locks open files.
"""
import sys, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

os.chdir(sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
wb = openpyxl.load_workbook('data/database.xlsx')
# Exactly one guide sheet: drop every old one (Instructions, Hướng dẫn, Hướng dẫn1, ...) before writing it again
for name in list(wb.sheetnames):
    if name.startswith('Hướng dẫn') or name.lower().startswith('instruction'):
        del wb[name]
ws = wb.create_sheet('Hướng dẫn', 0)

NAVY, BLUE, SOFT, GREY = '0A1730', '1B4DFF', 'EAF0FF', 'F4F6FA'
thin = Side(style='thin', color='DFE5F0')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical='top')
for col, w in zip('ABCD', (24, 12, 70, 44)):
    ws.column_dimensions[col].width = w
ws.sheet_view.showGridLines = False


def title(text):
    ws.append([text])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    c = ws.cell(r, 1)
    c.font, c.fill, c.alignment = Font(bold=True, size=16, color='FFFFFF'), PatternFill('solid', fgColor=NAVY), Alignment(vertical='center', indent=1)
    ws.row_dimensions[r].height = 32


def section(text, sub=''):
    ws.append([])
    ws.append([text])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    c = ws.cell(r, 1)
    c.font, c.fill, c.alignment = Font(bold=True, size=13, color=BLUE), PatternFill('solid', fgColor=SOFT), Alignment(vertical='center', indent=1)
    ws.row_dimensions[r].height = 24
    if sub:
        para(sub, italic=True)


def para(text, bold=False, italic=False, height=None):
    ws.append([text])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    c = ws.cell(r, 1)
    c.font, c.alignment = Font(bold=bold, italic=italic, color='55627A' if italic else '0A1730'), Alignment(wrap_text=True, vertical='top', indent=1)
    # merged cells do not auto-fit: estimate the height from the text length (about 140 characters per line)
    lines = sum(max(1, -(-len(part) // 140)) for part in text.split('\n'))
    ws.row_dimensions[r].height = height or 15.5 * lines + 3


def table(rows):
    ws.append(['Cột', 'Bắt buộc?', 'Cách nhập', 'Ví dụ'])
    r = ws.max_row
    for i in range(1, 5):
        c = ws.cell(r, i)
        c.font, c.fill, c.border, c.alignment = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor=BLUE), BOX, Alignment(vertical='center')
    for k, row in enumerate(rows):
        ws.append(list(row))
        r = ws.max_row
        for i in range(1, 5):
            c = ws.cell(r, i)
            c.alignment, c.border = WRAP, BOX
            if k % 2:
                c.fill = PatternFill('solid', fgColor=GREY)
        ws.cell(r, 1).font = Font(bold=True)
        lines = max(max(1, -(-len(part) // w)) for t, w in ((row[2], 66), (row[3], 40), (row[0], 22)) for part in str(t).split('\n'))
        ws.row_dimensions[r].height = 15.5 * lines + 4


# ---------------------------------------------------------------- content
title('HƯỚNG DẪN SỬ DỤNG FILE DATABASE CỦA WEBSITE')
para('File này (data/database.xlsx) là "cơ sở dữ liệu" của cả website. Trang chủ (index.html) và trang tuyển sinh viên nghiên cứu (research.html) đều đọc trực tiếp từ file này mỗi khi có người mở web. '
     'Bạn KHÔNG cần sửa code: chỉ cần sửa file Excel, lưu lại và đẩy lên GitHub.')

section('1. CÁC BƯỚC CẬP NHẬT WEBSITE')
para('Bước 1. Mở file E:\\Coding\\Resume\\data\\database.xlsx bằng Excel (đúng file nằm trong thư mục website, không phải bản sao ở nơi khác).')
para('Bước 2. Thêm, sửa hoặc xóa dòng trong sheet cần thay đổi (xem hướng dẫn từng sheet bên dưới).')
para('Bước 3. Lưu file (Ctrl + S), giữ nguyên định dạng .xlsx. Đóng Excel trước khi đẩy lên để file không bị khóa.')
para('Bước 4. Mở GitHub Desktop: ở ô Summary (góc dưới bên trái) ghi ngắn gọn đã sửa gì, ví dụ "Thêm đề tài về polyp", bấm Commit to main, rồi bấm Push origin.')
para('Bước 5. Đợi 1–2 phút để GitHub Pages cập nhật, rồi mở lại website (nhấn Ctrl + F5 để tải mới hoàn toàn).')

section('2. QUY TẮC CHUNG (áp dụng cho mọi sheet)')
para('• KHÔNG đổi tên sheet và KHÔNG đổi tên cột ở dòng 1 (dòng tiêu đề màu xanh đậm). Website tìm dữ liệu theo đúng các tên này.')
para('• Mỗi dòng là một bản ghi (một bài báo, một đề tài, một hội nghị...). Dòng trống bị bỏ qua, không gây lỗi.')
para('• Thứ tự dòng KHÔNG quan trọng: website tự sắp xếp (xem mục 9). Bạn có thể thêm dòng mới ở cuối bảng cho tiện.')
para('• Muốn xóa một bản ghi: xóa cả dòng (chuột phải vào số dòng → Delete). Đừng chỉ xóa chữ trong vài ô.')
para('• Các ô có mũi tên thả xuống (Direction, Status, Rank, Hybrid...) chỉ nhận giá trị trong danh sách: hãy chọn từ danh sách thay vì gõ tay để tránh sai chính tả.')
para('• Có thể thêm cột ghi chú riêng của bạn: website bỏ qua những cột nó không biết.')
para('• Một ô có nhiều dòng: nhấn Alt + Enter để xuống dòng trong cùng một ô.')
para('• LƯU Ý: file này công khai trên Internet (ai cũng tải được). Không ghi số điện thoại, điểm số hay thông tin riêng tư của sinh viên.')

section('3. SHEET "Publications": danh sách bài báo  →  trang chủ, mục Publications')
table([
    ('Type', 'Có', 'Loại bài: chọn Journal, Conference, Thesis hoặc Submission / Preprint.', 'Conference'),
    ('Status', 'Có', 'Tình trạng: Published (đã xuất bản), Accepted (đã được nhận), In Submission (đang nộp/đang phản biện), Thesis (luận văn).', 'Accepted'),
    ('Date', 'Có (ít nhất năm)', 'Ngày công bố. Nhập kiểu nào cũng được: dd-mm-yyyy, mm-yyyy, chỉ năm yyyy, hoặc ô ngày của Excel. Bài KHÔNG có năm sẽ bị ẩn khỏi web.', '15-03-2026  hoặc  03-2026  hoặc  2026'),
    ('Title', 'Có', 'Tên bài báo đầy đủ.', 'V-FUSE: Dynamic Gating Fusion ...'),
    ('Authors', 'Có', 'Các tác giả, cách nhau bằng dấu phẩy. Tên "Ha Anh Vu" tự động in đậm.', 'Hien Tran, Ha Anh Vu'),
    ('Venue', 'Có', 'Tên tạp chí/hội nghị. Có thể ghi ngắn hoặc đầy đủ: web tự nhận ra các cách ghi khác nhau của cùng một hội nghị (ví dụ "RIVF 2026" và "IEEE RIVF").', '20th IEEE-RIVF International Conference ... (RIVF 2026)'),
    ('DOI', 'Không', 'Mã DOI hoặc link đầy đủ. Để trống thì không hiện nút DOI.', '10.1007/s11760-025-01234-5'),
    ('PDF', 'Không', 'Link file PDF. Để trống thì không hiện nút.', 'https://arxiv.org/pdf/2501.01234'),
    ('Code', 'Không', 'Link mã nguồn (GitHub...). Để trống thì không hiện nút.', 'https://github.com/AnhVH-FPT/...'),
    ('Link ID', 'Không', 'Mã ngắn để nối bài báo với đề tài luận văn trong sheet Projects (xem mục 4).', 'AC1'),
])
para('Mã số bài (J1, C2, T1, S3...) do web tự đánh theo loại bài, không cần nhập.', italic=True)

section('4. SHEET "Projects": đề tài luận văn đang hướng dẫn  →  trang chủ, mục Undergraduate thesis projects')
table([
    ('Semester', 'Có', 'Học kỳ. Ghi Fall/Summer/Spring + năm. Cách viết tắt SP26, SU26, FA26 cũng được.', 'Fall 2026  hoặc  FA26'),
    ('Title', 'Có', 'Tên đề tài.', 'Ultra-Early Warning System for Tiny Fire and Smoke'),
    ('Members', 'Không', 'Tên sinh viên, cách nhau bằng dấu chấm phẩy.', 'Nguyen Van A; Tran Thi B'),
    ('Thesis Status', 'Có', 'Preparing = vừa chọn đề tài, sắp bắt đầu. In progress = đang làm. Defended = đã bảo vệ xong.', 'In progress'),
    ('Paper Status', 'Không', 'None (hoặc để trống) = chưa có bài báo. Submitted = đã nộp bài (hiện nhãn "Paper submitted"). Accepted = bài đã được nhận (hiện nhãn "Accepted, see C1", bấm vào sẽ nhảy tới bài báo).', 'Submitted'),
    ('Paper Venue', 'Không', 'Nơi đã nộp bài, hiện sau chữ "Paper submitted to".', 'RIVF 2026'),
    ('Link ID', 'Không', 'Dùng khi Paper Status = Accepted: ghi cùng một mã ở đây và ở dòng bài báo tương ứng trong sheet Publications. Tên đề tài và tên bài báo không cần giống nhau. Mỗi cặp đề tài–bài báo dùng một mã riêng.', 'AC1'),
])

section('5. SHEET "Directions": các hướng nghiên cứu  →  trang chủ (Research interests) và trang research.html')
table([
    ('Code', 'Có', 'Mã ngắn của hướng, viết hoa, không dấu, không khoảng trắng. Đã dùng rồi thì KHÔNG đổi, vì sheet Topics, Conferences và Journals đều tham chiếu tới mã này.', 'MI'),
    ('Direction', 'Có', 'Tên hướng hiển thị trên nút lọc của web. Có thể đổi tên thoải mái.', 'Medical Imaging'),
    ('Focus (current research trends)', 'Không', 'Các xu hướng chính, hiện thành dòng "Current trends" khi sinh viên chọn hướng này.', 'Medical SAM/foundation models, nnU-Net ...'),
    ('Summary', 'Không', 'Một câu mô tả ngắn (khoảng 10 từ, tiếng Anh) hiện trên thẻ hướng nghiên cứu ở mục Research interests của trang chủ. Mỗi thẻ bấm vào sẽ mở trang research.html lọc sẵn hướng đó, kèm số đề tài còn Open.', 'Deep learning for diagnosis and segmentation of medical scans.'),
    ('Topics', 'Tự động', 'Công thức Excel tự đếm số đề tài của hướng. Không cần sửa; khi thêm hướng mới thì copy công thức từ dòng trên xuống.', '=COUNTIF(Topics!$A:$A,A3)'),
])
para('Thêm một hướng mới: thêm một dòng ở sheet này (ví dụ Code = NLP, Direction = Natural Language Processing), sau đó dùng mã NLP trong cột Direction của sheet Topics. '
     'Thứ tự các dòng trong sheet này chính là thứ tự các thẻ Research interests ở trang chủ, các nút lọc và các nhóm đề tài trên web.')

section('6. SHEET "Topics": danh sách đề tài cho sinh viên  →  trang research.html, mục Research topics')
table([
    ('Direction', 'Có', 'Chọn mã hướng trong danh sách thả xuống (lấy từ sheet Directions).', 'MI'),
    ('Status', 'Có', 'Open = còn trống, sinh viên bấm Apply được. Taken = đã có người nhận: đề tài hiện mờ, ẩn nút Apply, và web hiện thêm ô lọc "Open topics only". Dòng Taken cũng tự hiện chữ xám trong Excel.', 'Open'),
    ('Vietnamese Title', 'Không', 'Tên đề tài tiếng Việt (hiện dưới tên tiếng Anh).', 'Phân vùng đa cơ quan ổ bụng trên ảnh CT ...'),
    ('English Title', 'Có', 'Tên đề tài tiếng Anh (tiêu đề chính). Dòng không có cột này sẽ bị bỏ qua.', 'Parameter-Efficient Adaptation of Medical SAM ...'),
    ('Dataset', 'Không', 'Mỗi bộ dữ liệu gồm 2 dòng trong cùng ô (Alt + Enter để xuống dòng): dòng 1 là "• Tên", dòng 2 là link. 3 bộ đầu tiên hiện thành nhãn nhỏ trên thẻ đề tài.',
     '• Kvasir-SEG\nhttps://datasets.simula.no/kvasir-seg/\n• CVC-ClinicDB\nhttps://polyp.grand-challenge.org/...'),
    ('Research Description', 'Có', 'Mô tả đề tài, PHẢI nhập theo mẫu ở mục 6a bên dưới. Gồm 6 mục, mỗi mục bắt đầu bằng đúng tên mục và dấu hai chấm ở đầu dòng. Mục Context cũng là câu tóm tắt hiện trên thẻ đề tài khi chưa mở ra.',
     'Context: Organ delineation on CT is required ...\nProblem formulation: Input: a 3D CT volume. Output: organ masks.\nGap: ...\nProposed approach: ...\nImplementation steps:\n(1) ...\n(2) ...\nKey knowledge to learn: ...'),
    ('Research Expectation', 'Có', 'Kết quả mong đợi, PHẢI nhập theo mẫu ở mục 6a bên dưới. Gồm 6 mục, cùng quy tắc như Research Description.',
     'Deliverables: ...\nExpected results:\n(1) ...\n(2) ...\nBaselines: ...\nMetrics: ...\nExperimental protocol: ...\nMinimum success criteria: ...'),
])
para('Mã đề tài (ví dụ MI-07) KHÔNG nhập trong Excel: web tự đánh mã = mã hướng + thứ tự của đề tài trong các dòng thuộc hướng đó, tính từ trên xuống. '
     'Dòng MI đầu tiên trong sheet là MI-01, dòng MI thứ hai là MI-02... (các dòng của hướng khác xen giữa không ảnh hưởng). Đề tài nằm trên thì số nhỏ hơn.', bold=True)
para('Hệ quả cần nhớ: XÓA hoặc CHÈN/DI CHUYỂN một dòng sẽ làm đổi số của các đề tài cùng hướng nằm phía dưới (ví dụ xóa MI-03 thì MI-04 cũ thành MI-03). '
     'Sinh viên đã đăng ký theo mã cũ hoặc link research.html#MI-04 đã gửi đi sẽ trỏ sang đề tài khác. Vì vậy: muốn thêm đề tài thì thêm ở CUỐI bảng; '
     'đề tài đã có người nhận thì đổi Status thành Taken thay vì xóa dòng.')
para('Ví dụ thêm đề tài: kéo xuống dòng trống cuối bảng, chọn Direction = MI, Status = Open, rồi điền tên, dataset, mô tả theo mẫu ở mục 6a. Nếu đang có 50 đề tài MI thì đề tài mới sẽ là MI-51.')

section('6a. MẪU NHẬP "Research Description" VÀ "Research Expectation" (bắt buộc để hiện trên web)')
para('Web chỉ hiện những mục viết ĐÚNG mẫu dưới đây. Cách nhập: trong một ô, mỗi mục bắt đầu ở đầu dòng bằng tên mục + dấu hai chấm, rồi đến nội dung. Xuống dòng trong ô bằng Alt + Enter. '
     'Mẹo: copy nguyên khung mẫu dưới đây vào ô rồi điền nội dung sau dấu hai chấm. Rê chuột lên ô tiêu đề cột trong sheet Topics cũng thấy nhắc lại mẫu.')
table([
    ('Research Description', 'Thứ tự', 'Tên mục phải gõ đúng (không phân biệt hoa thường)', 'Nội dung nên ghi'),
    ('', '1', 'Context:', 'Bối cảnh, vì sao bài toán quan trọng (1–2 câu ngắn, dùng làm câu tóm tắt trên thẻ).'),
    ('', '2', 'Problem formulation:', 'Đầu vào và đầu ra của bài toán.'),
    ('', '3', 'Gap:', 'Hạn chế của các phương pháp hiện tại.'),
    ('', '4', 'Proposed approach:', 'Hướng tiếp cận đề xuất.'),
    ('', '5', 'Implementation steps:', 'Các bước thực hiện: xuống dòng sau dấu hai chấm, mỗi bước một dòng bắt đầu bằng (1), (2), (3)...'),
    ('', '6', 'Key knowledge to learn:', 'Kiến thức sinh viên cần học.'),
    ('Research Expectation', 'Thứ tự', 'Tên mục phải gõ đúng (không phân biệt hoa thường)', 'Nội dung nên ghi'),
    ('', '1', 'Deliverables:', 'Sản phẩm cuối cùng cần nộp.'),
    ('', '2', 'Expected results:', 'Các kết quả cụ thể cần đạt: xuống dòng, mỗi kết quả một dòng bắt đầu bằng (1), (2)...'),
    ('', '3', 'Baselines:', 'Các phương pháp dùng để so sánh.'),
    ('', '4', 'Metrics:', 'Thước đo đánh giá.'),
    ('', '5', 'Experimental protocol:', 'Cách chia dữ liệu và cách đánh giá.'),
    ('', '6', 'Minimum success criteria:', 'Ngưỡng tối thiểu để coi là thành công.'),
])
para('Quy tắc:', bold=True)
para('• Web luôn hiện các mục theo đúng thứ tự trong bảng trên, dù trong ô bạn viết theo thứ tự khác.')
para('• Mục nào không có thì bỏ hẳn dòng đó (hoặc để trống sau dấu hai chấm): web sẽ không hiện mục đó.')
para('• Chữ nằm ngoài các mục trên (ví dụ dòng đầu không có tên mục, hoặc tên mục gõ sai như "Contex:" hay "Result:") sẽ KHÔNG hiện trên web. '
     'Muốn kiểm tra: mở trang research.html, nhấn F12, chọn tab Console: web liệt kê mã đề tài và các dòng bị bỏ qua.')
para('• Trong nội dung một mục vẫn được dùng dấu hai chấm (ví dụ "Problem formulation: Input: ... Output: ..."), vì chỉ 12 tên mục ở trên mới được hiểu là tên mục.')
para('• Một mục có thể dài nhiều dòng: các dòng tiếp theo (không bắt đầu bằng tên mục) thuộc về mục phía trên.')
para('Ví dụ một ô Research Description hoàn chỉnh:', bold=True)
para('Context: Organ delineation on CT is required for radiotherapy planning and volumetry.\n'
     'Problem formulation: Input: a 3D abdominal CT volume. Output: voxel-wise masks for multiple organs.\n'
     'Gap: MedSAM requires manual box prompts per slice; nnU-Net needs full training per task.\n'
     'Proposed approach: Adapt SAM/MedSAM with LoRA and an automatic prompt generator.\n'
     'Implementation steps:\n'
     '(1) Convert CT volumes to 2.5D slices with HU windowing.\n'
     '(2) Train nnU-Net as the reference upper bound.\n'
     '(3) Fine-tune SAM/MedSAM with LoRA.\n'
     'Key knowledge to learn: CT physics, MONAI, SAM/MedSAM, LoRA, Dice/NSD metrics.', italic=True)

section('7. SHEET "Conferences": hội nghị  →  trang research.html, mục Where to publish')
table([
    ('Conference', 'Có', 'Tên viết tắt của hội nghị (hiện in đậm, bấm vào mở website hội nghị).', 'MICCAI'),
    ('Full Name', 'Không', 'Tên đầy đủ.', 'Medical Image Computing and Computer Assisted Intervention'),
    ('Directions', 'Không', 'Các mã hướng phù hợp, cách nhau bằng dấu phẩy; hoặc All nếu hợp mọi hướng. Để trống cũng tính là All. Dùng cho bộ lọc Direction.', 'MI, XAI   hoặc   All'),
    ('CORE Rank', 'Không', 'Chọn A*, A, B, C hoặc Not ranked. Để trống sẽ hiện là Not ranked.', 'A'),
    ('Proceedings', 'Không', 'Nơi xuất bản kỷ yếu (chỉ để bạn tham khảo, web không hiện cột này).', 'Springer LNCS – Scopus'),
    ('Submission Deadline', 'Không', 'Hạn nộp bài, ghi tên tháng bằng TIẾNG ANH (viết đầy đủ hoặc 3 chữ đầu như Jan, Feb). Được dùng cho bộ lọc Deadline và để sắp xếp. '
     'Có thể thêm Early / Mid / Late; khoảng thời gian dùng dấu gạch ngang; ghi chú để trong ngoặc đơn (web bỏ qua phần trong ngoặc).',
     'Mid January\nFebruary–March\nMid December – January\nJuly (round 1) / September (round 2)'),
    ('Start Preparing', 'Không', 'Thời điểm nên bắt đầu làm để kịp hạn (hiện nguyên văn).', 'September (≈4 months before)'),
    ('Held In', 'Không', 'Tháng tổ chức hội nghị, cùng quy tắc ghi tháng như Submission Deadline. Dùng cho bộ lọc Held in.', 'October–November'),
    ('Presentation Format', 'Không', 'Hình thức trình bày (chỉ để tham khảo).', 'Hybrid – online/remote presentation accepted'),
    ('Hybrid', 'Không', 'Yes nếu cho phép trình bày online, No nếu bắt buộc đến trực tiếp.', 'Yes'),
    ('Website', 'Không', 'Link trang chủ hội nghị.', 'https://conferences.miccai.org/'),
])

section('8. SHEET "Journals" và "Venue Notes"  →  trang research.html, mục Where to publish')
table([
    ('Journal', 'Có', 'Tên tạp chí (bấm vào mở link).', 'Medical Image Analysis'),
    ('Publisher', 'Không', 'Nhà xuất bản.', 'Elsevier'),
    ('Rank', 'Không', 'Chọn Q1, Q2, Q3 hoặc Q4 (theo SCImago). Để trống thì xếp cuối danh sách.', 'Q1'),
    ('Directions', 'Không', 'Mã hướng cách nhau bằng dấu phẩy, hoặc All cho tạp chí AI/ML tổng quát. Lưu ý: cột này ban đầu được điền TỰ ĐỘNG theo tên tạp chí, nên rà lại.', 'MI   hoặc   CV, MI   hoặc   All'),
    ('Link', 'Không', 'Link trang tạp chí.', 'https://www.sciencedirect.com/journal/medical-image-analysis'),
])
para('Sheet "Venue Notes": mỗi dòng là một ghi chú hiện trong mục "Notes" dưới bảng. Cột Applies To chọn Conferences, Journals hoặc All (hiện ở cả hai bảng); cột Note là nội dung ghi chú.')

section('9. WEBSITE TỰ SẮP XẾP NHƯ THẾ NÀO?', 'Bạn có thể nhập dòng theo bất kỳ thứ tự nào: web luôn tự sắp xếp lại như sau.')
table([
    ('Publications', '', 'Chế độ By year: năm mới nhất lên trước. Trong mỗi năm: Published và Thesis trước, rồi Accepted, rồi In Submission. Các bài cùng một hội nghị/tạp chí được gom cạnh nhau, '
     'trong đó bài mới hơn lên trước; bài trùng ngày và trùng nơi đăng xếp theo tên A→Z. Chế độ By type: chia nhóm Journal, Conference, Thesis, Submission, cũng gom theo nơi đăng.', '3 năm gần nhất mở sẵn, các năm cũ thu gọn'),
    ('Projects', '', 'Học kỳ mới nhất lên trước (năm mới trước; trong một năm: Fall, rồi Summer, rồi Spring). 2 học kỳ mới nhất mở sẵn.', 'Fall 2026 → Summer 2026 → Spring 2026'),
    ('Topics', '', 'Theo thứ tự hướng trong sheet Directions, rồi theo thứ tự dòng trong sheet Topics (cũng chính là số trong mã đề tài).', 'CV-01, CV-02 ... MI-01, MI-02 ...'),
    ('Conferences', '', 'Theo tháng hạn nộp, từ tháng 1 đến tháng 12 (khoảng "Mid December – January" tính là tháng 12). Cùng tháng thì Early trước Mid, Mid trước Late; '
     'rồi theo CORE Rank (A* → A → B → C → Not ranked); rồi theo tên A→Z. Hội nghị không ghi hạn nộp xếp cuối.', 'Mid January → Late January → February ...'),
    ('Journals', '', 'Theo Rank: Q1 → Q2 → Q3 → Q4 (không có rank xếp cuối), rồi theo tên tạp chí A→Z.', 'Q1 Applied Soft Computing → Q1 Medical Image Analysis → Q2 ...'),
    ('Directions, Venue Notes', '', 'Giữ đúng thứ tự dòng trong sheet (bạn tự quyết thứ tự).', ''),
])

section('10. GẶP LỖI THÌ KIỂM TRA GÌ?')
para('• Web báo "Could not load" hoặc "was not found": kiểm tra file vẫn tên đúng là database.xlsx, nằm trong thư mục data, và đã Push lên GitHub.')
para('• Sửa rồi mà web chưa đổi: đợi thêm 1–2 phút rồi nhấn Ctrl + F5. Kiểm tra trên GitHub Desktop rằng đã Push (không còn nút "Push origin").')
para('• Một bài báo không hiện: cột Date chưa có năm.')
para('• Một đề tài không hiện: thiếu English Title, hoặc ô tìm kiếm/bộ lọc trên web đang chọn hướng khác.')
para('• Đề tài nằm sai nhóm: cột Direction ghi mã không có trong sheet Directions (web sẽ tạo nhóm riêng theo mã đó).')
para('• Hội nghị không hiện khi lọc theo tháng: cột Submission Deadline/Held In chưa có tên tháng tiếng Anh.')
para('• Mở file index.html bằng cách nhấp đúp sẽ không đọc được file Excel (trình duyệt chặn). Hãy xem trên website thật, hoặc chạy "python -m http.server" trong thư mục website rồi mở http://localhost:8000.')

ws.freeze_panes = 'A2'
order = ['Hướng dẫn', 'Publications', 'Projects', 'Directions', 'Topics', 'Conferences', 'Journals', 'Venue Notes']
wb._sheets = [wb[n] for n in order] + [s for s in wb._sheets if s.title not in order]
wb.active = 0
for s in wb._sheets:
    s.sheet_view.tabSelected = s.title == 'Hướng dẫn'
wb.save('data/database.xlsx')
print('rows', ws.max_row)
