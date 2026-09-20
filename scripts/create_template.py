# -*- coding: utf-8 -*-
"""
Tạo file Excel mẫu chuyên nghiệp để import danh sách user
- STT tự động (công thức Excel)
- Validation email & ngày
- Conditional formatting (hết hạn → đỏ, sắp hết → vàng)
- Freeze panes + Auto filter
- Số điện thoại format text (giữ số 0 đầu)
- Sheet 2: Hướng dẫn sử dụng tiếng Việt

Chạy: python scripts/create_template.py
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
import os
import sys

# ====== CẤU HÌNH ======
OUTPUT_FILE = "templates/users_template.xlsx"
FORCE_OVERWRITE = False  # True = luôn ghi đè, False = chỉ tạo nếu chưa có

# ====== MÀU SẮC ======
COLOR_HEADER_BG = "2563EB"       # Xanh dương đậm
COLOR_HEADER_TEXT = "FFFFFF"     # Trắng
COLOR_EMAIL_BG = "DBEAFE"        # Xanh nhạt
COLOR_NAME_BG = "F0F9FF"         # Xanh rất nhạt
COLOR_PHONE_BG = "FEF3C7"        # Vàng nhạt
COLOR_EXPIRES_BG = "DCFCE7"      # Xanh lá nhạt
COLOR_STT_BG = "F1F5F9"          # Xám nhạt
COLOR_ALT_ROW = "F8FAFC"         # Xen kẽ dòng
COLOR_NOTE_BG = "FEF9C3"         # Vàng note
COLOR_BORDER = "CBD5E1"
COLOR_SECTION_TITLE = "1E40AF"
COLOR_ERROR = "DC2626"
COLOR_WARNING = "F59E0B"

# ====== KIỂM TRA FILE ĐÃ TỒN TẠI ======
if os.path.exists(OUTPUT_FILE) and not FORCE_OVERWRITE:
    print(f"⚠️  File đã tồn tại: {OUTPUT_FILE}")
    print(f"   Để tạo lại, xóa file này hoặc đặt FORCE_OVERWRITE = True")
    print(f"   Hiện tại: BỎ QUA")
    sys.exit(0)

# ====== TẠO WORKBOOK ======
print(f"📄 Đang tạo file Excel mẫu...")
wb = openpyxl.Workbook()

# ============================================================
# SHEET 1: USERS
# ============================================================
ws = wb.active
ws.title = "Users"

# Định nghĩa cột: (tên, độ rộng, align)
COLUMNS = [
    ("STT", 6, "center"),
    ("email", 35, "left"),
    ("name", 25, "left"),
    ("phone", 16, "center"),
    ("expiresAt", 14, "center"),
    ("Ghi chú", 30, "left"),
]

# ====== STYLE CHUNG ======
thin_border = Border(
    left=Side(style='thin', color=COLOR_BORDER),
    right=Side(style='thin', color=COLOR_BORDER),
    top=Side(style='thin', color=COLOR_BORDER),
    bottom=Side(style='thin', color=COLOR_BORDER),
)

header_font = Font(name='Segoe UI', size=11, bold=True, color=COLOR_HEADER_TEXT)
header_fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type='solid')
header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

body_font = Font(name='Segoe UI', size=10, color='0F172A')
body_align_left = Alignment(horizontal='left', vertical='center', wrap_text=False)
body_align_center = Alignment(horizontal='center', vertical='center', wrap_text=False)

# ====== GHI HEADER ======
for col_idx, (col_name, width, align) in enumerate(COLUMNS, start=1):
    cell = ws.cell(row=1, column=col_idx, value=col_name)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_align
    cell.border = thin_border
    ws.column_dimensions[get_column_letter(col_idx)].width = width

ws.row_dimensions[1].height = 32

# ====== DỮ LIỆU MẪU (3 dòng ví dụ) ======
SAMPLE_DATA = [
    ("nguyenvana@gmail.com", "Nguyễn Văn A", "", "2026-12-31", "⚠️ Ví dụ - xóa dòng này khi dùng"),
    ("tranthib@gmail.com", "Trần Thị B", "", "", "⚠️ Ví dụ - để trống hạn = vĩnh viễn"),
    ("levanc@gmail.com", "Lê Văn C", "", "2027-06-15", "⚠️ Ví dụ - xóa trước khi import"),
]

START_ROW = 2

for i, (email, name, phone, expires, note) in enumerate(SAMPLE_DATA):
    row = START_ROW + i
    
    # STT - công thức tự động
    stt_cell = ws.cell(row=row, column=1)
    stt_cell.value = f'=IF(B{row}<>"",ROW()-1,"")'
    stt_cell.font = Font(name='Segoe UI', size=10, bold=True, color='64748B')
    stt_cell.alignment = body_align_center
    stt_cell.fill = PatternFill(start_color=COLOR_STT_BG, end_color=COLOR_STT_BG, fill_type='solid')
    stt_cell.border = thin_border
    
    # Email
    email_cell = ws.cell(row=row, column=2, value=email)
    email_cell.font = body_font
    email_cell.alignment = body_align_left
    email_cell.fill = PatternFill(start_color=COLOR_EMAIL_BG, end_color=COLOR_EMAIL_BG, fill_type='solid')
    email_cell.border = thin_border
    
    # Name
    name_cell = ws.cell(row=row, column=3, value=name)
    name_cell.font = body_font
    name_cell.alignment = body_align_left
    name_cell.fill = PatternFill(start_color=COLOR_NAME_BG, end_color=COLOR_NAME_BG, fill_type='solid')
    name_cell.border = thin_border
    
    # Phone - text format
    phone_cell = ws.cell(row=row, column=4, value=phone or None)
    phone_cell.font = body_font
    phone_cell.alignment = body_align_center
    phone_cell.fill = PatternFill(start_color=COLOR_PHONE_BG, end_color=COLOR_PHONE_BG, fill_type='solid')
    phone_cell.border = thin_border
    phone_cell.number_format = '@'
    
    # ExpiresAt
    exp_cell = ws.cell(row=row, column=5, value=expires or None)
    exp_cell.font = body_font
    exp_cell.alignment = body_align_center
    exp_cell.fill = PatternFill(start_color=COLOR_EXPIRES_BG, end_color=COLOR_EXPIRES_BG, fill_type='solid')
    exp_cell.border = thin_border
    if expires:
        exp_cell.number_format = 'YYYY-MM-DD'
    
    # Note
    note_cell = ws.cell(row=row, column=6, value=note)
    note_cell.font = Font(name='Segoe UI', size=9, italic=True, color='92400E')
    note_cell.alignment = body_align_left
    note_cell.fill = PatternFill(start_color=COLOR_NOTE_BG, end_color=COLOR_NOTE_BG, fill_type='solid')
    note_cell.border = thin_border
    
    ws.row_dimensions[row].height = 22

# ====== 50 DÒNG TRỐNG VỚI STT TỰ ĐỘNG ======
EMPTY_ROWS = 50

for i in range(EMPTY_ROWS):
    row = START_ROW + len(SAMPLE_DATA) + i
    
    # STT tự động
    stt_cell = ws.cell(row=row, column=1)
    stt_cell.value = f'=IF(B{row}<>"",ROW()-1,"")'
    stt_cell.font = Font(name='Segoe UI', size=10, color='94A3B8')
    stt_cell.alignment = body_align_center
    stt_cell.border = thin_border
    if i % 2 == 1:
        stt_cell.fill = PatternFill(start_color=COLOR_ALT_ROW, end_color=COLOR_ALT_ROW, fill_type='solid')
    
    # Các cột còn lại
    for col_idx in range(2, 7):
        cell = ws.cell(row=row, column=col_idx)
        cell.font = body_font
        cell.border = thin_border
        
        if col_idx == 4:  # Phone
            cell.alignment = body_align_center
            cell.number_format = '@'
        elif col_idx == 5:  # ExpiresAt
            cell.alignment = body_align_center
            cell.number_format = 'YYYY-MM-DD'
        elif col_idx == 6:  # Note
            cell.alignment = body_align_left
        else:
            cell.alignment = body_align_left
        
        if i % 2 == 1:
            cell.fill = PatternFill(start_color=COLOR_ALT_ROW, end_color=COLOR_ALT_ROW, fill_type='solid')
    
    ws.row_dimensions[row].height = 20

LAST_ROW = START_ROW + len(SAMPLE_DATA) + EMPTY_ROWS - 1

# ====== DATA VALIDATION ======

# Email: ít nhất 5 ký tự
email_dv = DataValidation(
    type="textLength",
    operator="greaterThanOrEqual",
    formula1="5",
    allow_blank=True,
    showErrorMessage=True,
    errorTitle="Email không hợp lệ",
    error="Email phải có ít nhất 5 ký tự và chứa @",
    promptTitle="Nhập email Google",
    prompt="Ví dụ: nguyenvana@gmail.com",
    showInputMessage=True
)
email_dv.add(f'B{START_ROW}:B{LAST_ROW}')
ws.add_data_validation(email_dv)

# ExpiresAt: phải là ngày hợp lệ
date_dv = DataValidation(
    type="date",
    operator="greaterThan",
    formula1="DATE(2020,1,1)",
    allow_blank=True,
    showErrorMessage=True,
    errorTitle="Ngày không hợp lệ",
    error="Ngày phải có định dạng YYYY-MM-DD và sau 01/01/2020",
    promptTitle="Hạn sử dụng",
    prompt="Định dạng: YYYY-MM-DD (VD: 2026-12-31)\nĐể trống = vĩnh viễn",
    showInputMessage=True
)
date_dv.add(f'E{START_ROW}:E{LAST_ROW}')
ws.add_data_validation(date_dv)

# ====== CONDITIONAL FORMATTING ======

# Email không có @ → tô đỏ
ws.conditional_formatting.add(
    f'B{START_ROW}:B{LAST_ROW}',
    CellIsRule(
        operator='notContainsText',
        formula=['"@"'],
        fill=PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid'),
        font=Font(color='DC2626', bold=True)
    )
)

# Ngày đã hết hạn → tô đỏ
ws.conditional_formatting.add(
    f'E{START_ROW}:E{LAST_ROW}',
    CellIsRule(
        operator='lessThan',
        formula=['TODAY()'],
        fill=PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid'),
        font=Font(color='DC2626', bold=True)
    )
)

# Ngày sắp hết (≤ 30 ngày) → tô vàng
ws.conditional_formatting.add(
    f'E{START_ROW}:E{LAST_ROW}',
    CellIsRule(
        operator='between',
        formula=['TODAY()', 'TODAY()+30'],
        fill=PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid'),
        font=Font(color='92400E', bold=True)
    )
)

# ====== FREEZE PANES + AUTO FILTER ======
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:F{LAST_ROW}"

# ============================================================
# SHEET 2: HƯỚNG DẪN
# ============================================================
ws2 = wb.create_sheet("Hướng dẫn")

ws2.column_dimensions['A'].width = 4
ws2.column_dimensions['B'].width = 28
ws2.column_dimensions['C'].width = 70

# Tiêu đề chính
ws2.merge_cells('A1:C1')
title_cell = ws2['A1']
title_cell.value = "📖  HƯỚNG DẪN IMPORT TÀI KHOẢN USER"
title_cell.font = Font(name='Segoe UI', size=16, bold=True, color='FFFFFF')
title_cell.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type='solid')
title_cell.alignment = Alignment(horizontal='center', vertical='center')
ws2.row_dimensions[1].height = 42

def add_section(row, title, color=COLOR_SECTION_TITLE):
    ws2.merge_cells(f'A{row}:C{row}')
    cell = ws2.cell(row=row, column=1, value=title)
    cell.font = Font(name='Segoe UI', size=12, bold=True, color=color)
    cell.fill = PatternFill(start_color='DBEAFE', end_color='DBEAFE', fill_type='solid')
    cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws2.row_dimensions[row].height = 28

def add_row(row, label, value, label_color='0F172A', value_color='475569'):
    lc = ws2.cell(row=row, column=2, value=label)
    lc.font = Font(name='Segoe UI', size=10, bold=True, color=label_color)
    lc.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    
    vc = ws2.cell(row=row, column=3, value=value)
    vc.font = Font(name='Segoe UI', size=10, color=value_color)
    vc.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True, indent=1)
    
    ws2.row_dimensions[row].height = 22

row = 3

# Section 1: Cột bắt buộc
add_section(row, "📌  CỘT BẮT BUỘC"); row += 1
add_row(row, "email", "✅ BẮT BUỘC - Email Google của user (VD: nguyenvana@gmail.com)", COLOR_ERROR); row += 1
add_row(row, "name", "⭕ Tùy chọn - Tên hiển thị. Để trống = lấy phần trước @ của email"); row += 1
add_row(row, "expiresAt", "⭕ Tùy chọn - Ngày hết hạn (YYYY-MM-DD). Để trống = vĩnh viễn"); row += 2

# Section 2: Cột không import
add_section(row, "🗑️  CỘT KHÔNG ĐƯỢC IMPORT (tự động bỏ qua)"); row += 1
add_row(row, "STT", "🔢 Tự động đánh số - KHÔNG cần sửa", '64748B'); row += 1
add_row(row, "phone", "📞 Chỉ để tham khảo - KHÔNG được import vào hệ thống", '64748B'); row += 1
add_row(row, "Ghi chú", "📝 Chỉ để tham khảo - KHÔNG được import vào hệ thống", '64748B'); row += 1
add_row(row, "Các cột khác", "➕ Có thể thêm tùy ý - hệ thống sẽ tự động bỏ qua", '64748B'); row += 2

# Section 3: Ví dụ
add_section(row, "💡  VÍ DỤ"); row += 1

# Bảng ví dụ
example_headers = ["", "Email", "Tên / Hạn sử dụng"]
for col_idx, header in enumerate(example_headers, start=1):
    if col_idx == 1:
        continue
    cell = ws2.cell(row=row, column=col_idx, value=header)
    cell.font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill(start_color='475569', end_color='475569', fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin_border
ws2.row_dimensions[row].height = 24
row += 1

examples = [
    ("nguyenvana@gmail.com", "Nguyễn Văn A → hết hạn 31/12/2026"),
    ("tranthib@gmail.com", "(không tên) → vĩnh viễn"),
    ("levanc@gmail.com", "Lê Văn C → hết hạn 15/06/2027"),
]

for email, desc in examples:
    ec = ws2.cell(row=row, column=2, value=email)
    ec.font = Font(name='Consolas', size=10, color='0F172A')
    ec.fill = PatternFill(start_color='F0F9FF', end_color='F0F9FF', fill_type='solid')
    ec.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ec.border = thin_border
    
    dc = ws2.cell(row=row, column=3, value=desc)
    dc.font = Font(name='Segoe UI', size=10, color='475569', italic=True)
    dc.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    dc.border = thin_border
    
    ws2.row_dimensions[row].height = 22
    row += 1

row += 1

# Section 4: Lưu ý quan trọng
add_section(row, "⚠️  LƯU Ý QUAN TRỌNG", COLOR_ERROR); row += 1

notes = [
    ("🚫 Không import admin", "Chỉ import user. Admin phải được thêm thủ công trong Admin Panel."),
    ("📅 Định dạng ngày", "Phải là YYYY-MM-DD. Ví dụ: 2026-12-31 (không dùng 31/12/2026)"),
    ("📧 Email trùng", "Nếu email đã tồn tại → sẽ CẬP NHẬT thông tin (name, expiresAt)"),
    ("🗑️ Xóa dòng ví dụ", "Nhớ xóa 3 dòng ví dụ (màu vàng) trước khi import thực tế"),
    ("➕ Số lượng", "Có thể thêm tối đa 400 dòng mỗi lần import"),
    ("💾 Lưu file", "Lưu dưới định dạng .xlsx hoặc .xls"),
]

for label, desc in notes:
    lc = ws2.cell(row=row, column=2, value=label)
    lc.font = Font(name='Segoe UI', size=10, bold=True, color=COLOR_ERROR)
    lc.alignment = Alignment(horizontal='left', vertical='top', indent=1)
    
    dc = ws2.cell(row=row, column=3, value=desc)
    dc.font = Font(name='Segoe UI', size=10, color='475569')
    dc.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True, indent=1)
    
    ws2.row_dimensions[row].height = 30
    row += 1

row += 1

# Section 5: Các bước import
add_section(row, "🚀  CÁC BƯỚC IMPORT"); row += 1

steps = [
    "1. Mở file Excel này, điền thông tin user vào sheet 'Users'",
    "2. Xóa các dòng ví dụ (màu vàng) và các dòng trống không dùng",
    "3. Lưu file với định dạng .xlsx",
    "4. Đăng nhập vào web với tài khoản Admin",
    "5. Click vào avatar → chọn 'Quản lý tài khoản'",
    "6. Nhấn nút 'Import' → chọn file vừa lưu",
    "7. Kiểm tra preview → nhấn 'Import X user' để hoàn tất",
]

for step in steps:
    ws2.merge_cells(f'B{row}:C{row}')
    cell = ws2.cell(row=row, column=2, value=step)
    cell.font = Font(name='Segoe UI', size=10, color='0F172A')
    cell.alignment = Alignment(horizontal='left', vertical='center', indent=2)
    ws2.row_dimensions[row].height = 24
    row += 1

# ============================================================
# LƯU FILE
# ============================================================
os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
wb.save(OUTPUT_FILE)

size_kb = os.path.getsize(OUTPUT_FILE) / 1024

print(f"\n🎉 Đã tạo file Excel mẫu thành công!")
print(f"📁 File: {OUTPUT_FILE}")
print(f"📦 Kích thước: {size_kb:.1f} KB")
print(f"\n📋 Nội dung:")
print(f"   📄 Sheet 1 - 'Users': Bảng nhập user")
print(f"      • 6 cột: STT (tự động) | email | name | phone | expiresAt | Ghi chú")
print(f"      • 3 dòng ví dụ + 50 dòng trống sẵn sàng điền")
print(f"   📄 Sheet 2 - 'Hướng dẫn': Hướng dẫn chi tiết tiếng Việt")
print(f"\n✨ Tính năng chuyên nghiệp:")
print(f"   ✅ STT tự động (công thức Excel)")
print(f"   ✅ Validation email (≥ 5 ký tự)")
print(f"   ✅ Validation ngày (YYYY-MM-DD, sau 2020)")
print(f"   ✅ Conditional formatting (hết hạn → đỏ, sắp hết → vàng)")
print(f"   ✅ Freeze panes + Auto filter")
print(f"   ✅ Số điện thoại format text (giữ số 0 đầu)")
print(f"   ✅ Màu sắc phân biệt cột bắt buộc và tùy chọn")
