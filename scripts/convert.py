# -*- coding: utf-8 -*-
"""
Chuyển file Excel → HTML tự chứa dữ liệu
- Font chữ Trung tối ưu: PingFang SC (iOS) / Microsoft YaHei (Win) / Noto Sans SC
- Font-weight 500 thanh mảnh, dễ đọc
- 3 nút toggle (FAB) cố định góc dưới phải
- Nhấn vào câu → phóng to CHỦ ĐIỂM + TIẾNG TRUNG (GIỮ NGUYÊN MÀU)
- Pinyin dark mode dễ đọc
- Header sticky + Search + Filters luôn dính trên cùng
- Phát âm bằng Web Speech API
"""
import openpyxl
import json
import os
import sys

# ====== CẤU HÌNH ======
EXCEL_FILE = "data/input.xlsx"
OUTPUT_HTML = "index.html"
SHEET_INDEX = 0

# ====== ĐỌC EXCEL ======
print(f"📖 Đang đọc file: {EXCEL_FILE}")
if not os.path.exists(EXCEL_FILE):
    print(f"❌ Không tìm thấy file {EXCEL_FILE}")
    sys.exit(1)

wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
ws = wb.worksheets[SHEET_INDEX]
print(f"📊 Sheet: {ws.title} - {ws.max_row} dòng")

def clean(s):
    if s is None:
        return ""
    return str(s).replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').replace('\\', '\\\\')

data = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if not row or len(row) < 6:
        continue
    stt = row[0] if row[0] is not None else ""
    hsk = clean(row[1]) if row[1] else ""
    topic = clean(row[2]) if row[2] else ""
    subject = clean(row[3]) if row[3] else ""
    vi = clean(row[4]) if row[4] else ""
    zh = clean(row[5]) if row[5] else ""
    pinyin = clean(row[6]) if len(row) > 6 and row[6] else ""
    if not vi and not zh:
        continue
    data.append({
        "stt": str(stt),
        "hsk": hsk,
        "topic": topic,
        "subject": subject,
        "vi": vi,
        "zh": zh,
        "pinyin": pinyin
    })

print(f"✅ Đã đọc {len(data)} câu")

# ====== CHUYỂN SANG JSON AN TOÀN ======
json_data = json.dumps(data, ensure_ascii=True, separators=(',', ':'))
json_data = json_data.replace('</', '<\\/')

# ====== TEMPLATE HTML ======
html_template = r'''<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, viewport-fit=cover">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<title>Học tiếng Trung · VP & CX</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
<style>
/* ========== RESET & THEME ========== */
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
:root{
    --bg:#f0f4f8;
    --surface:#ffffff;
    --surface-2:#f8fafc;
    --border:#e2e8f0;
    --border-strong:#cbd5e1;
    --text:#0f172a;
    --text-2:#475569;
    --text-3:#94a3b8;
    --primary:#2563eb;
    --primary-dark:#1d4ed8;
    --primary-light:#dbeafe;
    --success:#16a34a;
    --danger:#dc2626;
    --danger-light:#fee2e2;
    --amber:#f59e0b;
    --amber-light:#fef3c7;
    --shadow-sm:0 1px 2px rgba(15,23,42,.04);
    --shadow:0 4px 12px rgba(15,23,42,.06);
    --shadow-fab:0 8px 24px rgba(15,23,42,.18);
    --radius:14px;
    --radius-sm:10px;
    --radius-full:999px;
    --font-zh: 'PingFang SC','Hiragino Sans GB','Microsoft YaHei','Noto Sans SC','Source Han Sans SC',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
}
[data-theme="dark"]{
    --bg:#0f172a;
    --surface:#1e293b;
    --surface-2:#334155;
    --border:#334155;
    --border-strong:#475569;
    --text:#f1f5f9;
    --text-2:#cbd5e1;
    --text-3:#94a3b8;
    --primary:#3b82f6;
    --primary-dark:#2563eb;
    --primary-light:#1e3a8a;
    --success-light:#14532d;
    --danger-light:#7f1d1d;
    --amber-light:#78350f;
    --shadow-sm:0 1px 2px rgba(0,0,0,.3);
    --shadow:0 4px 12px rgba(0,0,0,.3);
    --shadow-fab:0 8px 24px rgba(0,0,0,.5);
}
html,body{height:100%}
body{
    font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',sans-serif;
    background:var(--bg);
    color:var(--text);
    line-height:1.5;
    font-size:15px;
    padding-bottom:calc(90px + env(safe-area-inset-bottom));
    transition:background .2s,color .2s;
}
.container{max-width:1400px;margin:0 auto;padding:0 1rem}

/* ========== STICKY TOP ========== */
.sticky-top{
    position:sticky;
    top:0;
    z-index:100;
    background:var(--bg);
    padding:.5rem 0 .6rem 0;
    transition:background .2s, box-shadow .2s, border-color .2s;
    border-bottom:1px solid transparent;
}
.sticky-top.scrolled{
    background:var(--surface);
    border-bottom-color:var(--border);
    box-shadow:0 4px 16px -8px rgba(15,23,42,.15);
}
[data-theme="dark"] .sticky-top.scrolled{
    box-shadow:0 4px 16px -8px rgba(0,0,0,.5);
}

/* ========== HEADER ========== */
.header{background:transparent;border:none}
.header-inner{
    display:flex;
    align-items:center;
    gap:.5rem;
    margin-bottom:.5rem;
}
.logo{
    display:flex;
    align-items:center;
    gap:.5rem;
    font-weight:700;
    flex:1;
    min-width:0;
}
.logo-icon{
    width:34px;height:34px;
    background:linear-gradient(135deg,#2563eb,#7c3aed);
    border-radius:10px;
    display:flex;
    align-items:center;
    justify-content:center;
    color:#fff;
    font-size:.95rem;
    flex-shrink:0;
    box-shadow:0 4px 10px rgba(37,99,235,.25);
}
.logo-text{
    overflow:hidden;
    text-overflow:ellipsis;
    white-space:nowrap;
    display:flex;
    flex-direction:column;
    line-height:1.15;
}
.logo-text .title{font-size:.9rem;font-weight:700;color:var(--text)}
.logo-text .subtitle{font-size:.68rem;color:var(--text-3);font-weight:500}
.header-actions{display:flex;gap:.35rem;align-items:center;flex-shrink:0}

.icon-btn{
    width:34px;height:34px;
    border-radius:8px;
    border:1px solid var(--border);
    background:var(--surface);
    color:var(--text-3);
    cursor:pointer;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:.85rem;
    transition:.15s;
    -webkit-tap-highlight-color:transparent;
    position:relative;
    flex-shrink:0;
}
.icon-btn:hover,.icon-btn:active{
    background:var(--surface-2);
    color:var(--primary);
    border-color:var(--primary);
}
.icon-btn.hidden{display:none}
.icon-btn.reset-btn:hover,.icon-btn.reset-btn:active{
    background:var(--danger-light);
    color:var(--danger);
    border-color:var(--danger);
}
.icon-btn .badge{
    position:absolute;
    top:-4px;
    right:-4px;
    min-width:16px;
    height:16px;
    border-radius:50%;
    background:var(--danger);
    color:#fff;
    font-size:.6rem;
    font-weight:700;
    display:flex;
    align-items:center;
    justify-content:center;
    padding:0 4px;
    border:2px solid var(--surface);
}
.icon-btn:not(.has-badge) .badge{display:none}

/* ========== SEARCH ========== */
.search-bar{position:relative;margin-bottom:.55rem}
.search-bar i.fa-search{
    position:absolute;
    left:14px;
    top:50%;
    transform:translateY(-50%);
    color:var(--text-3);
    font-size:.9rem;
    pointer-events:none;
}
.search-bar input{
    width:100%;
    padding:.7rem 2.6rem .7rem 2.5rem;
    border-radius:var(--radius-full);
    border:1.5px solid var(--border);
    background:var(--surface);
    color:var(--text);
    font-size:.92rem;
    outline:none;
    transition:.15s;
    box-shadow:var(--shadow-sm);
    -webkit-appearance:none;
    font-family:inherit;
}
.search-bar input:focus{
    border-color:var(--primary);
    box-shadow:0 0 0 4px rgba(37,99,235,.15);
}
.search-bar input::placeholder{color:var(--text-3)}
.search-clear{
    position:absolute;
    right:8px;
    top:50%;
    transform:translateY(-50%);
    width:30px;height:30px;
    border-radius:50%;
    border:none;
    background:var(--surface-2);
    color:var(--text-2);
    cursor:pointer;
    display:none;
    align-items:center;
    justify-content:center;
    font-size:.8rem;
}
.search-clear.show{display:flex}

/* ========== FILTERS ========== */
.filters{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:.5rem;
    max-width:600px;
}
.chip{
    display:flex;
    align-items:center;
    gap:.4rem;
    padding:.55rem .95rem;
    border-radius:var(--radius-full);
    border:1.5px solid var(--border);
    background:var(--surface);
    color:var(--text);
    font-size:.85rem;
    font-weight:500;
    cursor:pointer;
    transition:.15s;
    outline:none;
    font-family:inherit;
    min-width:0;
    box-shadow:var(--shadow-sm);
    -webkit-appearance:none;
    text-align:left;
    position:relative;
    overflow:hidden;
}
.chip:active{transform:scale(.98)}
.chip.has-value{
    background:var(--primary);
    color:#fff;
    border-color:var(--primary);
    box-shadow:0 4px 12px rgba(37,99,235,.3);
}
.chip.has-value .chip-label{color:#fff;opacity:.85}
.chip-label{
    font-size:.7rem;
    color:var(--text-3);
    text-transform:uppercase;
    letter-spacing:.3px;
    font-weight:700;
    flex-shrink:0;
}
.chip-value{
    overflow:hidden;
    text-overflow:ellipsis;
    white-space:nowrap;
    flex:1;
    min-width:0;
    color:inherit;
}
.chip-arrow{color:inherit;opacity:.5;font-size:.7rem;flex-shrink:0}
.chip select{
    position:absolute;
    inset:0;
    opacity:0;
    cursor:pointer;
    font-size:1rem;
    -webkit-appearance:none;
    appearance:none;
    width:100%;
    height:100%;
}

/* ========== FAB GROUP ========== */
.fab-group{
    position:fixed;
    bottom:calc(20px + env(safe-area-inset-bottom));
    right:20px;
    z-index:1000;
    display:flex;
    flex-direction:column;
    gap:.5rem;
    align-items:flex-end;
    pointer-events:none;
}
.fab-group > *{pointer-events:auto}

.fab-btn{
    width:52px;height:52px;
    border-radius:50%;
    border:none;
    background:var(--surface);
    color:var(--text-2);
    cursor:pointer;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:1.15rem;
    box-shadow:var(--shadow-fab);
    transition:transform .2s, background .2s, color .2s;
    -webkit-tap-highlight-color:transparent;
    position:relative;
    border:2px solid var(--surface);
}
.fab-btn:hover,.fab-btn:active{
    transform:scale(1.08);
    background:var(--primary-light);
    color:var(--primary-dark);
}
.fab-btn.active{
    background:var(--primary);
    color:#fff;
    border-color:var(--primary);
    box-shadow:0 8px 24px rgba(37,99,235,.4);
}
.fab-btn.active:hover,.fab-btn.active:active{
    background:var(--primary-dark);
    color:#fff;
}
.fab-btn::after{
    content:attr(data-label);
    position:absolute;
    right:calc(100% + 10px);
    top:50%;
    transform:translateY(-50%);
    background:var(--text);
    color:var(--surface);
    padding:.4rem .7rem;
    border-radius:8px;
    font-size:.75rem;
    font-weight:600;
    white-space:nowrap;
    opacity:0;
    pointer-events:none;
    transition:opacity .15s;
    font-family:inherit;
}
.fab-btn:hover::after{opacity:1}
@media(max-width:768px){.fab-btn::after{display:none}}

.fab-main{
    width:56px;height:56px;
    font-size:1.3rem;
    background:linear-gradient(135deg,#2563eb,#7c3aed);
    color:#fff;
    border:none;
    box-shadow:0 8px 24px rgba(37,99,235,.4);
}
.fab-main:hover,.fab-main:active{
    background:linear-gradient(135deg,#1d4ed8,#6d28d9);
    color:#fff;
    transform:scale(1.08) rotate(15deg);
}
.fab-main i{transition:transform .3s}
.fab-group.open .fab-main i{transform:rotate(180deg)}
.fab-sub{
    opacity:0;
    transform:translateY(10px) scale(.8);
    pointer-events:none !important;
    transition:opacity .2s, transform .25s;
}
.fab-group.open .fab-sub{
    opacity:1;
    transform:translateY(0) scale(1);
    pointer-events:auto !important;
}
.fab-group.open .fab-sub:nth-child(1){transition-delay:.05s}
.fab-group.open .fab-sub:nth-child(2){transition-delay:.1s}
.fab-group.open .fab-sub:nth-child(3){transition-delay:.15s}

/* ========== TOGGLE CLASSES ========== */
body:not(.show-pinyin) .col-pinyin,
body:not(.show-pinyin) .card-pinyin{display:none!important}
body:not(.show-vi) .col-vi,
body:not(.show-vi) .card-vi{display:none!important}
body:not(.show-practice) .col-practice,
body:not(.show-practice) .card-practice{display:none!important}

/* ========== MAIN ========== */
.main{padding:.25rem 0 3rem}

/* ========== DESKTOP TABLE ========== */
.desktop-view{display:block}
.table-card{
    background:var(--surface);
    border-radius:var(--radius);
    box-shadow:var(--shadow);
    border:1px solid var(--border);
    overflow:hidden;
}
.table-scroll{
    overflow-x:auto;
    overflow-y:auto;
    max-height:calc(100vh - 220px);
    -webkit-overflow-scrolling:touch;
}
table{width:100%;border-collapse:collapse;font-size:.85rem}
thead th{
    background:var(--surface-2);
    color:var(--text-2);
    font-weight:700;
    font-size:.7rem;
    text-transform:uppercase;
    letter-spacing:.5px;
    padding:.7rem .8rem;
    text-align:left;
    border-bottom:1.5px solid var(--border);
    position:sticky;
    top:0;
    z-index:10;
    white-space:nowrap;
}
tbody td{
    padding:.65rem .8rem;
    border-bottom:1px solid var(--border);
    vertical-align:middle;
    color:var(--text);
    transition:background .25s, padding .25s;
}
tbody tr:nth-child(even) td{background:var(--surface-2)}
tbody tr:hover td{background:rgba(37,99,235,.06)}
tbody tr:last-child td{border-bottom:none}
.stt{font-weight:700;color:var(--text-3);font-size:.75rem;text-align:center;width:44px}
.hsk-badge{
    display:inline-block;
    padding:.2rem .5rem;
    border-radius:var(--radius-full);
    background:var(--primary-light);
    color:var(--primary-dark);
    font-size:.7rem;
    font-weight:700;
    white-space:nowrap;
}
.topic-cell{
    font-size:.78rem;
    color:var(--text-2);
    font-weight:500;
    transition:font-size .3s cubic-bezier(.34,1.56,.64,1), font-weight .3s;
}
.subject-cell{font-size:.8rem;color:var(--text-2)}
.vi-cell{font-size:.85rem;color:var(--text)}

/* ✅ TIẾNG TRUNG - FONT CHUYÊN CHO CHỮ HÁN */
.zh-cell{
    font-size:1rem;
    font-weight:500;
    color:var(--text);
    font-family:var(--font-zh);
    letter-spacing:.02em;
    line-height:1.55;
    transition:font-size .3s cubic-bezier(.34,1.56,.64,1), font-weight .3s;
}
.pinyin{
    display:inline-block;
    padding:.15rem .45rem;
    background:var(--surface-2);
    color:var(--primary-dark);
    border-radius:6px;
    font-size:.78rem;
    font-style:italic;
    white-space:nowrap;
}
/* ✅ Pinyin dark mode: chữ sáng + nền xanh mờ */
[data-theme="dark"] .pinyin{
    background:rgba(59,130,246,.18);
    color:#93c5fd;
    font-weight:500;
    font-style:italic;
}
.audio-btn{
    width:32px;height:32px;
    border-radius:50%;
    border:none;
    background:var(--primary-light);
    color:var(--primary-dark);
    cursor:pointer;
    display:inline-flex;
    align-items:center;
    justify-content:center;
    font-size:.85rem;
    transition:.15s;
    -webkit-tap-highlight-color:transparent;
}
.audio-btn:hover,.audio-btn:active{background:var(--primary);color:#fff;transform:scale(1.08)}
.audio-btn.speaking{
    background:var(--danger);
    color:#fff;
    animation:pulse 1s infinite;
}
@keyframes pulse{
    0%,100%{box-shadow:0 0 0 0 rgba(220,38,38,.6)}
    50%{box-shadow:0 0 0 10px rgba(220,38,38,0)}
}
.practice-input{
    width:100%;
    min-width:140px;
    padding:.45rem .7rem;
    border-radius:var(--radius-full);
    border:1.5px solid var(--border);
    background:var(--surface);
    color:var(--text);
    font-size:.85rem;
    outline:none;
    transition:.15s;
    font-family:var(--font-zh);
    -webkit-appearance:none;
}
.practice-input:focus{
    border-color:var(--primary);
    box-shadow:0 0 0 3px rgba(37,99,235,.15);
}
.check-cell{font-weight:700;font-size:.78rem;white-space:nowrap;text-align:center;width:80px}
.check-correct{color:var(--success)}
.check-wrong{color:var(--danger)}

/* ========== FOCUS MODE ========== */
tr.focused td{
    background:var(--primary-light) !important;
    padding-top:1.1rem !important;
    padding-bottom:1.1rem !important;
    box-shadow:inset 3px 0 0 var(--primary);
}
[data-theme="dark"] tr.focused td{
    background:rgba(59,130,246,.15) !important;
}
/* ✅ CHỈ PHÓNG TO CHỦ ĐIỂM - GIỮ NGUYÊN MÀU */
tr.focused .topic-cell{
    font-size:1.15rem !important;
    font-weight:700 !important;
}
/* ✅ CHỈ PHÓNG TO TIẾNG TRUNG - FONT HÁN, WEIGHT 500 */
tr.focused .zh-cell{
    font-size:1.75rem !important;
    font-weight:500 !important;
    letter-spacing:.02em;
    line-height:1.55;
}

/* Mobile: card được chọn */
.card.focused{
    transform:scale(1.02);
    box-shadow:0 12px 32px rgba(37,99,235,.2);
    border-color:var(--primary);
    background:linear-gradient(135deg, var(--surface) 0%, rgba(37,99,235,.06) 100%);
}
[data-theme="dark"] .card.focused{
    background:linear-gradient(135deg, var(--surface) 0%, rgba(59,130,246,.15) 100%);
    box-shadow:0 12px 32px rgba(59,130,246,.3);
}
/* ✅ CHỈ PHÓNG TO TIẾNG TRUNG - FONT HÁN, WEIGHT 500 */
.card.focused .card-zh{
    font-size:2rem;
    font-weight:500;
    letter-spacing:.02em;
    line-height:1.5;
    animation:zoomIn .35s cubic-bezier(.34,1.56,.64,1);
}
/* ✅ CHỈ PHÓNG TO TAG CHỦ ĐIỂM - GIỮ NGUYÊN MÀU */
.card.focused .card-tag.topic{
    font-size:.9rem;
    padding:.28rem .7rem;
    font-weight:700;
    animation:zoomIn .35s cubic-bezier(.34,1.56,.64,1);
}

@keyframes zoomIn{
    0%{transform:scale(.85);opacity:.5}
    60%{transform:scale(1.08)}
    100%{transform:scale(1);opacity:1}
}

@keyframes tapPulse{
    0%{box-shadow:0 0 0 0 rgba(37,99,235,.4)}
    70%{box-shadow:0 0 0 14px rgba(37,99,235,0)}
    100%{box-shadow:0 0 0 0 rgba(37,99,235,0)}
}
.card.tapped{
    animation:tapPulse .6s;
}
tr.tapped td{
    animation:tapPulse .6s;
}

.card,.desktop-view tbody tr{
    cursor:pointer;
    user-select:none;
    transition:transform .25s, box-shadow .25s, border-color .25s, background .25s;
}
.card{
    background:var(--surface);
    border-radius:var(--radius);
    border:1px solid var(--border);
    padding:.9rem;
    margin-bottom:.6rem;
    box-shadow:var(--shadow-sm);
    transition:transform .25s cubic-bezier(.34,1.56,.64,1), box-shadow .25s, border-color .25s, background .25s;
}

.practice-input, .audio-btn, .card-practice, .col-practice{
    cursor:auto;
}

/* ========== MOBILE CARDS ========== */
.mobile-view{display:none}
.card-header{
    display:flex;
    align-items:center;
    gap:.4rem;
    margin-bottom:.6rem;
    padding-bottom:.6rem;
    border-bottom:1px dashed var(--border);
}
.card-stt{
    width:26px;height:26px;
    border-radius:50%;
    background:var(--surface-2);
    color:var(--text-3);
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:.7rem;
    font-weight:700;
    flex-shrink:0;
}
.card-meta{
    display:flex;
    gap:.3rem;
    align-items:center;
    flex:1;
    min-width:0;
    flex-wrap:wrap;
}
.card-tag{
    display:inline-block;
    padding:.12rem .45rem;
    border-radius:var(--radius-full);
    background:var(--surface-2);
    color:var(--text-2);
    font-size:.65rem;
    font-weight:600;
    white-space:nowrap;
    transition:font-size .35s cubic-bezier(.34,1.56,.64,1), padding .35s cubic-bezier(.34,1.56,.64,1);
}
.card-tag.hsk{background:var(--primary-light);color:var(--primary-dark)}
.card-tag.topic{background:var(--amber-light);color:#92400e}
[data-theme="dark"] .card-tag.topic{color:#fde68a}
.card-body{margin-bottom:.6rem}
.card-vi{
    font-size:.85rem;
    color:var(--text-2);
    margin-bottom:.35rem;
    line-height:1.4;
}

/* ✅ TIẾNG TRUNG CARD - FONT CHUYÊN CHO CHỮ HÁN */
.card-zh{
    font-size:1.2rem;
    font-weight:500;
    color:var(--text);
    margin-bottom:.4rem;
    line-height:1.5;
    font-family:var(--font-zh);
    letter-spacing:.02em;
    transition:font-size .3s cubic-bezier(.34,1.56,.64,1), font-weight .3s;
}
.card-pinyin{
    font-size:.78rem;
    font-style:italic;
    color:var(--primary-dark);
    background:var(--surface-2);
    padding:.2rem .45rem;
    border-radius:6px;
    display:inline-block;
}
/* ✅ Pinyin card dark mode: chữ sáng + nền xanh mờ */
[data-theme="dark"] .card-pinyin{
    background:rgba(59,130,246,.18);
    color:#93c5fd;
    font-weight:500;
    font-style:italic;
}
.card-practice{
    display:flex;
    align-items:center;
    gap:.4rem;
    padding-top:.6rem;
    border-top:1px dashed var(--border);
}
.card-practice .practice-input{flex:1;min-width:0}
.card-check{
    font-size:.75rem;
    font-weight:700;
    white-space:nowrap;
    min-width:55px;
    text-align:center;
}

/* ========== LOAD MORE ========== */
.load-more{
    display:block;
    width:100%;
    padding:.9rem;
    margin-top:.8rem;
    border-radius:var(--radius);
    border:1.5px dashed var(--border-strong);
    background:var(--surface);
    color:var(--primary);
    font-weight:700;
    font-size:.88rem;
    cursor:pointer;
    transition:.15s;
    font-family:inherit;
}
.load-more:hover,.load-more:active{background:var(--primary-light);border-color:var(--primary)}
.end-note{
    text-align:center;
    padding:1rem;
    color:var(--text-3);
    font-size:.82rem;
}
.end-note i{color:var(--success);margin-right:.35rem}
.no-data{
    text-align:center;
    padding:3rem 1rem;
    color:var(--text-3);
    background:var(--surface);
    border-radius:var(--radius);
    border:1px solid var(--border);
}
.no-data i{font-size:2.5rem;margin-bottom:.75rem;color:var(--border-strong);display:block}
.error-box{
    text-align:center;
    padding:2rem 1rem;
    color:var(--danger);
    background:var(--danger-light);
    border-radius:var(--radius);
    border:1px solid rgba(220,38,38,.3);
}
.error-box i{font-size:2rem;margin-bottom:.5rem;display:block}

/* ========== RESPONSIVE ========== */
@media(max-width:768px){
    .desktop-view{display:none}
    .mobile-view{display:block}
    
    .container{padding:0 .7rem}
    .header-inner{gap:.35rem;margin-bottom:.5rem}
    .logo-icon{width:30px;height:30px;font-size:.85rem}
    .logo-text .title{font-size:.82rem}
    .logo-text .subtitle{font-size:.6rem}
    .icon-btn{width:32px;height:32px;font-size:.8rem}
    
    .main{padding:.15rem 0 2rem}
    .search-bar input{padding:.65rem 2.5rem .65rem 2.4rem;font-size:.88rem}
    
    .filters{max-width:100%}
    .chip{padding:.5rem .75rem;font-size:.8rem}
    .chip-label{font-size:.65rem}
    
    .fab-group{
        right:16px;
        bottom:calc(16px + env(safe-area-inset-bottom));
        gap:.45rem;
    }
    .fab-main{width:52px;height:52px;font-size:1.2rem}
    .fab-btn{width:48px;height:48px;font-size:1.05rem}
    
    .card.focused .card-zh{
        font-size:2.2rem;
    }
}
@media(max-width:400px){
    .logo-text .subtitle{display:none}
    .icon-btn{width:30px;height:30px;font-size:.75rem}
}
</style>
</head>
<body>

<!-- ============ STICKY TOP ============ -->
<div class="sticky-top" id="stickyTop">
    <div class="container">

        <header class="header">
            <div class="header-inner">
                <div class="logo">
                    <div class="logo-icon"><i class="fas fa-language"></i></div>
                    <div class="logo-text">
                        <div class="title">Học tiếng Trung</div>
                        <div class="subtitle">Văn phòng & Công xưởng</div>
                    </div>
                </div>
                <div class="header-actions">
                    <button class="icon-btn reset-btn hidden" id="resetBtn" title="Đặt lại bộ lọc">
                        <i class="fas fa-undo-alt"></i>
                        <span class="badge" id="resetBadge">0</span>
                    </button>
                    <button class="icon-btn" id="themeToggle" title="Đổi giao diện">
                        <i class="fas fa-moon"></i>
                    </button>
                </div>
            </div>
        </header>

        <div class="search-bar">
            <i class="fas fa-search"></i>
            <input type="text" id="searchInput" placeholder="Tìm kiếm..." autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">
            <button class="search-clear" id="clearSearchBtn" aria-label="Xóa">
                <i class="fas fa-times"></i>
            </button>
        </div>

        <div class="filters">
            <div class="chip" id="hskChip">
                <span class="chip-label">HSK</span>
                <span class="chip-value" id="hskValue">Tất cả</span>
                <i class="fas fa-chevron-down chip-arrow"></i>
                <select id="hskFilter">
                    <option value="">Tất cả</option>
                    <option value="HSK1">HSK1</option>
                    <option value="HSK2">HSK2</option>
                    <option value="HSK3">HSK3</option>
                    <option value="HSK4">HSK4</option>
                    <option value="HSK5">HSK5</option>
                    <option value="HSK6">HSK6</option>
                </select>
            </div>
            <div class="chip" id="subjectChip">
                <span class="chip-label">Chủ đề</span>
                <span class="chip-value" id="subjectValue">Tất cả</span>
                <i class="fas fa-chevron-down chip-arrow"></i>
                <select id="subjectFilter"><option value="">Tất cả chủ đề</option></select>
            </div>
        </div>

    </div>
</div>

<!-- ============ FAB GROUP ============ -->
<div class="fab-group" id="fabGroup">
    <button class="fab-btn fab-sub" id="toggleViBtn" data-label="Tiếng Việt" title="Ẩn/hiện Tiếng Việt">
        <i class="fas fa-language"></i>
    </button>
    <button class="fab-btn fab-sub" id="togglePinyinBtn" data-label="Pinyin" title="Ẩn/hiện Pinyin">
        <i class="fas fa-spell-check"></i>
    </button>
    <button class="fab-btn fab-sub" id="togglePracticeBtn" data-label="Ô nhập tiếng Trung" title="Ẩn/hiện Ô nhập">
        <i class="fas fa-keyboard"></i>
    </button>
    <button class="fab-btn fab-main" id="fabMainBtn" title="Tùy chọn hiển thị">
        <i class="fas fa-sliders-h"></i>
    </button>
</div>

<!-- ============ MAIN ============ -->
<main class="main">
    <div class="container">
        <div class="desktop-view">
            <div class="table-card">
                <div class="table-scroll" id="desktopWrapper">
                    <div class="no-data"><i class="fas fa-spinner fa-pulse"></i>Đang tải dữ liệu...</div>
                </div>
            </div>
        </div>
        <div class="mobile-view" id="mobileWrapper"></div>
    </div>
</main>

<script>
/* ========== DỮ LIỆU ========== */
var RAW_DATA = __DATA__;

if (!Array.isArray(RAW_DATA)) {
    document.getElementById('desktopWrapper').innerHTML =
        '<div class="error-box"><i class="fas fa-exclamation-triangle"></i>Dữ liệu không hợp lệ.</div>';
} else {
    console.log('✅ Đã load', RAW_DATA.length, 'câu');
}

var filtered = RAW_DATA.slice();
var state = { search:'', hsk:'', subject:'' };
var PAGE_SIZE = 300;
var renderedCount = 0;
var focusedStt = null;

var $ = function(id) { return document.getElementById(id); };
var desktopWrapper = $('desktopWrapper');
var mobileWrapper = $('mobileWrapper');

/* ========== SCROLL DETECTION ========== */
(function initScroll() {
    var sticky = document.getElementById('stickyTop');
    if (!sticky) return;
    var ticking = false;
    function update() {
        if (window.scrollY > 5) sticky.classList.add('scrolled');
        else sticky.classList.remove('scrolled');
        ticking = false;
    }
    window.addEventListener('scroll', function() {
        if (!ticking) {
            window.requestAnimationFrame(update);
            ticking = true;
        }
    }, { passive: true });
    update();
})();

/* ========== FAB GROUP ========== */
var fabGroup = $('fabGroup');
var fabMainBtn = $('fabMainBtn');
var fabOpen = false;
try {
    var savedFab = localStorage.getItem('fabOpen');
    if (savedFab === 'true') fabOpen = true;
} catch(e) {}
if (fabOpen) fabGroup.classList.add('open');

fabMainBtn.addEventListener('click', function(e) {
    e.stopPropagation();
    fabOpen = !fabOpen;
    fabGroup.classList.toggle('open', fabOpen);
    try { localStorage.setItem('fabOpen', fabOpen ? 'true' : 'false'); } catch(e) {}
});

document.addEventListener('click', function(e) {
    if (!fabGroup.contains(e.target) && fabOpen && window.innerWidth > 768) {
        fabOpen = false;
        fabGroup.classList.remove('open');
        try { localStorage.setItem('fabOpen', 'false'); } catch(e) {}
    }
});

/* ========== DARK MODE ========== */
(function initTheme() {
    try {
        var saved = localStorage.getItem('theme');
        if (saved) document.documentElement.setAttribute('data-theme', saved);
        else if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            document.documentElement.setAttribute('data-theme', 'dark');
        }
    } catch(e) {}
    updateThemeIcon();
})();

function updateThemeIcon() {
    var isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    var icon = $('themeToggle').querySelector('i');
    icon.className = isDark ? 'fas fa-sun' : 'fas fa-moon';
}

$('themeToggle').addEventListener('click', function() {
    var isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    var newTheme = isDark ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', newTheme);
    try { localStorage.setItem('theme', newTheme); } catch(e) {}
    updateThemeIcon();
});

/* ========== TOGGLE DISPLAY ========== */
var displayState = { vi: false, pinyin: false, practice: false };

(function initDisplay() {
    try {
        var saved = localStorage.getItem('displayState');
        if (saved) {
            var parsed = JSON.parse(saved);
            displayState.vi = !!parsed.vi;
            displayState.pinyin = !!parsed.pinyin;
            displayState.practice = !!parsed.practice;
        }
    } catch(e) {}
    if (displayState.practice) displayState.pinyin = false;
    applyDisplayState();
    updateToggleButtons();
})();

function applyDisplayState() {
    document.body.classList.toggle('show-vi', displayState.vi);
    document.body.classList.toggle('show-pinyin', displayState.pinyin);
    document.body.classList.toggle('show-practice', displayState.practice);
}

function saveDisplayState() {
    try { localStorage.setItem('displayState', JSON.stringify(displayState)); } catch(e) {}
}

function updateToggleButtons() {
    $('toggleViBtn').classList.toggle('active', displayState.vi);
    $('togglePinyinBtn').classList.toggle('active', displayState.pinyin);
    $('togglePracticeBtn').classList.toggle('active', displayState.practice);
}

$('toggleViBtn').addEventListener('click', function(e) {
    e.stopPropagation();
    displayState.vi = !displayState.vi;
    applyDisplayState();
    saveDisplayState();
    updateToggleButtons();
});

$('togglePinyinBtn').addEventListener('click', function(e) {
    e.stopPropagation();
    displayState.pinyin = !displayState.pinyin;
    if (displayState.pinyin && displayState.practice) displayState.practice = false;
    applyDisplayState();
    saveDisplayState();
    updateToggleButtons();
});

$('togglePracticeBtn').addEventListener('click', function(e) {
    e.stopPropagation();
    displayState.practice = !displayState.practice;
    if (displayState.practice && displayState.pinyin) displayState.pinyin = false;
    applyDisplayState();
    saveDisplayState();
    updateToggleButtons();
});

/* ========== FOCUS MODE ========== */
window.toggleFocus = function(stt, element) {
    if (focusedStt === stt) {
        clearFocus();
        return;
    }
    clearFocus();
    focusedStt = stt;
    element.classList.add('focused', 'tapped');
    setTimeout(function() {
        if (element) element.classList.remove('tapped');
    }, 600);
};

window.clearFocus = function() {
    document.querySelectorAll('.focused').forEach(function(el) {
        el.classList.remove('focused', 'tapped');
    });
    focusedStt = null;
};

document.addEventListener('click', function(e) {
    if (e.target.closest('.card') || e.target.closest('.desktop-view tbody tr')) {
        return;
    }
    if (e.target.closest('.practice-input') || 
        e.target.closest('.audio-btn') || 
        e.target.closest('.chip') || 
        e.target.closest('.fab-group') || 
        e.target.closest('.icon-btn') ||
        e.target.closest('.search-bar') ||
        e.target.closest('.filters')) {
        return;
    }
    clearFocus();
}, true);

/* ========== WEB SPEECH API ========== */
var currentBtn = null;

function getChineseVoice() {
    if (!('speechSynthesis' in window)) return null;
    var voices = speechSynthesis.getVoices();
    if (!voices.length) return null;
    var priorities = [
        function(v){ return v.lang === 'zh-CN' && /Ting-?Ting/i.test(v.name); },
        function(v){ return v.lang === 'zh-CN' && /Siri/i.test(v.name); },
        function(v){ return v.lang === 'zh-CN' && v.localService; },
        function(v){ return v.lang === 'zh-CN'; },
        function(v){ return v.lang === 'zh-TW'; },
        function(v){ return v.lang && v.lang.indexOf('zh') === 0; }
    ];
    for (var i = 0; i < priorities.length; i++) {
        var found = voices.find(priorities[i]);
        if (found) return found;
    }
    return null;
}

if ('speechSynthesis' in window) {
    speechSynthesis.getVoices();
    if (speechSynthesis.onvoiceschanged !== undefined) {
        speechSynthesis.onvoiceschanged = function(){ getChineseVoice(); };
    }
}

window.speakText = function(text, btn, evt) {
    if (evt) evt.stopPropagation();
    if (!('speechSynthesis' in window)) {
        alert('Trình duyệt không hỗ trợ phát âm.');
        return;
    }
    speechSynthesis.cancel();
    if (currentBtn) currentBtn.classList.remove('speaking');
    if (btn) {
        btn.classList.add('speaking');
        currentBtn = btn;
    }
    var utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = 'zh-CN';
    utterance.rate = 0.85;
    utterance.pitch = 1.0;
    utterance.volume = 1.0;
    var voice = getChineseVoice();
    if (voice) utterance.voice = voice;
    utterance.onend = utterance.onerror = function() {
        if (currentBtn) {
            currentBtn.classList.remove('speaking');
            currentBtn = null;
        }
    };
    setTimeout(function(){ speechSynthesis.speak(utterance); }, 50);
};

document.addEventListener('visibilitychange', function() {
    if (document.hidden && 'speechSynthesis' in window) {
        speechSynthesis.cancel();
        if (currentBtn) { currentBtn.classList.remove('speaking'); currentBtn = null; }
    }
});

/* ========== ESCAPE ========== */
function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function escapeJs(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '\\"')
        .replace(/\n/g, '\\n').replace(/\r/g, '');
}

/* ========== BUILD FILTERS ========== */
function buildFilters() {
    var subjectSet = {};
    RAW_DATA.forEach(function(r) {
        if (r.subject) subjectSet[r.subject] = 1;
    });
    $('subjectFilter').innerHTML = '<option value="">Tất cả chủ đề</option>' +
        Object.keys(subjectSet).sort().map(function(v){ return '<option value="'+escapeHtml(v)+'">'+escapeHtml(v)+'</option>'; }).join('');
}

/* ========== UPDATE FILTER UI ========== */
function updateFilterUI() {
    var hsk = $('hskFilter').value;
    var subject = $('subjectFilter').value;
    
    $('hskValue').textContent = hsk || 'Tất cả';
    $('subjectValue').textContent = subject || 'Tất cả';
    
    $('hskChip').classList.toggle('has-value', !!hsk);
    $('subjectChip').classList.toggle('has-value', !!subject);
    
    var count = 0;
    if (state.search) count++;
    if (hsk) count++;
    if (subject) count++;
    
    var resetBtn = $('resetBtn');
    var badge = $('resetBadge');
    
    if (count > 0) {
        resetBtn.classList.remove('hidden');
        resetBtn.classList.add('has-badge');
        badge.textContent = count;
    } else {
        resetBtn.classList.add('hidden');
    }
}

/* ========== RENDER ========== */
function render(reset) {
    if (reset) {
        renderedCount = 0;
        focusedStt = null;
    }
    
    if (!filtered.length) {
        var html = '<div class="no-data"><i class="fas fa-search"></i>Không tìm thấy câu nào phù hợp</div>';
        desktopWrapper.innerHTML = html;
        mobileWrapper.innerHTML = html;
        return;
    }
    
    if (reset) {
        desktopWrapper.innerHTML = '<table><thead><tr>' +
            '<th>STT</th><th>HSK</th><th>Chủ điểm</th><th>Chủ đề</th>' +
            '<th class="col-vi">Tiếng Việt</th>' +
            '<th>Tiếng Trung</th>' +
            '<th class="col-pinyin">Pinyin</th>' +
            '<th></th>' +
            '<th class="col-practice">Luyện tập</th>' +
            '<th>Check</th>' +
            '</tr></thead><tbody id="desktopBody"></tbody></table>';
        mobileWrapper.innerHTML = '';
    }
    
    var desktopBody = $('desktopBody');
    var end = Math.min(renderedCount + PAGE_SIZE, filtered.length);
    var deskHtml = '';
    var mobHtml = '';
    
    for (var i = renderedCount; i < end; i++) {
        var r = filtered[i];
        var zhJs = escapeJs(r.zh);
        var zhHtml = escapeHtml(r.zh);
        var sttSafe = escapeHtml(r.stt);
        var sttJs = escapeJs(r.stt);
        var audio = r.zh
            ? '<button class="audio-btn" onclick="speakText(\'' + zhJs + '\', this, event)" title="Nghe"><i class="fas fa-volume-up"></i></button>'
            : '';
        
        deskHtml += '<tr onclick="toggleFocus(\'' + sttJs + '\', this)" data-stt="' + sttSafe + '">' +
            '<td class="stt">' + sttSafe + '</td>' +
            '<td><span class="hsk-badge">' + escapeHtml(r.hsk) + '</span></td>' +
            '<td class="topic-cell">' + escapeHtml(r.topic) + '</td>' +
            '<td class="subject-cell">' + escapeHtml(r.subject) + '</td>' +
            '<td class="vi-cell col-vi">' + escapeHtml(r.vi) + '</td>' +
            '<td class="zh-cell">' + zhHtml + '</td>' +
            '<td class="col-pinyin"><span class="pinyin">' + escapeHtml(r.pinyin) + '</span></td>' +
            '<td style="text-align:center" onclick="event.stopPropagation()">' + audio + '</td>' +
            '<td class="col-practice" onclick="event.stopPropagation()"><input type="text" class="practice-input" placeholder="Nhập tiếng Trung..." data-answer="' + zhHtml + '" data-stt="' + sttSafe + '" oninput="checkInput(this)" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false"></td>' +
            '<td class="check-cell" data-check-stt="' + sttSafe + '"></td>' +
            '</tr>';
        
        mobHtml += '<div class="card" onclick="toggleFocus(\'' + sttJs + '\', this)" data-stt="' + sttSafe + '">' +
            '<div class="card-header">' +
                '<div class="card-stt">' + sttSafe + '</div>' +
                '<div class="card-meta">' +
                    (r.hsk ? '<span class="card-tag hsk">' + escapeHtml(r.hsk) + '</span>' : '') +
                    (r.topic ? '<span class="card-tag topic">' + escapeHtml(r.topic) + '</span>' : '') +
                    (r.subject ? '<span class="card-tag">' + escapeHtml(r.subject) + '</span>' : '') +
                '</div>' +
                '<div onclick="event.stopPropagation()">' + audio + '</div>' +
            '</div>' +
            '<div class="card-body">' +
                (r.vi ? '<div class="card-vi">' + escapeHtml(r.vi) + '</div>' : '') +
                '<div class="card-zh">' + zhHtml + '</div>' +
                (r.pinyin ? '<div class="card-pinyin">' + escapeHtml(r.pinyin) + '</div>' : '') +
            '</div>' +
            '<div class="card-practice" onclick="event.stopPropagation()">' +
                '<input type="text" class="practice-input" placeholder="Nhập tiếng Trung..." data-answer="' + zhHtml + '" data-stt="' + sttSafe + '" oninput="checkInput(this)" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">' +
                '<div class="card-check" data-check-stt="' + sttSafe + '"></div>' +
            '</div>' +
            '</div>';
    }
    
    if (desktopBody) desktopBody.insertAdjacentHTML('beforeend', deskHtml);
    mobileWrapper.insertAdjacentHTML('beforeend', mobHtml);
    renderedCount = end;
    
    var oldDesktopBtn = desktopWrapper.parentElement.querySelector('.load-more');
    if (oldDesktopBtn) oldDesktopBtn.remove();
    var oldMobileBtn = mobileWrapper.querySelector('.load-more');
    if (oldMobileBtn) oldMobileBtn.remove();
    var oldDesktopEnd = desktopWrapper.parentElement.querySelector('.end-note');
    if (oldDesktopEnd) oldDesktopEnd.remove();
    var oldMobileEnd = mobileWrapper.querySelector('.end-note');
    if (oldMobileEnd) oldMobileEnd.remove();
    
    if (renderedCount < filtered.length) {
        var btnDesktop = document.createElement('button');
        btnDesktop.className = 'load-more';
        btnDesktop.innerHTML = '<i class="fas fa-chevron-down"></i> Xem thêm (' + renderedCount + '/' + filtered.length + ')';
        btnDesktop.onclick = function() { render(false); };
        desktopWrapper.parentElement.appendChild(btnDesktop);
        
        var btnMobile = document.createElement('button');
        btnMobile.className = 'load-more';
        btnMobile.innerHTML = '<i class="fas fa-chevron-down"></i> Xem thêm (' + renderedCount + '/' + filtered.length + ')';
        btnMobile.onclick = function() { render(false); };
        mobileWrapper.appendChild(btnMobile);
    } else if (filtered.length > PAGE_SIZE) {
        var endNote = '<div class="end-note"><i class="fas fa-check-circle"></i>Đã hiển thị tất cả ' + filtered.length + ' câu</div>';
        desktopWrapper.parentElement.insertAdjacentHTML('beforeend', endNote);
        mobileWrapper.insertAdjacentHTML('beforeend', endNote);
    }
}

/* ========== CHECK ========== */
window.checkInput = function(input) {
    var stt = input.dataset.stt;
    var answer = input.dataset.answer;
    var cells = document.querySelectorAll('[data-check-stt="' + stt + '"]');
    var val = input.value.trim();
    var norm = function(s){ return s.replace(/[。，！？、；：""''（）\s.,!?;:'"()\[\]{}]/g, ''); };
    
    var result = '';
    if (val) {
        if (norm(val) === norm(answer)) {
            result = '<span class="check-correct">✅ ĐÚNG</span>';
        } else {
            result = '<span class="check-wrong">❌ SAI</span>';
        }
    }
    cells.forEach(function(c) { c.innerHTML = result; });
};

/* ========== FILTER ========== */
function applyFilter() {
    state.search = $('searchInput').value.trim().toLowerCase();
    state.hsk = $('hskFilter').value;
    state.subject = $('subjectFilter').value;
    
    updateFilterUI();
    
    var clearBtn = $('clearSearchBtn');
    if (state.search) clearBtn.classList.add('show');
    else clearBtn.classList.remove('show');
    
    filtered = RAW_DATA.filter(function(r) {
        if (state.search) {
            var s = state.search;
            var inVi = (r.vi || '').toLowerCase().indexOf(s) !== -1;
            var inZh = (r.zh || '').toLowerCase().indexOf(s) !== -1;
            var inPinyin = (r.pinyin || '').toLowerCase().indexOf(s) !== -1;
            var inTopic = (r.topic || '').toLowerCase().indexOf(s) !== -1;
            var inSubject = (r.subject || '').toLowerCase().indexOf(s) !== -1;
            if (!inVi && !inZh && !inPinyin && !inTopic && !inSubject) return false;
        }
        if (state.hsk && r.hsk !== state.hsk) return false;
        if (state.subject && r.subject !== state.subject) return false;
        return true;
    });
    render(true);
}

/* ========== EVENTS ========== */
$('searchInput').addEventListener('input', applyFilter);
$('hskFilter').addEventListener('change', applyFilter);
$('subjectFilter').addEventListener('change', applyFilter);
$('resetBtn').addEventListener('click', function() {
    $('searchInput').value = '';
    $('hskFilter').value = '';
    $('subjectFilter').value = '';
    applyFilter();
});
$('clearSearchBtn').addEventListener('click', function() {
    $('searchInput').value = '';
    $('searchInput').focus();
    applyFilter();
});

/* ========== KHỞI TẠO ========== */
try {
    buildFilters();
    applyFilter();
} catch (e) {
    console.error('Lỗi khởi tạo:', e);
    desktopWrapper.innerHTML = '<div class="error-box"><i class="fas fa-exclamation-triangle"></i>Lỗi: ' + escapeHtml(e.message) + '</div>';
}
</script>
</body>
</html>'''

# ====== GHI FILE HTML ======
html_output = html_template.replace("__DATA__", json_data)
with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(html_output)

size_kb = os.path.getsize(OUTPUT_HTML) / 1024
print(f"🎉 Đã tạo: {OUTPUT_HTML}")
print(f"📦 Kích thước: {size_kb:.1f} KB")
print(f"📚 Tổng số câu: {len(data)}")
