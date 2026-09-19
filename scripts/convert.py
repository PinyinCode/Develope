# -*- coding: utf-8 -*-
"""
Chuyển file Excel → HTML tự chứa dữ liệu
- Demo mode cho khách chưa đăng nhập (giới hạn tính năng)
- Hệ thống đăng nhập Firebase
- Trang admin quản lý user
- Hanzi Writer, font chữ Trung tối ưu, 3 FAB, focus mode
"""
import openpyxl
import json
import os
import sys

# ====== CẤU HÌNH ======
EXCEL_FILE = "data/input.xlsx"
OUTPUT_HTML = "index.html"
SHEET_INDEX = 0
DEMO_LIMIT = 50
DEMO_AUDIO_LIMIT = 10

# ✅ Firebase config của bạn (đã nhúng sẵn)
FIREBASE_CONFIG = {
    "apiKey": "AIzaSyA8VlFQvE_om-51pxffgd7hd2Ud6hQm11k",
    "authDomain": "hoc-tieng-trung-64b8d.firebaseapp.com",
    "projectId": "hoc-tieng-trung-64b8d",
    "storageBucket": "hoc-tieng-trung-64b8d.firebasestorage.app",
    "messagingSenderId": "614326539390",
    "appId": "1:614326539390:web:57344a021d7cbc7b18b42f"
}

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
        "stt": str(stt), "hsk": hsk, "topic": topic, "subject": subject,
        "vi": vi, "zh": zh, "pinyin": pinyin
    })

print(f"✅ Đã đọc {len(data)} câu")

json_data = json.dumps(data, ensure_ascii=True, separators=(',', ':'))
json_data = json_data.replace('</', '<\\/')
firebase_config_json = json.dumps(FIREBASE_CONFIG, ensure_ascii=False)

# ====== TEMPLATE HTML ======
html_template = r'''<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, viewport-fit=cover">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<title>Học tiếng Trung · VP & CX</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
<script src="https://www.gstatic.com/firebasejs/10.7.0/firebase-app-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.7.0/firebase-auth-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.7.0/firebase-firestore-compat.js"></script>
<script src="https://cdn.jsdelivr.net/npm/hanzi-writer@3.5.0/dist/hanzi-writer.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
:root{
    --bg:#f0f4f8;--surface:#fff;--surface-2:#f8fafc;--border:#e2e8f0;--border-strong:#cbd5e1;
    --text:#0f172a;--text-2:#475569;--text-3:#94a3b8;
    --primary:#2563eb;--primary-dark:#1d4ed8;--primary-light:#dbeafe;
    --success:#16a34a;--danger:#dc2626;--danger-light:#fee2e2;
    --amber:#f59e0b;--amber-light:#fef3c7;
    --shadow-sm:0 1px 2px rgba(15,23,42,.04);--shadow:0 4px 12px rgba(15,23,42,.06);
    --shadow-fab:0 8px 24px rgba(15,23,42,.18);
    --radius:14px;--radius-sm:10px;--radius-full:999px;
    --font-zh:'PingFang SC','Hiragino Sans GB','Microsoft YaHei','Noto Sans SC',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
}
[data-theme="dark"]{
    --bg:#0f172a;--surface:#1e293b;--surface-2:#334155;--border:#334155;--border-strong:#475569;
    --text:#f1f5f9;--text-2:#cbd5e1;--text-3:#94a3b8;
    --primary:#3b82f6;--primary-dark:#2563eb;--primary-light:#1e3a8a;
    --danger-light:#7f1d1d;--amber-light:#78350f;
    --shadow-sm:0 1px 2px rgba(0,0,0,.3);--shadow:0 4px 12px rgba(0,0,0,.3);
    --shadow-fab:0 8px 24px rgba(0,0,0,.5);
}
html,body{height:100%}
body{
    font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',sans-serif;
    background:var(--bg);color:var(--text);line-height:1.5;font-size:15px;
    padding-bottom:calc(90px + env(safe-area-inset-bottom));
    transition:background .2s,color .2s;
}
.container{max-width:1400px;margin:0 auto;padding:0 1rem}

/* ========== LOADING ========== */
.loading-screen{
    position:fixed;inset:0;background:var(--bg);
    display:flex;align-items:center;justify-content:center;
    z-index:9998;flex-direction:column;gap:1rem;color:var(--text-2);
}
.loading-screen.hidden{display:none}
.loading-screen i{font-size:2.5rem;color:var(--primary);animation:spin 1s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}

/* ========== STICKY TOP ========== */
.sticky-top{
    position:sticky;top:0;z-index:100;background:var(--bg);
    padding:.5rem 0 .6rem 0;
    transition:background .2s, box-shadow .2s, border-color .2s;
    border-bottom:1px solid transparent;
}
.sticky-top.scrolled{
    background:var(--surface);border-bottom-color:var(--border);
    box-shadow:0 4px 16px -8px rgba(15,23,42,.15);
}
[data-theme="dark"] .sticky-top.scrolled{box-shadow:0 4px 16px -8px rgba(0,0,0,.5)}

.header{background:transparent;border:none}
.header-inner{display:flex;align-items:center;gap:.5rem;margin-bottom:.5rem}
.logo{display:flex;align-items:center;gap:.5rem;font-weight:700;flex:1;min-width:0}
.logo-icon{
    width:34px;height:34px;background:linear-gradient(135deg,#2563eb,#7c3aed);
    border-radius:10px;display:flex;align-items:center;justify-content:center;
    color:#fff;font-size:.95rem;flex-shrink:0;box-shadow:0 4px 10px rgba(37,99,235,.25);
}
.logo-text{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:flex;flex-direction:column;line-height:1.15}
.logo-text .title{font-size:.9rem;font-weight:700;color:var(--text)}
.logo-text .subtitle{font-size:.68rem;color:var(--text-3);font-weight:500}
.header-actions{display:flex;gap:.35rem;align-items:center;flex-shrink:0}

.icon-btn{
    width:34px;height:34px;border-radius:8px;border:1px solid var(--border);
    background:var(--surface);color:var(--text-3);cursor:pointer;
    display:flex;align-items:center;justify-content:center;font-size:.85rem;
    transition:.15s;position:relative;flex-shrink:0;
}
.icon-btn:hover,.icon-btn:active{background:var(--surface-2);color:var(--primary);border-color:var(--primary)}
.icon-btn.hidden{display:none}
.icon-btn.reset-btn:hover,.icon-btn.reset-btn:active{background:var(--danger-light);color:var(--danger);border-color:var(--danger)}
.icon-btn .badge{
    position:absolute;top:-4px;right:-4px;min-width:16px;height:16px;border-radius:50%;
    background:var(--danger);color:#fff;font-size:.6rem;font-weight:700;
    display:flex;align-items:center;justify-content:center;padding:0 4px;
    border:2px solid var(--surface);
}
.icon-btn:not(.has-badge) .badge{display:none}

/* Demo badge */
.demo-badge{
    display:flex;align-items:center;gap:.35rem;
    padding:.35rem .7rem;border-radius:50px;
    background:var(--amber-light);color:#92400e;
    font-size:.7rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.3px;
    border:1px solid rgba(245,158,11,.4);
}
[data-theme="dark"] .demo-badge{color:#fcd34d}

/* User menu */
.user-menu{position:relative}
.user-avatar{
    width:34px;height:34px;border-radius:50%;border:2px solid var(--border);
    cursor:pointer;object-fit:cover;transition:.15s;display:block;
}
.user-avatar:hover{border-color:var(--primary);transform:scale(1.05)}
.user-dropdown{
    position:absolute;top:calc(100% + .5rem);right:0;
    background:var(--surface);border:1px solid var(--border);
    border-radius:var(--radius);box-shadow:0 10px 30px rgba(0,0,0,.15);
    padding:.5rem;min-width:260px;display:none;z-index:200;
}
.user-dropdown.show{display:block}
.user-info{padding:.75rem;border-bottom:1px solid var(--border);margin-bottom:.5rem}
.user-info .name{font-weight:700;font-size:.9rem;color:var(--text);margin-bottom:.2rem}
.user-info .email{font-size:.75rem;color:var(--text-3);word-break:break-all}
.user-info .role{
    display:inline-block;margin-top:.4rem;padding:.15rem .5rem;
    background:var(--primary-light);color:var(--primary-dark);
    border-radius:50px;font-size:.68rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.3px;
}
.user-info .role.admin{background:var(--amber-light);color:#92400e}
.dropdown-item{
    display:flex;align-items:center;gap:.5rem;width:100%;
    padding:.65rem .75rem;border:none;border-radius:var(--radius-sm);
    background:transparent;color:var(--text);font-size:.85rem;font-weight:600;
    cursor:pointer;transition:.15s;font-family:inherit;text-align:left;
}
.dropdown-item:hover{background:var(--surface-2)}
.dropdown-item.danger{color:var(--danger)}
.dropdown-item.danger:hover{background:var(--danger-light)}

/* Login button */
.btn-login-header{
    display:flex;align-items:center;gap:.4rem;
    padding:.5rem .9rem;border-radius:50px;
    background:var(--primary);color:#fff;border:none;
    font-size:.8rem;font-weight:700;cursor:pointer;
    transition:.15s;font-family:inherit;
    box-shadow:0 4px 12px rgba(37,99,235,.3);
    white-space:nowrap;
}
.btn-login-header:hover,.btn-login-header:active{background:var(--primary-dark);transform:translateY(-1px)}

/* Search */
.search-bar{position:relative;margin-bottom:.55rem}
.search-bar i.fa-search{
    position:absolute;left:14px;top:50%;transform:translateY(-50%);
    color:var(--text-3);font-size:.9rem;pointer-events:none;
}
.search-bar input{
    width:100%;padding:.7rem 2.6rem .7rem 2.5rem;
    border-radius:var(--radius-full);border:1.5px solid var(--border);
    background:var(--surface);color:var(--text);font-size:.92rem;
    outline:none;transition:.15s;box-shadow:var(--shadow-sm);
    -webkit-appearance:none;font-family:inherit;
}
.search-bar input:focus{border-color:var(--primary);box-shadow:0 0 0 4px rgba(37,99,235,.15)}
.search-bar input::placeholder{color:var(--text-3)}
.search-clear{
    position:absolute;right:8px;top:50%;transform:translateY(-50%);
    width:30px;height:30px;border-radius:50%;border:none;
    background:var(--surface-2);color:var(--text-2);
    cursor:pointer;display:none;align-items:center;justify-content:center;
    font-size:.8rem;
}
.search-clear.show{display:flex}

/* Filters */
.filters{display:grid;grid-template-columns:1fr 1fr;gap:.5rem;max-width:600px}
.chip{
    display:flex;align-items:center;gap:.4rem;
    padding:.55rem .95rem;border-radius:var(--radius-full);
    border:1.5px solid var(--border);background:var(--surface);
    color:var(--text);font-size:.85rem;font-weight:500;
    cursor:pointer;transition:.15s;outline:none;font-family:inherit;
    min-width:0;box-shadow:var(--shadow-sm);-webkit-appearance:none;
    text-align:left;position:relative;overflow:hidden;
}
.chip:active{transform:scale(.98)}
.chip.has-value{
    background:var(--primary);color:#fff;border-color:var(--primary);
    box-shadow:0 4px 12px rgba(37,99,235,.3);
}
.chip.has-value .chip-label{color:#fff;opacity:.85}
.chip-label{
    font-size:.7rem;color:var(--text-3);text-transform:uppercase;
    letter-spacing:.3px;font-weight:700;flex-shrink:0;
}
.chip-value{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;min-width:0;color:inherit}
.chip-arrow{color:inherit;opacity:.5;font-size:.7rem;flex-shrink:0}
.chip select{
    position:absolute;inset:0;opacity:0;cursor:pointer;font-size:1rem;
    -webkit-appearance:none;appearance:none;width:100%;height:100%;
}
.chip.locked{opacity:.6;cursor:not-allowed;background:var(--surface-2);border-style:dashed}
.chip.locked::after{
    content:'\f023';
    font-family:'Font Awesome 6 Free';
    font-weight:900;
    position:absolute;right:12px;top:50%;transform:translateY(-50%);
    color:var(--text-3);font-size:.85rem;
    pointer-events:none;
}
.chip.locked select{pointer-events:none}

/* FAB */
.fab-group{
    position:fixed;bottom:calc(20px + env(safe-area-inset-bottom));
    right:20px;z-index:1000;display:flex;flex-direction:column;
    gap:.5rem;align-items:flex-end;pointer-events:none;
}
.fab-group > *{pointer-events:auto}
.fab-btn{
    width:52px;height:52px;border-radius:50%;border:none;
    background:var(--surface);color:var(--text-2);cursor:pointer;
    display:flex;align-items:center;justify-content:center;
    font-size:1.15rem;box-shadow:var(--shadow-fab);
    transition:transform .2s, background .2s, color .2s;
    position:relative;border:2px solid var(--surface);
}
.fab-btn:hover,.fab-btn:active{transform:scale(1.08);background:var(--primary-light);color:var(--primary-dark)}
.fab-btn.active{background:var(--primary);color:#fff;border-color:var(--primary);box-shadow:0 8px 24px rgba(37,99,235,.4)}
.fab-btn.active:hover,.fab-btn.active:active{background:var(--primary-dark);color:#fff}
.fab-btn::after{
    content:attr(data-label);position:absolute;right:calc(100% + 10px);
    top:50%;transform:translateY(-50%);background:var(--text);color:var(--surface);
    padding:.4rem .7rem;border-radius:8px;font-size:.75rem;font-weight:600;
    white-space:nowrap;opacity:0;pointer-events:none;transition:opacity .15s;font-family:inherit;
}
.fab-btn:hover::after{opacity:1}
@media(max-width:768px){.fab-btn::after{display:none}}
.fab-main{
    width:56px;height:56px;font-size:1.3rem;
    background:linear-gradient(135deg,#2563eb,#7c3aed);
    color:#fff;border:none;box-shadow:0 8px 24px rgba(37,99,235,.4);
}
.fab-main:hover,.fab-main:active{
    background:linear-gradient(135deg,#1d4ed8,#6d28d9);
    color:#fff;transform:scale(1.08) rotate(15deg);
}
.fab-main i{transition:transform .3s}
.fab-group.open .fab-main i{transform:rotate(180deg)}
.fab-sub{
    opacity:0;transform:translateY(10px) scale(.8);
    pointer-events:none !important;
    transition:opacity .2s, transform .25s;
}
.fab-group.open .fab-sub{
    opacity:1;transform:translateY(0) scale(1);
    pointer-events:auto !important;
}
.fab-group.open .fab-sub:nth-child(1){transition-delay:.05s}
.fab-group.open .fab-sub:nth-child(2){transition-delay:.1s}
.fab-group.open .fab-sub:nth-child(3){transition-delay:.15s}
.fab-btn.locked{opacity:.5}
.fab-btn.locked::before{
    content:'\f023';
    font-family:'Font Awesome 6 Free';
    font-weight:900;
    position:absolute;top:-4px;right:-4px;
    width:20px;height:20px;border-radius:50%;
    background:var(--amber);color:#fff;
    font-size:.6rem;display:flex;align-items:center;justify-content:center;
    border:2px solid var(--surface);
}

/* Toggle classes */
body:not(.show-pinyin) .col-pinyin,
body:not(.show-pinyin) .card-pinyin{display:none!important}
body:not(.show-vi) .col-vi,
body:not(.show-vi) .card-vi{display:none!important}
body:not(.show-practice) .col-practice,
body:not(.show-practice) .card-practice{display:none!important}

/* Main */
.main{padding:.25rem 0 3rem}

/* Demo banner */
.demo-banner{
    background:linear-gradient(135deg, #fef3c7, #fde68a);
    border:1.5px solid #f59e0b;
    border-radius:var(--radius);
    padding:.9rem 1.1rem;
    margin-bottom:1rem;
    display:flex;align-items:center;gap:.75rem;
    flex-wrap:wrap;
}
[data-theme="dark"] .demo-banner{
    background:linear-gradient(135deg, rgba(245,158,11,.15), rgba(245,158,11,.25));
    border-color:#f59e0b;
}
.demo-banner-icon{
    width:36px;height:36px;border-radius:50%;
    background:var(--amber);color:#fff;
    display:flex;align-items:center;justify-content:center;
    font-size:1rem;flex-shrink:0;
}
.demo-banner-text{flex:1;min-width:200px}
.demo-banner-text .title{font-weight:700;font-size:.9rem;color:#92400e;margin-bottom:.15rem}
[data-theme="dark"] .demo-banner-text .title{color:#fcd34d}
.demo-banner-text .desc{font-size:.78rem;color:#78350f;line-height:1.4}
[data-theme="dark"] .demo-banner-text .desc{color:#fde68a}
.demo-banner-btn{
    padding:.5rem .9rem;border-radius:50px;border:none;
    background:var(--amber);color:#fff;
    font-size:.8rem;font-weight:700;cursor:pointer;
    transition:.15s;font-family:inherit;
    display:flex;align-items:center;gap:.35rem;
    white-space:nowrap;
}
.demo-banner-btn:hover{background:#d97706;transform:translateY(-1px)}

/* Desktop table */
.desktop-view{display:block}
.table-card{
    background:var(--surface);border-radius:var(--radius);
    box-shadow:var(--shadow);border:1px solid var(--border);overflow:hidden;
}
.table-scroll{
    overflow-x:auto;overflow-y:auto;
    max-height:calc(100vh - 220px);-webkit-overflow-scrolling:touch;
}
table{width:100%;border-collapse:collapse;font-size:.85rem}
thead th{
    background:var(--surface-2);color:var(--text-2);font-weight:700;
    font-size:.7rem;text-transform:uppercase;letter-spacing:.5px;
    padding:.7rem .8rem;text-align:left;
    border-bottom:1.5px solid var(--border);
    position:sticky;top:0;z-index:10;white-space:nowrap;
}
tbody td{
    padding:.65rem .8rem;border-bottom:1px solid var(--border);
    vertical-align:middle;color:var(--text);
    transition:background .25s, padding .25s;
}
tbody tr:nth-child(even) td{background:var(--surface-2)}
tbody tr:hover td{background:rgba(37,99,235,.06)}
tbody tr:last-child td{border-bottom:none}
.stt{font-weight:700;color:var(--text-3);font-size:.75rem;text-align:center;width:44px}
.hsk-badge{
    display:inline-block;padding:.2rem .5rem;border-radius:var(--radius-full);
    background:var(--primary-light);color:var(--primary-dark);
    font-size:.7rem;font-weight:700;white-space:nowrap;
}
.topic-cell{font-size:.78rem;color:var(--text-2);font-weight:500;transition:font-size .3s cubic-bezier(.34,1.56,.64,1), font-weight .3s}
.subject-cell{font-size:.8rem;color:var(--text-2)}
.vi-cell{font-size:.85rem;color:var(--text)}
.zh-cell{
    font-size:1rem;font-weight:500;color:var(--text);
    font-family:var(--font-zh);letter-spacing:.02em;line-height:1.55;
    transition:font-size .3s cubic-bezier(.34,1.56,.64,1), font-weight .3s;
}
.pinyin{
    display:inline-block;padding:.15rem .45rem;background:var(--surface-2);
    color:var(--primary-dark);border-radius:6px;font-size:.78rem;
    font-style:italic;white-space:nowrap;
}
[data-theme="dark"] .pinyin{background:rgba(59,130,246,.18);color:#93c5fd;font-weight:500;font-style:italic}
.audio-btn{
    width:32px;height:32px;border-radius:50%;border:none;
    background:var(--primary-light);color:var(--primary-dark);
    cursor:pointer;display:inline-flex;align-items:center;justify-content:center;
    font-size:.85rem;transition:.15s;
}
.audio-btn:hover,.audio-btn:active{background:var(--primary);color:#fff;transform:scale(1.08)}
.audio-btn.speaking{background:var(--danger);color:#fff;animation:pulse 1s infinite}
@keyframes pulse{
    0%,100%{box-shadow:0 0 0 0 rgba(220,38,38,.6)}
    50%{box-shadow:0 0 0 10px rgba(220,38,38,0)}
}
.write-btn{
    width:32px;height:32px;border-radius:50%;border:none;
    background:var(--amber-light);color:#92400e;cursor:pointer;
    display:inline-flex;align-items:center;justify-content:center;
    font-size:.8rem;transition:.15s;
}
.write-btn:hover,.write-btn:active{background:var(--amber);color:#fff;transform:scale(1.08)}
[data-theme="dark"] .write-btn{background:rgba(245,158,11,.25);color:#fcd34d}
[data-theme="dark"] .write-btn:hover{background:var(--amber);color:#fff}
.audio-btn.locked,.write-btn.locked{opacity:.5;cursor:not-allowed}
.audio-btn.locked::after,.write-btn.locked::after{
    content:'\f023';
    font-family:'Font Awesome 6 Free';
    font-weight:900;
    position:absolute;bottom:-2px;right:-2px;
    width:14px;height:14px;border-radius:50%;
    background:var(--amber);color:#fff;
    font-size:.5rem;display:flex;align-items:center;justify-content:center;
}
.action-group{display:flex;gap:.3rem;justify-content:center;align-items:center;position:relative}
.practice-input{
    width:100%;min-width:140px;padding:.45rem .7rem;
    border-radius:var(--radius-full);border:1.5px solid var(--border);
    background:var(--surface);color:var(--text);font-size:.85rem;
    outline:none;transition:.15s;font-family:var(--font-zh);-webkit-appearance:none;
}
.practice-input:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(37,99,235,.15)}
.check-cell{font-weight:700;font-size:.78rem;white-space:nowrap;text-align:center;width:80px}
.check-correct{color:var(--success)}
.check-wrong{color:var(--danger)}

/* Focus mode */
tr.focused td{
    background:var(--primary-light) !important;
    padding-top:1.1rem !important;padding-bottom:1.1rem !important;
    box-shadow:inset 3px 0 0 var(--primary);
}
[data-theme="dark"] tr.focused td{background:rgba(59,130,246,.15) !important}
tr.focused .topic-cell{font-size:1.15rem !important;font-weight:700 !important}
tr.focused .zh-cell{font-size:1.75rem !important;font-weight:500 !important;letter-spacing:.02em;line-height:1.55}
.card.focused{
    transform:scale(1.02);box-shadow:0 12px 32px rgba(37,99,235,.2);
    border-color:var(--primary);
    background:linear-gradient(135deg, var(--surface) 0%, rgba(37,99,235,.06) 100%);
}
[data-theme="dark"] .card.focused{
    background:linear-gradient(135deg, var(--surface) 0%, rgba(59,130,246,.15) 100%);
    box-shadow:0 12px 32px rgba(59,130,246,.3);
}
.card.focused .card-zh{
    font-size:2rem;font-weight:500;letter-spacing:.02em;line-height:1.5;
    animation:zoomIn .35s cubic-bezier(.34,1.56,.64,1);
}
.card.focused .card-tag.topic{
    font-size:.9rem;padding:.28rem .7rem;font-weight:700;
    animation:zoomIn .35s cubic-bezier(.34,1.56,.64,1);
}
@keyframes zoomIn{0%{transform:scale(.85);opacity:.5}60%{transform:scale(1.08)}100%{transform:scale(1);opacity:1}}
@keyframes tapPulse{
    0%{box-shadow:0 0 0 0 rgba(37,99,235,.4)}
    70%{box-shadow:0 0 0 14px rgba(37,99,235,0)}
    100%{box-shadow:0 0 0 0 rgba(37,99,235,0)}
}
.card.tapped{animation:tapPulse .6s}
tr.tapped td{animation:tapPulse .6s}
.card,.desktop-view tbody tr{
    cursor:pointer;user-select:none;
    transition:transform .25s, box-shadow .25s, border-color .25s, background .25s;
}
.card{
    background:var(--surface);border-radius:var(--radius);
    border:1px solid var(--border);padding:.9rem;margin-bottom:.6rem;
    box-shadow:var(--shadow-sm);
    transition:transform .25s cubic-bezier(.34,1.56,.64,1), box-shadow .25s, border-color .25s, background .25s;
}
.practice-input, .audio-btn, .write-btn, .card-practice, .col-practice{cursor:auto}

/* Mobile cards */
.mobile-view{display:none}
.card-header{
    display:flex;align-items:center;gap:.4rem;
    margin-bottom:.6rem;padding-bottom:.6rem;
    border-bottom:1px dashed var(--border);
}
.card-stt{
    width:26px;height:26px;border-radius:50%;
    background:var(--surface-2);color:var(--text-3);
    display:flex;align-items:center;justify-content:center;
    font-size:.7rem;font-weight:700;flex-shrink:0;
}
.card-meta{display:flex;gap:.3rem;align-items:center;flex:1;min-width:0;flex-wrap:wrap}
.card-tag{
    display:inline-block;padding:.12rem .45rem;border-radius:var(--radius-full);
    background:var(--surface-2);color:var(--text-2);
    font-size:.65rem;font-weight:600;white-space:nowrap;
    transition:font-size .35s cubic-bezier(.34,1.56,.64,1), padding .35s cubic-bezier(.34,1.56,.64,1);
}
.card-tag.hsk{background:var(--primary-light);color:var(--primary-dark)}
.card-tag.topic{background:var(--amber-light);color:#92400e}
[data-theme="dark"] .card-tag.topic{color:#fde68a}
.card-body{margin-bottom:.6rem}
.card-vi{font-size:.85rem;color:var(--text-2);margin-bottom:.35rem;line-height:1.4}
.card-zh{
    font-size:1.2rem;font-weight:500;color:var(--text);
    margin-bottom:.4rem;line-height:1.5;
    font-family:var(--font-zh);letter-spacing:.02em;
    transition:font-size .3s cubic-bezier(.34,1.56,.64,1), font-weight .3s;
}
.card-pinyin{
    font-size:.78rem;font-style:italic;color:var(--primary-dark);
    background:var(--surface-2);padding:.2rem .45rem;border-radius:6px;display:inline-block;
}
[data-theme="dark"] .card-pinyin{background:rgba(59,130,246,.18);color:#93c5fd;font-weight:500;font-style:italic}
.card-practice{display:flex;align-items:center;gap:.4rem;padding-top:.6rem;border-top:1px dashed var(--border)}
.card-practice .practice-input{flex:1;min-width:0}
.card-check{font-size:.75rem;font-weight:700;white-space:nowrap;min-width:55px;text-align:center}

/* Load more */
.load-more{
    display:block;width:100%;padding:.9rem;margin-top:.8rem;
    border-radius:var(--radius);border:1.5px dashed var(--border-strong);
    background:var(--surface);color:var(--primary);
    font-weight:700;font-size:.88rem;cursor:pointer;
    transition:.15s;font-family:inherit;
}
.load-more:hover,.load-more:active{background:var(--primary-light);border-color:var(--primary)}
.load-more.locked{border-color:var(--amber);color:#92400e;background:var(--amber-light)}
[data-theme="dark"] .load-more.locked{color:#fcd34d;background:rgba(245,158,11,.15)}
.end-note{text-align:center;padding:1rem;color:var(--text-3);font-size:.82rem}
.end-note i{color:var(--success);margin-right:.35rem}
.no-data{
    text-align:center;padding:3rem 1rem;color:var(--text-3);
    background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);
}
.no-data i{font-size:2.5rem;margin-bottom:.75rem;color:var(--border-strong);display:block}
.error-box{
    text-align:center;padding:2rem 1rem;color:var(--danger);
    background:var(--danger-light);border-radius:var(--radius);
    border:1px solid rgba(220,38,38,.3);
}
.error-box i{font-size:2rem;margin-bottom:.5rem;display:block}

/* Login modal */
.login-modal{
    position:fixed;inset:0;
    background:rgba(15,23,42,.8);
    backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);
    z-index:3000;display:none;align-items:center;justify-content:center;
    padding:1.5rem;animation:fadeIn .2s;
}
.login-modal.show{display:flex}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
.login-box{
    background:#fff;border-radius:20px;padding:2.5rem 2rem;
    max-width:440px;width:100%;
    box-shadow:0 20px 60px rgba(0,0,0,.3);
    text-align:center;position:relative;
    animation:slideUp .3s cubic-bezier(.34,1.56,.64,1);
}
@keyframes slideUp{
    from{transform:translateY(30px) scale(.95);opacity:0}
    to{transform:translateY(0) scale(1);opacity:1}
}
.login-close{
    position:absolute;top:12px;right:12px;
    width:34px;height:34px;border-radius:50%;border:none;
    background:#f1f5f9;color:#475569;cursor:pointer;font-size:1rem;
    display:flex;align-items:center;justify-content:center;transition:.15s;
}
.login-close:hover{background:#fee2e2;color:#dc2626}
.login-logo{
    width:70px;height:70px;
    background:linear-gradient(135deg,#2563eb,#7c3aed);
    border-radius:20px;display:flex;align-items:center;justify-content:center;
    color:#fff;font-size:2rem;margin:0 auto 1.5rem;
    box-shadow:0 8px 20px rgba(37,99,235,.35);
}
.login-box h2{font-size:1.4rem;color:#0f172a;margin-bottom:.5rem;font-weight:700}
.login-box p{color:#64748b;font-size:.9rem;margin-bottom:2rem;line-height:1.5}
.btn-google{
    display:flex;align-items:center;justify-content:center;gap:.75rem;
    width:100%;padding:.9rem 1.5rem;border-radius:50px;
    border:2px solid #e2e8f0;background:#fff;color:#0f172a;
    font-size:1rem;font-weight:600;cursor:pointer;transition:.15s;font-family:inherit;
}
.btn-google:hover{border-color:#2563eb;background:#f0f7ff;transform:translateY(-1px);box-shadow:0 4px 12px rgba(37,99,235,.15)}
.btn-google img{width:22px;height:22px}
.login-error{
    background:#fee2e2;color:#dc2626;padding:.85rem 1rem;
    border-radius:10px;font-size:.85rem;margin-top:1rem;
    display:none;text-align:left;line-height:1.4;
}
.login-error.show{display:block}
.login-footer{
    margin-top:1.5rem;padding-top:1.5rem;
    border-top:1px solid #e2e8f0;font-size:.78rem;color:#94a3b8;line-height:1.5;
}

/* Writer modal */
.writer-modal{
    position:fixed;inset:0;background:rgba(15,23,42,.7);
    backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);
    z-index:2000;display:none;align-items:center;justify-content:center;
    padding:1rem;animation:fadeIn .2s;
}
.writer-modal.show{display:flex}
.writer-box{
    background:var(--surface);border-radius:20px;padding:1.5rem 1.25rem;
    max-width:420px;width:100%;max-height:calc(100vh - 2rem);overflow-y:auto;
    box-shadow:0 20px 60px rgba(0,0,0,.3);position:relative;
    animation:slideUp .3s cubic-bezier(.34,1.56,.64,1);
}
.writer-close{
    position:absolute;top:10px;right:10px;
    width:34px;height:34px;border-radius:50%;border:none;
    background:var(--surface-2);color:var(--text-2);cursor:pointer;
    font-size:1rem;display:flex;align-items:center;justify-content:center;
    transition:.15s;z-index:5;
}
.writer-close:hover{background:var(--danger-light);color:var(--danger)}
.writer-char-info{text-align:center;margin-bottom:.75rem}
.writer-char-info .vi-small{font-size:.85rem;color:var(--text-2);margin-bottom:.3rem;line-height:1.4}
.writer-char-info .pinyin-small{
    font-size:.8rem;font-style:italic;color:var(--primary-dark);
    background:var(--surface-2);padding:.2rem .6rem;border-radius:6px;display:inline-block;
}
[data-theme="dark"] .writer-char-info .pinyin-small{background:rgba(59,130,246,.18);color:#93c5fd}
.writer-chars{display:flex;gap:.4rem;justify-content:center;flex-wrap:wrap;margin-bottom:.75rem}
.writer-char-btn{
    width:42px;height:42px;border-radius:10px;border:1.5px solid var(--border);
    background:var(--surface-2);color:var(--text);
    font-family:var(--font-zh);font-size:1.3rem;font-weight:500;
    cursor:pointer;transition:.15s;display:flex;align-items:center;justify-content:center;padding:0;
}
.writer-char-btn:hover{border-color:var(--primary)}
.writer-char-btn.active{background:var(--primary);color:#fff;border-color:var(--primary);box-shadow:0 4px 10px rgba(37,99,235,.3)}
.writer-target{
    width:280px;height:280px;margin:0 auto;background:#fff;
    border-radius:14px;position:relative;
    box-shadow:inset 0 0 0 2px var(--border);overflow:hidden;
    background-image:
        linear-gradient(to right, transparent calc(50% - 0.5px), #e2e8f0 calc(50% - 0.5px), #e2e8f0 calc(50% + 0.5px), transparent calc(50% + 0.5px)),
        linear-gradient(to bottom, transparent calc(50% - 0.5px), #e2e8f0 calc(50% - 0.5px), #e2e8f0 calc(50% + 0.5px), transparent calc(50% + 0.5px)),
        linear-gradient(45deg, transparent calc(50% - 0.5px), #e2e8f0 calc(50% - 0.5px), #e2e8f0 calc(50% + 0.5px), transparent calc(50% + 0.5px)),
        linear-gradient(-45deg, transparent calc(50% - 0.5px), #e2e8f0 calc(50% - 0.5px), #e2e8f0 calc(50% + 0.5px), transparent calc(50% + 0.5px));
}
.writer-target svg{display:block;width:100%;height:100%;position:relative;z-index:1}
.writer-controls{display:flex;gap:.4rem;justify-content:center;margin-top:1rem;flex-wrap:wrap}
.writer-btn{
    padding:.6rem .95rem;border-radius:10px;border:1.5px solid var(--border);
    background:var(--surface);color:var(--text);font-size:.82rem;font-weight:600;
    cursor:pointer;transition:.15s;display:flex;align-items:center;gap:.35rem;
    font-family:inherit;
}
.writer-btn:hover{background:var(--primary-light);border-color:var(--primary);color:var(--primary-dark)}
.writer-btn.primary{background:var(--primary);color:#fff;border-color:var(--primary)}
.writer-btn.primary:hover{background:var(--primary-dark);color:#fff}
.writer-loading{
    display:flex;flex-direction:column;align-items:center;justify-content:center;
    height:100%;color:var(--text-3);font-size:.85rem;gap:.5rem;padding:1rem;text-align:center;
}
.writer-loading i{font-size:1.8rem;color:var(--primary)}
.writer-score{text-align:center;margin-top:.6rem;font-size:.82rem;color:var(--text-2);min-height:1.2em}
.writer-score.success{color:var(--success);font-weight:600}
.writer-score.error{color:var(--danger);font-weight:600}

/* Admin modal */
.admin-modal{
    position:fixed;inset:0;background:rgba(15,23,42,.75);
    backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);
    z-index:3000;display:none;align-items:center;justify-content:center;
    padding:1rem;animation:fadeIn .2s;
}
.admin-modal.show{display:flex}
.admin-box{
    background:var(--surface);border-radius:20px;width:100%;max-width:900px;
    max-height:calc(100vh - 2rem);overflow:hidden;
    box-shadow:0 20px 60px rgba(0,0,0,.3);
    display:flex;flex-direction:column;
    animation:slideUp .3s cubic-bezier(.34,1.56,.64,1);
}
.admin-header{
    padding:1.25rem 1.5rem;border-bottom:1px solid var(--border);
    display:flex;align-items:center;justify-content:space-between;gap:1rem;flex-wrap:wrap;
}
.admin-header h2{font-size:1.15rem;color:var(--text);display:flex;align-items:center;gap:.5rem;font-weight:700}
.admin-header h2 i{color:var(--amber)}
.admin-header-actions{display:flex;gap:.5rem;align-items:center}
.admin-body{padding:1.25rem 1.5rem;overflow-y:auto;flex:1}
.admin-stats{
    display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));
    gap:.75rem;margin-bottom:1.25rem;
}
.stat-card{
    background:var(--surface-2);border:1px solid var(--border);
    border-radius:var(--radius);padding:1rem;text-align:center;
}
.stat-card .num{font-size:1.8rem;font-weight:800;color:var(--primary);line-height:1;margin-bottom:.3rem}
.stat-card .label{font-size:.75rem;color:var(--text-3);text-transform:uppercase;letter-spacing:.3px;font-weight:600}
.admin-section-title{
    font-size:.75rem;font-weight:700;text-transform:uppercase;
    letter-spacing:.5px;color:var(--text-3);margin-bottom:.6rem;
    display:flex;align-items:center;justify-content:space-between;
}
.btn-add{
    padding:.45rem .85rem;border-radius:8px;border:none;
    background:var(--primary);color:#fff;font-size:.78rem;font-weight:600;
    cursor:pointer;display:inline-flex;align-items:center;gap:.35rem;
    transition:.15s;font-family:inherit;
}
.btn-add:hover{background:var(--primary-dark)}
.user-list{display:flex;flex-direction:column;gap:.5rem}
.user-row{
    display:flex;align-items:center;gap:.75rem;padding:.75rem;
    background:var(--surface-2);border:1px solid var(--border);
    border-radius:var(--radius-sm);transition:.15s;
}
.user-row:hover{border-color:var(--primary)}
.user-row .u-info{flex:1;min-width:0}
.user-row .u-name{font-weight:700;font-size:.88rem;color:var(--text);margin-bottom:.15rem}
.user-row .u-email{font-size:.75rem;color:var(--text-3);word-break:break-all}
.user-row .u-role{
    padding:.15rem .5rem;border-radius:50px;font-size:.65rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.3px;white-space:nowrap;
}
.user-row .u-role.admin{background:var(--amber-light);color:#92400e}
.user-row .u-role.user{background:var(--primary-light);color:var(--primary-dark)}
.user-row .u-actions{display:flex;gap:.3rem}
.u-btn{
    width:32px;height:32px;border-radius:8px;border:1px solid var(--border);
    background:var(--surface);color:var(--text-2);cursor:pointer;
    display:flex;align-items:center;justify-content:center;
    font-size:.8rem;transition:.15s;
}
.u-btn:hover{background:var(--surface-2);color:var(--primary);border-color:var(--primary)}
.u-btn.danger:hover{background:var(--danger-light);color:var(--danger);border-color:var(--danger)}
.admin-close{
    width:34px;height:34px;border-radius:8px;border:1px solid var(--border);
    background:var(--surface);color:var(--text-2);cursor:pointer;
    display:flex;align-items:center;justify-content:center;
    font-size:.9rem;transition:.15s;
}
.admin-close:hover{background:var(--danger-light);color:var(--danger);border-color:var(--danger)}
.add-user-form{
    background:var(--surface-2);border:1px solid var(--border);
    border-radius:var(--radius-sm);padding:1rem;margin-bottom:1rem;display:none;
}
.add-user-form.show{display:block}
.add-user-form h3{font-size:.85rem;color:var(--text);margin-bottom:.75rem;font-weight:700}
.form-group{margin-bottom:.75rem}
.form-group label{display:block;font-size:.75rem;font-weight:600;color:var(--text-2);margin-bottom:.3rem}
.form-group input,.form-group select{
    width:100%;padding:.6rem .85rem;border-radius:8px;
    border:1.5px solid var(--border);background:var(--surface);
    color:var(--text);font-size:.85rem;outline:none;transition:.15s;font-family:inherit;
}
.form-group input:focus,.form-group select:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(37,99,235,.15)}
.form-actions{display:flex;gap:.5rem;justify-content:flex-end;margin-top:.75rem}
.btn{
    padding:.55rem 1rem;border-radius:8px;border:1.5px solid var(--border);
    background:var(--surface);color:var(--text);font-size:.82rem;font-weight:600;
    cursor:pointer;transition:.15s;font-family:inherit;
    display:inline-flex;align-items:center;gap:.35rem;
}
.btn:hover{background:var(--surface-2)}
.btn.primary{background:var(--primary);color:#fff;border-color:var(--primary)}
.btn.primary:hover{background:var(--primary-dark)}
.logs-list{
    max-height:200px;overflow-y:auto;background:var(--surface-2);
    border:1px solid var(--border);border-radius:var(--radius-sm);padding:.5rem;
}
.log-item{
    display:flex;gap:.5rem;padding:.4rem .5rem;font-size:.75rem;
    color:var(--text-2);border-bottom:1px solid var(--border);
}
.log-item:last-child{border-bottom:none}
.log-item .log-time{color:var(--text-3);flex-shrink:0;font-family:monospace;font-size:.7rem}
.log-item .log-msg{flex:1;word-break:break-word}

/* Responsive */
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
    .fab-group{right:16px;bottom:calc(16px + env(safe-area-inset-bottom));gap:.45rem}
    .fab-main{width:52px;height:52px;font-size:1.2rem}
    .fab-btn{width:48px;height:48px;font-size:1.05rem}
    .card.focused .card-zh{font-size:2.2rem}
    .writer-box{padding:1.2rem 1rem;max-width:340px}
    .writer-target{width:240px;height:240px}
    .writer-char-btn{width:38px;height:38px;font-size:1.15rem}
    .writer-btn{padding:.55rem .8rem;font-size:.78rem}
    .login-box{padding:2rem 1.5rem}
    .login-logo{width:60px;height:60px;font-size:1.6rem}
    .login-box h2{font-size:1.2rem}
    .admin-box{max-height:calc(100vh - 1rem);border-radius:16px}
    .admin-header{padding:1rem}
    .admin-header h2{font-size:1rem}
    .admin-body{padding:1rem}
    .user-dropdown{min-width:220px}
    .btn-login-header{padding:.45rem .7rem;font-size:.75rem}
    .btn-login-header span{display:none}
}
@media(max-width:400px){
    .logo-text .subtitle{display:none}
    .icon-btn{width:30px;height:30px;font-size:.75rem}
}
</style>
</head>
<body>

<!-- LOADING -->
<div class="loading-screen" id="loadingScreen">
    <i class="fas fa-spinner"></i>
    <div>Đang tải...</div>
</div>

<!-- STICKY TOP -->
<div class="sticky-top" id="stickyTop" style="display:none">
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
                    <div class="demo-badge" id="demoBadge" style="display:none">
                        <i class="fas fa-eye"></i> Demo
                    </div>
                    <button class="btn-login-header" id="headerLoginBtn" style="display:none">
                        <i class="fas fa-sign-in-alt"></i> <span>Đăng nhập</span>
                    </button>
                    <button class="icon-btn reset-btn hidden" id="resetBtn" title="Đặt lại bộ lọc">
                        <i class="fas fa-undo-alt"></i>
                        <span class="badge" id="resetBadge">0</span>
                    </button>
                    <button class="icon-btn" id="themeToggle" title="Đổi giao diện">
                        <i class="fas fa-moon"></i>
                    </button>
                    <div class="user-menu" id="userMenu" style="display:none">
                        <img class="user-avatar" id="userAvatar" src="" alt="Avatar">
                        <div class="user-dropdown" id="userDropdown">
                            <div class="user-info">
                                <div class="name" id="userName">-</div>
                                <div class="email" id="userEmail">-</div>
                                <span class="role" id="userRole">user</span>
                            </div>
                            <button class="dropdown-item" id="openAdminBtn" style="display:none">
                                <i class="fas fa-shield-alt"></i> Quản lý tài khoản
                            </button>
                            <button class="dropdown-item danger" id="logoutBtn">
                                <i class="fas fa-sign-out-alt"></i> Đăng xuất
                            </button>
                        </div>
                    </div>
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

<!-- FAB GROUP -->
<div class="fab-group" id="fabGroup" style="display:none">
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

<!-- MAIN -->
<main class="main" id="mainContent" style="display:none">
    <div class="container">
        <div class="demo-banner" id="demoBanner" style="display:none">
            <div class="demo-banner-icon"><i class="fas fa-gift"></i></div>
            <div class="demo-banner-text">
                <div class="title">Bạn đang dùng bản Demo</div>
                <div class="desc">
                    Chỉ xem được <b id="demoLimitText">50</b> câu đầu, không luyện viết, 
                    không dùng bộ lọc, phát âm giới hạn <b id="demoAudioText">10</b> lần/ngày.
                    Đăng nhập Google để mở khóa toàn bộ!
                </div>
            </div>
            <button class="demo-banner-btn" onclick="showLoginModal()">
                <i class="fas fa-sign-in-alt"></i> Đăng nhập ngay
            </button>
        </div>

        <div class="desktop-view">
            <div class="table-card">
                <div class="table-scroll" id="desktopWrapper">
                    <div class="no-data"><i class="fas fa-spinner fa-pulse"></i>Đang tải...</div>
                </div>
            </div>
        </div>
        <div class="mobile-view" id="mobileWrapper"></div>
    </div>
</main>

<!-- LOGIN MODAL -->
<div class="login-modal" id="loginModal">
    <div class="login-box">
        <button class="login-close" id="loginClose"><i class="fas fa-times"></i></button>
        <div class="login-logo"><i class="fas fa-language"></i></div>
        <h2>Đăng nhập để mở khóa</h2>
        <p>Đăng nhập bằng Google để sử dụng <b>toàn bộ 1751 câu</b>, luyện viết chữ Hán, phát âm không giới hạn và nhiều tính năng khác.</p>
        <button class="btn-google" id="googleLoginBtn">
            <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" alt="Google">
            Đăng nhập bằng Google
        </button>
        <div class="login-error" id="loginError"></div>
        <div class="login-footer">
            <i class="fas fa-shield-alt"></i> Chỉ tài khoản được cấp phép mới truy cập được.
        </div>
    </div>
</div>

<!-- WRITER MODAL -->
<div class="writer-modal" id="writerModal">
    <div class="writer-box">
        <button class="writer-close" id="writerClose" aria-label="Đóng"><i class="fas fa-times"></i></button>
        <div class="writer-char-info">
            <div class="vi-small" id="writerViSmall"></div>
            <div class="pinyin-small" id="writerPinyinSmall"></div>
        </div>
        <div class="writer-chars" id="writerChars"></div>
        <div class="writer-target" id="writerTarget"></div>
        <div class="writer-score" id="writerScore"></div>
        <div class="writer-controls">
            <button class="writer-btn primary" id="writerAnimate"><i class="fas fa-play"></i> Viết</button>
            <button class="writer-btn" id="writerQuiz"><i class="fas fa-pen"></i> Tự viết</button>
            <button class="writer-btn" id="writerReset"><i class="fas fa-undo-alt"></i> Xóa</button>
        </div>
    </div>
</div>

<!-- ADMIN MODAL -->
<div class="admin-modal" id="adminModal">
    <div class="admin-box">
        <div class="admin-header">
            <h2><i class="fas fa-shield-alt"></i> Quản lý tài khoản</h2>
            <div class="admin-header-actions">
                <button class="admin-close" id="adminClose"><i class="fas fa-times"></i></button>
            </div>
        </div>
        <div class="admin-body">
            <div class="admin-stats" id="adminStats"></div>
            <div class="admin-section-title">
                <span>Danh sách tài khoản</span>
                <button class="btn-add" id="showAddUserBtn">
                    <i class="fas fa-plus"></i> Thêm tài khoản
                </button>
            </div>
            <div class="add-user-form" id="addUserForm">
                <h3>Thêm tài khoản mới</h3>
                <div class="form-group">
                    <label>Email Google</label>
                    <input type="email" id="newUserEmail" placeholder="user@gmail.com">
                </div>
                <div class="form-group">
                    <label>Tên hiển thị</label>
                    <input type="text" id="newUserName" placeholder="Nguyễn Văn A">
                </div>
                <div class="form-group">
                    <label>Vai trò</label>
                    <select id="newUserRole">
                        <option value="user">User (chỉ học)</option>
                        <option value="admin">Admin (quản trị)</option>
                    </select>
                </div>
                <div class="form-actions">
                    <button class="btn" id="cancelAddUser">Hủy</button>
                    <button class="btn primary" id="confirmAddUser"><i class="fas fa-check"></i> Thêm</button>
                </div>
            </div>
            <div class="user-list" id="userList">
                <div class="no-data"><i class="fas fa-spinner fa-pulse"></i>Đang tải...</div>
            </div>
            <div class="admin-section-title" style="margin-top:1.5rem">
                <span>Lịch sử đăng nhập (gần đây)</span>
            </div>
            <div class="logs-list" id="logsList">
                <div class="no-data" style="padding:1rem;font-size:.8rem"><i class="fas fa-spinner fa-pulse"></i>Đang tải...</div>
            </div>
        </div>
    </div>
</div>

<script>
var RAW_DATA = __DATA__;
var FIREBASE_CONFIG = __FIREBASE_CONFIG__;
var DEMO_LIMIT = 50;
var DEMO_AUDIO_LIMIT = 10;

var currentUser = null;
var isDemo = true;
var auth, db;
var filtered = [];
var state = { search:'', hsk:'', subject:'' };
var PAGE_SIZE = 300;
var renderedCount = 0;
var focusedStt = null;
var appInitialized = false;
var usersCache = [];
var usersUnsubscribe = null;

var $ = function(id) { return document.getElementById(id); };
var desktopWrapper, mobileWrapper;

function getDemoAudioUsed() {
    try {
        var today = new Date().toDateString();
        var data = JSON.parse(localStorage.getItem('demo_audio') || '{}');
        if (data.date !== today) { data = { date: today, count: 0 }; localStorage.setItem('demo_audio', JSON.stringify(data)); }
        return data.count;
    } catch(e) { return 0; }
}
function incDemoAudioUsed() {
    try {
        var today = new Date().toDateString();
        var data = JSON.parse(localStorage.getItem('demo_audio') || '{}');
        if (data.date !== today) data = { date: today, count: 0 };
        data.count++;
        localStorage.setItem('demo_audio', JSON.stringify(data));
    } catch(e) {}
}
function canUseAudio() {
    if (!isDemo) return true;
    return getDemoAudioUsed() < DEMO_AUDIO_LIMIT;
}

try {
    firebase.initializeApp(FIREBASE_CONFIG);
    auth = firebase.auth();
    db = firebase.firestore();
    auth.onAuthStateChanged(handleAuthChange);
} catch(e) {
    console.error('Firebase init error:', e);
    enterDemoMode();
}

async function handleAuthChange(user) {
    if (!user) {
        currentUser = null; isDemo = true;
        applyUserUI(); enterDemoMode();
        return;
    }
    try {
        var email = (user.email || '').toLowerCase();
        var doc = await db.collection('allowed_users').doc(email).get();
        if (!doc.exists) {
            await auth.signOut();
            showLoginError('Tài khoản <b>' + email + '</b> chưa được cấp quyền.');
            enterDemoMode();
            return;
        }
        var data = doc.data();
        currentUser = {
            email: email,
            name: data.name || user.displayName || email.split('@')[0],
            role: data.role || 'user',
            photo: user.photoURL || ''
        };
        isDemo = false;
        applyUserUI();
        logLogin(currentUser);
        if (!appInitialized) { initApp(); appInitialized = true; }
        else { refreshApp(); }
    } catch(e) {
        console.error('Auth check error:', e);
        isDemo = true; enterDemoMode();
    }
}

function enterDemoMode() {
    currentUser = null; isDemo = true;
    applyUserUI();
    if (!appInitialized) { initApp(); appInitialized = true; }
    else { refreshApp(); }
    $('loadingScreen').classList.add('hidden');
    $('stickyTop').style.display = 'block';
    $('fabGroup').style.display = 'flex';
    $('mainContent').style.display = 'block';
}

function applyUserUI() {
    var demoBadge = $('demoBadge');
    var headerLoginBtn = $('headerLoginBtn');
    var userMenu = $('userMenu');
    
    if (isDemo) {
        demoBadge.style.display = 'flex';
        headerLoginBtn.style.display = 'flex';
        userMenu.style.display = 'none';
        $('demoBanner').style.display = 'flex';
    } else {
        demoBadge.style.display = 'none';
        headerLoginBtn.style.display = 'none';
        userMenu.style.display = 'block';
        $('demoBanner').style.display = 'none';
        
        $('userName').textContent = currentUser.name;
        $('userEmail').textContent = currentUser.email;
        $('userRole').textContent = currentUser.role;
        $('userRole').className = 'role' + (currentUser.role === 'admin' ? ' admin' : '');
        $('openAdminBtn').style.display = currentUser.role === 'admin' ? 'flex' : 'none';
        
        var avatar = $('userAvatar');
        if (currentUser.photo) avatar.src = currentUser.photo;
        else {
            avatar.src = 'data:image/svg+xml;utf8,' + encodeURIComponent(
                '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">' +
                '<rect fill="#2563eb" width="100" height="100"/>' +
                '<text x="50" y="65" font-size="45" fill="#fff" text-anchor="middle" font-family="sans-serif" font-weight="bold">' +
                currentUser.name.charAt(0).toUpperCase() +
                '</text></svg>'
            );
        }
    }
    
    var hskChip = $('hskChip');
    var subjectChip = $('subjectChip');
    if (isDemo) {
        hskChip.classList.add('locked');
        subjectChip.classList.add('locked');
    } else {
        hskChip.classList.remove('locked');
        subjectChip.classList.remove('locked');
    }
    
    $('demoLimitText').textContent = DEMO_LIMIT;
    $('demoAudioText').textContent = DEMO_AUDIO_LIMIT;
}

function refreshApp() {
    applyFilter();
    applyDisplayState();
    updateToggleButtons();
}

window.showLoginModal = function() {
    $('loginModal').classList.add('show');
    $('loginError').classList.remove('show');
};
function hideLoginModal() { $('loginModal').classList.remove('show'); }

$('headerLoginBtn').addEventListener('click', showLoginModal);
$('loginClose').addEventListener('click', hideLoginModal);
$('loginModal').addEventListener('click', function(e) {
    if (e.target === this) hideLoginModal();
});

$('googleLoginBtn').addEventListener('click', async function() {
    var provider = new firebase.auth.GoogleAuthProvider();
    provider.setCustomParameters({ prompt: 'select_account' });
    try {
        await auth.signInWithPopup(provider);
        hideLoginModal();
    } catch(e) {
        if (e.code === 'auth/popup-blocked') {
            await auth.signInWithRedirect(provider);
        } else if (e.code !== 'auth/popup-closed-by-user') {
            showLoginError('Lỗi: ' + e.message);
        }
    }
});

function showLoginError(msg) {
    var el = $('loginError');
    el.innerHTML = '<i class="fas fa-exclamation-triangle"></i> ' + msg;
    el.classList.add('show');
}

$('logoutBtn').addEventListener('click', function() {
    if (confirm('Đăng xuất?')) auth.signOut();
});

$('userAvatar').addEventListener('click', function(e) {
    e.stopPropagation();
    $('userDropdown').classList.toggle('show');
});
document.addEventListener('click', function(e) {
    var dd = $('userDropdown');
    if (dd && !dd.contains(e.target) && !$('userAvatar').contains(e.target)) {
        dd.classList.remove('show');
    }
});

function logLogin(u) {
    try {
        db.collection('login_logs').add({
            email: u.email, name: u.name, role: u.role,
            time: firebase.firestore.FieldValue.serverTimestamp(),
            userAgent: navigator.userAgent.substring(0, 100)
        }).catch(function() {});
    } catch(e) {}
}

function initApp() {
    desktopWrapper = $('desktopWrapper');
    mobileWrapper = $('mobileWrapper');
    
    $('loadingScreen').classList.add('hidden');
    $('stickyTop').style.display = 'block';
    $('fabGroup').style.display = 'flex';
    $('mainContent').style.display = 'block';
    
    initScrollDetection();
    initFabGroup();
    initTheme();
    initDisplayState();
    initSpeech();
    initWriter();
    
    $('searchInput').addEventListener('input', applyFilter);
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
    
    $('hskFilter').addEventListener('change', function() {
        if (isDemo) { alert('Vui lòng đăng nhập để sử dụng bộ lọc!'); this.value = ''; return; }
        applyFilter();
    });
    $('subjectFilter').addEventListener('change', function() {
        if (isDemo) { alert('Vui lòng đăng nhập để sử dụng bộ lọc!'); this.value = ''; return; }
        applyFilter();
    });
    
    try { buildFilters(); applyFilter(); }
    catch(e) { console.error('Init error:', e); }
}

function initScrollDetection() {
    var sticky = $('stickyTop');
    if (!sticky) return;
    var ticking = false;
    function update() {
        if (window.scrollY > 5) sticky.classList.add('scrolled');
        else sticky.classList.remove('scrolled');
        ticking = false;
    }
    window.addEventListener('scroll', function() {
        if (!ticking) { window.requestAnimationFrame(update); ticking = true; }
    }, { passive: true });
    update();
}

function initFabGroup() {
    var fabGroup = $('fabGroup');
    var fabMainBtn = $('fabMainBtn');
    var fabOpen = false;
    try { var savedFab = localStorage.getItem('fabOpen'); if (savedFab === 'true') fabOpen = true; } catch(e) {}
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
        }
    });
}

function initTheme() {
    try {
        var saved = localStorage.getItem('theme');
        if (saved) document.documentElement.setAttribute('data-theme', saved);
        else if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            document.documentElement.setAttribute('data-theme', 'dark');
        }
    } catch(e) {}
    updateThemeIcon();
    $('themeToggle').addEventListener('click', function() {
        var isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        var newTheme = isDark ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', newTheme);
        try { localStorage.setItem('theme', newTheme); } catch(e) {}
        updateThemeIcon();
    });
}
function updateThemeIcon() {
    var isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    var icon = $('themeToggle').querySelector('i');
    icon.className = isDark ? 'fas fa-sun' : 'fas fa-moon';
}

var displayState = { vi: true, pinyin: false, practice: false };
function initDisplayState() {
    try {
        var saved = localStorage.getItem('displayState');
        if (saved) {
            var parsed = JSON.parse(saved);
            displayState.vi = parsed.vi !== false;
            displayState.pinyin = !!parsed.pinyin;
            displayState.practice = !!parsed.practice;
        }
    } catch(e) {}
    if (displayState.practice) displayState.pinyin = false;
    applyDisplayState();
    updateToggleButtons();
    
    $('toggleViBtn').addEventListener('click', function(e) {
        e.stopPropagation();
        displayState.vi = !displayState.vi;
        applyDisplayState(); saveDisplayState(); updateToggleButtons();
    });
    $('togglePinyinBtn').addEventListener('click', function(e) {
        e.stopPropagation();
        if (isDemo) { alert('Vui lòng đăng nhập để xem Pinyin!'); return; }
        displayState.pinyin = !displayState.pinyin;
        if (displayState.pinyin && displayState.practice) displayState.practice = false;
        applyDisplayState(); saveDisplayState(); updateToggleButtons();
    });
    $('togglePracticeBtn').addEventListener('click', function(e) {
        e.stopPropagation();
        if (isDemo) { alert('Vui lòng đăng nhập để dùng ô luyện tập!'); return; }
        displayState.practice = !displayState.practice;
        if (displayState.practice && displayState.pinyin) displayState.pinyin = false;
        applyDisplayState(); saveDisplayState(); updateToggleButtons();
    });
}
function applyDisplayState() {
    document.body.classList.toggle('show-vi', displayState.vi);
    document.body.classList.toggle('show-pinyin', displayState.pinyin);
    document.body.classList.toggle('show-practice', displayState.practice);
}
function saveDisplayState() {
    if (isDemo) return;
    try { localStorage.setItem('displayState', JSON.stringify(displayState)); } catch(e) {}
}
function updateToggleButtons() {
    $('toggleViBtn').classList.toggle('active', displayState.vi);
    $('togglePinyinBtn').classList.toggle('active', displayState.pinyin);
    $('togglePracticeBtn').classList.toggle('active', displayState.practice);
    $('togglePinyinBtn').classList.toggle('locked', isDemo);
    $('togglePracticeBtn').classList.toggle('locked', isDemo);
}

window.toggleFocus = function(stt, element) {
    if (focusedStt === stt) { clearFocus(); return; }
    clearFocus();
    focusedStt = stt;
    element.classList.add('focused', 'tapped');
    setTimeout(function() { if (element) element.classList.remove('tapped'); }, 600);
};
window.clearFocus = function() {
    document.querySelectorAll('.focused').forEach(function(el) {
        el.classList.remove('focused', 'tapped');
    });
    focusedStt = null;
};
document.addEventListener('click', function(e) {
    if (e.target.closest('.card') || e.target.closest('.desktop-view tbody tr')) return;
    if (e.target.closest('.practice-input') || e.target.closest('.audio-btn') || 
        e.target.closest('.write-btn') || e.target.closest('.chip') || 
        e.target.closest('.fab-group') || e.target.closest('.icon-btn') ||
        e.target.closest('.search-bar') || e.target.closest('.filters') ||
        e.target.closest('.writer-modal') || e.target.closest('.user-menu') ||
        e.target.closest('.login-modal') || e.target.closest('.admin-modal')) return;
    clearFocus();
}, true);

var currentBtn = null;
function initSpeech() {
    if ('speechSynthesis' in window) {
        speechSynthesis.getVoices();
        if (speechSynthesis.onvoiceschanged !== undefined) {
            speechSynthesis.onvoiceschanged = function(){ getChineseVoice(); };
        }
    }
}
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
window.speakText = function(text, btn, evt) {
    if (evt) evt.stopPropagation();
    if (isDemo && !canUseAudio()) {
        if (confirm('Bạn đã dùng hết ' + DEMO_AUDIO_LIMIT + ' lần phát âm miễn phí hôm nay.\n\nĐăng nhập để tiếp tục phát âm không giới hạn?')) {
            showLoginModal();
        }
        return;
    }
    if (!('speechSynthesis' in window)) { alert('Trình duyệt không hỗ trợ phát âm.'); return; }
    if (isDemo) incDemoAudioUsed();
    speechSynthesis.cancel();
    if (currentBtn) currentBtn.classList.remove('speaking');
    if (btn) { btn.classList.add('speaking'); currentBtn = btn; }
    var utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = 'zh-CN';
    utterance.rate = 0.85;
    utterance.pitch = 1.0;
    utterance.volume = 1.0;
    var voice = getChineseVoice();
    if (voice) utterance.voice = voice;
    utterance.onend = utterance.onerror = function() {
        if (currentBtn) { currentBtn.classList.remove('speaking'); currentBtn = null; }
    };
    setTimeout(function(){ speechSynthesis.speak(utterance); }, 50);
};
document.addEventListener('visibilitychange', function() {
    if (document.hidden && 'speechSynthesis' in window) {
        speechSynthesis.cancel();
        if (currentBtn) { currentBtn.classList.remove('speaking'); currentBtn = null; }
    }
});

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

function buildFilters() {
    var subjectSet = {};
    RAW_DATA.forEach(function(r) { if (r.subject) subjectSet[r.subject] = 1; });
    $('subjectFilter').innerHTML = '<option value="">Tất cả chủ đề</option>' +
        Object.keys(subjectSet).sort().map(function(v){ return '<option value="'+escapeHtml(v)+'">'+escapeHtml(v)+'</option>'; }).join('');
}

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

function render(reset) {
    if (reset) { renderedCount = 0; focusedStt = null; }
    if (!filtered.length) {
        var html = '<div class="no-data"><i class="fas fa-search"></i>Không tìm thấy câu nào</div>';
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
            '<th>Thao tác</th>' +
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
        var viJs = escapeJs(r.vi);
        var pinyinJs = escapeJs(r.pinyin);
        var zhHtml = escapeHtml(r.zh);
        var sttSafe = escapeHtml(r.stt);
        var sttJs = escapeJs(r.stt);
        
        var audio = r.zh ? '<button class="audio-btn" onclick="speakText(\'' + zhJs + '\', this, event)" title="Nghe"><i class="fas fa-volume-up"></i></button>' : '';
        var writeBtn = '';
        if (r.zh) {
            if (isDemo) {
                writeBtn = '<button class="write-btn locked" onclick="event.stopPropagation(); if(confirm(\'Đăng nhập để luyện viết chữ Hán?\')) showLoginModal();" title="Đăng nhập để dùng"><i class="fas fa-pen-fancy"></i></button>';
            } else {
                writeBtn = '<button class="write-btn" onclick="openWriter(\'' + zhJs + '\', \'' + viJs + '\', \'' + pinyinJs + '\', event)" title="Luyện viết"><i class="fas fa-pen-fancy"></i></button>';
            }
        }
        var actionGroup = '<div class="action-group">' + audio + writeBtn + '</div>';
        var practiceInput = '<input type="text" class="practice-input" placeholder="Nhập tiếng Trung..." data-answer="' + zhHtml + '" data-stt="' + sttSafe + '" oninput="checkInput(this)" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">';
        
        deskHtml += '<tr onclick="toggleFocus(\'' + sttJs + '\', this)" data-stt="' + sttSafe + '">' +
            '<td class="stt">' + sttSafe + '</td>' +
            '<td><span class="hsk-badge">' + escapeHtml(r.hsk) + '</span></td>' +
            '<td class="topic-cell">' + escapeHtml(r.topic) + '</td>' +
            '<td class="subject-cell">' + escapeHtml(r.subject) + '</td>' +
            '<td class="vi-cell col-vi">' + escapeHtml(r.vi) + '</td>' +
            '<td class="zh-cell">' + zhHtml + '</td>' +
            '<td class="col-pinyin"><span class="pinyin">' + escapeHtml(r.pinyin) + '</span></td>' +
            '<td onclick="event.stopPropagation()" style="text-align:center">' + actionGroup + '</td>' +
            '<td class="col-practice" onclick="event.stopPropagation()">' + practiceInput + '</td>' +
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
                '<div onclick="event.stopPropagation()" class="action-group">' + audio + writeBtn + '</div>' +
            '</div>' +
            '<div class="card-body">' +
                (r.vi ? '<div class="card-vi">' + escapeHtml(r.vi) + '</div>' : '') +
                '<div class="card-zh">' + zhHtml + '</div>' +
                (r.pinyin ? '<div class="card-pinyin">' + escapeHtml(r.pinyin) + '</div>' : '') +
            '</div>' +
            '<div class="card-practice" onclick="event.stopPropagation()">' +
                practiceInput +
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
    
    if (renderedCount < filtered.length) {
        var isDemoLock = isDemo && renderedCount >= DEMO_LIMIT;
        if (!isDemoLock) {
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
        } else {
            var lockedBtn = document.createElement('button');
            lockedBtn.className = 'load-more locked';
            lockedBtn.innerHTML = '<i class="fas fa-lock"></i> Đăng nhập để xem toàn bộ ' + RAW_DATA.length + ' câu';
            lockedBtn.onclick = function() { showLoginModal(); };
            desktopWrapper.parentElement.appendChild(lockedBtn);
            mobileWrapper.appendChild(lockedBtn.cloneNode(true));
            mobileWrapper.querySelector('.load-more.locked').onclick = function() { showLoginModal(); };
        }
    }
}

window.checkInput = function(input) {
    if (isDemo) {
        input.value = '';
        if (confirm('Đăng nhập để sử dụng ô luyện tập?')) showLoginModal();
        return;
    }
    var stt = input.dataset.stt;
    var answer = input.dataset.answer;
    var cells = document.querySelectorAll('[data-check-stt="' + stt + '"]');
    var val = input.value.trim();
    var norm = function(s){ return s.replace(/[。，！？、；：""''（）\s.,!?;:'"()\[\]{}]/g, ''); };
    var result = '';
    if (val) {
        if (norm(val) === norm(answer)) result = '<span class="check-correct">✅ ĐÚNG</span>';
        else result = '<span class="check-wrong">❌ SAI</span>';
    }
    cells.forEach(function(c) { c.innerHTML = result; });
};

function applyFilter() {
    state.search = $('searchInput').value.trim().toLowerCase();
    state.hsk = isDemo ? '' : $('hskFilter').value;
    state.subject = isDemo ? '' : $('subjectFilter').value;
    
    updateFilterUI();
    var clearBtn = $('clearSearchBtn');
    if (state.search) clearBtn.classList.add('show');
    else clearBtn.classList.remove('show');
    
    var baseData = isDemo ? RAW_DATA.slice(0, DEMO_LIMIT) : RAW_DATA;
    
    filtered = baseData.filter(function(r) {
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

var writerInstance = null;
var currentWriteZh = '';
var currentWriteVi = '';
var currentWritePinyin = '';
var currentCharIndex = 0;

function initWriter() {
    $('writerAnimate').addEventListener('click', function() {
        if (!writerInstance) return;
        $('writerScore').textContent = '';
        $('writerScore').className = 'writer-score';
        writerInstance.cancelQuiz();
        writerInstance.animateCharacter();
    });
    $('writerQuiz').addEventListener('click', function() {
        if (!writerInstance) return;
        $('writerScore').textContent = 'Vẽ chữ bằng ngón tay...';
        $('writerScore').className = 'writer-score';
        writerInstance.quiz({
            onMistake: function(strokeData) {
                $('writerScore').textContent = 'Sai nét ' + (strokeData.strokeNum + 1) + ' - thử lại';
                $('writerScore').className = 'writer-score error';
            },
            onComplete: function(summary) {
                if (summary.totalMistakes === 0) {
                    $('writerScore').textContent = '🎉 Tuyệt vời! Viết đúng tất cả các nét!';
                    $('writerScore').className = 'writer-score success';
                } else {
                    $('writerScore').textContent = 'Hoàn thành! Số nét sai: ' + summary.totalMistakes;
                    $('writerScore').className = 'writer-score';
                }
            }
        });
    });
    $('writerReset').addEventListener('click', function() {
        if (!writerInstance) return;
        $('writerScore').textContent = '';
        $('writerScore').className = 'writer-score';
        writerInstance.cancelQuiz();
        var chars = currentWriteZh.split('').filter(function(c) { return /[\u4e00-\u9fa5]/.test(c); });
        var currentChar = chars[currentCharIndex];
        if (currentChar) showWriterChar(currentChar);
    });
    $('writerClose').addEventListener('click', closeWriter);
    $('writerModal').addEventListener('click', function(e) {
        if (e.target === this) closeWriter();
    });
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') closeWriter();
    });
}

window.openWriter = function(zh, vi, pinyin, evt) {
    if (evt) { evt.stopPropagation(); if (evt.preventDefault) evt.preventDefault(); }
    if (isDemo) {
        if (confirm('Đăng nhập để luyện viết chữ Hán?')) showLoginModal();
        return;
    }
    if (typeof HanziWriter === 'undefined') { alert('Thư viện chưa tải xong.'); return; }
    currentWriteZh = zh || '';
    currentWriteVi = vi || '';
    currentWritePinyin = pinyin || '';
    currentCharIndex = 0;
    var chars = currentWriteZh.split('').filter(function(c) { return /[\u4e00-\u9fa5]/.test(c); });
    if (chars.length === 0) { alert('Câu này không có chữ Hán.'); return; }
    $('writerModal').classList.add('show');
    $('writerScore').textContent = '';
    $('writerScore').className = 'writer-score';
    $('writerViSmall').textContent = currentWriteVi;
    $('writerPinyinSmall').textContent = currentWritePinyin;
    renderWriterChars(chars);
    showWriterChar(chars[0]);
};

window.closeWriter = function() {
    $('writerModal').classList.remove('show');
    writerInstance = null;
};

function renderWriterChars(chars) {
    var container = $('writerChars');
    if (chars.length <= 1) { container.innerHTML = ''; return; }
    container.innerHTML = chars.map(function(c, i) {
        return '<button class="writer-char-btn' + (i === 0 ? ' active' : '') + '" data-idx="' + i + '" data-char="' + c + '">' + c + '</button>';
    }).join('');
    container.querySelectorAll('.writer-char-btn').forEach(function(btn) {
        btn.addEventListener('click', function() {
            var idx = parseInt(this.dataset.idx);
            var ch = this.dataset.char;
            container.querySelectorAll('.writer-char-btn').forEach(function(b) { b.classList.remove('active'); });
            this.classList.add('active');
            currentCharIndex = idx;
            $('writerScore').textContent = '';
            $('writerScore').className = 'writer-score';
            showWriterChar(ch);
        });
    });
}

function showWriterChar(char) {
    var target = $('writerTarget');
    target.innerHTML = '<div class="writer-loading"><i class="fas fa-spinner fa-pulse"></i>Đang tải...</div>';
    writerInstance = null;
    setTimeout(function() {
        try {
            target.innerHTML = '';
            var targetSize = target.offsetWidth || 280;
            var padSize = Math.round(targetSize * 0.06);
            var drawWidth = Math.round(targetSize * 0.07);
            writerInstance = HanziWriter.create('writerTarget', char, {
                width: targetSize, height: targetSize, padding: padSize,
                strokeColor: '#1e293b', radicalColor: '#2563eb',
                highlightColor: '#f59e0b', outlineColor: '#cbd5e1',
                drawingColor: '#2563eb', drawingWidth: drawWidth,
                showOutline: true, strokeAnimationSpeed: 1, delayBetweenStrokes: 250,
                charDataLoader: function(ch, onComplete, onError) {
                    fetch('https://cdn.jsdelivr.net/npm/hanzi-writer-data@2.0/' + encodeURIComponent(ch) + '.json')
                        .then(function(res) { if (!res.ok) throw new Error('Không có dữ liệu'); return res.json(); })
                        .then(onComplete)
                        .catch(function(err) {
                            if (onError) onError(err);
                            target.innerHTML = '<div class="writer-loading"><i class="fas fa-exclamation-triangle" style="color:#dc2626"></i>Không tải được dữ liệu.</div>';
                        });
                }
            });
        } catch(e) {
            target.innerHTML = '<div class="writer-loading"><i class="fas fa-exclamation-triangle"></i>Lỗi tạo khung vẽ</div>';
        }
    }, 100);
}

$('openAdminBtn').addEventListener('click', function() {
    $('userDropdown').classList.remove('show');
    openAdminPanel();
});
$('adminClose').addEventListener('click', function() {
    $('adminModal').classList.remove('show');
    if (usersUnsubscribe) { usersUnsubscribe(); usersUnsubscribe = null; }
});
$('adminModal').addEventListener('click', function(e) {
    if (e.target === this) {
        $('adminModal').classList.remove('show');
        if (usersUnsubscribe) { usersUnsubscribe(); usersUnsubscribe = null; }
    }
});

function openAdminPanel() {
    if (!currentUser || currentUser.role !== 'admin') return;
    $('adminModal').classList.add('show');
    loadUsers();
    loadLogs();
}

function loadUsers() {
    if (usersUnsubscribe) usersUnsubscribe();
    usersUnsubscribe = db.collection('allowed_users').onSnapshot(function(snapshot) {
        usersCache = [];
        snapshot.forEach(function(doc) {
            usersCache.push({ email: doc.id, ...doc.data() });
        });
        usersCache.sort(function(a, b) { return (a.email || '').localeCompare(b.email || ''); });
        renderUsers();
        renderAdminStats();
    }, function(err) {
        $('userList').innerHTML = '<div class="no-data" style="color:#dc2626;padding:1rem"><i class="fas fa-exclamation-triangle"></i>Lỗi: ' + err.message + '</div>';
    });
}

function renderAdminStats() {
    var total = usersCache.length;
    var admins = usersCache.filter(function(u) { return u.role === 'admin'; }).length;
    var users = total - admins;
    $('adminStats').innerHTML =
        '<div class="stat-card"><div class="num">' + total + '</div><div class="label">Tổng</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#f59e0b">' + admins + '</div><div class="label">Admin</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#16a34a">' + users + '</div><div class="label">User</div></div>';
}

function renderUsers() {
    var list = $('userList');
    if (!usersCache.length) {
        list.innerHTML = '<div class="no-data" style="padding:1.5rem;font-size:.85rem">Chưa có tài khoản</div>';
        return;
    }
    list.innerHTML = usersCache.map(function(u) {
        var isMe = u.email === currentUser.email;
        return '<div class="user-row">' +
            '<div class="u-info">' +
                '<div class="u-name">' + escapeHtml(u.name || u.email.split('@')[0]) + (isMe ? ' <span style="color:#94a3b8;font-size:.7rem">(bạn)</span>' : '') + '</div>' +
                '<div class="u-email">' + escapeHtml(u.email) + '</div>' +
            '</div>' +
            '<span class="u-role ' + (u.role === 'admin' ? 'admin' : 'user') + '">' + (u.role || 'user') + '</span>' +
            '<div class="u-actions">' +
                (u.role === 'admin' ? 
                    '<button class="u-btn" onclick="changeRole(\'' + escapeJs(u.email) + '\', \'user\')" title="Hạ xuống User"><i class="fas fa-user"></i></button>' :
                    '<button class="u-btn" onclick="changeRole(\'' + escapeJs(u.email) + '\', \'admin\')" title="Nâng lên Admin"><i class="fas fa-shield-alt"></i></button>'
                ) +
                (!isMe ? '<button class="u-btn danger" onclick="deleteUser(\'' + escapeJs(u.email) + '\')" title="Xóa"><i class="fas fa-trash"></i></button>' : '') +
            '</div>' +
        '</div>';
    }).join('');
}

window.changeRole = async function(email, newRole) {
    if (!confirm('Đổi vai trò của ' + email + ' thành "' + newRole + '"?')) return;
    try { await db.collection('allowed_users').doc(email).update({ role: newRole }); }
    catch(e) { alert('Lỗi: ' + e.message); }
};

window.deleteUser = async function(email) {
    if (!confirm('XÓA tài khoản ' + email + '?')) return;
    try { await db.collection('allowed_users').doc(email).delete(); }
    catch(e) { alert('Lỗi: ' + e.message); }
};

$('showAddUserBtn').addEventListener('click', function() {
    $('addUserForm').classList.toggle('show');
    if ($('addUserForm').classList.contains('show')) $('newUserEmail').focus();
});
$('cancelAddUser').addEventListener('click', function() {
    $('addUserForm').classList.remove('show');
    $('newUserEmail').value = '';
    $('newUserName').value = '';
    $('newUserRole').value = 'user';
});
$('confirmAddUser').addEventListener('click', async function() {
    var email = $('newUserEmail').value.trim().toLowerCase();
    var name = $('newUserName').value.trim();
    var role = $('newUserRole').value;
    if (!email || !email.includes('@')) { alert('Email không hợp lệ'); return; }
    if (!name) name = email.split('@')[0];
    try {
        var docRef = db.collection('allowed_users').doc(email);
        var doc = await docRef.get();
        if (doc.exists) { alert('Email này đã tồn tại!'); return; }
        await docRef.set({
            name: name, role: role,
            addedAt: firebase.firestore.FieldValue.serverTimestamp(),
            addedBy: currentUser.email
        });
        $('addUserForm').classList.remove('show');
        $('newUserEmail').value = '';
        $('newUserName').value = '';
        $('newUserRole').value = 'user';
    } catch(e) { alert('Lỗi: ' + e.message); }
});

function loadLogs() {
    db.collection('login_logs').orderBy('time', 'desc').limit(30).get()
        .then(function(snapshot) {
            if (snapshot.empty) {
                $('logsList').innerHTML = '<div class="no-data" style="padding:1rem;font-size:.8rem">Chưa có log</div>';
                return;
            }
            var html = '';
            snapshot.forEach(function(doc) {
                var d = doc.data();
                var time = d.time ? new Date(d.time.toDate()).toLocaleString('vi-VN') : 'N/A';
                html += '<div class="log-item">' +
                    '<span class="log-time">' + time + '</span>' +
                    '<span class="log-msg"><b>' + escapeHtml(d.name || d.email) + '</b> (' + escapeHtml(d.role || 'user') + ')</span>' +
                '</div>';
            });
            $('logsList').innerHTML = html;
        })
        .catch(function(err) {
            $('logsList').innerHTML = '<div class="no-data" style="padding:1rem;font-size:.8rem;color:#dc2626">Lỗi: ' + err.message + '</div>';
        });
}

setTimeout(function() {
    if (!appInitialized) {
        console.warn('Auth timeout, entering demo mode');
        enterDemoMode();
    }
}, 5000);
</script>
</body>
</html>'''

# ====== GHI FILE HTML ======
html_output = html_template.replace("__DATA__", json_data).replace("__FIREBASE_CONFIG__", firebase_config_json)
with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(html_output)

size_kb = os.path.getsize(OUTPUT_HTML) / 1024
print(f"🎉 Đã tạo: {OUTPUT_HTML}")
print(f"📦 Kích thước: {size_kb:.1f} KB")
print(f"📚 Tổng số câu: {len(data)}")
print(f"🎁 Demo: {DEMO_LIMIT} câu + {DEMO_AUDIO_LIMIT} lần phát âm/ngày")
print(f"🔥 Firebase: {FIREBASE_CONFIG['projectId']}")
