# -*- coding: utf-8 -*-
"""
Chuyển file Excel → HTML tự chứa dữ liệu
- Giao diện hiện đại, responsive (card view trên mobile, table trên desktop)
- Dark mode toggle
- Phát âm bằng Web Speech API (giọng AI tích hợp iOS/Android)
- Hỗ trợ iPhone notch (safe-area-inset)
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
    """Loại bỏ ký tự gây lỗi JSON"""
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
/* ========== RESET & BASE ========== */
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
    --accent:#f59e0b;
    --success:#16a34a;
    --success-light:#dcfce7;
    --danger:#dc2626;
    --danger-light:#fee2e2;
    --shadow-sm:0 1px 2px rgba(15,23,42,.04);
    --shadow:0 4px 12px rgba(15,23,42,.06);
    --shadow-lg:0 10px 30px rgba(15,23,42,.08);
    --radius:14px;
    --radius-sm:10px;
    --radius-full:999px;
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
    --shadow-sm:0 1px 2px rgba(0,0,0,.3);
    --shadow:0 4px 12px rgba(0,0,0,.3);
    --shadow-lg:0 10px 30px rgba(0,0,0,.4);
}
html,body{height:100%}
body{
    font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',sans-serif;
    background:var(--bg);
    color:var(--text);
    line-height:1.5;
    font-size:15px;
    padding-bottom:env(safe-area-inset-bottom);
    transition:background .2s,color .2s;
}
.container{max-width:1400px;margin:0 auto;padding:0 1rem}

/* ========== HEADER ========== */
.header{
    position:sticky;
    top:0;
    z-index:100;
    background:var(--surface);
    border-bottom:1px solid var(--border);
    backdrop-filter:saturate(180%) blur(12px);
    -webkit-backdrop-filter:saturate(180%) blur(12px);
}
.header-inner{
    max-width:1400px;
    margin:0 auto;
    padding:.75rem 1rem;
    display:flex;
    align-items:center;
    gap:.75rem;
}
.logo{
    display:flex;
    align-items:center;
    gap:.6rem;
    font-weight:700;
    font-size:1rem;
    color:var(--text);
    flex:1;
    min-width:0;
}
.logo-icon{
    width:36px;
    height:36px;
    background:linear-gradient(135deg,#2563eb,#7c3aed);
    border-radius:10px;
    display:flex;
    align-items:center;
    justify-content:center;
    color:#fff;
    font-size:1rem;
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
.logo-text .title{font-size:.95rem;font-weight:700;color:var(--text)}
.logo-text .subtitle{font-size:.72rem;color:var(--text-3);font-weight:500}
.header-actions{display:flex;gap:.4rem;align-items:center;flex-shrink:0}
.icon-btn{
    width:38px;height:38px;
    border-radius:var(--radius-sm);
    border:1px solid var(--border);
    background:var(--surface);
    color:var(--text-2);
    cursor:pointer;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:.95rem;
    transition:.15s;
    -webkit-tap-highlight-color:transparent;
}
.icon-btn:hover,.icon-btn:active{background:var(--surface-2);color:var(--primary);border-color:var(--primary)}
.stats-pill{
    display:inline-flex;
    align-items:center;
    gap:.4rem;
    background:var(--primary-light);
    color:var(--primary-dark);
    padding:.45rem .85rem;
    border-radius:var(--radius-full);
    font-size:.78rem;
    font-weight:600;
    white-space:nowrap;
}
.stats-pill b{font-weight:800}

/* ========== MAIN ========== */
.main{padding:1rem 0 3rem}

/* ========== SEARCH ========== */
.search-bar{
    position:relative;
    margin-bottom:.75rem;
}
.search-bar i.fa-search{
    position:absolute;
    left:16px;
    top:50%;
    transform:translateY(-50%);
    color:var(--text-3);
    font-size:.95rem;
    pointer-events:none;
}
.search-bar input{
    width:100%;
    padding:.85rem 3rem .85rem 2.85rem;
    border-radius:var(--radius-full);
    border:1.5px solid var(--border);
    background:var(--surface);
    color:var(--text);
    font-size:.95rem;
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
    right:10px;
    top:50%;
    transform:translateY(-50%);
    width:32px;height:32px;
    border-radius:50%;
    border:none;
    background:var(--surface-2);
    color:var(--text-2);
    cursor:pointer;
    display:none;
    align-items:center;
    justify-content:center;
    font-size:.85rem;
    transition:.15s;
}
.search-clear.show{display:flex}
.search-clear:hover{background:var(--danger-light);color:var(--danger)}

/* ========== FILTERS ========== */
.filters{
    display:grid;
    grid-template-columns:repeat(3,1fr) auto;
    gap:.5rem;
    margin-bottom:1rem;
}
.chip{
    display:flex;
    align-items:center;
    gap:.4rem;
    padding:.65rem 1rem;
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
    font-size:.72rem;
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
.btn-reset{
    display:flex;
    align-items:center;
    justify-content:center;
    gap:.4rem;
    padding:.65rem 1.1rem;
    border-radius:var(--radius-full);
    border:1.5px solid var(--border);
    background:var(--surface);
    color:var(--text-2);
    font-size:.85rem;
    font-weight:600;
    cursor:pointer;
    transition:.15s;
    font-family:inherit;
    box-shadow:var(--shadow-sm);
    -webkit-tap-highlight-color:transparent;
    white-space:nowrap;
}
.btn-reset:hover,.btn-reset:active{
    background:var(--danger-light);
    color:var(--danger);
    border-color:var(--danger);
}

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
    max-height:calc(100vh - 260px);
    -webkit-overflow-scrolling:touch;
}
table{width:100%;border-collapse:collapse;font-size:.85rem;min-width:1100px}
thead th{
    background:var(--surface-2);
    color:var(--text-2);
    font-weight:700;
    font-size:.7rem;
    text-transform:uppercase;
    letter-spacing:.5px;
    padding:.75rem .85rem;
    text-align:left;
    border-bottom:1.5px solid var(--border);
    position:sticky;
    top:0;
    z-index:10;
    white-space:nowrap;
}
tbody td{
    padding:.7rem .85rem;
    border-bottom:1px solid var(--border);
    vertical-align:middle;
    color:var(--text);
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
.topic-cell,.subject-cell{font-size:.82rem;color:var(--text-2)}
.vi-cell{font-size:.85rem;color:var(--text)}
.zh-cell{font-size:.95rem;font-weight:600;color:var(--text)}
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
.audio-btn{
    width:34px;height:34px;
    border-radius:50%;
    border:none;
    background:var(--primary-light);
    color:var(--primary-dark);
    cursor:pointer;
    display:inline-flex;
    align-items:center;
    justify-content:center;
    font-size:.9rem;
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
    padding:.5rem .75rem;
    border-radius:var(--radius-full);
    border:1.5px solid var(--border);
    background:var(--surface);
    color:var(--text);
    font-size:.83rem;
    outline:none;
    transition:.15s;
    font-family:inherit;
    -webkit-appearance:none;
}
.practice-input:focus{
    border-color:var(--primary);
    box-shadow:0 0 0 3px rgba(37,99,235,.15);
}
.check-cell{font-weight:700;font-size:.8rem;white-space:nowrap;text-align:center;width:90px}
.check-correct{color:var(--success)}
.check-wrong{color:var(--danger)}

/* ========== MOBILE CARD VIEW ========== */
.mobile-view{display:none}
.card{
    background:var(--surface);
    border-radius:var(--radius);
    border:1px solid var(--border);
    padding:1rem;
    margin-bottom:.75rem;
    box-shadow:var(--shadow-sm);
    position:relative;
}
.card-header{
    display:flex;
    align-items:center;
    gap:.5rem;
    margin-bottom:.75rem;
    padding-bottom:.75rem;
    border-bottom:1px dashed var(--border);
}
.card-stt{
    width:28px;height:28px;
    border-radius:50%;
    background:var(--surface-2);
    color:var(--text-3);
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:.75rem;
    font-weight:700;
    flex-shrink:0;
}
.card-meta{
    display:flex;
    gap:.35rem;
    align-items:center;
    flex:1;
    min-width:0;
    flex-wrap:wrap;
}
.card-tag{
    display:inline-block;
    padding:.15rem .5rem;
    border-radius:var(--radius-full);
    background:var(--surface-2);
    color:var(--text-2);
    font-size:.68rem;
    font-weight:600;
    white-space:nowrap;
}
.card-tag.hsk{background:var(--primary-light);color:var(--primary-dark)}
.card-body{margin-bottom:.75rem}
.card-vi{
    font-size:.85rem;
    color:var(--text-2);
    margin-bottom:.4rem;
    line-height:1.4;
}
.card-zh{
    font-size:1.15rem;
    font-weight:700;
    color:var(--text);
    margin-bottom:.4rem;
    line-height:1.35;
}
.card-pinyin{
    font-size:.8rem;
    font-style:italic;
    color:var(--primary-dark);
    background:var(--surface-2);
    padding:.25rem .5rem;
    border-radius:6px;
    display:inline-block;
}
.card-footer{
    display:flex;
    align-items:center;
    gap:.5rem;
    padding-top:.75rem;
    border-top:1px dashed var(--border);
}
.card-footer .practice-input{flex:1;min-width:0}
.card-check{
    font-size:.8rem;
    font-weight:700;
    white-space:nowrap;
    min-width:60px;
    text-align:center;
}

/* ========== LOAD MORE & STATES ========== */
.load-more{
    display:block;
    width:100%;
    padding:1rem;
    margin-top:1rem;
    border-radius:var(--radius);
    border:1.5px dashed var(--border-strong);
    background:var(--surface);
    color:var(--primary);
    font-weight:700;
    font-size:.9rem;
    cursor:pointer;
    transition:.15s;
    font-family:inherit;
    -webkit-tap-highlight-color:transparent;
}
.load-more:hover,.load-more:active{background:var(--primary-light);border-color:var(--primary)}
.end-note{
    text-align:center;
    padding:1.25rem 1rem;
    color:var(--text-3);
    font-size:.85rem;
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
@media(max-width:900px){
    .filters{grid-template-columns:repeat(3,1fr);gap:.4rem}
    .btn-reset{grid-column:1 / -1;padding:.6rem;font-size:.82rem}
}
@media(max-width:768px){
    .desktop-view{display:none}
    .mobile-view{display:block}
    
    .container{padding:0 .75rem}
    .header-inner{padding:.6rem .75rem;gap:.5rem}
    .logo-icon{width:32px;height:32px;font-size:.9rem}
    .logo-text .title{font-size:.85rem}
    .logo-text .subtitle{font-size:.65rem}
    .stats-pill{padding:.35rem .6rem;font-size:.7rem}
    .icon-btn{width:34px;height:34px;font-size:.85rem}
    
    .main{padding:.75rem 0 2rem}
    
    .search-bar input{padding:.75rem 2.75rem .75rem 2.6rem;font-size:.9rem}
    
    .filters{grid-template-columns:1fr 1fr;gap:.4rem}
    .chip{padding:.55rem .75rem;font-size:.78rem}
    .chip-label{font-size:.65rem}
    .btn-reset{grid-column:1 / -1}
}
@media(max-width:400px){
    .logo-text .subtitle{display:none}
    .chip-label{display:none}
}
</style>
</head>
<body>

<!-- ============ HEADER ============ -->
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
            <div class="stats-pill" id="statsDisplay">
                <i class="fas fa-book-open"></i> <span id="statsText">Đang tải...</span>
            </div>
            <button class="icon-btn" id="themeToggle" title="Đổi giao diện sáng/tối">
                <i class="fas fa-moon"></i>
            </button>
        </div>
    </div>
</header>

<!-- ============ MAIN ============ -->
<main class="main">
    <div class="container">

        <!-- SEARCH -->
        <div class="search-bar">
            <i class="fas fa-search"></i>
            <input type="text" id="searchInput" placeholder="Tìm tiếng Việt, tiếng Trung hoặc pinyin..." autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">
            <button class="search-clear" id="clearSearchBtn" aria-label="Xóa">
                <i class="fas fa-times"></i>
            </button>
        </div>

        <!-- FILTERS -->
        <div class="filters">
            <div class="chip" id="hskChip">
                <span class="chip-label">HSK</span>
                <span class="chip-value" id="hskValue">Tất cả</span>
                <i class="fas fa-chevron-down chip-arrow"></i>
                <select id="hskFilter"><option value="">Tất cả HSK</option></select>
            </div>
            <div class="chip" id="topicChip">
                <span class="chip-label">Chủ điểm</span>
                <span class="chip-value" id="topicValue">Tất cả</span>
                <i class="fas fa-chevron-down chip-arrow"></i>
                <select id="topicFilter"><option value="">Tất cả chủ điểm</option></select>
            </div>
            <div class="chip" id="subjectChip">
                <span class="chip-label">Chủ đề</span>
                <span class="chip-value" id="subjectValue">Tất cả</span>
                <i class="fas fa-chevron-down chip-arrow"></i>
                <select id="subjectFilter"><option value="">Tất cả chủ đề</option></select>
            </div>
            <button class="btn-reset" id="resetBtn">
                <i class="fas fa-undo-alt"></i> Đặt lại
            </button>
        </div>

        <!-- DESKTOP TABLE -->
        <div class="desktop-view">
            <div class="table-card">
                <div class="table-scroll" id="desktopWrapper">
                    <div class="no-data"><i class="fas fa-spinner fa-pulse"></i>Đang tải dữ liệu...</div>
                </div>
            </div>
        </div>

        <!-- MOBILE CARDS -->
        <div class="mobile-view" id="mobileWrapper"></div>

    </div>
</main>

<script>
/* ========== DỮ LIỆU ========== */
var RAW_DATA = __DATA__;

(function() {
    if (!Array.isArray(RAW_DATA)) {
        document.getElementById('desktopWrapper').innerHTML =
            '<div class="error-box"><i class="fas fa-exclamation-triangle"></i>Dữ liệu không hợp lệ.</div>';
        return;
    }
    console.log('✅ Đã load', RAW_DATA.length, 'câu');
})();

var filtered = RAW_DATA.slice();
var state = { search:'', hsk:'', topic:'', subject:'' };
var PAGE_SIZE = 300;
var renderedCount = 0;

var $ = function(id) { return document.getElementById(id); };
var desktopWrapper = $('desktopWrapper');
var mobileWrapper = $('mobileWrapper');
var statsText = $('statsText');

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

window.speakText = function(text, btn) {
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
    var hskSet = {}, topicSet = {}, subjectSet = {};
    RAW_DATA.forEach(function(r) {
        if (r.hsk) hskSet[r.hsk] = 1;
        if (r.topic) topicSet[r.topic] = 1;
        if (r.subject) subjectSet[r.subject] = 1;
    });
    $('hskFilter').innerHTML = '<option value="">Tất cả HSK</option>' +
        Object.keys(hskSet).sort().map(function(v){ return '<option value="'+escapeHtml(v)+'">'+escapeHtml(v)+'</option>'; }).join('');
    $('topicFilter').innerHTML = '<option value="">Tất cả chủ điểm</option>' +
        Object.keys(topicSet).sort().map(function(v){ return '<option value="'+escapeHtml(v)+'">'+escapeHtml(v)+'</option>'; }).join('');
    $('subjectFilter').innerHTML = '<option value="">Tất cả chủ đề</option>' +
        Object.keys(subjectSet).sort().map(function(v){ return '<option value="'+escapeHtml(v)+'">'+escapeHtml(v)+'</option>'; }).join('');
}

/* ========== UPDATE FILTER UI ========== */
function updateFilterUI() {
    var hsk = $('hskFilter').value;
    var topic = $('topicFilter').value;
    var subject = $('subjectFilter').value;
    
    $('hskValue').textContent = hsk || 'Tất cả';
    $('topicValue').textContent = topic || 'Tất cả';
    $('subjectValue').textContent = subject || 'Tất cả';
    
    $('hskChip').classList.toggle('has-value', !!hsk);
    $('topicChip').classList.toggle('has-value', !!topic);
    $('subjectChip').classList.toggle('has-value', !!subject);
}

/* ========== RENDER ========== */
function render(reset) {
    if (reset) renderedCount = 0;
    
    if (!filtered.length) {
        var html = '<div class="no-data"><i class="fas fa-search"></i>Không tìm thấy câu nào phù hợp</div>';
        desktopWrapper.innerHTML = html;
        mobileWrapper.innerHTML = html;
        updateStats();
        return;
    }
    
    if (reset) {
        desktopWrapper.innerHTML = '<table><thead><tr>' +
            '<th>STT</th><th>HSK</th><th>Chủ điểm</th><th>Chủ đề</th>' +
            '<th>Tiếng Việt</th><th>Tiếng Trung</th><th>Pinyin</th>' +
            '<th></th><th>Luyện tập</th><th>Check</th>' +
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
        var audio = r.zh
            ? '<button class="audio-btn" onclick="speakText(\'' + zhJs + '\', this)" title="Nghe"><i class="fas fa-volume-up"></i></button>'
            : '';
        
        // Desktop row
        deskHtml += '<tr>' +
            '<td class="stt">' + escapeHtml(r.stt) + '</td>' +
            '<td><span class="hsk-badge">' + escapeHtml(r.hsk) + '</span></td>' +
            '<td class="topic-cell">' + escapeHtml(r.topic) + '</td>' +
            '<td class="subject-cell">' + escapeHtml(r.subject) + '</td>' +
            '<td class="vi-cell">' + escapeHtml(r.vi) + '</td>' +
            '<td class="zh-cell">' + zhHtml + '</td>' +
            '<td><span class="pinyin">' + escapeHtml(r.pinyin) + '</span></td>' +
            '<td style="text-align:center">' + audio + '</td>' +
            '<td><input type="text" class="practice-input" placeholder="Nhập tiếng Trung..." data-answer="' + zhHtml + '" data-stt="' + escapeHtml(r.stt) + '" oninput="checkInput(this)" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false"></td>' +
            '<td class="check-cell" data-check-stt="' + escapeHtml(r.stt) + '"></td>' +
            '</tr>';
        
        // Mobile card
        mobHtml += '<div class="card">' +
            '<div class="card-header">' +
                '<div class="card-stt">' + escapeHtml(r.stt) + '</div>' +
                '<div class="card-meta">' +
                    (r.hsk ? '<span class="card-tag hsk">' + escapeHtml(r.hsk) + '</span>' : '') +
                    (r.topic ? '<span class="card-tag">' + escapeHtml(r.topic) + '</span>' : '') +
                    (r.subject ? '<span class="card-tag">' + escapeHtml(r.subject) + '</span>' : '') +
                '</div>' +
                audio +
            '</div>' +
            '<div class="card-body">' +
                '<div class="card-vi">' + escapeHtml(r.vi) + '</div>' +
                '<div class="card-zh">' + zhHtml + '</div>' +
                '<div class="card-pinyin">' + escapeHtml(r.pinyin) + '</div>' +
            '</div>' +
            '<div class="card-footer">' +
                '<input type="text" class="practice-input" placeholder="Nhập tiếng Trung..." data-answer="' + zhHtml + '" data-stt="' + escapeHtml(r.stt) + '" oninput="checkInput(this)" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">' +
                '<div class="card-check" data-check-stt="' + escapeHtml(r.stt) + '"></div>' +
            '</div>' +
            '</div>';
    }
    
    // Append
    if (desktopBody) desktopBody.insertAdjacentHTML('beforeend', deskHtml);
    mobileWrapper.insertAdjacentHTML('beforeend', mobHtml);
    renderedCount = end;
    
    // Remove old buttons
    var oldDesktopBtn = desktopWrapper.parentElement.querySelector('.load-more');
    if (oldDesktopBtn) oldDesktopBtn.remove();
    var oldMobileBtn = mobileWrapper.querySelector('.load-more');
    if (oldMobileBtn) oldMobileBtn.remove();
    var oldDesktopEnd = desktopWrapper.parentElement.querySelector('.end-note');
    if (oldDesktopEnd) oldDesktopEnd.remove();
    var oldMobileEnd = mobileWrapper.querySelector('.end-note');
    if (oldMobileEnd) oldMobileEnd.remove();
    
    // Add "load more" button if needed
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
    
    updateStats();
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
    state.topic = $('topicFilter').value;
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
            if (!inVi && !inZh && !inPinyin) return false;
        }
        if (state.hsk && r.hsk !== state.hsk) return false;
        if (state.topic && r.topic !== state.topic) return false;
        if (state.subject && r.subject !== state.subject) return false;
        return true;
    });
    render(true);
}

function updateStats() {
    statsText.innerHTML = '<b>' + renderedCount + '</b> / ' + filtered.length;
}

/* ========== EVENTS ========== */
$('searchInput').addEventListener('input', applyFilter);
$('hskFilter').addEventListener('change', applyFilter);
$('topicFilter').addEventListener('change', applyFilter);
$('subjectFilter').addEventListener('change', applyFilter);
$('resetBtn').addEventListener('click', function() {
    $('searchInput').value = '';
    $('hskFilter').value = '';
    $('topicFilter').value = '';
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
    render(true);
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
