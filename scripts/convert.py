# -*- coding: utf-8 -*-
"""
Chuyển file Excel → HTML tự chứa dữ liệu
- Fix lỗi "Đang tải dữ liệu..." do JSON bị vỡ
- Phát âm bằng Web Speech API
- Fix dòng cuối bị thanh trình duyệt che
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

data = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if not row or len(row) < 6:
        continue
    stt = row[0] if row[0] is not None else ""
    hsk = str(row[1]).strip() if row[1] else ""
    topic = str(row[2]).strip() if row[2] else ""
    subject = str(row[3]).strip() if row[3] else ""
    vi = str(row[4]).strip() if row[4] else ""
    zh = str(row[5]).strip() if row[5] else ""
    pinyin = str(row[6]).strip() if len(row) > 6 and row[6] else ""
    if not vi and not zh:
        continue
    # ✅ Loại bỏ ký tự điều khiển và xuống dòng gây lỗi JSON
    def clean(s):
        return s.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').replace('\\', '\\\\')
    data.append({
        "stt": str(stt),
        "hsk": clean(hsk),
        "topic": clean(topic),
        "subject": clean(subject),
        "vi": clean(vi),
        "zh": clean(zh),
        "pinyin": clean(pinyin)
    })

print(f"✅ Đã đọc {len(data)} câu")

# ====== CHUYỂN SANG JSON AN TOÀN ======
# ✅ ensure_ascii=True: escape toàn bộ Unicode (an toàn cho mọi trình duyệt)
json_data = json.dumps(data, ensure_ascii=True, separators=(',', ':'))
# ✅ Escape </ để không phá thẻ <script>
json_data = json_data.replace('</', '<\\/')

# ====== TEMPLATE HTML ======
html_template = r'''<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, viewport-fit=cover">
<title>Học tiếng Trung · Văn phòng & Công xưởng</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
<style>
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
html,body{height:100%}
body{
    font-family:-apple-system,'Segoe UI',Roboto,sans-serif;
    background:#f5f7fb;
    color:#1e2a3a;
    padding:1rem;
    padding-bottom:calc(1rem + env(safe-area-inset-bottom));
    line-height:1.5;
}
.container{max-width:1500px;margin:0 auto}
header{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:.8rem;margin-bottom:1rem}
h1{font-size:1.3rem;font-weight:600;color:#0b2b4a;display:flex;align-items:center;gap:8px}
h1 i{color:#c0392b}
.stats{background:#fff;border-radius:40px;padding:.4rem 1rem;box-shadow:0 2px 8px rgba(0,0,0,.03);font-size:.85rem;border:1px solid #e9edf4}
.stats span{font-weight:700;color:#c0392b}
.filters{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:1rem;align-items:center}
.search-box{flex:2;min-width:200px;position:relative}
.search-box i{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#8a9aa8;font-size:.9rem}
.search-box input{width:100%;padding:11px 16px 11px 38px;border-radius:50px;border:1px solid #dbe1e9;font-size:.95rem;background:#fff;outline:none;transition:.2s}
.search-box input:focus{border-color:#3b7cbf;box-shadow:0 0 0 3px rgba(59,124,191,.1)}
.filter-select{padding:10px 14px;border-radius:30px;border:1px solid #dbe1e9;background:#fff;font-size:.85rem;outline:none;cursor:pointer;min-width:120px}
.reset-btn{background:#fff;border:1px solid #dbe1e9;padding:10px 14px;border-radius:30px;cursor:pointer;font-size:.85rem;display:flex;align-items:center;gap:6px;color:#5b6f82;transition:.2s}
.reset-btn:hover{background:#eef3fa}
.table-wrapper{
    background:#fff;
    border-radius:20px;
    box-shadow:0 12px 30px rgba(0,0,0,.05);
    overflow-y:auto;
    overflow-x:auto;
    max-height:calc(100vh - 240px);
    border:1px solid #eef2f7;
    -webkit-overflow-scrolling:touch;
    padding-bottom:calc(120px + env(safe-area-inset-bottom));
}
table{width:100%;border-collapse:collapse;font-size:.88rem;min-width:1200px}
th{background:#f0f5fc;color:#1e3b5c;font-weight:600;font-size:.75rem;text-transform:uppercase;letter-spacing:.3px;padding:12px 10px;text-align:left;border-bottom:1px solid #d8e2ee;position:sticky;top:0;z-index:10;white-space:nowrap}
td{padding:10px;border-bottom:1px solid #ecf1f7;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f9fcff}
th:nth-child(1),td:nth-child(1){width:50px;text-align:center}
th:nth-child(2),td:nth-child(2){width:70px}
th:nth-child(3),td:nth-child(3){width:130px}
th:nth-child(4),td:nth-child(4){width:110px}
th:nth-child(5),td:nth-child(5){width:200px}
th:nth-child(6),td:nth-child(6){width:200px}
th:nth-child(7),td:nth-child(7){width:160px}
th:nth-child(8),td:nth-child(8){width:55px;text-align:center}
th:nth-child(9),td:nth-child(9){width:180px}
th:nth-child(10),td:nth-child(10){width:100px;text-align:center}
.audio-btn{
    display:inline-flex;
    align-items:center;
    justify-content:center;
    background:#e8f0fe;
    color:#1a5a9c;
    width:36px;height:36px;
    border-radius:50%;
    border:none;
    cursor:pointer;
    transition:.2s;
    font-size:1rem;
    padding:0;
    font-family:inherit;
    -webkit-tap-highlight-color:transparent;
}
.audio-btn:hover,.audio-btn:active{background:#1a5a9c;color:#fff;transform:scale(.95)}
.audio-btn.speaking{
    background:#c0392b;
    color:#fff;
    animation:pulse 1s infinite;
}
@keyframes pulse{
    0%,100%{transform:scale(1);box-shadow:0 0 0 0 rgba(192,57,43,.5)}
    50%{transform:scale(1.1);box-shadow:0 0 0 8px rgba(192,57,43,0)}
}
.pinyin{color:#2c6b9e;font-style:italic;font-size:.82rem;background:#f2f8ff;padding:2px 6px;border-radius:12px;display:inline-block}
.check-cell{font-weight:600;font-size:.85rem;white-space:nowrap}
.check-correct{color:#1e7a4a}
.check-wrong{color:#b33a3a}
.practice-input{width:100%;padding:7px 10px;border-radius:30px;border:1px solid #dbe1e9;font-size:.85rem;background:#fbfdff;outline:none;transition:.15s}
.practice-input:focus{border-color:#3b7cbf;box-shadow:0 0 0 3px rgba(59,124,191,.1)}
.no-data{text-align:center;padding:2.5rem 1rem;color:#7a8b9f;font-size:1rem}
.no-data i{font-size:2.5rem;margin-bottom:.8rem;color:#b8ccdf;display:block}
.hsk-badge{background:#eef3fa;padding:3px 9px;border-radius:20px;font-size:.75rem;font-weight:600;white-space:nowrap}
.load-more{
    text-align:center;
    padding:1.3rem 1rem;
    margin:1rem 1rem calc(4rem + env(safe-area-inset-bottom)) 1rem;
    color:#3b7cbf;
    cursor:pointer;
    background:#eef3fa;
    border-radius:14px;
    font-weight:600;
    font-size:1rem;
    border:1px dashed #b8ccdf;
    transition:.2s;
    position:relative;
    z-index:5;
}
.load-more:hover,.load-more:active{background:#dbe8f7;border-color:#3b7cbf}
.load-more i{margin-right:6px}
.end-note{
    text-align:center;
    padding:1.5rem 1rem calc(4rem + env(safe-area-inset-bottom)) 1rem;
    color:#7a8b9f;
    font-size:.9rem;
}
.end-note i{color:#1e7a4a;margin-right:6px}
.error-box{
    text-align:center;
    padding:2rem 1rem;
    color:#c0392b;
    font-size:.95rem;
    background:#fef5f5;
    border-radius:20px;
    border:1px solid #f5c6c6;
    margin:1rem;
}
.error-box i{font-size:2rem;margin-bottom:.5rem;display:block}
@media(max-width:600px){
    body{padding:.7rem;padding-bottom:calc(.7rem + env(safe-area-inset-bottom))}
    h1{font-size:1.1rem}
    .filter-select{min-width:100px;font-size:.8rem;padding:8px 10px}
    .search-box input{font-size:.9rem;padding:10px 14px 10px 34px}
    .table-wrapper{max-height:calc(100vh - 280px);padding-bottom:calc(140px + env(safe-area-inset-bottom))}
    .load-more{margin-bottom:calc(5rem + env(safe-area-inset-bottom))}
    .audio-btn{width:34px;height:34px;font-size:.95rem}
}
</style>
</head>
<body>
<div class="container">
<header>
<h1><i class="fas fa-language"></i> Học tiếng Trung · Văn phòng & Công xưởng</h1>
<div class="stats" id="statsDisplay"><i class="fas fa-book-open"></i> Đang tải...</div>
</header>

<div class="filters">
    <div class="search-box">
        <i class="fas fa-search"></i>
        <input type="text" id="searchInput" placeholder="Tìm tiếng Việt, tiếng Trung, pinyin...">
    </div>
    <select id="hskFilter" class="filter-select"><option value="">Tất cả HSK</option></select>
    <select id="topicFilter" class="filter-select"><option value="">Tất cả chủ điểm</option></select>
    <select id="subjectFilter" class="filter-select"><option value="">Tất cả chủ đề</option></select>
    <button class="reset-btn" id="resetBtn"><i class="fas fa-undo-alt"></i> Đặt lại</button>
</div>

<div class="table-wrapper" id="tableWrapper">
    <div class="no-data"><i class="fas fa-spinner fa-pulse"></i>Đang tải dữ liệu...</div>
</div>
</div>

<script>
/* ========== DỮ LIỆU ========== */
var RAW_DATA = __DATA__;

/* ========== KIỂM TRA DỮ LIỆU ========== */
(function() {
    // Nếu JSON lỗi, hiển thị thông báo rõ ràng
    if (!Array.isArray(RAW_DATA)) {
        document.getElementById('tableWrapper').innerHTML =
            '<div class="error-box"><i class="fas fa-exclamation-triangle"></i>' +
            'Dữ liệu không hợp lệ. Vui lòng kiểm tra lại file Excel.</div>';
        return;
    }
    console.log('✅ Đã load', RAW_DATA.length, 'câu');
})();

var filtered = RAW_DATA.slice();
var state = { search:'', hsk:'', topic:'', subject:'' };
var PAGE_SIZE = 300;
var renderedCount = 0;

var $ = function(id) { return document.getElementById(id); };
var tableWrapper = $('tableWrapper');
var statsDisplay = $('statsDisplay');

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
        alert('Trình duyệt không hỗ trợ phát âm. Vui lòng dùng Chrome hoặc Safari mới.');
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
        if (currentBtn) {
            currentBtn.classList.remove('speaking');
            currentBtn = null;
        }
    }
});

/* ========== ESCAPE ========== */
function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
function escapeJs(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\\'")
        .replace(/"/g, '\\"')
        .replace(/\n/g, '\\n')
        .replace(/\r/g, '');
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

/* ========== RENDER ========== */
function render(reset) {
    if (reset) renderedCount = 0;
    if (!filtered.length) {
        tableWrapper.innerHTML = '<div class="no-data"><i class="fas fa-search"></i>Không tìm thấy câu nào phù hợp.</div>';
        updateStats();
        return;
    }
    if (reset) {
        tableWrapper.innerHTML = '<table id="dataTable"><thead><tr>' +
            '<th>STT</th><th>HSK</th><th>Chủ điểm</th><th>Chủ đề</th>' +
            '<th>Tiếng Việt</th><th>Tiếng Trung</th><th>Pinyin</th>' +
            '<th>Nghe</th><th>Luyện tập</th><th>Check</th>' +
            '</tr></thead><tbody id="dataBody"></tbody></table>';
    }
    var tbody = $('dataBody');
    if (!tbody) {
        console.error('Không tìm thấy tbody');
        return;
    }
    var end = Math.min(renderedCount + PAGE_SIZE, filtered.length);
    var html = '';
    for (var i = renderedCount; i < end; i++) {
        var r = filtered[i];
        var zhJs = escapeJs(r.zh);
        var zhHtml = escapeHtml(r.zh);
        var audio = r.zh
            ? '<button class="audio-btn" onclick="speakText(\'' + zhJs + '\', this)" title="Nghe"><i class="fas fa-volume-up"></i></button>'
            : '';
        html += '<tr>' +
            '<td>' + escapeHtml(r.stt) + '</td>' +
            '<td><span class="hsk-badge">' + escapeHtml(r.hsk) + '</span></td>' +
            '<td>' + escapeHtml(r.topic) + '</td>' +
            '<td>' + escapeHtml(r.subject) + '</td>' +
            '<td>' + escapeHtml(r.vi) + '</td>' +
            '<td><strong>' + zhHtml + '</strong></td>' +
            '<td><span class="pinyin">' + escapeHtml(r.pinyin) + '</span></td>' +
            '<td style="text-align:center">' + audio + '</td>' +
            '<td><input type="text" class="practice-input" placeholder="Nhập..." data-answer="' + zhHtml + '" data-stt="' + escapeHtml(r.stt) + '" oninput="checkInput(this)" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false"></td>' +
            '<td class="check-cell" data-check-stt="' + escapeHtml(r.stt) + '"></td>' +
            '</tr>';
    }
    var oldBtn = tableWrapper.querySelector('.load-more');
    if (oldBtn) oldBtn.remove();
    var oldEnd = tableWrapper.querySelector('.end-note');
    if (oldEnd) oldEnd.remove();
    tbody.insertAdjacentHTML('beforeend', html);
    renderedCount = end;
    if (renderedCount < filtered.length) {
        var btn = document.createElement('div');
        btn.className = 'load-more';
        btn.innerHTML = '<i class="fas fa-chevron-down"></i> Xem thêm (' + renderedCount + '/' + filtered.length + ')';
        btn.onclick = function() {
            render(false);
            setTimeout(function() {
                var newBtn = tableWrapper.querySelector('.load-more');
                if (newBtn) newBtn.scrollIntoView({ behavior:'smooth', block:'center' });
            }, 100);
        };
        tableWrapper.appendChild(btn);
    } else if (filtered.length > PAGE_SIZE) {
        var endNote = document.createElement('div');
        endNote.className = 'end-note';
        endNote.innerHTML = '<i class="fas fa-check-circle"></i> Đã hiển thị tất cả ' + filtered.length + ' câu';
        tableWrapper.appendChild(endNote);
    }
    updateStats();
}

/* ========== CHECK ========== */
window.checkInput = function(input) {
    var stt = input.dataset.stt;
    var answer = input.dataset.answer;
    var cell = document.querySelector('[data-check-stt="' + stt + '"]');
    if (!cell) return;
    var val = input.value.trim();
    if (!val) { cell.innerHTML = ''; return; }
    var norm = function(s){ return s.replace(/[。，！？、；：""''（）\s.,!?;:'"()\[\]{}]/g, ''); };
    if (norm(val) === norm(answer)) {
        cell.innerHTML = '<span class="check-correct">✅ ĐÚNG</span>';
    } else {
        cell.innerHTML = '<span class="check-wrong">❌ SAI</span>';
    }
};

/* ========== FILTER ========== */
function applyFilter() {
    state.search = $('searchInput').value.trim().toLowerCase();
    state.hsk = $('hskFilter').value;
    state.topic = $('topicFilter').value;
    state.subject = $('subjectFilter').value;
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
    statsDisplay.innerHTML = '<i class="fas fa-book-open"></i> Hiển thị <span>' + renderedCount + '</span> / ' + filtered.length + ' câu';
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

/* ========== KHỞI TẠO ========== */
try {
    buildFilters();
    render(true);
} catch (e) {
    console.error('Lỗi khởi tạo:', e);
    tableWrapper.innerHTML =
        '<div class="error-box"><i class="fas fa-exclamation-triangle"></i>' +
        'Lỗi khởi tạo: ' + escapeHtml(e.message) + '</div>';
}
</script>
</body>
</html>'''

html_output = html_template.replace("__DATA__", json_data)
with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(html_output)

size_kb = os.path.getsize(OUTPUT_HTML) / 1024
print(f"🎉 Đã tạo: {OUTPUT_HTML}")
print(f"📦 Kích thước: {size_kb:.1f} KB")
print(f"📚 Tổng số câu: {len(data)}")
