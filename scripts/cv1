# -*- coding: utf-8 -*-
"""
Chuyển file Excel → HTML tự chứa dữ liệu
- Demo mode + Full màn hình luyện tập (search + filter + dropdown chọn câu)
- Nút Gợi ý (mặc định TẮT) → bật hiện ghost text mờ
- Đáp án tách theo PINYIN viết liền/rời
- Nút Xem đáp án: toggle ẩn/hiện
- Desktop hover phóng to, mobile click vừa đọc vừa phóng to
- Click ký tự sai → con trỏ nhảy về + bôi đen để gõ đè
- ✅ Chủ đề demo: mở khóa lên đầu, khóa xuống dưới
- ✅ Dropdown chọn nhanh câu trong modal luyện tập
- Theme mặc định LIGHT MODE
"""
import openpyxl
import json
import os
import sys

CONFIG_FILE = "scripts/config.json"

if not os.path.exists(CONFIG_FILE):
    print(f"❌ Không tìm thấy file cấu hình: {CONFIG_FILE}")
    sys.exit(1)

with open(CONFIG_FILE, "r", encoding="utf-8") as f:
    CONFIG = json.load(f)

EXCEL_FILE = CONFIG.get("excel_file", "data/input.xlsx")
OUTPUT_HTML = CONFIG.get("output_html", "index.html")
SHEET_INDEX = CONFIG.get("sheet_index", 0)
DEMO_LIMIT = CONFIG.get("demo_limit", 50)
DEMO_DAILY_LIMIT = CONFIG.get("demo_daily_limit", 100)
DEMO_HSK_MAX = CONFIG.get("demo_hsk_max", 3)
TARGET_ADMINS = CONFIG.get("target_admins", 2)
ZALO_PHONE = CONFIG.get("zalo_phone", "")
ZALO_NAME = CONFIG.get("zalo_name", "Hỗ trợ")
FIREBASE_CONFIG = CONFIG.get("firebase_config", {})

DEFAULT_SYNONYMS = {
    "我": ["俺", "本人", "咱"], "你": ["您", "阁下"], "他": ["她", "它"],
    "是": ["系", "为"], "的": ["之"], "不": ["没", "未"],
    "很": ["非常", "十分", "特别"], "好": ["棒", "优秀", "不错"],
    "说": ["讲", "谈"], "看": ["瞧", "望"], "吃": ["食", "用"],
    "给": ["送", "赠"], "想要": ["想", "要"],
    "越南": ["越南"], "中国": ["中华"],
    "谢谢": ["感谢", "多谢"], "对不起": ["抱歉", "不好意思"],
    "再见": ["拜拜", "再会"], "请": ["麻烦", "拜托"],
}
SYNONYMS = CONFIG.get("synonyms", DEFAULT_SYNONYMS)
FILLER_WORDS = CONFIG.get("filler_words", ["了", "的", "吗", "呢", "吧", "啊", "呀", "哦", "嘛", "哈", "哪", "着", "过"])

if not FIREBASE_CONFIG.get("apiKey"):
    print(f"❌ Firebase config chưa được cấu hình trong {CONFIG_FILE}")
    sys.exit(1)

print(f"⚙️  Đã đọc cấu hình từ: {CONFIG_FILE}")
print(f"   📞 Zalo: {ZALO_PHONE} ({ZALO_NAME})")
print(f"   🎁 Demo: {DEMO_LIMIT} câu + HSK1-{DEMO_HSK_MAX} + {DEMO_DAILY_LIMIT} lượt nghe/viết")
print(f"   🧠 Chấm điểm: So khớp thông minh")
print(f"   ☀️  Theme mặc định: Light mode")
print(f"   💡 Nút Gợi ý (mặc định TẮT) + Ghost text mờ")
print(f"   📝 Đáp án tách theo PINYIN viết liền/rời")
print(f"   🎯 Click ký tự sai → bôi đen để gõ đè")
print(f"   📂 Chủ đề demo: mở khóa lên đầu, khóa xuống dưới")
print(f"   🔍 Search + Filter + Dropdown chọn câu trong modal")
print(f"   👑 Target admins: {TARGET_ADMINS}")

print(f"\n📖 Đang đọc file: {EXCEL_FILE}")
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
synonyms_json = json.dumps(SYNONYMS, ensure_ascii=True, separators=(',', ':'))
fillers_json = json.dumps(FILLER_WORDS, ensure_ascii=True, separators=(',', ':'))

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
    --zalo:#0068ff;
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
.container{max-width:1100px;margin:0 auto;padding:0 1.5rem}
@media(min-width:1200px){.container{max-width:1050px}}
@media(min-width:1600px){.container{max-width:1200px}}
@media(min-width:2000px){.container{max-width:1300px}}

.loading-screen{
    position:fixed;inset:0;background:var(--bg);
    display:flex;align-items:center;justify-content:center;
    z-index:9998;flex-direction:column;gap:1rem;color:var(--text-2);
}
.loading-screen.hidden{display:none}
.loading-screen i{font-size:2.5rem;color:var(--primary);animation:spin 1s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}

.sticky-top{
    position:sticky;top:0;z-index:150;background:var(--bg);
    padding:.5rem 0 .6rem 0;transition:background .2s, box-shadow .2s, border-color .2s;
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

.demo-badge{
    display:flex;align-items:center;gap:.35rem;
    padding:.35rem .7rem;border-radius:50px;
    background:var(--amber-light);color:#92400e;
    font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:.3px;
    border:1px solid rgba(245,158,11,.4);
}
[data-theme="dark"] .demo-badge{color:#fcd34d}

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

.btn-login-header{
    display:flex;align-items:center;gap:.4rem;
    padding:.5rem .9rem;border-radius:50px;
    background:var(--primary);color:#fff;border:none;
    font-size:.8rem;font-weight:700;cursor:pointer;
    transition:.15s;font-family:inherit;
    box-shadow:0 4px 12px rgba(37,99,235,.3);white-space:nowrap;
}
.btn-login-header:hover,.btn-login-header:active{background:var(--primary-dark);transform:translateY(-1px)}

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
    cursor:pointer;display:none;align-items:center;justify-content:center;font-size:.8rem;
}
.search-clear.show{display:flex}

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
.chip.has-value{background:var(--primary);color:#fff;border-color:var(--primary);box-shadow:0 4px 12px rgba(37,99,235,.3)}
.chip.has-value .chip-label{color:#fff;opacity:.85}
.chip-label{font-size:.7rem;color:var(--text-3);text-transform:uppercase;letter-spacing:.3px;font-weight:700;flex-shrink:0}
.chip-value{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;min-width:0;color:inherit}
.chip-arrow{color:inherit;opacity:.5;font-size:.7rem;flex-shrink:0}
.chip select{
    position:absolute;inset:0;opacity:0;cursor:pointer;font-size:1rem;
    -webkit-appearance:none;appearance:none;width:100%;height:100%;
}
.chip.demo-limited{border-color:var(--amber)}
.chip.demo-limited::before{
    content:'\f023';font-family:'Font Awesome 6 Free';font-weight:900;
    position:absolute;top:4px;right:6px;
    color:var(--amber);font-size:.6rem;pointer-events:none;z-index:2;
}

.result-count{
    display:none;align-items:center;gap:.4rem;margin-top:.5rem;
    padding:.45rem .85rem;border-radius:var(--radius-full);
    background:var(--surface-2);border:1px solid var(--border);
    color:var(--text-2);font-size:.8rem;font-weight:600;
    width:fit-content;box-shadow:var(--shadow-sm);transition:.2s;
}
.result-count.show{display:inline-flex}
.result-count i{color:var(--primary);font-size:.85rem}
.result-count b{color:var(--primary);font-weight:800}
.result-count.empty{background:var(--danger-light);border-color:rgba(220,38,38,.3);color:var(--danger)}
.result-count.empty i,.result-count.empty b{color:var(--danger)}

.zalo-btn{
    position:fixed;bottom:calc(20px + env(safe-area-inset-bottom));
    left:20px;z-index:1000;display:flex;align-items:center;gap:.5rem;
    padding:.7rem 1.1rem;border-radius:50px;
    background:linear-gradient(135deg, #0068ff, #0084ff);
    color:#fff;text-decoration:none;font-weight:700;font-size:.85rem;
    box-shadow:0 8px 24px rgba(0,104,255,.4);
    transition:all .35s cubic-bezier(.34,1.56,.64,1);
    -webkit-tap-highlight-color:transparent;white-space:nowrap;
    border:2px solid #fff;font-family:inherit;overflow:hidden;
}
.zalo-btn:hover,.zalo-btn:active{transform:scale(1.05);box-shadow:0 12px 32px rgba(0,104,255,.55);color:#fff}
.zalo-btn i{font-size:1.2rem;flex-shrink:0;line-height:1;position:relative;z-index:2;transition:font-size .3s}
.zalo-btn .zalo-text{
    line-height:1.15;display:flex;flex-direction:column;
    position:relative;z-index:2;transition:opacity .2s, max-width .35s;
    max-width:200px;overflow:hidden;
}
.zalo-btn .zalo-label{font-size:.65rem;opacity:.85;font-weight:500;white-space:nowrap}
.zalo-btn .zalo-name{font-size:.85rem;font-weight:700;white-space:nowrap}
.zalo-btn::before{
    content:'';position:absolute;inset:0;border-radius:50px;
    background:linear-gradient(135deg, #0068ff, #0084ff);
    opacity:.5;z-index:1;animation:zaloPulse 2s infinite;
}
@keyframes zaloPulse{
    0%{transform:scale(1);opacity:.5}
    50%{transform:scale(1.08);opacity:0}
    100%{transform:scale(1);opacity:0}
}
.zalo-btn.compact{width:52px;height:52px;padding:0;border-radius:50%;justify-content:center;gap:0}
.zalo-btn.compact .zalo-text{opacity:0;max-width:0}
.zalo-btn.compact i{font-size:1.35rem}
.zalo-btn.compact::before{border-radius:50%}
.zalo-btn.compact::after{
    content:'Zalo: ' attr(data-phone);
    position:absolute;left:calc(100% + 10px);top:50%;
    transform:translateY(-50%) scale(.9);
    background:var(--text);color:var(--surface);
    padding:.4rem .7rem;border-radius:8px;
    font-size:.75rem;font-weight:600;white-space:nowrap;
    opacity:0;pointer-events:none;
    transition:opacity .15s, transform .15s;
    font-family:inherit;z-index:3;
}
.zalo-btn.compact:hover::after{opacity:1;transform:translateY(-50%) scale(1)}

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

body:not(.show-pinyin) .col-pinyin,
body:not(.show-pinyin) .card-pinyin{display:none!important}
body:not(.show-vi) .col-vi,
body:not(.show-vi) .card-vi{display:none!important}
body:not(.show-practice) .col-practice,
body:not(.show-practice) .card-practice{display:none!important}

body.show-practice .card-zh,
body.show-practice .card-pinyin{display:none!important}
body.show-practice .card-vi{
    display:block!important;font-size:1rem;font-weight:600;
    color:var(--text);margin-bottom:.55rem;line-height:1.4;
}
body.show-practice .card-body{
    background:linear-gradient(135deg, var(--surface-2), rgba(37,99,235,.06));
    padding:.75rem .85rem;border-radius:10px;border-left:3px solid var(--primary);
}

.ai-correct{color:var(--success);font-weight:700}
.ai-partial{color:var(--amber);font-weight:700}
.ai-wrong{color:var(--danger);font-weight:700}
.ai-reason{
    display:block;font-size:.68rem;color:var(--text-3);
    font-weight:400;margin-top:.2rem;font-style:italic;line-height:1.3;
}

.main{padding:.5rem 0 3rem}

.demo-banner{
    background:linear-gradient(135deg, #fef3c7, #fde68a);
    border:1.5px solid #f59e0b;border-radius:var(--radius);
    padding:.9rem 1.1rem;margin-bottom:1rem;
    display:flex;align-items:center;gap:.75rem;flex-wrap:wrap;
}
[data-theme="dark"] .demo-banner{
    background:linear-gradient(135deg, rgba(245,158,11,.15), rgba(245,158,11,.25));
    border-color:#f59e0b;
}
.demo-banner-icon{
    width:36px;height:36px;border-radius:50%;
    background:var(--amber);color:#fff;
    display:flex;align-items:center;justify-content:center;font-size:1rem;flex-shrink:0;
}
.demo-banner-text{flex:1;min-width:200px}
.demo-banner-text .title{font-weight:700;font-size:.9rem;color:#92400e;margin-bottom:.15rem}
[data-theme="dark"] .demo-banner-text .title{color:#fcd34d}
.demo-banner-text .desc{font-size:.78rem;color:#78350f;line-height:1.5}
[data-theme="dark"] .demo-banner-text .desc{color:#fde68a}
.demo-banner-text .desc b{color:#dc2626}
[data-theme="dark"] .demo-banner-text .desc b{color:#fca5a5}
.demo-banner-btn{
    padding:.5rem .9rem;border-radius:50px;border:none;
    background:var(--amber);color:#fff;
    font-size:.8rem;font-weight:700;cursor:pointer;
    transition:.15s;font-family:inherit;
    display:flex;align-items:center;gap:.35rem;white-space:nowrap;
}
.demo-banner-btn:hover{background:#d97706;transform:translateY(-1px)}

.mobile-view{display:grid;grid-template-columns:1fr;gap:.8rem;max-width:100%}
@media(min-width:769px){.mobile-view{grid-template-columns:1fr 1fr;gap:1.2rem;}}
@media(min-width:1800px){.mobile-view{grid-template-columns:1fr 1fr 1fr;}}

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
}
.card-pinyin{
    font-size:.78rem;font-style:italic;color:var(--primary-dark);
    background:var(--surface-2);padding:.2rem .45rem;border-radius:6px;display:inline-block;
}
[data-theme="dark"] .card-pinyin{background:rgba(59,130,246,.18);color:#93c5fd;font-weight:500;font-style:italic}
.card-practice{
    display:flex;align-items:center;gap:.4rem;
    padding-top:.6rem;border-top:1px dashed var(--border);flex-wrap:wrap;
}
.card-practice .practice-input{flex:1;min-width:120px}
.card-check{font-size:.75rem;font-weight:700;white-space:nowrap;min-width:55px;text-align:center}

.audio-btn{
    width:32px;height:32px;border-radius:50%;border:none;
    background:var(--primary-light);color:var(--primary-dark);
    cursor:pointer;display:inline-flex;align-items:center;justify-content:center;
    font-size:.85rem;transition:.15s;position:relative;
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

.practice-full-btn{
    width:32px;height:32px;border-radius:50%;border:none;
    background:var(--primary-light);color:var(--primary-dark);
    cursor:pointer;display:inline-flex;align-items:center;justify-content:center;
    font-size:.8rem;transition:.15s;
}
.practice-full-btn:hover,.practice-full-btn:active{background:var(--primary);color:#fff;transform:scale(1.08)}

.action-group{display:flex;gap:.3rem;justify-content:center;align-items:center;position:relative}
.practice-input{
    width:100%;min-width:120px;padding:.5rem .8rem;
    border-radius:var(--radius-full);border:1.5px solid var(--border);
    background:var(--surface);color:var(--text);font-size:.9rem;
    outline:none;transition:.15s;font-family:var(--font-zh);-webkit-appearance:none;
}
.practice-input:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(37,99,235,.15)}

.card{
    background:var(--surface);border-radius:var(--radius);
    border:1px solid var(--border);padding:.9rem;
    box-shadow:var(--shadow-sm);
    transition:transform .25s, box-shadow .25s, border-color .25s;
    cursor:pointer;user-select:none;
}
.card.tapped{animation:tapPulse .6s}
@keyframes tapPulse{
    0%{box-shadow:0 0 0 0 rgba(37,99,235,.4)}
    70%{box-shadow:0 0 0 14px rgba(37,99,235,0)}
    100%{box-shadow:0 0 0 0 rgba(37,99,235,0)}
}
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
.card.focused .card-zh{font-size:2rem;font-weight:500;letter-spacing:.02em;line-height:1.5}
.practice-input, .audio-btn, .write-btn, .card-practice{cursor:auto}

.load-more{
    grid-column:1 / -1;display:block;width:100%;padding:.9rem;margin-top:.5rem;
    border-radius:var(--radius);border:1.5px dashed var(--border-strong);
    background:var(--surface);color:var(--primary);
    font-weight:700;font-size:.88rem;cursor:pointer;
    transition:.15s;font-family:inherit;
}
.load-more:hover,.load-more:active{background:var(--primary-light);border-color:var(--primary)}
.load-more.locked{border-color:var(--amber);color:#92400e;background:var(--amber-light)}
[data-theme="dark"] .load-more.locked{color:#fcd34d;background:rgba(245,158,11,.15)}
.end-note{
    grid-column:1 / -1;text-align:center;padding:1rem;
    color:var(--text-3);font-size:.82rem;
}
.end-note i{color:var(--success);margin-right:.35rem}
.no-data{
    grid-column:1 / -1;text-align:center;padding:3rem 1rem;color:var(--text-3);
    background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);
}
.no-data i{font-size:2.5rem;margin-bottom:.75rem;color:var(--border-strong);display:block}

/* ✅ MODAL LUYỆN TẬP FULL MÀN HÌNH */
.practice-full-modal{
    position:fixed;inset:0;background:var(--bg);z-index:2500;
    display:none;flex-direction:column;animation:fadeIn .2s;
}
.practice-full-modal.show{display:flex}

.practice-full-header{
    display:flex;align-items:center;gap:.75rem;
    padding:.85rem 1.25rem;background:var(--surface);
    border-bottom:1px solid var(--border);
    flex-shrink:0;box-shadow:0 2px 8px rgba(15,23,42,.04);
}
.practice-full-header .pf-counter{
    font-size:.85rem;font-weight:700;color:var(--text-2);
    background:var(--surface-2);padding:.35rem .75rem;border-radius:50px;
    white-space:nowrap;
}
.practice-full-header .pf-tags{
    display:flex;gap:.35rem;flex:1;min-width:0;flex-wrap:wrap;
}
.practice-full-header .pf-close{
    width:38px;height:38px;border-radius:50%;border:none;
    background:var(--surface-2);color:var(--text-2);cursor:pointer;
    font-size:1rem;display:flex;align-items:center;justify-content:center;
    transition:.15s;flex-shrink:0;
}
.practice-full-header .pf-close:hover{background:var(--danger-light);color:var(--danger)}

/* ✅ Search + Filter + Quick nav trong modal full */
.pf-filters{
    padding:.6rem 1.25rem .5rem 1.25rem;
    background:var(--surface);
    border-bottom:1px solid var(--border);
    flex-shrink:0;
}
.pf-search-wrap{position:relative;margin-bottom:.5rem;}
.pf-search-wrap i.fa-search{
    position:absolute;left:14px;top:50%;transform:translateY(-50%);
    color:var(--text-3);font-size:.85rem;pointer-events:none;
}
.pf-search-wrap input{
    width:100%;padding:.6rem 2.4rem .6rem 2.4rem;
    border-radius:var(--radius-full);border:1.5px solid var(--border);
    background:var(--bg);color:var(--text);font-size:.88rem;
    outline:none;transition:.15s;font-family:inherit;
    -webkit-appearance:none;
}
.pf-search-wrap input:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(37,99,235,.15)}
.pf-search-wrap input::placeholder{color:var(--text-3)}
.pf-search-clear{
    position:absolute;right:8px;top:50%;transform:translateY(-50%);
    width:28px;height:28px;border-radius:50%;border:none;
    background:var(--surface-2);color:var(--text-2);
    cursor:pointer;display:none;align-items:center;justify-content:center;font-size:.75rem;
}
.pf-search-clear.show{display:flex}

.pf-filter-row{display:grid;grid-template-columns:1fr 1fr;gap:.5rem;}
.pf-chip{
    display:flex;align-items:center;gap:.4rem;
    padding:.45rem .8rem;border-radius:var(--radius-full);
    border:1.5px solid var(--border);background:var(--bg);
    color:var(--text);font-size:.8rem;font-weight:500;
    cursor:pointer;transition:.15s;outline:none;font-family:inherit;
    min-width:0;position:relative;overflow:hidden;
}
.pf-chip:active{transform:scale(.98)}
.pf-chip.has-value{background:var(--primary);color:#fff;border-color:var(--primary)}
.pf-chip.has-value .pf-chip-label{color:#fff;opacity:.85}
.pf-chip-label{font-size:.65rem;color:var(--text-3);text-transform:uppercase;letter-spacing:.3px;font-weight:700;flex-shrink:0}
.pf-chip-value{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;min-width:0}
.pf-chip-arrow{color:inherit;opacity:.5;font-size:.65rem;flex-shrink:0}
.pf-chip select{
    position:absolute;inset:0;opacity:0;cursor:pointer;font-size:1rem;
    -webkit-appearance:none;appearance:none;width:100%;height:100%;
}
.pf-result-count{
    display:none;align-items:center;gap:.4rem;margin-top:.45rem;
    padding:.35rem .7rem;border-radius:var(--radius-full);
    background:var(--bg);border:1px solid var(--border);
    color:var(--text-2);font-size:.75rem;font-weight:600;
    width:fit-content;
}
.pf-result-count.show{display:inline-flex}
.pf-result-count i{color:var(--primary);font-size:.8rem}
.pf-result-count b{color:var(--primary);font-weight:800}
.pf-result-count.empty{background:var(--danger-light);color:var(--danger)}
.pf-result-count.empty i,.pf-result-count.empty b{color:var(--danger)}

/* ✅ Dropdown chọn nhanh câu */
.pf-quick-nav{
    margin-top:.5rem;
    display:flex;
    align-items:center;
    gap:.5rem;
}
.pf-quick-nav-label{
    font-size:.7rem;
    font-weight:700;
    color:var(--text-3);
    text-transform:uppercase;
    letter-spacing:.3px;
    white-space:nowrap;
}
.pf-quick-nav-select{
    flex:1;
    min-width:0;
    padding:.5rem 2rem .5rem .8rem;
    border-radius:var(--radius-full);
    border:1.5px solid var(--border);
    background:var(--bg);
    color:var(--text);
    font-size:.82rem;
    font-family:inherit;
    outline:none;
    cursor:pointer;
    transition:.15s;
    -webkit-appearance:none;
    appearance:none;
    background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 12 12'><path fill='%2394a3b8' d='M6 9L1 4h10z'/></svg>");
    background-repeat:no-repeat;
    background-position:right 12px center;
    background-size:10px;
}
.pf-quick-nav-select:focus{
    border-color:var(--primary);
    box-shadow:0 0 0 3px rgba(37,99,235,.15);
}

.practice-full-body{
    flex:1;overflow-y:auto;padding:2rem 1.25rem;
    display:flex;flex-direction:column;
    align-items:center;justify-content:center;min-height:0;
}
.practice-full-content{
    width:100%;max-width:700px;
    display:flex;flex-direction:column;gap:1.5rem;
}

.practice-full-vi{
    font-size:1.75rem;font-weight:600;color:var(--text);
    text-align:center;line-height:1.45;
    padding:1.25rem 1rem;
    background:linear-gradient(135deg, var(--surface-2), rgba(37,99,235,.06));
    border-radius:16px;border-left:4px solid var(--primary);
}
@media(min-width:769px){.practice-full-vi{font-size:2.25rem;}}

.practice-full-input-wrap{display:flex;flex-direction:column;gap:.75rem;}
.practice-full-input{
    width:100%;padding:1rem 1.25rem;
    font-size:1.5rem;font-family:var(--font-zh);
    border-radius:14px;border:2px solid var(--border);
    background:var(--surface);color:var(--text);
    outline:none;transition:.15s;
    text-align:center;letter-spacing:.05em;
    -webkit-appearance:none;
}
.practice-full-input:focus{
    border-color:var(--primary);
    box-shadow:0 0 0 4px rgba(37,99,235,.15);
}
@media(min-width:769px){.practice-full-input{font-size:1.85rem;padding:1.15rem 1.5rem;}}

.char-preview{
    display:flex;justify-content:center;flex-wrap:wrap;
    gap:.5rem;min-height:2.5rem;padding:.75rem 1rem;
    background:var(--surface-2);border-radius:12px;
    border:1px dashed var(--border);
    user-select:none;
    -webkit-user-select:none;
}
.char-preview:empty{display:none}
.char-slot{
    font-family:var(--font-zh);font-size:1.6rem;font-weight:500;
    display:inline-flex;align-items:center;justify-content:center;
    min-width:1.8rem;height:2.4rem;padding:0 .35rem;
    border-radius:8px;transition:.15s;line-height:1;
}
@media(min-width:769px){.char-slot{font-size:2rem;min-width:2.2rem;height:2.8rem;}}

.char-slot.correct{
    color:var(--text);
    font-weight:700;
    background:rgba(22,163,74,.12);
}
.char-slot.wrong{
    color:#fff;background:var(--danger);
    animation:shakeWrong .3s;
    box-shadow:0 2px 8px rgba(220,38,38,.35);
    cursor:pointer;
    position:relative;
}
.char-slot.wrong:hover{
    transform:scale(1.15);
    box-shadow:0 4px 12px rgba(220,38,38,.5);
    z-index:5;
}
.char-slot.wrong:active{transform:scale(.95);}
.char-slot.wrong.highlight{
    animation:blinkHighlight 0.6s ease-in-out 2;
    box-shadow:0 0 0 4px rgba(220,38,38,.4);
}
@keyframes blinkHighlight{
    0%,100%{transform:scale(1.15);background:var(--danger);}
    50%{transform:scale(1.25);background:#ef4444;box-shadow:0 0 0 8px rgba(220,38,38,.3);}
}
.char-slot.ghost{
    color:var(--text);
    opacity:.12;
    font-weight:400;
    background:transparent;
    user-select:none;
    pointer-events:none;
    filter:blur(0.3px);
}
[data-theme="dark"] .char-slot.ghost{
    color:var(--text);
    opacity:.15;
}
.char-slot.extra{
    color:#fff;background:var(--amber);
    box-shadow:0 2px 8px rgba(245,158,11,.35);
    cursor:pointer;
    transition:.15s;
}
.char-slot.extra:hover{
    transform:scale(1.15);
    box-shadow:0 4px 12px rgba(245,158,11,.5);
    z-index:5;
}
@keyframes shakeWrong{
    0%,100%{transform:translateX(0)}
    25%{transform:translateX(-3px)}
    75%{transform:translateX(3px)}
}

.inline-char-preview{
    display:flex;flex-wrap:wrap;gap:.25rem;
    margin-top:.4rem;width:100%;
}
.inline-char-preview .char-slot{
    font-size:.95rem;min-width:1.1rem;height:1.5rem;
    padding:0 .25rem;border-radius:5px;
}

.practice-full-status{
    text-align:center;font-size:1rem;font-weight:700;min-height:1.5rem;
}
.practice-full-status.correct{color:var(--success);}
.practice-full-status.partial{color:var(--amber);}
.practice-full-status.wrong{color:var(--danger);}

.reveal-actions{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:.75rem;
}
.reveal-actions button{
    width:100%;padding:.85rem;
    border-radius:14px;border:2px dashed var(--border-strong);
    background:var(--surface);color:var(--text-2);
    font-size:.9rem;font-weight:600;cursor:pointer;
    transition:all .2s ease;
    font-family:inherit;
    display:flex;align-items:center;justify-content:center;gap:.5rem;
}
.reveal-actions button:hover{
    border-color:var(--primary);color:var(--primary);background:var(--primary-light);
}
.reveal-actions button.hidden{display:none}

#pfHintBtn.active{
    background:var(--amber);color:#fff;border-color:var(--amber);border-style:solid;
    box-shadow:0 4px 12px rgba(245,158,11,.3);
}
#pfHintBtn.active:hover{background:#d97706}

#pfRevealBtn.revealed{
    background:var(--success);color:#fff;border-color:var(--success);border-style:solid;
    box-shadow:0 4px 12px rgba(22,163,74,.3);
}
#pfRevealBtn.revealed:hover{background:#15803d}

.answer-reveal{
    display:none;flex-direction:column;gap:.75rem;
    padding:1.25rem;background:var(--surface-2);
    border-radius:14px;border:1px solid var(--border);
}
.answer-reveal.show{display:flex}
.answer-reveal .ar-label{
    font-size:.75rem;text-transform:uppercase;letter-spacing:.5px;
    color:var(--text-3);font-weight:700;text-align:center;
}
.answer-chars{
    display:flex;justify-content:center;flex-wrap:wrap;gap:.5rem;
}

.answer-phrase-btn{
    font-family:var(--font-zh);font-size:1.5rem;font-weight:500;
    padding:.5rem .9rem;
    border-radius:12px;border:2px solid var(--border);
    background:var(--surface);color:var(--text);
    cursor:pointer;
    transition:transform .25s cubic-bezier(.34,1.56,.64,1),
               background .2s,
               color .2s,
               border-color .2s,
               box-shadow .2s;
    display:inline-flex;align-items:center;justify-content:center;
    -webkit-appearance:none;
    transform-origin:center center;
}

@media(hover:hover) and (pointer:fine){
    .answer-phrase-btn:hover{
        transform:scale(1.35);
        background:var(--primary);
        color:#fff;
        border-color:var(--primary);
        box-shadow:0 8px 24px rgba(37,99,235,.4);
        z-index:10;
    }
}

.answer-phrase-btn.zoom-in{
    transform:scale(1.35);
    background:var(--primary);
    color:#fff;
    border-color:var(--primary);
    box-shadow:0 8px 24px rgba(37,99,235,.4);
    z-index:10;
}

.answer-phrase-btn.speaking{
    background:var(--primary);color:#fff;border-color:var(--primary);
    animation:pulse 1s infinite;
}

@media(min-width:769px){
    .answer-phrase-btn{font-size:1.8rem;padding:.6rem 1.1rem;}
}

.answer-pinyin{
    text-align:center;font-size:.95rem;font-style:italic;
    color:var(--primary-dark);font-weight:500;
}
.answer-actions{
    display:flex;justify-content:center;gap:.5rem;flex-wrap:wrap;
}
.answer-actions button{
    padding:.6rem 1.1rem;border-radius:50px;border:1.5px solid var(--border);
    background:var(--surface);color:var(--text);
    font-size:.85rem;font-weight:600;cursor:pointer;
    transition:.15s;font-family:inherit;
    display:inline-flex;align-items:center;gap:.4rem;
}
.answer-actions button:hover{background:var(--surface-2);border-color:var(--primary);color:var(--primary)}
.answer-actions button.primary{background:var(--primary);color:#fff;border-color:var(--primary)}
.answer-actions button.primary:hover{background:var(--primary-dark)}

.practice-full-nav{
    display:flex;gap:.75rem;padding:1rem 1.25rem;
    background:var(--surface);border-top:1px solid var(--border);
    flex-shrink:0;justify-content:center;
}
.pf-nav-btn{
    flex:1;max-width:220px;padding:.85rem 1rem;
    border-radius:14px;border:1.5px solid var(--border);
    background:var(--surface);color:var(--text);
    font-size:.9rem;font-weight:700;cursor:pointer;
    transition:.15s;font-family:inherit;
    display:inline-flex;align-items:center;justify-content:center;gap:.5rem;
}
.pf-nav-btn:hover:not(:disabled){
    background:var(--primary-light);border-color:var(--primary);color:var(--primary-dark);
    transform:translateY(-1px);
}
.pf-nav-btn:disabled{opacity:.35;cursor:not-allowed}
.pf-nav-btn.primary{
    background:var(--primary);color:#fff;border-color:var(--primary);
    box-shadow:0 4px 12px rgba(37,99,235,.3);
}
.pf-nav-btn.primary:hover:not(:disabled){background:var(--primary-dark)}

.login-modal{
    position:fixed;inset:0;background:rgba(15,23,42,.8);
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
.u-btn:disabled{opacity:.35;cursor:not-allowed}
.u-btn:disabled:hover{background:var(--surface);color:var(--text-2);border-color:var(--border)}
.u-btn.danger:disabled:hover{background:var(--surface);color:var(--text-2);border-color:var(--border)}
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

[data-theme="dark"] .hsk-badge{
    background:rgba(59,130,246,.25);color:#93c5fd;font-weight:800;
    border:1px solid rgba(59,130,246,.4);
}
[data-theme="dark"] .card-tag.hsk{
    background:rgba(59,130,246,.25);color:#93c5fd;font-weight:700;
    border:1px solid rgba(59,130,246,.4);
}
[data-theme="dark"] .audio-btn{
    background:rgba(59,130,246,.22);color:#93c5fd;
    border:1px solid rgba(59,130,246,.35);
}
[data-theme="dark"] .audio-btn:hover,
[data-theme="dark"] .audio-btn:active{background:#3b82f6;color:#fff;border-color:#3b82f6;}

@media(max-width:768px){
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
    .result-count{font-size:.75rem;padding:.4rem .7rem;margin-top:.45rem}
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
    
    .zalo-btn{
        left:16px;bottom:calc(16px + env(safe-area-inset-bottom));
        padding:.6rem .9rem;font-size:.8rem;gap:.4rem;
    }
    .zalo-btn i{font-size:1.05rem}
    .zalo-btn .zalo-label{font-size:.6rem}
    .zalo-btn .zalo-name{font-size:.78rem;max-width:80px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    
    .zalo-btn.compact{width:48px;height:48px}
    .zalo-btn.compact i{font-size:1.25rem}
    .zalo-btn.compact::after{display:none}
    
    .practice-full-header{padding:.65rem .85rem;gap:.5rem}
    .practice-full-header .pf-counter{font-size:.75rem;padding:.3rem .6rem}
    .pf-filters{padding:.5rem .85rem .4rem .85rem}
    .pf-search-wrap input{padding:.55rem 2.2rem .55rem 2.2rem;font-size:.85rem}
    .pf-search-wrap i.fa-search{left:12px;font-size:.8rem}
    .pf-search-clear{width:26px;height:26px;font-size:.7rem}
    .pf-chip{padding:.4rem .7rem;font-size:.75rem}
    .pf-chip-label{font-size:.6rem}
    .pf-result-count{font-size:.7rem;padding:.3rem .6rem}
    .pf-quick-nav-select{font-size:.78rem;padding:.45rem 1.8rem .45rem .7rem}
    .pf-quick-nav-label{font-size:.65rem}
    .practice-full-body{padding:1.25rem .85rem}
    .practice-full-content{gap:1.1rem}
    .practice-full-vi{font-size:1.35rem;padding:1rem .75rem}
    .practice-full-input{font-size:1.35rem;padding:.85rem 1rem}
    .char-slot{font-size:1.4rem;min-width:1.6rem;height:2.1rem}
    .answer-phrase-btn{font-size:1.3rem;padding:.45rem .75rem;}
    .pf-nav-btn{padding:.75rem .75rem;font-size:.82rem}
    .practice-full-nav{padding:.75rem .85rem;gap:.5rem}
    .reveal-actions button{padding:.7rem;font-size:.82rem;}
}
@media(max-width:400px){
    .logo-text .subtitle{display:none}
    .icon-btn{width:30px;height:30px;font-size:.75rem}
    .zalo-btn:not(.compact) .zalo-text{display:none}
    .zalo-btn:not(.compact){padding:0;border-radius:50%;width:48px;height:48px;justify-content:center}
    .zalo-btn:not(.compact) i{font-size:1.2rem}
}
</style>
</head>
<body>

<div class="loading-screen" id="loadingScreen">
    <i class="fas fa-spinner"></i>
    <div>Đang tải...</div>
</div>

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

        <div class="result-count" id="resultCount">
            <i class="fas fa-list-ul"></i>
            <span>Tìm thấy <b id="resultCountNum">0</b> kết quả</span>
        </div>
    </div>
</div>

<a class="zalo-btn" id="zaloBtn" href="#" target="_blank" rel="noopener noreferrer" title="Liên hệ Zalo hỗ trợ">
    <i class="fas fa-comment-dots"></i>
    <div class="zalo-text">
        <span class="zalo-label">Liên hệ Zalo</span>
        <span class="zalo-name">Hỗ trợ</span>
    </div>
</a>

<div class="fab-group" id="fabGroup" style="display:none">
    <button class="fab-btn fab-sub" id="toggleViBtn" data-label="Tiếng Việt" title="Ẩn/hiện Tiếng Việt">
        <i class="fas fa-language"></i>
    </button>
    <button class="fab-btn fab-sub" id="togglePinyinBtn" data-label="Pinyin" title="Ẩn/hiện Pinyin">
        <i class="fas fa-spell-check"></i>
    </button>
    <button class="fab-btn fab-sub" id="togglePracticeBtn" data-label="Luyện dịch Việt → Trung" title="Ẩn/hiện Ô luyện dịch">
        <i class="fas fa-keyboard"></i>
    </button>
    <button class="fab-btn fab-main" id="fabMainBtn" title="Tùy chọn hiển thị">
        <i class="fas fa-sliders-h"></i>
    </button>
</div>

<main class="main" id="mainContent" style="display:none">
    <div class="container">
        <div class="demo-banner" id="demoBanner" style="display:none">
            <div class="demo-banner-icon"><i class="fas fa-gift"></i></div>
            <div class="demo-banner-text">
                <div class="title">Bạn đang dùng bản Demo</div>
                <div class="desc">
                    Xem <b id="demoLimitText">50</b> câu đầu (HSK1-<span id="demoHskMaxText">3</span>).
                    Bộ lọc HSK chỉ từ HSK1-<span id="demoHskMaxText2">3</span>, chủ đề giới hạn theo 50 câu đầu.
                    Nghe + Luyện viết giới hạn <b id="demoDailyText">100</b> lượt/ngày
                    (còn lại: <b id="demoRemainingText">100</b> lượt).
                    Đăng nhập để mở khóa toàn bộ!
                </div>
            </div>
            <button class="demo-banner-btn" onclick="showLoginModal()">
                <i class="fas fa-sign-in-alt"></i> Đăng nhập ngay
            </button>
        </div>

        <div class="mobile-view" id="mobileWrapper">
            <div class="no-data"><i class="fas fa-spinner fa-pulse"></i>Đang tải...</div>
        </div>
    </div>
</main>

<div class="login-modal" id="loginModal">
    <div class="login-box">
        <button class="login-close" id="loginClose"><i class="fas fa-times"></i></button>
        <div class="login-logo"><i class="fas fa-language"></i></div>
        <h2>Đăng nhập để mở khóa</h2>
        <p>Đăng nhập bằng Google để sử dụng <b>toàn bộ câu</b>, tất cả bộ lọc HSK1-6, chủ đề đầy đủ, luyện viết không giới hạn và nhiều tính năng khác.</p>
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

<!-- ✅ MODAL LUYỆN TẬP FULL MÀN HÌNH -->
<div class="practice-full-modal" id="practiceFullModal">
    <div class="practice-full-header">
        <div class="pf-counter" id="pfCounter">Câu 1 / 1</div>
        <div class="pf-tags" id="pfTags"></div>
        <button class="pf-close" id="pfClose" aria-label="Đóng">
            <i class="fas fa-times"></i>
        </button>
    </div>
    
    <!-- ✅ Search + Filter + Quick nav trong modal full -->
    <div class="pf-filters">
        <div class="pf-search-wrap">
            <i class="fas fa-search"></i>
            <input type="text" id="pfSearchInput" placeholder="Tìm kiếm..." autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">
            <button class="pf-search-clear" id="pfClearSearchBtn" aria-label="Xóa">
                <i class="fas fa-times"></i>
            </button>
        </div>
        <div class="pf-filter-row">
            <div class="pf-chip" id="pfHskChip">
                <span class="pf-chip-label">HSK</span>
                <span class="pf-chip-value" id="pfHskValue">Tất cả</span>
                <i class="fas fa-chevron-down pf-chip-arrow"></i>
                <select id="pfHskFilter">
                    <option value="">Tất cả</option>
                    <option value="HSK1">HSK1</option>
                    <option value="HSK2">HSK2</option>
                    <option value="HSK3">HSK3</option>
                    <option value="HSK4">HSK4</option>
                    <option value="HSK5">HSK5</option>
                    <option value="HSK6">HSK6</option>
                </select>
            </div>
            <div class="pf-chip" id="pfSubjectChip">
                <span class="pf-chip-label">Chủ đề</span>
                <span class="pf-chip-value" id="pfSubjectValue">Tất cả</span>
                <i class="fas fa-chevron-down pf-chip-arrow"></i>
                <select id="pfSubjectFilter"><option value="">Tất cả chủ đề</option></select>
            </div>
        </div>
        <div class="pf-result-count" id="pfResultCount">
            <i class="fas fa-list-ul"></i>
            <span>Tìm thấy <b id="pfResultCountNum">0</b> kết quả</span>
        </div>
        <div class="pf-quick-nav">
            <span class="pf-quick-nav-label">Câu:</span>
            <select class="pf-quick-nav-select" id="pfQuickNav">
                <option value="">-- Chọn câu --</option>
            </select>
        </div>
    </div>
    
    <div class="practice-full-body">
        <div class="practice-full-content">
            <div class="practice-full-vi" id="pfVi">-</div>
            
            <div class="practice-full-input-wrap">
                <input type="text" class="practice-full-input" id="pfInput" 
                    placeholder="Gõ tiếng Trung..." autocomplete="off" 
                    autocorrect="off" autocapitalize="off" spellcheck="false">
                <div class="char-preview" id="pfPreview"></div>
                <div class="practice-full-status" id="pfStatus"></div>
            </div>
            
            <div class="reveal-actions">
                <button id="pfHintBtn">
                    <i class="fas fa-lightbulb"></i> Gợi ý
                </button>
                <button id="pfRevealBtn">
                    <i class="fas fa-eye"></i> Xem đáp án
                </button>
            </div>
            
            <div class="answer-reveal" id="pfAnswer">
                <div class="ar-label">Đáp án</div>
                <div class="answer-chars" id="pfAnswerChars"></div>
                <div class="answer-pinyin" id="pfAnswerPinyin"></div>
                <div class="answer-actions">
                    <button class="primary" onclick="speakFullSentence()">
                        <i class="fas fa-volume-up"></i> Đọc cả câu
                    </button>
                </div>
            </div>
        </div>
    </div>
    
    <div class="practice-full-nav">
        <button class="pf-nav-btn" id="pfPrevBtn">
            <i class="fas fa-chevron-left"></i> Câu trước
        </button>
        <button class="pf-nav-btn primary" id="pfNextBtn">
            Câu sau <i class="fas fa-chevron-right"></i>
        </button>
    </div>
</div>

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
var DEMO_LIMIT = __DEMO_LIMIT__;
var DEMO_DAILY_LIMIT = __DEMO_DAILY_LIMIT__;
var DEMO_HSK_MAX = __DEMO_HSK_MAX__;
var TARGET_ADMINS = __TARGET_ADMINS__;
var ZALO_PHONE = "__ZALO_PHONE__";
var ZALO_NAME = "__ZALO_NAME__";
var SYNONYMS = __SYNONYMS__;
var FILLER_WORDS = __FILLER_WORDS__;

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
var mobileWrapper;

function getDemoData() { return RAW_DATA.slice(0, DEMO_LIMIT); }
function getDemoHskList() {
    var list = [];
    for (var i = 1; i <= DEMO_HSK_MAX; i++) list.push('HSK' + i);
    return list;
}
function getDemoSubjectList() {
    var set = {};
    getDemoData().forEach(function(r) { if (r.subject) set[r.subject] = 1; });
    return Object.keys(set).sort();
}

function getDemoUsage() {
    try {
        var today = new Date().toDateString();
        var data = JSON.parse(localStorage.getItem('demo_usage') || '{}');
        if (data.date !== today) {
            data = { date: today, count: 0 };
            localStorage.setItem('demo_usage', JSON.stringify(data));
        }
        return data.count;
    } catch(e) { return 0; }
}
function incDemoUsage() {
    try {
        var today = new Date().toDateString();
        var data = JSON.parse(localStorage.getItem('demo_usage') || '{}');
        if (data.date !== today) data = { date: today, count: 0 };
        data.count++;
        localStorage.setItem('demo_usage', JSON.stringify(data));
    } catch(e) {}
}
function getDemoRemaining() { return Math.max(0, DEMO_DAILY_LIMIT - getDemoUsage()); }
function canUseFeature() {
    if (!isDemo) return true;
    return getDemoUsage() < DEMO_DAILY_LIMIT;
}
function updateDemoRemaining() {
    var el = $('demoRemainingText');
    if (!el) return;
    var remaining = getDemoRemaining();
    el.textContent = remaining;
    if (remaining < 20) el.style.color = '#dc2626';
    else if (remaining < 50) el.style.color = '#f59e0b';
    else el.style.color = '#16a34a';
}
function showLimitMessage() {
    if (confirm('🔒 Bạn đã dùng hết ' + DEMO_DAILY_LIMIT + ' lượt miễn phí hôm nay.\n\n(Bao gồm cả NGHE và LUYỆN VIẾT)\n\nĐăng nhập Google để dùng KHÔNG GIỚI HẠN!')) {
        showLoginModal();
    }
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
        hskChip.classList.add('demo-limited');
        subjectChip.classList.add('demo-limited');
    } else {
        hskChip.classList.remove('demo-limited');
        subjectChip.classList.remove('demo-limited');
    }
    
    var zaloBtn = $('zaloBtn');
    if (zaloBtn) {
        if (isDemo) zaloBtn.classList.remove('compact');
        else zaloBtn.classList.add('compact');
    }
    
    $('demoLimitText').textContent = DEMO_LIMIT;
    $('demoDailyText').textContent = DEMO_DAILY_LIMIT;
    var hskMaxEl1 = $('demoHskMaxText');
    var hskMaxEl2 = $('demoHskMaxText2');
    if (hskMaxEl1) hskMaxEl1.textContent = DEMO_HSK_MAX;
    if (hskMaxEl2) hskMaxEl2.textContent = DEMO_HSK_MAX;
    updateDemoRemaining();
}

function refreshApp() {
    buildFilters();
    applyFilter();
    applyDisplayState();
    updateToggleButtons();
    updateResultCount();
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
    mobileWrapper = $('mobileWrapper');
    
    $('loadingScreen').classList.add('hidden');
    $('stickyTop').style.display = 'block';
    $('fabGroup').style.display = 'flex';
    $('mainContent').style.display = 'block';
    
    initZaloButton();
    initScrollDetection();
    initFabGroup();
    initTheme();
    initDisplayState();
    initSpeech();
    initWriter();
    initPracticeFull();
    
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
        if (isDemo) {
            var val = this.value;
            var allowed = getDemoHskList();
            if (val && allowed.indexOf(val) === -1) {
                alert('🔒 Bản Demo chỉ cho phép lọc HSK1-' + DEMO_HSK_MAX + '.\n\nĐăng nhập Google để mở khóa tất cả HSK!');
                this.value = '';
                applyFilter();
                return;
            }
        }
        applyFilter();
    });
    $('subjectFilter').addEventListener('change', function() {
        if (isDemo) {
            var val = this.value;
            var allowed = getDemoSubjectList();
            if (val && allowed.indexOf(val) === -1) {
                alert('🔒 Chủ đề này chưa có trong ' + DEMO_LIMIT + ' câu Demo.\n\nĐăng nhập Google để mở khóa tất cả chủ đề!');
                this.value = '';
                applyFilter();
                return;
            }
        }
        applyFilter();
    });
    
    try { buildFilters(); applyFilter(); }
    catch(e) { console.error('Init error:', e); }
}

function initZaloButton() {
    var zaloBtn = $('zaloBtn');
    if (!zaloBtn) return;
    
    var phone = (ZALO_PHONE || '').replace(/\D/g, '');
    
    if (phone) {
        zaloBtn.href = 'https://zalo.me/' + phone;
        zaloBtn.title = 'Liên hệ Zalo: ' + phone;
        zaloBtn.setAttribute('data-phone', phone);
    } else {
        zaloBtn.href = '#';
        zaloBtn.onclick = function(e) {
            e.preventDefault();
            alert('Chưa cấu hình số Zalo.');
        };
    }
    
    var nameEl = zaloBtn.querySelector('.zalo-name');
    if (nameEl && ZALO_NAME) nameEl.textContent = ZALO_NAME;
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
        else document.documentElement.setAttribute('data-theme', 'light');
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
    if (displayState.practice) {
        displayState.pinyin = false;
        displayState.vi = true;
    }
    applyDisplayState();
    updateToggleButtons();
    
    $('toggleViBtn').addEventListener('click', function(e) {
        e.stopPropagation();
        displayState.vi = !displayState.vi;
        applyDisplayState(); saveDisplayState(); updateToggleButtons();
    });
    $('togglePinyinBtn').addEventListener('click', function(e) {
        e.stopPropagation();
        displayState.pinyin = !displayState.pinyin;
        if (displayState.pinyin && displayState.practice) displayState.practice = false;
        applyDisplayState(); saveDisplayState(); updateToggleButtons();
    });
    $('togglePracticeBtn').addEventListener('click', function(e) {
        e.stopPropagation();
        displayState.practice = !displayState.practice;
        if (displayState.practice) {
            displayState.pinyin = false;
            displayState.vi = true;
        }
        applyDisplayState(); saveDisplayState(); updateToggleButtons();
    });
}
function applyDisplayState() {
    document.body.classList.toggle('show-vi', displayState.vi);
    document.body.classList.toggle('show-pinyin', displayState.pinyin);
    document.body.classList.toggle('show-practice', displayState.practice);
    
    if (displayState.practice) {
        document.querySelectorAll('.card-check').forEach(function(c) { c.innerHTML = ''; });
        document.querySelectorAll('.practice-input').forEach(function(i) { 
            i.value = '';
            var answer = i.dataset.answer || '';
            updateInlinePreview(i, answer);
        });
    }
}
function saveDisplayState() {
    try { localStorage.setItem('displayState', JSON.stringify(displayState)); } catch(e) {}
}
function updateToggleButtons() {
    $('toggleViBtn').classList.toggle('active', displayState.vi);
    $('togglePinyinBtn').classList.toggle('active', displayState.pinyin);
    $('togglePracticeBtn').classList.toggle('active', displayState.practice);
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
    if (e.target.closest('.card')) return;
    if (e.target.closest('.practice-input') || e.target.closest('.audio-btn') || 
        e.target.closest('.write-btn') || e.target.closest('.chip') || 
        e.target.closest('.fab-group') || e.target.closest('.icon-btn') ||
        e.target.closest('.search-bar') || e.target.closest('.filters') ||
        e.target.closest('.writer-modal') || e.target.closest('.user-menu') ||
        e.target.closest('.login-modal') || e.target.closest('.admin-modal') ||
        e.target.closest('.practice-full-modal') ||
        e.target.closest('.zalo-btn')) return;
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
    if (isDemo && !canUseFeature()) { showLimitMessage(); return; }
    if (!('speechSynthesis' in window)) { alert('Trình duyệt không hỗ trợ phát âm.'); return; }
    if (isDemo) { incDemoUsage(); updateDemoRemaining(); }
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

/* ✅ BUILD FILTERS — Chủ đề demo: mở khóa lên đầu, khóa xuống dưới */
function buildFilters() {
    var hskSelect = $('hskFilter');
    var subjectSelect = $('subjectFilter');
    
    var allSubjectSet = {};
    RAW_DATA.forEach(function(r) { if (r.subject) allSubjectSet[r.subject] = 1; });
    var allSubjects = Object.keys(allSubjectSet).sort();
    
    if (isDemo) {
        var demoHsk = getDemoHskList();
        var hskHtml = '<option value="">Tất cả (HSK1-' + DEMO_HSK_MAX + ')</option>';
        demoHsk.forEach(function(h) { hskHtml += '<option value="' + h + '">' + h + '</option>'; });
        var allHskList = ['HSK1','HSK2','HSK3','HSK4','HSK5','HSK6'];
        allHskList.forEach(function(h) {
            if (demoHsk.indexOf(h) === -1) hskHtml += '<option value="' + h + '" disabled>🔒 ' + h + ' (đăng nhập)</option>';
        });
        hskSelect.innerHTML = hskHtml;
        
        var demoSubjects = getDemoSubjectList();
        var demoSubjectMap = {};
        demoSubjects.forEach(function(s) { demoSubjectMap[s] = 1; });
        
        var unlockedSubjects = [];
        var lockedSubjects = [];
        allSubjects.forEach(function(s) {
            if (demoSubjectMap[s]) unlockedSubjects.push(s);
            else lockedSubjects.push(s);
        });
        
        var subjHtml = '<option value="">Tất cả chủ đề</option>';
        unlockedSubjects.forEach(function(s) {
            subjHtml += '<option value="' + escapeHtml(s) + '">' + escapeHtml(s) + '</option>';
        });
        lockedSubjects.forEach(function(s) {
            subjHtml += '<option value="' + escapeHtml(s) + '" disabled>🔒 ' + escapeHtml(s) + ' (đăng nhập)</option>';
        });
        subjectSelect.innerHTML = subjHtml;
    } else {
        hskSelect.innerHTML = 
            '<option value="">Tất cả</option>' +
            '<option value="HSK1">HSK1</option>' +
            '<option value="HSK2">HSK2</option>' +
            '<option value="HSK3">HSK3</option>' +
            '<option value="HSK4">HSK4</option>' +
            '<option value="HSK5">HSK5</option>' +
            '<option value="HSK6">HSK6</option>';
        
        subjectSelect.innerHTML = '<option value="">Tất cả chủ đề</option>' +
            allSubjects.map(function(v){ return '<option value="'+escapeHtml(v)+'">'+escapeHtml(v)+'</option>'; }).join('');
    }
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
    updateResultCount();
}

function updateResultCount() {
    var el = $('resultCount');
    if (!el) return;
    var total = filtered ? filtered.length : 0;
    var hasFilter = !!(state.search || state.hsk || state.subject);
    if (!hasFilter) { el.classList.remove('show', 'empty'); return; }
    el.classList.add('show');
    el.classList.toggle('empty', total === 0);
    var spanEl = el.querySelector('span');
    if (spanEl) {
        if (total === 0) spanEl.innerHTML = 'Không tìm thấy kết quả nào';
        else spanEl.innerHTML = 'Tìm thấy <b>' + total + '</b> kết quả';
    }
}

function render(reset) {
    if (reset) { renderedCount = 0; focusedStt = null; }
    if (!filtered.length) {
        mobileWrapper.innerHTML = '<div class="no-data"><i class="fas fa-search"></i>Không tìm thấy câu nào</div>';
        return;
    }
    if (reset) mobileWrapper.innerHTML = '';
    var end = Math.min(renderedCount + PAGE_SIZE, filtered.length);
    var mobHtml = '';
    for (var i = renderedCount; i < end; i++) {
        var r = filtered[i];
        var zhJs = escapeJs(r.zh);
        var viJs = escapeJs(r.vi);
        var pinyinJs = escapeJs(r.pinyin);
        var zhHtml = escapeHtml(r.zh);
        var viHtml = escapeHtml(r.vi);
        var sttSafe = escapeHtml(r.stt);
        var sttJs = escapeJs(r.stt);
        
        var audio = r.zh ? '<button class="audio-btn" onclick="speakText(\'' + zhJs + '\', this, event)" title="Nghe"><i class="fas fa-volume-up"></i></button>' : '';
        var writeBtn = '';
        if (r.zh) {
            writeBtn = '<button class="write-btn" onclick="openWriter(\'' + zhJs + '\', \'' + viJs + '\', \'' + pinyinJs + '\', event)" title="Luyện viết"><i class="fas fa-pen-fancy"></i></button>';
        }
        var fullBtn = '';
        if (r.zh) {
            fullBtn = '<button class="practice-full-btn" onclick="openPracticeFull(\'' + sttJs + '\', event)" title="Luyện tập full màn hình"><i class="fas fa-expand"></i></button>';
        }
        var practiceInput = '<input type="text" class="practice-input" placeholder="Gõ tiếng Trung..." data-answer="' + zhHtml + '" data-vi-hint="' + viHtml + '" data-stt="' + sttSafe + '" oninput="checkInput(this)" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">';
        
        mobHtml += '<div class="card" onclick="toggleFocus(\'' + sttJs + '\', this)" data-stt="' + sttSafe + '">' +
            '<div class="card-header">' +
                '<div class="card-stt">' + sttSafe + '</div>' +
                '<div class="card-meta">' +
                    (r.hsk ? '<span class="card-tag hsk">' + escapeHtml(r.hsk) + '</span>' : '') +
                    (r.topic ? '<span class="card-tag topic">' + escapeHtml(r.topic) + '</span>' : '') +
                    (r.subject ? '<span class="card-tag">' + escapeHtml(r.subject) + '</span>' : '') +
                '</div>' +
                '<div onclick="event.stopPropagation()" class="action-group">' + audio + writeBtn + fullBtn + '</div>' +
            '</div>' +
            '<div class="card-body">' +
                (r.vi ? '<div class="card-vi">' + viHtml + '</div>' : '') +
                '<div class="card-zh">' + zhHtml + '</div>' +
                (r.pinyin ? '<div class="card-pinyin">' + escapeHtml(r.pinyin) + '</div>' : '') +
            '</div>' +
            '<div class="card-practice" onclick="event.stopPropagation()">' +
                practiceInput +
                '<div class="card-check" data-check-stt="' + sttSafe + '"></div>' +
            '</div>' +
            '</div>';
    }
    mobileWrapper.insertAdjacentHTML('beforeend', mobHtml);
    renderedCount = end;
    
    var oldMobileBtn = mobileWrapper.querySelector('.load-more');
    if (oldMobileBtn) oldMobileBtn.remove();
    var oldEndNote = mobileWrapper.querySelector('.end-note');
    if (oldEndNote) oldEndNote.remove();
    
    if (renderedCount < filtered.length) {
        var isDemoLock = isDemo && renderedCount >= DEMO_LIMIT;
        if (!isDemoLock) {
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
            mobileWrapper.appendChild(lockedBtn);
        }
    } else if (filtered.length > PAGE_SIZE) {
        var endNote = document.createElement('div');
        endNote.className = 'end-note';
        endNote.innerHTML = '<i class="fas fa-check-circle"></i> Đã hiển thị tất cả ' + filtered.length + ' câu';
        mobileWrapper.appendChild(endNote);
    }
}

/* ============================================================
   🧠 SO KHỚP THÔNG MINH
   ============================================================ */
function normalizeAnswer(str) {
    if (!str) return '';
    return String(str)
        .replace(/[。，！？、；：""''「」『』（）《》〈〉【】〔〕]/g, '')
        .replace(/[.,!?;:'"()\[\]{}\-~`@#$%^&*+=|\\/<>]/g, '')
        .replace(/\s+/g, '')
        .toLowerCase()
        .trim();
}
function removeTones(str) {
    if (!str) return '';
    var map = {
        'ā':'a','á':'a','ǎ':'a','à':'a','ē':'e','é':'e','ě':'e','è':'e',
        'ī':'i','í':'i','ǐ':'i','ì':'i','ō':'o','ó':'o','ǒ':'o','ò':'o',
        'ū':'u','ú':'u','ǔ':'u','ù':'u','ǖ':'v','ǘ':'v','ǚ':'v','ǜ':'v','ü':'v'
    };
    return str.replace(/[āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü]/g, function(c) { return map[c] || c; });
}
function removeFillers(str) {
    if (!str) return '';
    var result = str;
    FILLER_WORDS.forEach(function(w) { result = result.split(w).join(''); });
    return result;
}
function expandSynonyms(str) {
    var results = [str];
    var keys = Object.keys(SYNONYMS);
    for (var i = 0; i < keys.length; i++) {
        var key = keys[i];
        if (str.indexOf(key) !== -1) {
            var values = SYNONYMS[key];
            for (var j = 0; j < values.length; j++) results.push(str.split(key).join(values[j]));
        }
    }
    return results;
}
function levenshtein(a, b) {
    if (a === b) return 0;
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    var matrix = [];
    for (var i = 0; i <= b.length; i++) matrix[i] = [i];
    for (var j = 0; j <= a.length; j++) matrix[0][j] = j;
    for (var i = 1; i <= b.length; i++) {
        for (var j = 1; j <= a.length; j++) {
            if (b.charAt(i-1) === a.charAt(j-1)) matrix[i][j] = matrix[i-1][j-1];
            else matrix[i][j] = Math.min(matrix[i-1][j-1] + 1, matrix[i][j-1] + 1, matrix[i-1][j] + 1);
        }
    }
    return matrix[b.length][a.length];
}
function similarity(a, b) {
    var maxLen = Math.max(a.length, b.length);
    if (maxLen === 0) return 1;
    return 1 - (levenshtein(a, b) / maxLen);
}
function smartCheck(userAnswer, correctAnswer) {
    var user = normalizeAnswer(userAnswer);
    var correct = normalizeAnswer(correctAnswer);
    if (!user) return { status: 'wrong', reason: '' };
    if (user === correct) return { status: 'correct', reason: 'Chính xác' };
    var userNoTone = removeTones(user);
    var correctNoTone = removeTones(correct);
    if (userNoTone === correctNoTone) return { status: 'correct', reason: 'Đúng (thiếu dấu thanh)' };
    var userNoFill = removeFillers(user);
    var correctNoFill = removeFillers(correct);
    if (userNoFill === correctNoFill) return { status: 'correct', reason: 'Đúng (bỏ qua từ phụ)' };
    var uNF = removeTones(userNoFill);
    var cNF = removeTones(correctNoFill);
    if (uNF === cNF) return { status: 'correct', reason: 'Đúng (từ phụ + dấu thanh)' };
    var userVariants = expandSynonyms(userNoFill);
    var correctVariants = expandSynonyms(correctNoFill);
    for (var i = 0; i < userVariants.length; i++) {
        for (var j = 0; j < correctVariants.length; j++) {
            if (userVariants[i] === correctVariants[j]) return { status: 'correct', reason: 'Đúng (từ đồng nghĩa)' };
        }
    }
    var maxSim = 0;
    for (var k = 0; k < correctVariants.length; k++) {
        var sim = similarity(userNoFill, correctVariants[k]);
        if (sim > maxSim) maxSim = sim;
    }
    for (var m = 0; m < userVariants.length; m++) {
        var sim2 = similarity(userVariants[m], correctNoFill);
        if (sim2 > maxSim) maxSim = sim2;
    }
    if (maxSim >= 0.85) return { status: 'partial', reason: 'Gần đúng (' + Math.round(maxSim * 100) + '%)' };
    if (user.indexOf(correct) !== -1 || correct.indexOf(user) !== -1) return { status: 'partial', reason: 'Thiếu/thừa từ' };
    return { status: 'wrong', reason: 'Không khớp' };
}

/* ============================================================
   🎯 TÁCH CỤM TỪ THEO PINYIN
   ============================================================ */
function countSyllables(pinyinWord) {
    if (!pinyinWord) return 0;
    var cleaned = pinyinWord.replace(/[.,!?;:'"()\[\]{}\-~`@#$%^&*+=|\\/<>。，！？、；：]/g, '').toLowerCase();
    if (!cleaned) return 0;
    var map = {
        'ā':'a','á':'a','ǎ':'a','à':'a',
        'ē':'e','é':'e','ě':'e','è':'e',
        'ī':'i','í':'i','ǐ':'i','ì':'i',
        'ō':'o','ó':'o','ǒ':'o','ò':'o',
        'ū':'u','ú':'u','ǔ':'u','ù':'u',
        'ǖ':'v','ǘ':'v','ǚ':'v','ǜ':'v','ü':'v'
    };
    cleaned = cleaned.replace(/[āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü]/g, function(c) {
        return map[c] || c;
    });
    var vowels = 'aeiou';
    var count = 0;
    var prevIsVowel = false;
    for (var i = 0; i < cleaned.length; i++) {
        var c = cleaned[i];
        var isVowel = vowels.indexOf(c) !== -1;
        if (isVowel && !prevIsVowel) count++;
        prevIsVowel = isVowel;
    }
    return count || 1;
}

function splitByPinyin(zh, pinyin) {
    if (!zh) return [];
    var hanziChars = [];
    for (var i = 0; i < zh.length; i++) {
        var c = zh[i];
        if (/[\u4e00-\u9fa5]/.test(c)) hanziChars.push(c);
    }
    if (hanziChars.length === 0) return [];
    if (!pinyin || !pinyin.trim()) {
        return hanziChars.map(function(c) { return { text: c, type: 'phrase' }; });
    }
    var pinyinGroups = pinyin.trim()
        .replace(/[.,!?;:'"()\[\]{}\-~`@#$%^&*+=|\\/<>。，！？、；：]/g, ' ')
        .split(/\s+/)
        .filter(function(w) { return w.length > 0; });
    if (pinyinGroups.length === 0) {
        return hanziChars.map(function(c) { return { text: c, type: 'phrase' }; });
    }
    var syllableCounts = pinyinGroups.map(function(w) { return countSyllables(w); });
    var totalSyllables = syllableCounts.reduce(function(a, b) { return a + b; }, 0);
    if (totalSyllables !== hanziChars.length) {
        return hanziChars.map(function(c) { return { text: c, type: 'phrase' }; });
    }
    var result = [];
    var idx = 0;
    for (var j = 0; j < syllableCounts.length; j++) {
        var cnt = syllableCounts[j];
        if (cnt <= 0) continue;
        var phrase = hanziChars.slice(idx, idx + cnt).join('');
        if (phrase) result.push({ text: phrase, type: 'phrase' });
        idx += cnt;
    }
    if (idx < hanziChars.length) {
        var remaining = hanziChars.slice(idx).join('');
        if (result.length > 0) result[result.length - 1].text += remaining;
        else result.push({ text: remaining, type: 'phrase' });
    }
    return result;
}

function updateInlinePreview(input, answer) {
    var wrapper = input.parentElement;
    var preview = wrapper.querySelector('.inline-char-preview');
    if (!preview) {
        preview = document.createElement('div');
        preview.className = 'inline-char-preview';
        input.insertAdjacentElement('afterend', preview);
    }
    var userVal = input.value.replace(/\s+/g, '');
    var cleanAnswer = (answer || '').replace(/\s+/g, '');
    if (!cleanAnswer) { preview.innerHTML = ''; return; }
    var html = '';
    var maxLen = Math.max(userVal.length, cleanAnswer.length);
    for (var i = 0; i < maxLen; i++) {
        var userChar = userVal[i] || '';
        var answerChar = cleanAnswer[i] || '';
        var cls = 'char-slot';
        var display = '';
        if (userChar && answerChar) {
            if (userChar === answerChar) { cls += ' correct'; display = userChar; }
            else { cls += ' wrong'; display = userChar; }
        } else if (!userChar && answerChar) { continue; }
        else if (userChar && !answerChar) { cls += ' extra'; display = userChar; }
        else { continue; }
        html += '<span class="' + cls + '">' + escapeHtml(display) + '</span>';
    }
    preview.innerHTML = html;
}

window.checkInput = function(input) {
    var stt = input.dataset.stt;
    var answer = input.dataset.answer;
    var cells = document.querySelectorAll('[data-check-stt="' + stt + '"]');
    var val = input.value.trim();
    updateInlinePreview(input, answer);
    if (!val) { cells.forEach(function(c) { c.innerHTML = ''; }); return; }
    var result = smartCheck(val, answer);
    var html = '';
    if (result.status === 'correct') html = '<span class="ai-correct">✅ ĐÚNG</span>';
    else if (result.status === 'partial') html = '<span class="ai-partial">⚠️ GẦN ĐÚNG</span>';
    else html = '<span class="ai-wrong">❌ SAI</span>';
    if (result.reason) html += '<span class="ai-reason">' + escapeHtml(result.reason) + '</span>';
    cells.forEach(function(c) { c.innerHTML = html; });
};

function applyFilter() {
    state.search = $('searchInput').value.trim().toLowerCase();
    state.hsk = $('hskFilter').value;
    state.subject = $('subjectFilter').value;
    updateFilterUI();
    var clearBtn = $('clearSearchBtn');
    if (state.search) clearBtn.classList.add('show');
    else clearBtn.classList.remove('show');
    var baseData = isDemo ? getDemoData() : RAW_DATA;
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
    updateResultCount();
    render(true);
}

/* ============================================================
   🎯 CHẾ ĐỘ LUYỆN TẬP FULL MÀN HÌNH
   ============================================================ */
var pfCurrentStt = null;
var pfCurrentAnswer = '';
var pfCurrentVi = '';
var pfCurrentPinyin = '';
var pfHintEnabled = false;

window.openPracticeFull = function(stt, evt) {
    if (evt) { evt.stopPropagation(); if (evt.preventDefault) evt.preventDefault(); }
    
    // ✅ Build filter options + dropdown cho modal trước khi mở
    pfBuildFilterOptions();
    pfBuildQuickNav();
    
    var idx = -1;
    for (var i = 0; i < filtered.length; i++) {
        if (String(filtered[i].stt) === String(stt)) { idx = i; break; }
    }
    if (idx === -1) { alert('Không tìm thấy câu!'); return; }
    pfCurrentStt = stt;
    $('practiceFullModal').classList.add('show');
    document.body.style.overflow = 'hidden';
    loadPracticeFull(stt);
};

window.closePracticeFull = function() {
    $('practiceFullModal').classList.remove('show');
    document.body.style.overflow = '';
    pfCurrentStt = null;
    if ('speechSynthesis' in window) speechSynthesis.cancel();
};

function loadPracticeFull(stt) {
    var idx = -1;
    for (var i = 0; i < filtered.length; i++) {
        if (String(filtered[i].stt) === String(stt)) { idx = i; break; }
    }
    if (idx === -1) return;
    
    // ✅ Đồng bộ filter modal nếu chưa có
    if (!$('pfSearchInput').value && !$('pfHskFilter').value && !$('pfSubjectFilter').value) {
        $('pfSearchInput').value = $('searchInput').value;
        $('pfHskFilter').value = $('hskFilter').value;
        $('pfSubjectFilter').value = $('subjectFilter').value;
    }
    pfUpdateFilterUI();
    
    var r = filtered[idx];
    pfCurrentStt = stt;
    pfCurrentAnswer = r.zh || '';
    pfCurrentVi = r.vi || '';
    pfCurrentPinyin = r.pinyin || '';
    
    $('pfCounter').textContent = 'Câu ' + (idx + 1) + ' / ' + filtered.length;
    
    var tagsHtml = '';
    if (r.hsk) tagsHtml += '<span class="card-tag hsk">' + escapeHtml(r.hsk) + '</span>';
    if (r.topic) tagsHtml += '<span class="card-tag topic">' + escapeHtml(r.topic) + '</span>';
    if (r.subject) tagsHtml += '<span class="card-tag">' + escapeHtml(r.subject) + '</span>';
    $('pfTags').innerHTML = tagsHtml;
    
    $('pfVi').textContent = pfCurrentVi;
    $('pfInput').value = '';
    $('pfStatus').textContent = '';
    $('pfStatus').className = 'practice-full-status';
    
    $('pfAnswer').classList.remove('show');
    var revealBtn = $('pfRevealBtn');
    revealBtn.classList.remove('revealed', 'hidden');
    revealBtn.innerHTML = '<i class="fas fa-eye"></i> Xem đáp án';
    
    pfHintEnabled = false;
    $('pfHintBtn').classList.remove('active');
    
    updateCharPreview();
    
    $('pfPrevBtn').disabled = (idx === 0);
    $('pfNextBtn').disabled = (idx === filtered.length - 1);
    
    // ✅ Cập nhật dropdown chọn câu
    var quickNav = $('pfQuickNav');
    if (quickNav && quickNav.value !== stt) {
        quickNav.value = stt;
    }
    
    setTimeout(function() { $('pfInput').focus(); }, 200);
}

window.pfNext = function() {
    if (!pfCurrentStt) return;
    var idx = -1;
    for (var i = 0; i < filtered.length; i++) {
        if (String(filtered[i].stt) === String(pfCurrentStt)) { idx = i; break; }
    }
    if (idx === -1 || idx >= filtered.length - 1) return;
    loadPracticeFull(filtered[idx + 1].stt);
};

window.pfPrev = function() {
    if (!pfCurrentStt) return;
    var idx = -1;
    for (var i = 0; i < filtered.length; i++) {
        if (String(filtered[i].stt) === String(pfCurrentStt)) { idx = i; break; }
    }
    if (idx <= 0) return;
    loadPracticeFull(filtered[idx - 1].stt);
};

/* ✅ Build dropdown chọn nhanh câu — hiển thị tiếng Việt */
function pfBuildQuickNav() {
    var sel = $('pfQuickNav');
    if (!sel) return;
    
    var html = '<option value="">-- Chọn câu (' + filtered.length + ') --</option>';
    filtered.forEach(function(r, i) {
        var vi = (r.vi || '').substring(0, 45);
        var label = 'Câu ' + (i + 1) + ': ' + vi;
        html += '<option value="' + escapeHtml(r.stt) + '">' + escapeHtml(label) + '</option>';
    });
    sel.innerHTML = html;
    
    // Nếu đang ở 1 câu cụ thể → set selected
    if (pfCurrentStt) {
        sel.value = pfCurrentStt;
    }
}
/* ✅ Cập nhật dropdown khi user chọn câu */
function pfQuickNavChange() {
    var sel = $('pfQuickNav');
    if (!sel) return;
    var stt = sel.value;
    if (!stt) return;
    loadPracticeFull(stt);
}

/* ✅ Build filter options cho modal full (khớp với filter chính) */
function pfBuildFilterOptions() {
    var hskSel = $('pfHskFilter');
    var subjSel = $('pfSubjectFilter');
    
    var allSubjectSet = {};
    RAW_DATA.forEach(function(r) { if (r.subject) allSubjectSet[r.subject] = 1; });
    var allSubjects = Object.keys(allSubjectSet).sort();
    
    if (isDemo) {
        var demoHsk = getDemoHskList();
        var hskHtml = '<option value="">Tất cả</option>';
        demoHsk.forEach(function(h) { hskHtml += '<option value="' + h + '">' + h + '</option>'; });
        var allHskList = ['HSK1','HSK2','HSK3','HSK4','HSK5','HSK6'];
        allHskList.forEach(function(h) {
            if (demoHsk.indexOf(h) === -1) hskHtml += '<option value="' + h + '" disabled>🔒 ' + h + '</option>';
        });
        hskSel.innerHTML = hskHtml;
        
        var demoSubjects = getDemoSubjectList();
        var demoSubjectMap = {};
        demoSubjects.forEach(function(s) { demoSubjectMap[s] = 1; });
        var unlocked = [];
        var locked = [];
        allSubjects.forEach(function(s) {
            if (demoSubjectMap[s]) unlocked.push(s);
            else locked.push(s);
        });
        var subjHtml = '<option value="">Tất cả chủ đề</option>';
        unlocked.forEach(function(s) { subjHtml += '<option value="' + escapeHtml(s) + '">' + escapeHtml(s) + '</option>'; });
        locked.forEach(function(s) { subjHtml += '<option value="' + escapeHtml(s) + '" disabled>🔒 ' + escapeHtml(s) + '</option>'; });
        subjSel.innerHTML = subjHtml;
    } else {
        hskSel.innerHTML = 
            '<option value="">Tất cả</option>' +
            '<option value="HSK1">HSK1</option>' +
            '<option value="HSK2">HSK2</option>' +
            '<option value="HSK3">HSK3</option>' +
            '<option value="HSK4">HSK4</option>' +
            '<option value="HSK5">HSK5</option>' +
            '<option value="HSK6">HSK6</option>';
        subjSel.innerHTML = '<option value="">Tất cả chủ đề</option>' +
            allSubjects.map(function(v){ return '<option value="'+escapeHtml(v)+'">'+escapeHtml(v)+'</option>'; }).join('');
    }
    
    // Đồng bộ với filter chính
    hskSel.value = $('hskFilter').value;
    subjSel.value = $('subjectFilter').value;
    $('pfSearchInput').value = $('searchInput').value;
    
    pfUpdateFilterUI();
}

/* ✅ Cập nhật UI filter trong modal */
function pfUpdateFilterUI() {
    var hsk = $('pfHskFilter').value;
    var subject = $('pfSubjectFilter').value;
    $('pfHskValue').textContent = hsk || 'Tất cả';
    $('pfSubjectValue').textContent = subject || 'Tất cả';
    $('pfHskChip').classList.toggle('has-value', !!hsk);
    $('pfSubjectChip').classList.toggle('has-value', !!subject);
    
    var search = $('pfSearchInput').value.trim();
    if (search) $('pfClearSearchBtn').classList.add('show');
    else $('pfClearSearchBtn').classList.remove('show');
    
    pfUpdateResultCount();
}

/* ✅ Hiển thị số kết quả trong modal */
function pfUpdateResultCount() {
    var el = $('pfResultCount');
    if (!el) return;
    var total = filtered ? filtered.length : 0;
    var search = $('pfSearchInput').value.trim();
    var hsk = $('pfHskFilter').value;
    var subject = $('pfSubjectFilter').value;
    var hasFilter = !!(search || hsk || subject);
    if (!hasFilter) { el.classList.remove('show', 'empty'); return; }
    el.classList.add('show');
    el.classList.toggle('empty', total === 0);
    var spanEl = el.querySelector('span');
    if (spanEl) {
        if (total === 0) spanEl.innerHTML = 'Không tìm thấy kết quả';
        else spanEl.innerHTML = 'Tìm thấy <b>' + total + '</b> kết quả';
    }
}

/* ✅ Áp dụng filter trong modal */
function pfApplyFilter() {
    $('searchInput').value = $('pfSearchInput').value;
    $('hskFilter').value = $('pfHskFilter').value;
    $('subjectFilter').value = $('pfSubjectFilter').value;
    
    state.search = $('pfSearchInput').value.trim().toLowerCase();
    state.hsk = $('pfHskFilter').value;
    state.subject = $('pfSubjectFilter').value;
    
    var baseData = isDemo ? getDemoData() : RAW_DATA;
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
    
    updateFilterUI();
    var clearBtn = $('clearSearchBtn');
    if (state.search) clearBtn.classList.add('show');
    else clearBtn.classList.remove('show');
    
    pfUpdateFilterUI();
    pfBuildQuickNav();
    render(true);
    
    if (filtered.length > 0) {
        loadPracticeFull(filtered[0].stt);
    } else {
        pfCurrentStt = null;
        pfCurrentAnswer = '';
        pfCurrentVi = 'Không tìm thấy câu nào';
        pfCurrentPinyin = '';
        $('pfVi').textContent = 'Không tìm thấy câu nào';
        $('pfInput').value = '';
        $('pfCounter').textContent = 'Câu 0 / 0';
        $('pfTags').innerHTML = '';
        $('pfAnswer').classList.remove('show');
        $('pfPreview').innerHTML = '';
        $('pfStatus').textContent = '';
        $('pfPrevBtn').disabled = true;
        $('pfNextBtn').disabled = true;
    }
}

function updateCharPreview() {
    var input = $('pfInput');
    var preview = $('pfPreview');
    var userVal = input.value;
    var cleanUser = userVal.replace(/\s+/g, '');
    var cleanAnswer = pfCurrentAnswer.replace(/\s+/g, '');
    
    if (!cleanAnswer) { preview.innerHTML = ''; return; }
    
    var html = '';
    var maxLen = Math.max(cleanUser.length, cleanAnswer.length);
    
    for (var i = 0; i < maxLen; i++) {
        var userChar = cleanUser[i] || '';
        var answerChar = cleanAnswer[i] || '';
        var cls = 'char-slot';
        var display = '';
        var clickable = false;
        
        if (userChar && answerChar) {
            if (userChar === answerChar) {
                cls += ' correct';
                display = userChar;
            } else {
                cls += ' wrong';
                display = userChar;
                clickable = true;
            }
        } else if (!userChar && answerChar) {
            if (pfHintEnabled) {
                cls += ' ghost';
                display = answerChar;
            } else {
                continue;
            }
        } else if (userChar && !answerChar) {
            cls += ' extra';
            display = userChar;
            clickable = true;
        } else {
            continue;
        }
        
        if (clickable) {
            html += '<span class="' + cls + '" data-idx="' + i + '" onclick="fixCharAt(' + i + ', this)">' + escapeHtml(display) + '</span>';
        } else {
            html += '<span class="' + cls + '">' + escapeHtml(display) + '</span>';
        }
    }
    
    preview.innerHTML = html;
}

window.fixCharAt = function(idx, el) {
    var input = $('pfInput');
    if (!input) return;
    input.focus();
    setTimeout(function() {
        try {
            input.setSelectionRange(idx, idx + 1);
        } catch(e) {
            input.selectionStart = idx;
            input.selectionEnd = idx + 1;
        }
        if (el) {
            el.classList.add('highlight');
            setTimeout(function() {
                el.classList.remove('highlight');
            }, 1200);
        }
    }, 10);
};

function toggleHint() {
    pfHintEnabled = !pfHintEnabled;
    var btn = $('pfHintBtn');
    if (pfHintEnabled) btn.classList.add('active');
    else btn.classList.remove('active');
    updateCharPreview();
}

function checkFullAnswer() {
    var input = $('pfInput');
    var statusEl = $('pfStatus');
    var val = input.value.trim();
    if (!val) {
        statusEl.textContent = '';
        statusEl.className = 'practice-full-status';
        return;
    }
    var result = smartCheck(val, pfCurrentAnswer);
    if (result.status === 'correct') {
        statusEl.textContent = '✅ ĐÚNG';
        statusEl.className = 'practice-full-status correct';
    } else if (result.status === 'partial') {
        statusEl.textContent = '⚠️ ' + (result.reason || 'GẦN ĐÚNG');
        statusEl.className = 'practice-full-status partial';
    } else {
        statusEl.textContent = '❌ SAI';
        statusEl.className = 'practice-full-status wrong';
    }
}

function revealFullAnswer() {
    var answerEl = $('pfAnswer');
    var revealBtn = $('pfRevealBtn');
    
    if (answerEl.classList.contains('show')) {
        answerEl.classList.remove('show');
        revealBtn.classList.remove('revealed');
        revealBtn.innerHTML = '<i class="fas fa-eye"></i> Xem đáp án';
        return;
    }
    
    var charsEl = $('pfAnswerChars');
    var pinyinEl = $('pfAnswerPinyin');
    
    charsEl.innerHTML = '';
    var phrases = splitByPinyin(pfCurrentAnswer, pfCurrentPinyin);
    
    if (phrases.length === 0) {
        pfCurrentAnswer.split('').forEach(function(c) {
            if (/[\u4e00-\u9fa5]/.test(c)) phrases.push({ text: c, type: 'phrase' });
        });
    }
    
    phrases.forEach(function(item) {
        var btn = document.createElement('button');
        btn.className = 'answer-phrase-btn';
        btn.textContent = item.text;
        btn.title = 'Nhấn để đọc: ' + item.text;
        
        btn.onclick = (function(text, el) {
            return function(e) {
                e.stopPropagation();
                el.classList.add('zoom-in');
                setTimeout(function() {
                    el.classList.remove('zoom-in');
                }, 700);
                speakPhrase(text, el);
            };
        })(item.text, btn);
        
        charsEl.appendChild(btn);
    });
    
    pinyinEl.textContent = pfCurrentPinyin;
    answerEl.classList.add('show');
    revealBtn.classList.add('revealed');
    revealBtn.innerHTML = '<i class="fas fa-eye-slash"></i> Ẩn đáp án';
}

window.speakPhrase = function(phrase, btn) {
    if (isDemo && !canUseFeature()) { showLimitMessage(); return; }
    if (!('speechSynthesis' in window)) { alert('Trình duyệt không hỗ trợ phát âm.'); return; }
    if (isDemo) { incDemoUsage(); updateDemoRemaining(); }
    speechSynthesis.cancel();
    document.querySelectorAll('.answer-phrase-btn.speaking').forEach(function(b) { b.classList.remove('speaking'); });
    btn.classList.add('speaking');
    var utterance = new SpeechSynthesisUtterance(phrase);
    utterance.lang = 'zh-CN';
    utterance.rate = 0.75;
    var voice = getChineseVoice();
    if (voice) utterance.voice = voice;
    utterance.onend = utterance.onerror = function() { btn.classList.remove('speaking'); };
    setTimeout(function(){ speechSynthesis.speak(utterance); }, 30);
};

window.speakFullSentence = function() {
    if (isDemo && !canUseFeature()) { showLimitMessage(); return; }
    if (!('speechSynthesis' in window)) return;
    if (isDemo) { incDemoUsage(); updateDemoRemaining(); }
    speechSynthesis.cancel();
    var utterance = new SpeechSynthesisUtterance(pfCurrentAnswer);
    utterance.lang = 'zh-CN';
    utterance.rate = 0.85;
    var voice = getChineseVoice();
    if (voice) utterance.voice = voice;
    setTimeout(function(){ speechSynthesis.speak(utterance); }, 30);
};

function initPracticeFull() {
    $('pfClose').addEventListener('click', closePracticeFull);
    $('pfPrevBtn').addEventListener('click', pfPrev);
    $('pfNextBtn').addEventListener('click', pfNext);
    $('pfRevealBtn').addEventListener('click', revealFullAnswer);
    $('pfHintBtn').addEventListener('click', toggleHint);
    $('pfInput').addEventListener('input', function() {
        updateCharPreview();
        checkFullAnswer();
    });
    
    // ✅ Dropdown chọn câu
    $('pfQuickNav').addEventListener('change', pfQuickNavChange);
    
    // ✅ Search trong modal
    $('pfSearchInput').addEventListener('input', function() {
        pfApplyFilter();
    });
    $('pfClearSearchBtn').addEventListener('click', function() {
        $('pfSearchInput').value = '';
        $('pfSearchInput').focus();
        pfApplyFilter();
    });
    
    // ✅ Filter HSK + Chủ đề trong modal
    $('pfHskFilter').addEventListener('change', function() {
        if (isDemo) {
            var val = this.value;
            var allowed = getDemoHskList();
            if (val && allowed.indexOf(val) === -1) {
                alert('🔒 Bản Demo chỉ cho phép lọc HSK1-' + DEMO_HSK_MAX + '.');
                this.value = '';
                return;
            }
        }
        pfApplyFilter();
    });
    $('pfSubjectFilter').addEventListener('change', function() {
        if (isDemo) {
            var val = this.value;
            var allowed = getDemoSubjectList();
            if (val && allowed.indexOf(val) === -1) {
                alert('🔒 Chủ đề này chưa có trong Demo.');
                this.value = '';
                return;
            }
        }
        pfApplyFilter();
    });
    
    document.addEventListener('keydown', function(e) {
        if (!$('practiceFullModal').classList.contains('show')) return;
        if (e.key === 'Escape') closePracticeFull();
        if (e.key === 'ArrowRight' && e.ctrlKey) pfNext();
        if (e.key === 'ArrowLeft' && e.ctrlKey) pfPrev();
    });
    
    var modal = $('practiceFullModal');
    var touchStartX = 0;
    modal.addEventListener('touchstart', function(e) {
        touchStartX = e.touches[0].clientX;
    }, { passive: true });
    modal.addEventListener('touchend', function(e) {
        var dx = e.changedTouches[0].clientX - touchStartX;
        if (Math.abs(dx) > 100) {
            if (dx < 0) pfNext();
            else pfPrev();
        }
    }, { passive: true });
}

/* ============================================================
   ✍️ LUYỆN VIẾT
   ============================================================ */
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
    if (isDemo && !canUseFeature()) { showLimitMessage(); return; }
    if (typeof HanziWriter === 'undefined') { alert('Thư viện chưa tải xong.'); return; }
    if (isDemo) { incDemoUsage(); updateDemoRemaining(); }
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

/* ============================================================
   👑 ADMIN PANEL
   ============================================================ */
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
        snapshot.forEach(function(doc) { usersCache.push({ email: doc.id, ...doc.data() }); });
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
    var adminCount = usersCache.filter(function(u) { return u.role === 'admin'; }).length;
    
    list.innerHTML = usersCache.map(function(u) {
        var isMe = u.email === currentUser.email;
        var isAdmin = u.role === 'admin';
        
        var roleBtn = '';
        if (isAdmin) {
            var reason = isMe ? 'Không thể tự hạ quyền chính mình' : 'Phải giữ đúng ' + TARGET_ADMINS + ' admin';
            roleBtn = '<button class="u-btn" disabled title="' + escapeHtml(reason) + '"><i class="fas fa-user"></i></button>';
        } else {
            if (adminCount < TARGET_ADMINS) roleBtn = '<button class="u-btn" onclick="changeRole(\'' + escapeJs(u.email) + '\', \'admin\')" title="Nâng lên Admin"><i class="fas fa-shield-alt"></i></button>';
            else roleBtn = '<button class="u-btn" disabled title="Đã đủ ' + TARGET_ADMINS + ' admin"><i class="fas fa-shield-alt"></i></button>';
        }
        
        var deleteBtn = '';
        if (isMe) deleteBtn = '<button class="u-btn danger" disabled title="Không thể tự xóa chính mình"><i class="fas fa-trash"></i></button>';
        else if (isAdmin) deleteBtn = '<button class="u-btn danger" disabled title="Phải giữ đúng ' + TARGET_ADMINS + ' admin"><i class="fas fa-trash"></i></button>';
        else deleteBtn = '<button class="u-btn danger" onclick="deleteUser(\'' + escapeJs(u.email) + '\')" title="Xóa"><i class="fas fa-trash"></i></button>';
        
        return '<div class="user-row">' +
            '<div class="u-info">' +
                '<div class="u-name">' + escapeHtml(u.name || u.email.split('@')[0]) + (isMe ? ' <span style="color:#94a3b8;font-size:.7rem">(bạn)</span>' : '') + '</div>' +
                '<div class="u-email">' + escapeHtml(u.email) + '</div>' +
            '</div>' +
            '<span class="u-role ' + (isAdmin ? 'admin' : 'user') + '">' + (u.role || 'user') + '</span>' +
            '<div class="u-actions">' + roleBtn + deleteBtn + '</div>' +
        '</div>';
    }).join('');
}

window.changeRole = async function(email, newRole) {
    var target = usersCache.find(function(u) { return u.email === email; });
    if (!target) { alert('Không tìm thấy user!'); return; }
    var isMe = email === currentUser.email;
    var isAdmin = target.role === 'admin';
    var adminCount = usersCache.filter(function(u) { return u.role === 'admin'; }).length;
    
    if (isMe && newRole === 'user') { alert('⚠️ Không thể tự hạ quyền admin của chính mình!'); return; }
    if (isAdmin && newRole === 'user') { alert('⚠️ Không thể hạ quyền admin!\n\nHệ thống phải giữ đúng ' + TARGET_ADMINS + ' admin.'); return; }
    if (!isAdmin && newRole === 'admin' && adminCount >= TARGET_ADMINS) { alert('⚠️ Đã có đủ ' + TARGET_ADMINS + ' admin!'); return; }
    
    var action = newRole === 'admin' ? 'NÂNG LÊN ADMIN' : 'HẠ XUỐNG USER';
    if (!confirm(action + ' cho tài khoản:\n\n' + email + '\n\nBạn có chắc không?')) return;
    try { await db.collection('allowed_users').doc(email).update({ role: newRole }); }
    catch(e) { alert('Lỗi: ' + e.message); }
};

window.deleteUser = async function(email) {
    var target = usersCache.find(function(u) { return u.email === email; });
    if (!target) { alert('Không tìm thấy user!'); return; }
    var isMe = email === currentUser.email;
    var isAdmin = target.role === 'admin';
    
    if (isMe) { alert('⚠️ Không thể tự xóa tài khoản của chính mình!'); return; }
    if (isAdmin) { alert('⚠️ Không thể xóa admin!'); return; }
    if (!confirm('⚠️ XÓA TÀI KHOẢN\n\n' + email + '\n\nNgười này sẽ không đăng nhập được nữa.\n\nBạn có chắc không?')) return;
    
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
    
    if (role === 'admin') {
        var currentAdminCount = usersCache.filter(function(u) { return u.role === 'admin'; }).length;
        if (currentAdminCount >= TARGET_ADMINS) { alert('⚠️ Đã có đủ ' + TARGET_ADMINS + ' admin!'); return; }
    }
    
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
html_output = (html_template
    .replace("__DATA__", json_data)
    .replace("__FIREBASE_CONFIG__", firebase_config_json)
    .replace("__ZALO_PHONE__", ZALO_PHONE)
    .replace("__ZALO_NAME__", ZALO_NAME)
    .replace("__DEMO_LIMIT__", str(DEMO_LIMIT))
    .replace("__DEMO_DAILY_LIMIT__", str(DEMO_DAILY_LIMIT))
    .replace("__DEMO_HSK_MAX__", str(DEMO_HSK_MAX))
    .replace("__TARGET_ADMINS__", str(TARGET_ADMINS))
    .replace("__SYNONYMS__", synonyms_json)
    .replace("__FILLER_WORDS__", fillers_json)
)
with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(html_output)

size_kb = os.path.getsize(OUTPUT_HTML) / 1024
print(f"\n🎉 Đã tạo: {OUTPUT_HTML}")
print(f"📦 Kích thước: {size_kb:.1f} KB")
print(f"📚 Tổng số câu: {len(data)}")
print(f"🎁 Demo: {DEMO_LIMIT} câu + HSK1-{DEMO_HSK_MAX} + {DEMO_DAILY_LIMIT} lượt")
print(f"🔥 Firebase: {FIREBASE_CONFIG.get('projectId', 'N/A')}")
print(f"👑 Chế độ: Đúng {TARGET_ADMINS} admin")
print(f"🧠 Chấm điểm: So khớp thông minh")
print(f"💡 Nút Gợi ý (mặc định TẮT) + Ghost text mờ")
print(f"📝 Đáp án tách theo PINYIN viết liền/rời")
print(f"☀️  Theme mặc định: Light mode")
print(f"🎯 Desktop hover phóng to + Mobile click vừa đọc vừa phóng to")
print(f"👁️  Nút Xem đáp án: toggle ẩn/hiện")
print(f"✏️  Click ký tự sai → con trỏ về + bôi đen để gõ đè")
print(f"🔍 Search + Filter + Dropdown chọn câu trong modal full màn hình")
print(f"📂 Chủ đề demo: mở khóa lên đầu, khóa xuống dưới")
