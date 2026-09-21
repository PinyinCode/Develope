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
- ✅ Dropdown chọn nhanh câu (hiển thị tiếng Việt)
- ✅ Cache user doc 12h → giảm 90% Firestore reads
- ✅ Admin panel load 1 lần (không realtime)
- ✅ Login log 1 lần/user/ngày
- Theme mặc định LIGHT MODE
- ✅ Nút ẩn/hiện kết quả ở mỗi thẻ trang chính (kèm đáp án tiếng Trung)
- ✅ Click chữ sai inline → bôi đen ĐÚNG ký tự trong ô gõ
- ✅ Gõ search trong modal KHÔNG nhảy sang ô nhập tiếng Trung
- ✅ Export CHUYÊN NGHIỆP: 7 cột, 3 sheet, style màu, sắp xếp theo hạn
- ✅ Import CHỈ USER, KHÔNG ADMIN + tự động bỏ qua cột dư thừa
- ✅ Import lần 2, 3, ... không lỗi null
- ✅ User có hạn sử dụng (expiresAt) — tự động khóa khi hết hạn
- ✅ Banner cảnh báo sắp hết hạn (≤ 7 ngày)
- ✅ Click avatar user → hiển thị chi tiết tài khoản + ngày hết hạn + progress bar
- ✅ Nút Zalo liên hệ trong dropdown user
- ✅ PHÂN QUYỀN ADMIN:
    + CHỈ email "hoanginvest@gmail.com" → SUPER ADMIN: full quyền, thấy tất cả,
      7 ô thống kê, thấy lịch sử đăng nhập, HẠ QUYỀN / XÓA ADMIN THƯỜNG
    + TẤT CẢ admin khác → ADMIN THƯỜNG: full quyền nhưng ẩn admin khác,
      chỉ thấy 4 ô (Tổng User, Đang hoạt động, Sắp hết hạn, Hết hạn),
      ẩn lịch sử truy cập, KHÔNG được hạ quyền / xóa admin khác
- ✅ Tổng User KHÔNG tính admin
- ✅ Nút đổi tên hiển thị (user + admin đều có)
- ✅ Quản lý ngày hết hạn ngay trên giao diện admin
- ✅ Nút chỉnh hạn sử dụng = BIỂU TƯỢNG LỊCH (icon only + tooltip + pulse đỏ khi ≤7 ngày)
- ✅ Header LỚN + TikTok info bar + Header KHÔNG biến mất khi luyện tập
- ✅ Chế độ thường hiện Zalo + TikTok; Chế độ luyện tập chỉ hiện TikTok
- ✅ PC hover nút TikTok → hiện card info TikTok (avatar + nickname + username)
- ✅ Avatar TikTok nhập trong config.json
- ✅ Mobile 1 tap TikTok = card info; 2 tap = mở TikTok
- ✅ Nút FAB "Silent mode" (CHUÔNG BẬT 🔔 / CHUÔNG TẮT 🔕) ẩn/hiện Zalo + TikTok + TikTok bar
         (CHỈ áp dụng cho user/admin, KHÔNG áp dụng demo)
         → Bình thường: 🔔 XANH (active) — đang BẬT thông báo
         → Silent mode: 🔕 XÁM (chưa active) — đang TẮT thông báo
- ✅ FIX: Bật chế độ luyện tập full màn hình KHÔNG còn trùng lặp search/HSK/chủ đề
- ✅ FIX: Modal luyện tập set top 1 lần theo chiều cao header, không cập nhật liên tục
- ✅ FIX: Bỏ tooltip ::after của FAB → không còn chấm đen khi hover/nhấn

Chạy: python scripts/convert.py
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
SUPER_ADMIN = CONFIG.get("super_admin", "hoanginvest@gmail.com")
ZALO_PHONE = CONFIG.get("zalo_phone", "")
ZALO_NAME = CONFIG.get("zalo_name", "Hỗ trợ")
TIKTOK_USERNAME = CONFIG.get("tiktok_username", "thaonoizhongwen")
TIKTOK_NICKNAME = CONFIG.get("tiktok_nickname", "Thảo nói 中文")
TIKTOK_AVATAR = CONFIG.get("tiktok_avatar", "")
TIKTOK_URL = CONFIG.get("tiktok_url", f"https://www.tiktok.com/@{TIKTOK_USERNAME}")
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
print(f"   🎵 TikTok: @{TIKTOK_USERNAME} ({TIKTOK_NICKNAME})")
print(f"   🖼️  TikTok Avatar: {'Có' if TIKTOK_AVATAR else 'Không (dùng fallback)'}")
print(f"   🎁 Demo: {DEMO_LIMIT} câu + HSK1-{DEMO_HSK_MAX} + {DEMO_DAILY_LIMIT} lượt nghe/viết")
print(f"   🧠 Chấm điểm: So khớp thông minh")
print(f"   ☀️  Theme mặc định: Light mode")
print(f"   💡 Nút Gợi ý (mặc định TẮT) + Ghost text mờ")
print(f"   📝 Đáp án tách theo PINYIN viết liền/rời")
print(f"   🎯 Click ký tự sai → bôi đen để gõ đè")
print(f"   📂 Chủ đề demo: mở khóa lên đầu, khóa xuống dưới")
print(f"   🔍 Search + Filter + Dropdown chọn câu trong modal")
print(f"   💾 Cache user 12h + Log 1 lần/ngày")
print(f"   👑 Super admin: {SUPER_ADMIN}")
print(f"   🕵️  Admin thường: 4 ô thống kê + ẩn admin khác")
print(f"   👥 Tổng User KHÔNG tính admin")
print(f"   👁️  Nút ẩn/hiện kết quả + đáp án ở mỗi thẻ trang chính")
print(f"   📤 Export CHUYÊN NGHIỆP (7 cột + 3 sheet + style)")
print(f"   📥 Import CHỈ USER, KHÔNG ADMIN")
print(f"   ⏰ User có hạn sử dụng (expiresAt)")
print(f"   👤 Click avatar → hiển thị chi tiết + ngày hết hạn")
print(f"   ✏️  Nút đổi tên hiển thị (user + admin)")
print(f"   📅 Quản lý ngày hết hạn trên giao diện admin")
print(f"   🔔🔕 Nút FAB 'Silent mode' (chuông BẬT/TẮT) ẩn/hiện Zalo+TikTok+TikTokBar")
print(f"      → Bình thường: 🔔 XANH (active)")
print(f"      → Silent mode: 🔕 XÁM (chưa active)")
print(f"   🚫 FIX: Không còn trùng lặp search/HSK/chủ đề khi luyện tập full")
print(f"   🚫 FIX: Không còn chấm đen tooltip trên FAB")

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
<title>Học tiếng Trung · Văn phòng & Công xưởng</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
<script src="https://www.gstatic.com/firebasejs/10.7.0/firebase-app-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.7.0/firebase-auth-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.7.0/firebase-firestore-compat.js"></script>
<script src="https://cdn.jsdelivr.net/npm/hanzi-writer@3.5.0/dist/hanzi-writer.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
:root{
    --bg:#f0f4f8;--surface:#fff;--surface-2:#f8fafc;--border:#e2e8f0;--border-strong:#cbd5e1;
    --text:#0f172a;--text-2:#475569;--text-3:#94a3b8;
    --primary:#2563eb;--primary-dark:#1d4ed8;--primary-light:#dbeafe;
    --success:#16a34a;--danger:#dc2626;--danger-light:#fee2e2;
    --amber:#f59e0b;--amber-light:#fef3c7;
    --zalo:#0068ff;
    --tiktok:#000;
    --tiktok-pink:#fe2c55;
    --tiktok-cyan:#25f4ee;
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
    --tiktok:#fff;
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
    padding:.6rem 0 .7rem 0;transition:background .2s, box-shadow .2s, border-color .2s;
    border-bottom:1px solid transparent;
}
.sticky-top.scrolled{
    background:var(--surface);border-bottom-color:var(--border);
    box-shadow:0 4px 16px -8px rgba(15,23,42,.15);
}
[data-theme="dark"] .sticky-top.scrolled{box-shadow:0 4px 16px -8px rgba(0,0,0,.5)}

/* ============ HEADER LỚN ============ */
.header{background:transparent;border:none}
.header-inner{display:flex;align-items:center;gap:.75rem;margin-bottom:.5rem}
.logo{display:flex;align-items:center;gap:.85rem;flex:1;min-width:0}
.logo-icon{
    width:64px;height:64px;background:linear-gradient(135deg,#2563eb,#7c3aed 60%,#db2777);
    border-radius:16px;display:flex;align-items:center;justify-content:center;
    color:#fff;font-size:2rem;flex-shrink:0;
    box-shadow:0 8px 24px rgba(37,99,235,.4), inset 0 1px 0 rgba(255,255,255,.25);
    position:relative;overflow:hidden;
}
.logo-icon::after{
    content:'';position:absolute;inset:0;
    background:radial-gradient(circle at 30% 20%, rgba(255,255,255,.35), transparent 60%);
    pointer-events:none;
}
.logo-text{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:flex;flex-direction:column;line-height:1.15;min-width:0}
.logo-text .title{
    font-size:2rem;
    font-weight:900;
    letter-spacing:-.02em;
    line-height:1.15;
    color:var(--text);
}
.logo-text .subtitle{
    font-size:.95rem;
    color:var(--text-3);
    font-weight:700;
    margin-top:3px;
    letter-spacing:.01em;
    padding-left:1.1rem;
}
.header-actions{display:flex;gap:.4rem;align-items:center;flex-shrink:0}

/* ============ TIKTOK INFO BAR (dưới header) ============ */
.tiktok-bar{
    display:flex;align-items:center;gap:.6rem;
    padding:.5rem .85rem;margin-bottom:.6rem;
    background:linear-gradient(135deg, rgba(254,44,85,.06), rgba(37,244,238,.06));
    border:1.5px solid var(--border);border-radius:var(--radius-full);
    font-size:.8rem;color:var(--text-2);
    width:fit-content;max-width:100%;
    transition:.15s;box-shadow:var(--shadow-sm);
    text-decoration:none;
}
[data-theme="dark"] .tiktok-bar{
    background:linear-gradient(135deg, rgba(254,44,85,.1), rgba(37,244,238,.1));
}
.tiktok-bar:hover{border-color:var(--text-2);box-shadow:0 4px 12px rgba(0,0,0,.08);transform:translateY(-1px);}
.tiktok-bar-avatar{
    width:32px;height:32px;border-radius:50%;
    object-fit:cover;flex-shrink:0;
    border:2px solid var(--border);
    background:linear-gradient(135deg,#fe2c55,#25f4ee);
}
.tiktok-bar-info{display:flex;flex-direction:column;line-height:1.2;min-width:0;flex:1}
.tiktok-bar-info .tiktok-nick{font-weight:800;color:var(--text);font-size:.92rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tiktok-bar-info .tiktok-user{font-size:.75rem;color:var(--text-3);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:flex;align-items:center;gap:.25rem}
.tiktok-bar-info .tiktok-user i{font-size:.7rem;}
.tiktok-bar-link{
    display:inline-flex;align-items:center;gap:.35rem;
    padding:.4rem .8rem;border-radius:50px;
    background:linear-gradient(135deg,#000,#333);color:#fff !important;
    font-size:.75rem;font-weight:700;text-decoration:none;
    white-space:nowrap;transition:.15s;flex-shrink:0;
    box-shadow:0 2px 8px rgba(0,0,0,.2);
}
[data-theme="dark"] .tiktok-bar-link{background:linear-gradient(135deg,#fff,#e5e5e5);color:#000 !important;}
.tiktok-bar-link:hover{transform:translateY(-1px);box-shadow:0 6px 16px rgba(0,0,0,.3);}
.tiktok-bar-link i{font-size:.85rem;}

.icon-btn{
    width:38px;height:38px;border-radius:10px;border:1px solid var(--border);
    background:var(--surface);color:var(--text-3);cursor:pointer;
    display:flex;align-items:center;justify-content:center;font-size:.9rem;
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
    padding:.4rem .75rem;border-radius:50px;
    background:var(--amber-light);color:#92400e;
    font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.3px;
    border:1px solid rgba(245,158,11,.4);
}
[data-theme="dark"] .demo-badge{color:#fcd34d}

.user-menu{position:relative}
.user-avatar{
    width:38px;height:38px;border-radius:50%;border:2px solid var(--border);
    cursor:pointer;object-fit:cover;transition:.15s;display:block;
}
.user-avatar:hover{border-color:var(--primary);transform:scale(1.05)}
.user-dropdown{
    position:absolute;top:calc(100% + .5rem);right:0;
    background:var(--surface);border:1px solid var(--border);
    border-radius:var(--radius);box-shadow:0 10px 30px rgba(0,0,0,.15);
    padding:.5rem;min-width:290px;display:none;z-index:200;
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

.user-details{
    padding:.6rem .75rem .75rem .75rem;
    border-bottom:1px solid var(--border);
    margin-bottom:.5rem;
    display:flex;flex-direction:column;gap:.6rem;
}
.detail-row{display:flex;align-items:flex-start;gap:.65rem;}
.detail-icon{
    width:36px;height:36px;border-radius:10px;
    display:flex;align-items:center;justify-content:center;
    font-size:.95rem;flex-shrink:0;
    background:var(--surface-2);color:var(--text-2);
    transition:.2s;
}
.detail-icon.ok{background:rgba(22,163,74,.12);color:var(--success);}
.detail-icon.warn{background:rgba(245,158,11,.15);color:#d97706;}
.detail-icon.urgent{background:rgba(220,38,38,.15);color:var(--danger);}
.detail-icon.permanent{background:var(--primary-light);color:var(--primary-dark);}
.detail-icon.expired{background:rgba(220,38,38,.25);color:var(--danger);}

[data-theme="dark"] .detail-icon.ok{background:rgba(22,163,74,.25);color:#4ade80;}
[data-theme="dark"] .detail-icon.warn{background:rgba(245,158,11,.25);color:#fcd34d;}
[data-theme="dark"] .detail-icon.urgent{background:rgba(220,38,38,.3);color:#fca5a5;}
[data-theme="dark"] .detail-icon.permanent{background:rgba(59,130,246,.25);color:#93c5fd;}
[data-theme="dark"] .detail-icon.expired{background:rgba(220,38,38,.35);color:#fca5a5;}

.detail-content{flex:1;min-width:0;}
.detail-label{
    font-size:.65rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.3px;
    color:var(--text-3);margin-bottom:.15rem;
}
.detail-value{
    font-size:.85rem;font-weight:700;
    color:var(--text);line-height:1.3;
    word-break:break-word;
}
.detail-value.ok{color:var(--success);}
.detail-value.warn{color:#d97706;}
[data-theme="dark"] .detail-value.warn{color:#fcd34d;}
.detail-value.urgent{color:var(--danger);}
.detail-value.permanent{color:var(--primary-dark);}
[data-theme="dark"] .detail-value.permanent{color:#93c5fd;}
.detail-value.expired{color:var(--danger);text-decoration:line-through;}

.detail-sub{
    font-size:.7rem;color:var(--text-3);
    margin-top:.15rem;line-height:1.35;
}
.detail-sub b{color:var(--text-2);font-weight:700;}

.detail-progress{margin-top:.15rem;}
.progress-track{
    width:100%;height:6px;
    background:var(--surface-2);
    border-radius:50px;overflow:hidden;
    border:1px solid var(--border);
}
.progress-bar{
    height:100%;border-radius:50px;
    transition:width .4s ease, background .3s ease;
    background:linear-gradient(90deg, #16a34a, #22c55e);
}
.progress-bar.ok{background:linear-gradient(90deg, #16a34a, #22c55e);}
.progress-bar.warn{background:linear-gradient(90deg, #f59e0b, #fbbf24);}
.progress-bar.urgent{background:linear-gradient(90deg, #dc2626, #ef4444);}
.progress-bar.permanent{background:linear-gradient(90deg, #2563eb, #3b82f6);}

.dropdown-zalo{
    display:flex;align-items:center;gap:.5rem;width:100%;
    padding:.65rem .75rem;border-radius:var(--radius-sm);
    background:linear-gradient(135deg, #0068ff, #0084ff);
    color:#fff !important;font-size:.85rem;font-weight:700;
    text-decoration:none;transition:.15s;
    margin-top:.25rem;
}
.dropdown-zalo:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(0,104,255,.3);color:#fff !important;}
.dropdown-zalo i{font-size:1rem;}

.dropdown-tiktok{
    display:flex;align-items:center;gap:.5rem;width:100%;
    padding:.65rem .75rem;border-radius:var(--radius-sm);
    background:linear-gradient(135deg, #000, #333);
    color:#fff !important;font-size:.85rem;font-weight:700;
    text-decoration:none;transition:.15s;
    margin-top:.25rem;
}
[data-theme="dark"] .dropdown-tiktok{background:linear-gradient(135deg,#fff,#e5e5e5);color:#000 !important;}
.dropdown-tiktok:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(0,0,0,.3);}
.dropdown-tiktok i{font-size:1rem;}

.btn-login-header{
    display:flex;align-items:center;gap:.4rem;
    padding:.55rem 1rem;border-radius:50px;
    background:var(--primary);color:#fff;border:none;
    font-size:.85rem;font-weight:700;cursor:pointer;
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

/* ============ FLOATING LEFT GROUP (Zalo + TikTok) ============ */
.floating-left-group{
    position:fixed;bottom:calc(20px + env(safe-area-inset-bottom));
    left:20px;z-index:1000;display:flex;flex-direction:column;
    gap:.5rem;align-items:flex-start;
}

.zalo-btn{
    display:flex;align-items:center;gap:.5rem;
    padding:.7rem 1.1rem;border-radius:50px;
    background:linear-gradient(135deg, #0068ff, #0084ff);
    color:#fff;text-decoration:none;font-weight:700;font-size:.85rem;
    box-shadow:0 8px 24px rgba(0,104,255,.4);
    transition:all .35s cubic-bezier(.34,1.56,.64,1);
    -webkit-tap-highlight-color:transparent;white-space:nowrap;
    border:2px solid #fff;font-family:inherit;overflow:hidden;
    position:relative;
}
.zalo-btn:hover,.zalo-btn:active{transform:scale(1.05);box-shadow:0 12px 32px rgba(0,104,255,.55);color:#fff}
.zalo-btn i{font-size:1.2rem;flex-shrink:0;line-height:1;position:relative;z-index:2}
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

/* ============ TIKTOK FLOAT BUTTON + HOVER CARD ============ */
.tiktok-float-wrap{position:relative;}
.tiktok-float-btn{
    display:flex;align-items:center;gap:.5rem;
    padding:.7rem 1.1rem;border-radius:50px;
    background:linear-gradient(135deg, #000, #2d2d2d);
    color:#fff;text-decoration:none;font-weight:700;font-size:.85rem;
    box-shadow:0 8px 24px rgba(0,0,0,.4), 0 0 0 2px rgba(254,44,85,.4);
    transition:all .35s cubic-bezier(.34,1.56,.64,1);
    -webkit-tap-highlight-color:transparent;white-space:nowrap;
    border:2px solid #fff;font-family:inherit;overflow:hidden;
    position:relative;cursor:pointer;
}
.tiktok-float-btn:hover,.tiktok-float-btn:active{transform:scale(1.05);box-shadow:0 12px 32px rgba(0,0,0,.5), 0 0 0 3px rgba(254,44,85,.6);color:#fff}
.tiktok-float-btn i{font-size:1.2rem;flex-shrink:0;line-height:1;position:relative;z-index:2}
.tiktok-float-btn .tiktok-text{
    line-height:1.15;display:flex;flex-direction:column;
    position:relative;z-index:2;transition:opacity .2s, max-width .35s;
    max-width:200px;overflow:hidden;
}
.tiktok-float-btn .tiktok-label{font-size:.65rem;opacity:.85;font-weight:500;white-space:nowrap}
.tiktok-float-btn .tiktok-name{font-size:.85rem;font-weight:700;white-space:nowrap}
.tiktok-float-btn::before{
    content:'';position:absolute;inset:0;border-radius:50px;
    background:linear-gradient(135deg, #fe2c55, #25f4ee);
    opacity:.35;z-index:1;animation:tiktokPulse 2s infinite;
}
@keyframes tiktokPulse{
    0%{transform:scale(1);opacity:.35}
    50%{transform:scale(1.1);opacity:0}
    100%{transform:scale(1);opacity:0}
}
.tiktok-float-btn.compact{width:52px;height:52px;padding:0;border-radius:50%;justify-content:center;gap:0}
.tiktok-float-btn.compact .tiktok-text{opacity:0;max-width:0}
.tiktok-float-btn.compact i{font-size:1.35rem}
.tiktok-float-btn.compact::before{border-radius:50%}

/* ✅ HOVER CARD thông tin TikTok */
.tiktok-hover-card{
    position:absolute;
    bottom:calc(100% + 10px);
    left:0;
    width:300px;
    background:var(--surface);
    border:1px solid var(--border);
    border-radius:16px;
    box-shadow:0 20px 50px rgba(0,0,0,.25);
    padding:1rem;
    opacity:0;visibility:hidden;
    transform:translateY(8px) scale(.96);
    transition:opacity .25s, visibility .25s, transform .25s cubic-bezier(.34,1.56,.64,1);
    pointer-events:none;
    z-index:2000;
}
.tiktok-float-wrap:hover .tiktok-hover-card{
    opacity:1;visibility:visible;
    transform:translateY(0) scale(1);
    pointer-events:auto;
}
.tiktok-hover-card::after{
    content:'';position:absolute;top:100%;left:32px;
    width:16px;height:16px;
    background:var(--surface);
    border-right:1px solid var(--border);
    border-bottom:1px solid var(--border);
    transform:translateY(-8px) rotate(45deg);
}
/* 📱 Mobile: chỉ hiện khi có class .show-mobile */
@media(hover:none) and (pointer:coarse){
    .tiktok-float-wrap:hover .tiktok-hover-card{
        opacity:0;
        visibility:hidden;
        transform:translateY(8px) scale(.96);
        pointer-events:none;
    }
    .tiktok-float-wrap.show-mobile .tiktok-hover-card{
        opacity:1;
        visibility:visible;
        transform:translateY(0) scale(1);
        pointer-events:auto;
        left:-4px;
        width:min(300px, calc(100vw - 32px));
    }
}
.thc-header{display:flex;align-items:center;gap:.75rem;margin-bottom:.85rem}
.thc-avatar-wrap{
    position:relative;flex-shrink:0;
}
.thc-avatar{
    width:60px;height:60px;border-radius:50%;
    object-fit:cover;
    border:3px solid transparent;
    background:linear-gradient(var(--surface),var(--surface)) padding-box,
               linear-gradient(135deg,#fe2c55,#25f4ee) border-box;
    box-shadow:0 4px 12px rgba(0,0,0,.15);
}
.thc-avatar-fallback{
    width:60px;height:60px;border-radius:50%;
    background:linear-gradient(135deg,#fe2c55,#25f4ee);
    color:#fff;display:flex;align-items:center;justify-content:center;
    font-size:1.7rem;font-weight:800;flex-shrink:0;
    border:3px solid var(--surface);
    box-shadow:0 4px 12px rgba(0,0,0,.15);
}
.thc-verified{
    position:absolute;bottom:-2px;right:-2px;
    width:20px;height:20px;border-radius:50%;
    background:#20d5ec;color:#fff;
    display:flex;align-items:center;justify-content:center;
    font-size:.65rem;border:2.5px solid var(--surface);
}
.thc-info{flex:1;min-width:0}
.thc-nick{
    font-size:1rem;font-weight:800;color:var(--text);
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
    margin-bottom:.15rem;line-height:1.2;
}
.thc-user{
    font-size:.78rem;color:var(--text-3);
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
    display:flex;align-items:center;gap:.25rem;
}
.thc-user i{font-size:.72rem;color:var(--tiktok-pink)}
.thc-desc{
    font-size:.75rem;color:var(--text-2);
    padding:.6rem 0;border-top:1px solid var(--border);
    border-bottom:1px solid var(--border);
    margin-bottom:.75rem;line-height:1.5;
}
.thc-stats{
    display:flex;gap:1rem;margin-bottom:.75rem;
    padding-bottom:.75rem;border-bottom:1px solid var(--border);
}
.thc-stat{text-align:center;flex:1}
.thc-stat .num{font-size:.95rem;font-weight:800;color:var(--text);line-height:1;}
.thc-stat .label{font-size:.65rem;color:var(--text-3);text-transform:uppercase;letter-spacing:.3px;margin-top:.25rem;font-weight:600;}
.thc-btn{
    display:flex;align-items:center;justify-content:center;gap:.4rem;
    width:100%;padding:.7rem;
    border-radius:10px;border:none;
    background:linear-gradient(135deg,#fe2c55,#ff0050);
    color:#fff;font-size:.85rem;font-weight:700;
    text-decoration:none;cursor:pointer;
    transition:.15s;font-family:inherit;
    box-shadow:0 4px 12px rgba(254,44,85,.3);
}
.thc-btn:hover{transform:translateY(-2px);box-shadow:0 6px 18px rgba(254,44,85,.45);color:#fff}
.thc-btn i{font-size:.95rem}

/* ============ FAB GROUP (right) ============ */
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
.fab-btn:hover{transform:scale(1.08);background:var(--primary-light);color:var(--primary-dark)}
.fab-btn:active{transform:scale(0.95);background:var(--primary-light);color:var(--primary-dark)}
.fab-btn.active{background:var(--primary);color:#fff;border-color:var(--primary);box-shadow:0 8px 24px rgba(37,99,235,.4)}
.fab-btn.active:hover{background:var(--primary-dark);color:#fff}
.fab-btn.active:active{background:var(--primary-dark);color:#fff;transform:scale(0.95)}

/* ✅ ĐÃ BỎ TOOLTIP ::after → không còn chấm đen khi hover/nhấn */

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
.fab-group.open .fab-sub:nth-child(4){transition-delay:.2s}

/* ✅ Icon nút Silent nhỏ hơn một chút cho đẹp */
.fab-focus i {
    font-size: 1rem !important;
}

/* ✅ Nút Silent khi active DÙNG MÀU XANH GIỐNG các nút FAB khác
   (đã tự kế thừa từ .fab-btn.active phía trên, không cần override)
   → Bình thường: 🔔 XANH (active)
   → Silent mode: 🔕 XÁM (chưa active) */

/* ✅ Ẩn floating buttons + TikTok bar khi user bật chế độ Silent */
body.hide-floating .floating-left-group,
body.hide-floating .tiktok-bar {
    display: none !important;
}

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

.answer-inline-display{
    display:flex;align-items:center;justify-content:center;
    gap:.4rem;margin-top:.35rem;padding:.35rem .6rem;
    background:linear-gradient(135deg, rgba(37,99,235,.08), rgba(37,99,235,.04));
    border:1px dashed rgba(37,99,235,.3);
    border-radius:8px;flex-wrap:wrap;
}
[data-theme="dark"] .answer-inline-display{
    background:linear-gradient(135deg, rgba(59,130,246,.15), rgba(59,130,246,.08));
    border-color:rgba(59,130,246,.4);
}
.answer-inline-label{
    font-size:.7rem;font-weight:700;color:var(--primary-dark);
    text-transform:uppercase;letter-spacing:.3px;
    display:inline-flex;align-items:center;gap:.25rem;white-space:nowrap;
}
[data-theme="dark"] .answer-inline-label{color:#93c5fd;}
.answer-inline-label i{font-size:.75rem;}
.answer-inline-text{
    font-family:var(--font-zh);font-size:1.05rem;font-weight:600;
    color:var(--text);letter-spacing:.03em;word-break:break-all;
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

.expiry-banner{
    background:linear-gradient(135deg, #fef3c7, #fde68a);
    border:1.5px solid #f59e0b;border-radius:var(--radius);
    padding:.85rem 1.1rem;margin-bottom:1rem;
    display:flex;align-items:center;gap:.75rem;flex-wrap:wrap;
}
[data-theme="dark"] .expiry-banner{
    background:linear-gradient(135deg, rgba(245,158,11,.15), rgba(245,158,11,.25));
}
.expiry-banner.urgent{
    background:linear-gradient(135deg, #fecaca, #fca5a5);
    border-color:#dc2626;
    animation:pulseUrgent 2s infinite;
}
[data-theme="dark"] .expiry-banner.urgent{
    background:linear-gradient(135deg, rgba(220,38,38,.2), rgba(220,38,38,.3));
}
@keyframes pulseUrgent{
    0%,100%{box-shadow:0 0 0 0 rgba(220,38,38,.4);}
    50%{box-shadow:0 0 0 8px rgba(220,38,38,0);}
}
.expiry-banner-icon{
    width:36px;height:36px;border-radius:50%;
    background:#f59e0b;color:#fff;
    display:flex;align-items:center;justify-content:center;font-size:1rem;flex-shrink:0;
}
.expiry-banner.urgent .expiry-banner-icon{background:#dc2626;}
.expiry-banner-text{flex:1;min-width:200px}
.expiry-banner-text .title{font-weight:700;font-size:.9rem;color:#92400e;margin-bottom:.15rem;}
[data-theme="dark"] .expiry-banner-text .title{color:#fcd34d;}
.expiry-banner.urgent .expiry-banner-text .title{color:#991b1b;}
[data-theme="dark"] .expiry-banner.urgent .expiry-banner-text .title{color:#fca5a5;}
.expiry-banner-text .desc{font-size:.78rem;color:#78350f;line-height:1.5;}
[data-theme="dark"] .expiry-banner-text .desc{color:#fde68a;}
.expiry-banner-text b{color:#dc2626;}
.expiry-banner.urgent .expiry-banner-text b{color:#991b1b;}
.expiry-banner-btn{
    padding:.5rem .9rem;border-radius:50px;border:none;
    background:#f59e0b;color:#fff;text-decoration:none;
    font-size:.8rem;font-weight:700;cursor:pointer;transition:.15s;
    display:inline-flex;align-items:center;gap:.35rem;white-space:nowrap;
}
.expiry-banner-btn:hover{background:#d97706;color:#fff;transform:translateY(-1px);}
.expiry-banner.urgent .expiry-banner-btn{background:#dc2626;}
.expiry-banner.urgent .expiry-banner-btn:hover{background:#b91c1c;}

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
.card-check{
    font-size:.75rem;font-weight:700;
    min-width:55px;text-align:center;
    width:100%;
}

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

.toggle-check-btn{
    width:32px;height:32px;border-radius:50%;border:none;
    background:var(--surface-2);color:var(--text-2);
    cursor:pointer;display:inline-flex;align-items:center;justify-content:center;
    font-size:.8rem;transition:.15s;flex-shrink:0;
    border:1px solid var(--border);
}
.toggle-check-btn:hover,.toggle-check-btn:active{
    background:var(--primary-light);color:var(--primary-dark);
    border-color:var(--primary);
}
.toggle-check-btn.active{
    background:var(--primary);color:#fff;border-color:var(--primary);
}
.toggle-check-btn.active:hover{
    background:var(--primary-dark);color:#fff;
}

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

/* ============ PRACTICE FULL MODAL ============ */
.practice-full-modal{
    position:fixed;
    left:0;right:0;bottom:0;
    background:var(--bg);z-index:2500;
    display:none;flex-direction:column;animation:fadeIn .2s;
    top:0;
    overflow:hidden;
}
.practice-full-modal.show{display:flex}

body.practice-full-open .search-bar,
body.practice-full-open .filters,
body.practice-full-open .result-count {
    display: none !important;
}
body.practice-full-open .demo-banner,
body.practice-full-open .expiry-banner {
    display: none !important;
}

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

.pf-quick-nav{
    margin-top:.5rem;
    display:flex;
    align-items:center;
    gap:.5rem;
}
.pf-quick-nav-label{
    font-size:.7rem;font-weight:700;color:var(--text-3);
    text-transform:uppercase;letter-spacing:.3px;white-space:nowrap;
}
.pf-quick-nav-select{
    flex:1;min-width:0;padding:.5rem 2rem .5rem .8rem;
    border-radius:var(--radius-full);border:1.5px solid var(--border);
    background:var(--bg);color:var(--text);font-size:.82rem;
    font-family:inherit;outline:none;cursor:pointer;transition:.15s;
    -webkit-appearance:none;appearance:none;
    background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 12 12'><path fill='%2394a3b8' d='M6 9L1 4h10z'/></svg>");
    background-repeat:no-repeat;background-position:right 12px center;background-size:10px;
}
.pf-quick-nav-select:focus{
    border-color:var(--primary);box-shadow:0 0 0 3px rgba(37,99,235,.15);
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
    border-color:var(--primary);box-shadow:0 0 0 4px rgba(37,99,235,.15);
}
@media(min-width:769px){.practice-full-input{font-size:1.85rem;padding:1.15rem 1.5rem;}}

.char-preview{
    display:flex;justify-content:center;flex-wrap:wrap;
    gap:.5rem;min-height:2.5rem;padding:.75rem 1rem;
    background:var(--surface-2);border-radius:12px;
    border:1px dashed var(--border);
    user-select:none;-webkit-user-select:none;
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
    color:var(--text);font-weight:700;
    background:rgba(22,163,74,.12);
}
.char-slot.wrong{
    color:#fff;background:var(--danger);
    animation:shakeWrong .3s;
    box-shadow:0 2px 8px rgba(220,38,38,.35);
    cursor:pointer;position:relative;
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
    color:var(--text);opacity:.12;font-weight:400;
    background:transparent;user-select:none;
    pointer-events:none;filter:blur(0.3px);
}
[data-theme="dark"] .char-slot.ghost{color:var(--text);opacity:.15;}
.char-slot.extra{
    color:#fff;background:var(--amber);
    box-shadow:0 2px 8px rgba(245,158,11,.35);
    cursor:pointer;transition:.15s;
}
.char-slot.extra:hover{
    transform:scale(1.15);
    box-shadow:0 4px 12px rgba(245,158,11,.5);
    z-index:5;
}
.char-slot.ghost-missing{
    color:var(--text-3);opacity:.5;background:transparent;
    border:1px dashed var(--border);
    cursor:pointer;font-size:1.2rem;font-weight:400;
}
.char-slot.ghost-missing:hover{
    opacity:.9;border-color:var(--primary);
    color:var(--primary);transform:scale(1.1);
}
.char-slot.ghost-missing.highlight{
    animation:blinkHighlightMissing 0.6s ease-in-out 2;
}
@keyframes blinkHighlightMissing{
    0%,100%{transform:scale(1.1);border-color:var(--primary);color:var(--primary);}
    50%{transform:scale(1.25);border-color:var(--primary);color:var(--primary);box-shadow:0 0 0 6px rgba(37,99,235,.2);}
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
.inline-char-preview .char-slot.ghost-missing{
    font-size:.9rem;min-width:1.1rem;height:1.5rem;
}

.practice-full-status{
    text-align:center;font-size:1rem;font-weight:700;min-height:1.5rem;
}
.practice-full-status.correct{color:var(--success);}
.practice-full-status.partial{color:var(--amber);}
.practice-full-status.wrong{color:var(--danger);}

.reveal-actions{
    display:grid;grid-template-columns:1fr 1fr;gap:.75rem;
}
.reveal-actions button{
    width:100%;padding:.85rem;
    border-radius:14px;border:2px dashed var(--border-strong);
    background:var(--surface);color:var(--text-2);
    font-size:.9rem;font-weight:600;cursor:pointer;
    transition:all .2s ease;font-family:inherit;
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
               background .2s,color .2s,border-color .2s,box-shadow .2s;
    display:inline-flex;align-items:center;justify-content:center;
    -webkit-appearance:none;transform-origin:center center;
}
@media(hover:hover) and (pointer:fine){
    .answer-phrase-btn:hover{
        transform:scale(1.35);
        background:var(--primary);color:#fff;
        border-color:var(--primary);
        box-shadow:0 8px 24px rgba(37,99,235,.4);
        z-index:10;
    }
}
.answer-phrase-btn.zoom-in{
    transform:scale(1.35);
    background:var(--primary);color:#fff;
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
.admin-header-actions{display:flex;gap:.5rem;align-items:center;flex-wrap:wrap}
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
.admin-section-title i{margin-right:.3rem;}
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
.user-row .u-role.super{
    background:linear-gradient(135deg, #f59e0b, #d97706);
    color:#fff;
    box-shadow:0 2px 6px rgba(245,158,11,.4);
}
.user-row .u-actions{display:flex;gap:.3rem}
.u-btn{
    width:32px;height:32px;border-radius:8px;border:1px solid var(--border);
    background:var(--surface);color:var(--text-2);cursor:pointer;
    display:flex;align-items:center;justify-content:center;
    font-size:.8rem;transition:.15s;
}
.u-btn:hover{background:var(--surface-2);color:var(--primary);border-color:var(--primary)}
.u-btn.danger:hover{background:var(--danger-light);color:var(--danger);border-color:var(--danger)}

.u-btn.expiry{
    background:rgba(245,158,11,.1);
    color:#d97706;
    border-color:rgba(245,158,11,.4);
}
.u-btn.expiry:hover{
    background:var(--amber);
    color:#fff;
    border-color:var(--amber);
    transform:scale(1.08);
}
.u-btn.expiry:active{transform:scale(.95);}
[data-theme="dark"] .u-btn.expiry{
    background:rgba(245,158,11,.2);
    color:#fcd34d;
    border-color:rgba(245,158,11,.5);
}
[data-theme="dark"] .u-btn.expiry:hover{background:var(--amber);color:#fff;}
.u-btn.expiry.urgent{
    background:rgba(220,38,38,.15);
    color:var(--danger);
    border-color:rgba(220,38,38,.5);
    animation:expiryPulse 2s infinite;
}
.u-btn.expiry.urgent:hover{background:var(--danger);color:#fff;border-color:var(--danger);}
[data-theme="dark"] .u-btn.expiry.urgent{
    background:rgba(220,38,38,.3);
    color:#fca5a5;
    border-color:rgba(220,38,38,.6);
}
@keyframes expiryPulse{
    0%,100%{box-shadow:0 0 0 0 rgba(220,38,38,.4);}
    50%{box-shadow:0 0 0 6px rgba(220,38,38,0);}
}

.u-btn:disabled{opacity:.35;cursor:not-allowed}
.u-btn:disabled:hover{background:var(--surface);color:var(--text-2);border-color:var(--border);transform:none;}
.u-btn.danger:disabled:hover{background:var(--surface);color:var(--text-2);border-color:var(--border);}
.u-btn.expiry:disabled:hover{background:var(--surface);color:var(--text-2);border-color:var(--border);transform:none;}

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
.btn:disabled{opacity:.5;cursor:not-allowed;}
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

.u-last-login{
    font-size:.68rem;color:var(--text-3);
    display:flex;align-items:center;gap:.25rem;margin-top:.2rem;
}
.u-last-login.active{color:var(--success);}
.u-last-login.recent{color:var(--primary);}
.u-last-login i{font-size:.65rem;}
.u-expiry{
    font-size:.68rem;font-weight:600;
    display:inline-flex;align-items:center;gap:.25rem;
    margin-top:.2rem;padding:.15rem .45rem;border-radius:50px;
    cursor:pointer;transition:.15s;
}
.u-expiry:hover{opacity:.8;}
.u-expiry i{font-size:.6rem;}
.u-expiry.permanent{background:rgba(148,163,184,.15);color:var(--text-3);}
.u-expiry.ok{background:rgba(22,163,74,.12);color:var(--success);}
.u-expiry.warn{background:rgba(245,158,11,.15);color:#92400e;}
[data-theme="dark"] .u-expiry.warn{color:#fcd34d;}
.u-expiry.urgent{background:rgba(220,38,38,.15);color:var(--danger);}
.u-expiry.expired{background:rgba(220,38,38,.25);color:#fff;text-decoration:line-through;}

.edit-modal{
    position:fixed;inset:0;background:rgba(15,23,42,.85);
    backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);
    z-index:4000;display:none;align-items:center;justify-content:center;
    padding:1rem;animation:fadeIn .2s;
}
.edit-modal.show{display:flex}
.edit-box{
    background:var(--surface);border-radius:20px;
    width:100%;max-width:420px;box-shadow:0 20px 60px rgba(0,0,0,.4);
    padding:1.5rem;position:relative;
    animation:slideUp .3s cubic-bezier(.34,1.56,.64,1);
}
.edit-box h2{
    font-size:1.1rem;color:var(--text);font-weight:700;
    display:flex;align-items:center;gap:.5rem;margin-bottom:1.25rem;
}
.edit-box h2 i{color:var(--primary);}
.edit-close{
    position:absolute;top:12px;right:12px;
    width:32px;height:32px;border-radius:8px;border:1px solid var(--border);
    background:var(--surface);color:var(--text-2);cursor:pointer;
    display:flex;align-items:center;justify-content:center;
    font-size:.85rem;transition:.15s;
}
.edit-close:hover{background:var(--danger-light);color:var(--danger);border-color:var(--danger)}
.edit-user-info{
    padding:.75rem;background:var(--surface-2);
    border-radius:10px;margin-bottom:1rem;
    border:1px solid var(--border);
}
.edit-user-info .eu-name{font-weight:700;font-size:.9rem;color:var(--text);margin-bottom:.2rem}
.edit-user-info .eu-email{font-size:.75rem;color:var(--text-3);word-break:break-all}

.quick-expiry-btns{
    display:grid;grid-template-columns:repeat(3,1fr);
    gap:.4rem;margin-bottom:1rem;
}
.quick-expiry-btn{
    padding:.5rem .4rem;border-radius:8px;
    border:1.5px solid var(--border);background:var(--surface);
    color:var(--text-2);font-size:.72rem;font-weight:600;
    cursor:pointer;transition:.15s;font-family:inherit;
    display:flex;flex-direction:column;align-items:center;gap:.2rem;
    white-space:nowrap;
}
.quick-expiry-btn i{font-size:.85rem;}
.quick-expiry-btn:hover{
    background:var(--primary-light);border-color:var(--primary);
    color:var(--primary-dark);transform:translateY(-1px);
}
.quick-expiry-btn.danger:hover{
    background:var(--danger-light);border-color:var(--danger);color:var(--danger);
}

.import-modal{
    position:fixed;inset:0;background:rgba(15,23,42,.85);
    backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);
    z-index:3500;display:none;align-items:center;justify-content:center;
    padding:1rem;animation:fadeIn .2s;
}
.import-modal.show{display:flex}
.import-box{
    background:var(--surface);border-radius:20px;
    width:100%;max-width:900px;max-height:calc(100vh - 2rem);
    box-shadow:0 20px 60px rgba(0,0,0,.4);
    display:flex;flex-direction:column;overflow:hidden;
}
.import-header{
    padding:1.25rem 1.5rem;border-bottom:1px solid var(--border);
    display:flex;align-items:center;justify-content:space-between;
}
.import-header h2{
    font-size:1.15rem;color:var(--text);
    display:flex;align-items:center;gap:.5rem;font-weight:700;
}
.import-header h2 i{color:#16a34a;}
.import-body{padding:1.25rem 1.5rem;overflow-y:auto;flex:1;min-height:0;}
.import-summary{
    display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));
    gap:.6rem;margin-bottom:1rem;
}
.import-stat{
    padding:.65rem .85rem;border-radius:10px;text-align:center;
    border:1px solid var(--border);background:var(--surface-2);
}
.import-stat .num{font-size:1.5rem;font-weight:800;line-height:1;margin-bottom:.25rem;}
.import-stat .label{font-size:.7rem;color:var(--text-3);text-transform:uppercase;font-weight:600;}
.import-stat.ok .num{color:var(--success);}
.import-stat.update .num{color:var(--primary);}
.import-stat.warn .num{color:var(--amber);}
.import-stat.err .num{color:var(--danger);}
.import-preview-wrap{
    max-height:400px;overflow-y:auto;
    border:1px solid var(--border);border-radius:10px;
    background:var(--surface-2);
}
.import-table{width:100%;border-collapse:collapse;font-size:.82rem;}
.import-table th{
    padding:.6rem .8rem;text-align:left;
    background:var(--surface);color:var(--text-2);
    font-size:.7rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.3px;
    border-bottom:2px solid var(--border);
    position:sticky;top:0;z-index:2;
}
.import-table td{
    padding:.55rem .8rem;color:var(--text);
    border-bottom:1px solid var(--border);word-break:break-word;
}
.import-table tr:last-child td{border-bottom:none;}
.import-table tr.row-error{background:rgba(220,38,38,.08);}
.import-table tr.row-warn{background:rgba(245,158,11,.08);}
.import-table tr.row-new{background:rgba(22,163,74,.05);}
.import-table tr.row-update{background:rgba(37,99,235,.05);}
.import-table .status-badge{
    display:inline-flex;align-items:center;gap:.3rem;
    padding:.2rem .5rem;border-radius:50px;
    font-size:.68rem;font-weight:700;white-space:nowrap;
}
.import-table .status-badge.ok{background:rgba(22,163,74,.15);color:var(--success);}
.import-table .status-badge.update{background:rgba(37,99,235,.15);color:var(--primary);}
.import-table .status-badge.warn{background:rgba(245,158,11,.15);color:#92400e;}
.import-table .status-badge.err{background:rgba(220,38,38,.15);color:var(--danger);}
.import-table .role-badge{
    display:inline-block;padding:.15rem .5rem;border-radius:50px;
    font-size:.68rem;font-weight:700;text-transform:uppercase;
}
.import-table .role-badge.admin{background:var(--amber-light);color:#92400e;}
.import-table .role-badge.user{background:var(--primary-light);color:var(--primary-dark);}
.import-options{
    display:flex;gap:1rem;margin-top:1rem;flex-wrap:wrap;
}
.import-options label{
    display:flex;align-items:center;gap:.4rem;
    font-size:.82rem;font-weight:600;color:var(--text-2);
    cursor:pointer;user-select:none;
}
.import-options input[type="checkbox"]{
    width:16px;height:16px;accent-color:var(--primary);cursor:pointer;
}
.import-info{
    margin-top:1rem;padding:.65rem .85rem;
    border-radius:10px;
    background:var(--primary-light);color:var(--primary-dark);
    font-size:.78rem;line-height:1.6;
    display:flex;align-items:flex-start;gap:.5rem;
}
[data-theme="dark"] .import-info{color:#93c5fd;}
.import-info i{margin-top:.15rem;flex-shrink:0;}
.import-info b{font-weight:800;}
.import-footer{
    padding:1rem 1.5rem;border-top:1px solid var(--border);
    display:flex;gap:.5rem;justify-content:flex-end;align-items:center;
    background:var(--surface);
}

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

/* ============ RESPONSIVE ============ */
@media(max-width:768px){
    .container{padding:0 .7rem}
    .header-inner{gap:.5rem;margin-bottom:.4rem}
    .logo-icon{width:52px;height:52px;font-size:1.5rem;border-radius:13px}
    .logo-text .title{font-size:1.5rem}
    .logo-text .subtitle{font-size:.82rem;padding-left:.8rem}
    .icon-btn{width:34px;height:34px;font-size:.8rem}
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
    .user-dropdown{min-width:260px;right:-8px}
    .btn-login-header{padding:.45rem .7rem;font-size:.75rem}
    .btn-login-header span{display:none}
    
    .floating-left-group{
        left:16px;bottom:calc(16px + env(safe-area-inset-bottom));
        gap:.45rem;
    }
    .zalo-btn{padding:.6rem .9rem;font-size:.8rem;gap:.4rem;}
    .zalo-btn i{font-size:1.05rem}
    .zalo-btn .zalo-label{font-size:.6rem}
    .zalo-btn .zalo-name{font-size:.78rem;max-width:80px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .zalo-btn.compact{width:48px;height:48px}
    .zalo-btn.compact i{font-size:1.25rem}
    
    .tiktok-float-btn{
        padding:0;
        border-radius:50%;
        width:48px;
        height:48px;
        justify-content:center;
        gap:0;
    }
    .tiktok-float-btn i{font-size:1.35rem}
    .tiktok-float-btn .tiktok-text{display:none}
    .tiktok-float-btn::before{border-radius:50%}
    
    .tiktok-bar{padding:.4rem .7rem;gap:.5rem;}
    .tiktok-bar-avatar{width:28px;height:28px;}
    .tiktok-bar-info .tiktok-nick{font-size:.85rem;}
    .tiktok-bar-info .tiktok-user{font-size:.7rem;}
    .tiktok-bar-link{padding:.35rem .65rem;font-size:.7rem;}
    
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
    
    .import-box{max-height:calc(100vh - 1rem);border-radius:16px;}
    .import-header,.import-body,.import-footer{padding:1rem;}
    .import-table th,.import-table td{padding:.45rem .5rem;font-size:.75rem;}
    
    .quick-expiry-btns{grid-template-columns:repeat(3,1fr);gap:.35rem;}
    .quick-expiry-btn{padding:.45rem .3rem;font-size:.68rem;}
    .quick-expiry-btn i{font-size:.8rem;}
}
@media(max-width:400px){
    .logo-icon{width:46px;height:46px;font-size:1.35rem;border-radius:12px}
    .logo-text .title{font-size:1.35rem}
    .logo-text .subtitle{display:none}
    .icon-btn{width:32px;height:32px;font-size:.75rem}
    .zalo-btn .zalo-text{display:none}
    .zalo-btn{padding:0;border-radius:50%;width:48px;height:48px;justify-content:center}
    .zalo-btn i{font-size:1.2rem}
    .tiktok-bar-link span{display:none}
    .tiktok-bar-link{padding:.4rem .5rem;}
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
                        <div class="subtitle">Văn phòng &amp; Công xưởng</div>
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
                            
                            <div class="user-details" id="userDetails" style="display:none">
                                <div class="detail-row" id="expiryRow">
                                    <div class="detail-icon" id="expiryIconWrap">
                                        <i class="fas fa-calendar-check" id="expiryIcon"></i>
                                    </div>
                                    <div class="detail-content">
                                        <div class="detail-label">Hạn sử dụng</div>
                                        <div class="detail-value" id="expiryValue">-</div>
                                        <div class="detail-sub" id="expirySub"></div>
                                    </div>
                                </div>
                                
                                <div class="detail-progress" id="expiryProgressWrap" style="display:none">
                                    <div class="progress-track">
                                        <div class="progress-bar" id="expiryProgressBar"></div>
                                    </div>
                                </div>
                            </div>
                            
                            <button class="dropdown-item" id="changeNameBtn">
                                <i class="fas fa-user-edit"></i> Đổi tên hiển thị
                            </button>
                            
                            <button class="dropdown-item" id="openAdminBtn" style="display:none">
                                <i class="fas fa-shield-alt"></i> Quản lý tài khoản
                            </button>
                            <a class="dropdown-zalo" id="dropdownZaloBtn" href="#" target="_blank" rel="noopener noreferrer">
                                <i class="fas fa-comment-dots"></i> Liên hệ Zalo hỗ trợ
                            </a>
                            <a class="dropdown-tiktok" id="dropdownTiktokBtn" href="#" target="_blank" rel="noopener noreferrer">
                                <i class="fab fa-tiktok"></i> Theo dõi TikTok
                            </a>
                            <button class="dropdown-item danger" id="logoutBtn">
                                <i class="fas fa-sign-out-alt"></i> Đăng xuất
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </header>

        <!-- ✅ TIKTOK INFO BAR (dưới header) -->
        <a class="tiktok-bar" id="tiktokBar" href="#" target="_blank" rel="noopener noreferrer" title="Theo dõi TikTok">
            <img class="tiktok-bar-avatar" id="tiktokBarAvatar" src="" alt="TikTok">
            <div class="tiktok-bar-info">
                <div class="tiktok-nick" id="tiktokBarNick">Thảo nói 中文</div>
                <div class="tiktok-user"><i class="fab fa-tiktok"></i> <span id="tiktokBarUser">@thaonoizhongwen</span></div>
            </div>
            <span class="tiktok-bar-link">
                <i class="fab fa-tiktok"></i> <span>Theo dõi</span>
            </span>
        </a>

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

<!-- ✅ FLOATING LEFT GROUP: Zalo + TikTok -->
<div class="floating-left-group" id="floatingLeftGroup">
    <a class="zalo-btn" id="zaloBtn" href="#" target="_blank" rel="noopener noreferrer" title="Liên hệ Zalo hỗ trợ">
        <i class="fas fa-comment-dots"></i>
        <div class="zalo-text">
            <span class="zalo-label">Liên hệ Zalo</span>
            <span class="zalo-name">Hỗ trợ</span>
        </div>
    </a>
    
    <div class="tiktok-float-wrap" id="tiktokFloatWrap">
        <a class="tiktok-float-btn" id="tiktokFloatBtn" href="#" target="_blank" rel="noopener noreferrer" title="Theo dõi TikTok">
            <i class="fab fa-tiktok"></i>
            <div class="tiktok-text">
                <span class="tiktok-label">Theo dõi TikTok</span>
                <span class="tiktok-name" id="tiktokFloatName">Thảo nói 中文</span>
            </div>
        </a>
        
        <!-- ✅ HOVER CARD TikTok (PC hover / Mobile tap) -->
        <div class="tiktok-hover-card" id="tiktokHoverCard">
            <div class="thc-header">
                <div class="thc-avatar-wrap">
                    <img class="thc-avatar" id="thcAvatar" src="" alt="TikTok" style="display:none">
                    <div class="thc-avatar-fallback" id="thcAvatarFallback">T</div>
                    <div class="thc-verified"><i class="fas fa-check"></i></div>
                </div>
                <div class="thc-info">
                    <div class="thc-nick" id="thcNick">Thảo nói 中文</div>
                    <div class="thc-user"><i class="fab fa-tiktok"></i> <span id="thcUser">@thaonoizhongwen</span></div>
                </div>
            </div>
            <div class="thc-desc">Học tiếng Trung mỗi ngày cùng Thảo · Văn phòng & Công xưởng 🇨🇳</div>
            <div class="thc-stats">
                <div class="thc-stat">
                    <div class="num" id="thcFollowers">1.2K</div>
                    <div class="label">Followers</div>
                </div>
                <div class="thc-stat">
                    <div class="num" id="thcLikes">15.6K</div>
                    <div class="label">Likes</div>
                </div>
                <div class="thc-stat">
                    <div class="num" id="thcVideos">128</div>
                    <div class="label">Videos</div>
                </div>
            </div>
            <a class="thc-btn" id="thcFollowBtn" href="#" target="_blank" rel="noopener noreferrer">
                <i class="fab fa-tiktok"></i> Theo dõi ngay
            </a>
        </div>
    </div>
</div>

<div class="fab-group" id="fabGroup" style="display:none">
    <button class="fab-btn fab-sub" id="toggleViBtn" title="Ẩn/hiện Tiếng Việt">
        <i class="fas fa-language"></i>
    </button>
    <button class="fab-btn fab-sub" id="togglePinyinBtn" title="Ẩn/hiện Pinyin">
        <i class="fas fa-spell-check"></i>
    </button>
    <button class="fab-btn fab-sub" id="togglePracticeBtn" title="Ẩn/hiện Ô luyện dịch">
        <i class="fas fa-keyboard"></i>
    </button>
    <!-- ✅ Nút Silent: CHUÔNG BẬT 🔔 (XANH) / CHUÔNG TẮT 🔕 (XÁM) -->
    <button class="fab-btn fab-sub fab-focus" id="toggleFocusBtn" title="Click để tắt Zalo/TikTok (Silent mode)">
        <i class="fas fa-bell"></i>
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

        <div class="expiry-banner" id="expiryBanner" style="display:none">
            <div class="expiry-banner-icon"><i class="fas fa-hourglass-half"></i></div>
            <div class="expiry-banner-text">
                <div class="title">Tài khoản sắp hết hạn</div>
                <div class="desc">
                    Còn <b id="expiryDaysText">7</b> ngày (đến <b id="expiryDateText">-</b>). Liên hệ Admin để gia hạn!
                </div>
            </div>
            <a class="expiry-banner-btn" id="expiryContactBtn" href="#" target="_blank" rel="noopener noreferrer">
                <i class="fas fa-comment-dots"></i> Liên hệ
            </a>
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

<div class="edit-modal" id="changeNameModal">
    <div class="edit-box">
        <button class="edit-close" id="changeNameClose"><i class="fas fa-times"></i></button>
        <h2><i class="fas fa-user-edit"></i> Đổi tên hiển thị</h2>
        <div class="edit-user-info">
            <div class="eu-name" id="changeNameCurrent">-</div>
            <div class="eu-email" id="changeNameEmail">-</div>
        </div>
        <div class="form-group">
            <label>Tên mới</label>
            <input type="text" id="changeNameInput" placeholder="Nhập tên mới..." maxlength="50">
        </div>
        <div class="form-actions">
            <button class="btn" id="changeNameCancel">Hủy</button>
            <button class="btn primary" id="changeNameConfirm">
                <i class="fas fa-check"></i> Lưu
            </button>
        </div>
    </div>
</div>

<div class="edit-modal" id="editExpiryModal">
    <div class="edit-box">
        <button class="edit-close" id="editExpiryClose"><i class="fas fa-times"></i></button>
        <h2><i class="fas fa-calendar-edit"></i> Chỉnh hạn sử dụng</h2>
        <div class="edit-user-info">
            <div class="eu-name" id="editExpiryName">-</div>
            <div class="eu-email" id="editExpiryEmail">-</div>
        </div>
        
        <div class="quick-expiry-btns">
            <button class="quick-expiry-btn" onclick="setQuickExpiry(7)">
                <i class="fas fa-calendar-plus"></i> +7 ngày
            </button>
            <button class="quick-expiry-btn" onclick="setQuickExpiry(30)">
                <i class="fas fa-calendar-plus"></i> +30 ngày
            </button>
            <button class="quick-expiry-btn" onclick="setQuickExpiry(90)">
                <i class="fas fa-calendar-plus"></i> +90 ngày
            </button>
            <button class="quick-expiry-btn" onclick="setQuickExpiry(180)">
                <i class="fas fa-calendar-plus"></i> +6 tháng
            </button>
            <button class="quick-expiry-btn" onclick="setQuickExpiry(365)">
                <i class="fas fa-calendar-plus"></i> +1 năm
            </button>
            <button class="quick-expiry-btn danger" onclick="setQuickExpiryPermanent()">
                <i class="fas fa-infinity"></i> Vĩnh viễn
            </button>
        </div>
        
        <div class="form-group">
            <label>Hoặc chọn ngày cụ thể</label>
            <input type="date" id="editExpiryInput">
        </div>
        <div class="form-actions">
            <button class="btn" id="editExpiryCancel">Hủy</button>
            <button class="btn primary" id="editExpiryConfirm">
                <i class="fas fa-check"></i> Lưu
            </button>
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

<div class="practice-full-modal" id="practiceFullModal">
    <div class="practice-full-header">
        <div class="pf-counter" id="pfCounter">Câu 1 / 1</div>
        <div class="pf-tags" id="pfTags"></div>
        <button class="pf-close" id="pfClose" aria-label="Đóng">
            <i class="fas fa-times"></i>
        </button>
    </div>
    
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

<div class="import-modal" id="importModal">
    <div class="import-box">
        <div class="import-header">
            <h2><i class="fas fa-file-import"></i> Import danh sách user</h2>
            <button class="admin-close" id="importClose"><i class="fas fa-times"></i></button>
        </div>
        <div class="import-body">
            <div class="import-summary" id="importSummary"></div>
            <div class="import-preview-wrap">
                <table class="import-table">
                    <thead>
                        <tr>
                            <th style="width:45px">#</th>
                            <th>Email</th>
                            <th>Tên</th>
                            <th style="width:80px">Vai trò</th>
                            <th style="width:110px">Hạn dùng</th>
                            <th style="width:130px">Trạng thái</th>
                        </tr>
                    </thead>
                    <tbody id="importTableBody"></tbody>
                </table>
            </div>
            <div class="import-options">
                <label>
                    <input type="checkbox" id="importSkipDuplicates">
                    Bỏ qua user đã tồn tại (không update)
                </label>
                <label>
                    <input type="checkbox" id="importSkipInvalid" checked>
                    Bỏ qua dòng không hợp lệ
                </label>
            </div>
            <div class="import-info">
                <i class="fas fa-info-circle"></i>
                <div>
                    File Excel cần có cột: <b>email</b> (bắt buộc), <b>name</b> (tùy chọn), <b>expiresAt</b> (tùy chọn).
                    <br>• <b>email</b>: bắt buộc, phải hợp lệ
                    <br>• <b>name</b>: nếu thiếu sẽ lấy phần trước @ của email
                    <br>• <b>expiresAt</b>: định dạng <b>YYYY-MM-DD</b>, để trống = vĩnh viễn
                    <br>• <b>Các cột khác</b> (STT, ghi chú, phone...) sẽ được <b>tự động bỏ qua</b>
                    <br><b>⚠️ Lưu ý:</b> Chỉ import <b>user</b>, KHÔNG import admin.
                </div>
            </div>
        </div>
        <div class="import-footer">
            <button class="btn" id="importCancelBtn">Hủy</button>
            <button class="btn primary" id="importConfirmBtn">
                <i class="fas fa-check"></i> Import <span id="importCount">0</span> user
            </button>
        </div>
    </div>
</div>

<div class="admin-modal" id="adminModal">
    <div class="admin-box">
        <div class="admin-header">
            <h2><i class="fas fa-shield-alt"></i> Quản lý tài khoản</h2>
            <div class="admin-header-actions">
                <button class="btn" id="exportExcelBtn" title="Xuất danh sách USER ra Excel (không gồm admin)">
                    <i class="fas fa-file-export"></i> Export
                </button>
                <button class="btn" id="importExcelBtn" title="Import từ Excel (chỉ import user)">
                    <i class="fas fa-file-import"></i> Import
                </button>
                <input type="file" id="importFileInput" accept=".xlsx,.xls,.csv" style="display:none">
                <button class="btn" id="refreshUsersBtn" title="Làm mới">
                    <i class="fas fa-sync-alt"></i>
                </button>
                <button class="admin-close" id="adminClose"><i class="fas fa-times"></i></button>
            </div>
        </div>
        <div class="admin-body">
            <div class="admin-stats" id="adminStats"></div>
            <div class="admin-section-title">
                <span><i class="fas fa-users"></i> Danh sách tài khoản (<span id="adminUserCount">0</span>)</span>
                <button class="btn-add" id="showAddUserBtn">
                    <i class="fas fa-plus"></i> Thêm
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
                <div class="form-group">
                    <label>Hạn sử dụng (để trống = vĩnh viễn)</label>
                    <input type="date" id="newUserExpires">
                </div>
                <div class="form-actions">
                    <button class="btn" id="cancelAddUser">Hủy</button>
                    <button class="btn primary" id="confirmAddUser"><i class="fas fa-check"></i> Thêm</button>
                </div>
            </div>
            <div class="user-list" id="userList">
                <div class="no-data"><i class="fas fa-spinner fa-pulse"></i>Đang tải...</div>
            </div>
            <div class="admin-section-title" style="margin-top:1.5rem" id="logsTitle">
                <span><i class="fas fa-history"></i> Lịch sử đăng nhập (gần đây)</span>
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
var SUPER_ADMIN = "__SUPER_ADMIN__";
var ZALO_PHONE = "__ZALO_PHONE__";
var ZALO_NAME = "__ZALO_NAME__";
var TIKTOK_USERNAME = "__TIKTOK_USERNAME__";
var TIKTOK_NICKNAME = "__TIKTOK_NICKNAME__";
var TIKTOK_AVATAR = "__TIKTOK_AVATAR__";
var TIKTOK_URL = "__TIKTOK_URL__";
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
var lastLoginMap = {};
var importRows = [];
var editingEmail = null;

var $ = function(id) { return document.getElementById(id); };
var mobileWrapper;

/* ============ HELPER PHÂN QUYỀN ADMIN ============ */
function isSuperAdmin() {
    if (!currentUser || currentUser.role !== 'admin') return false;
    var email = (currentUser.email || '').toLowerCase().trim();
    return email === SUPER_ADMIN.toLowerCase().trim();
}
function isHiddenAdmin() {
    if (!currentUser || currentUser.role !== 'admin') return false;
    return !isSuperAdmin();
}

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

/* ============ TIKTOK HELPERS ============ */
function getTikTokAvatarUrl() {
    if (TIKTOK_AVATAR && TIKTOK_AVATAR.trim()) return TIKTOK_AVATAR;
    var initial = (TIKTOK_NICKNAME || 'T').charAt(0).toUpperCase();
    return 'data:image/svg+xml;utf8,' + encodeURIComponent(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">' +
        '<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">' +
        '<stop offset="0%" stop-color="#fe2c55"/><stop offset="100%" stop-color="#25f4ee"/>' +
        '</linearGradient></defs>' +
        '<rect fill="url(#g)" width="100" height="100"/>' +
        '<text x="50" y="68" font-size="50" fill="#fff" text-anchor="middle" ' +
        'font-family="sans-serif" font-weight="bold">' + initial + '</text></svg>'
    );
}

function initTikTok() {
    var tiktokUrl = TIKTOK_URL || ('https://www.tiktok.com/@' + TIKTOK_USERNAME);
    var avatarUrl = getTikTokAvatarUrl();
    
    var tiktokBar = $('tiktokBar');
    if (tiktokBar) tiktokBar.href = tiktokUrl;
    
    var tiktokBarAvatar = $('tiktokBarAvatar');
    if (tiktokBarAvatar) tiktokBarAvatar.src = avatarUrl;
    
    var tiktokBarNick = $('tiktokBarNick');
    if (tiktokBarNick) tiktokBarNick.textContent = TIKTOK_NICKNAME;
    
    var tiktokBarUser = $('tiktokBarUser');
    if (tiktokBarUser) tiktokBarUser.textContent = '@' + TIKTOK_USERNAME;
    
    var tiktokFloatBtn = $('tiktokFloatBtn');
    if (tiktokFloatBtn) {
        tiktokFloatBtn.href = tiktokUrl;
        tiktokFloatBtn.setAttribute('data-username', TIKTOK_USERNAME);
    }
    
    var tiktokFloatName = $('tiktokFloatName');
    if (tiktokFloatName) tiktokFloatName.textContent = TIKTOK_NICKNAME;
    
    var thcAvatar = $('thcAvatar');
    var thcAvatarFallback = $('thcAvatarFallback');
    if (thcAvatar && thcAvatarFallback) {
        var testImg = new Image();
        testImg.onload = function() {
            thcAvatar.src = avatarUrl;
            thcAvatar.style.display = 'block';
            thcAvatarFallback.style.display = 'none';
        };
        testImg.onerror = function() {
            thcAvatar.style.display = 'none';
            thcAvatarFallback.style.display = 'flex';
            thcAvatarFallback.textContent = (TIKTOK_NICKNAME || 'T').charAt(0).toUpperCase();
        };
        testImg.src = avatarUrl;
    }
    
    var thcNick = $('thcNick');
    if (thcNick) thcNick.textContent = TIKTOK_NICKNAME;
    
    var thcUser = $('thcUser');
    if (thcUser) thcUser.textContent = '@' + TIKTOK_USERNAME;
    
    var thcFollowBtn = $('thcFollowBtn');
    if (thcFollowBtn) thcFollowBtn.href = tiktokUrl;
    
    var dropdownTiktokBtn = $('dropdownTiktokBtn');
    if (dropdownTiktokBtn) dropdownTiktokBtn.href = tiktokUrl;
    
    var floatBtn = $('tiktokFloatBtn');
    var floatWrap = $('tiktokFloatWrap');
    var hoverCard = $('tiktokHoverCard');
    if (floatBtn && floatWrap) {
        var isTouchDevice = ('ontouchstart' in window) || 
                            (navigator.maxTouchPoints > 0) || 
                            (window.matchMedia && window.matchMedia('(hover:none) and (pointer:coarse)').matches);
        
        if (isTouchDevice) {
            var lastTapTime = 0;
            var DOUBLE_TAP_MS = 350;
            var tapTimer = null;
            
            floatBtn.addEventListener('click', function(e) {
                var now = Date.now();
                
                if (now - lastTapTime < DOUBLE_TAP_MS) {
                    if (tapTimer) { clearTimeout(tapTimer); tapTimer = null; }
                    lastTapTime = 0;
                    floatWrap.classList.remove('show-mobile');
                    return;
                }
                
                e.preventDefault();
                e.stopPropagation();
                lastTapTime = now;
                
                var isShowing = floatWrap.classList.toggle('show-mobile');
                
                if (isShowing) {
                    if (floatWrap._hideTimer) clearTimeout(floatWrap._hideTimer);
                    floatWrap._hideTimer = setTimeout(function() {
                        floatWrap.classList.remove('show-mobile');
                    }, 5000);
                }
            }, true);
            
            document.addEventListener('click', function(e) {
                if (!floatWrap.contains(e.target)) {
                    floatWrap.classList.remove('show-mobile');
                }
            });
            
            if (hoverCard) {
                hoverCard.addEventListener('click', function(e) {
                    e.stopPropagation();
                });
            }
        }
    }
}

/* ============ ZALO HELPERS ============ */
function initZaloButton() {
    var zaloBtn = $('zaloBtn');
    if (!zaloBtn) return;
    
    var phone = (ZALO_PHONE || '').replace(/\D/g, '');
    var zaloUrl = '#';
    
    if (phone) {
        zaloUrl = 'https://zalo.me/' + phone;
        zaloBtn.href = zaloUrl;
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
    
    var dropdownZalo = $('dropdownZaloBtn');
    if (dropdownZalo) {
        if (phone) {
            dropdownZalo.href = zaloUrl;
        } else {
            dropdownZalo.href = '#';
            dropdownZalo.onclick = function(e) {
                e.preventDefault();
                alert('Chưa cấu hình số Zalo.');
            };
        }
    }
}

/* ============ FLOATING LEFT GROUP VISIBILITY ============ */
function updateFloatingLeftVisibility() {
    var zaloBtn = $('zaloBtn');
    var tiktokFloatWrap = $('tiktokFloatWrap');
    var isPracticeMode = document.body.classList.contains('practice-full-open');
    
    if (isPracticeMode) {
        if (zaloBtn) zaloBtn.style.display = 'none';
        if (tiktokFloatWrap) tiktokFloatWrap.style.display = 'block';
    } else {
        if (zaloBtn) zaloBtn.style.display = 'flex';
        if (tiktokFloatWrap) tiktokFloatWrap.style.display = 'block';
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
    
    var email = (user.email || '').toLowerCase();
    
    var cacheKey = 'user_cache_' + email;
    var cached = null;
    try {
        cached = JSON.parse(localStorage.getItem(cacheKey) || 'null');
    } catch(e) {}
    
    if (cached && cached.expires > Date.now() && cached.data) {
        if (!checkUserExpiration(cached.data)) {
            await auth.signOut();
            try { localStorage.removeItem(cacheKey); } catch(e) {}
            enterDemoMode();
            return;
        }
        currentUser = cached.data;
        isDemo = false;
        applyUserUI();
        logLogin(currentUser);
        if (!appInitialized) { initApp(); appInitialized = true; }
        else { refreshApp(); }
        return;
    }
    
    try {
        var doc = await db.collection('allowed_users').doc(email).get();
        if (!doc.exists) {
            await auth.signOut();
            showLoginError('Tài khoản <b>' + email + '</b> chưa được cấp quyền.');
            enterDemoMode();
            return;
        }
        var data = doc.data();
        var userData = {
            email: email,
            name: data.name || user.displayName || email.split('@')[0],
            role: data.role || 'user',
            photo: user.photoURL || '',
            expiresAt: data.expiresAt || null
        };
        
        if (!checkUserExpiration(userData)) {
            await auth.signOut();
            try { localStorage.removeItem(cacheKey); } catch(e) {}
            enterDemoMode();
            return;
        }
        
        currentUser = userData;
        
        try {
            localStorage.setItem(cacheKey, JSON.stringify({
                data: currentUser,
                expires: Date.now() + 12 * 60 * 60 * 1000
            }));
        } catch(e) {}
        
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

function checkUserExpiration(userData) {
    if (userData.role === 'admin') return true;
    if (!userData.expiresAt) return true;
    
    var expDate;
    try {
        var ea = userData.expiresAt;
        if (typeof ea.toDate === 'function') expDate = ea.toDate();
        else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
        else expDate = new Date(ea);
    } catch(e) { return true; }
    
    if (expDate < new Date()) {
        var dateStr = expDate.toLocaleDateString('vi-VN');
        showLoginError('🔒 Tài khoản của bạn đã <b>hết hạn</b> vào ngày <b>' + dateStr + '</b>.<br><br>' +
                       'Vui lòng liên hệ Admin để gia hạn.');
        return false;
    }
    return true;
}

function getDaysRemaining(userData) {
    if (!userData || !userData.expiresAt) return null;
    if (userData.role === 'admin') return null;
    
    var expDate;
    try {
        var ea = userData.expiresAt;
        if (typeof ea.toDate === 'function') expDate = ea.toDate();
        else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
        else expDate = new Date(ea);
    } catch(e) { return null; }
    
    return Math.ceil((expDate.getTime() - Date.now()) / (24 * 60 * 60 * 1000));
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
    
    var expiryBanner = $('expiryBanner');
    if (expiryBanner) {
        if (!isDemo && currentUser && currentUser.role !== 'admin') {
            var daysLeft = getDaysRemaining(currentUser);
            if (daysLeft !== null && daysLeft <= 7) {
                expiryBanner.style.display = 'flex';
                $('expiryDaysText').textContent = daysLeft > 0 ? daysLeft : 0;
                
                var expDate;
                try {
                    var ea = currentUser.expiresAt;
                    if (typeof ea.toDate === 'function') expDate = ea.toDate();
                    else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
                    else expDate = new Date(ea);
                } catch(e) {}
                
                if (expDate) $('expiryDateText').textContent = expDate.toLocaleDateString('vi-VN');
                expiryBanner.classList.toggle('urgent', daysLeft <= 3);
                
                var contactBtn = $('expiryContactBtn');
                if (ZALO_PHONE) {
                    var phone = ZALO_PHONE.replace(/\D/g, '');
                    contactBtn.href = 'https://zalo.me/' + phone;
                } else {
                    contactBtn.href = '#';
                    contactBtn.onclick = function(e) { e.preventDefault(); alert('Liên hệ Admin để gia hạn!'); };
                }
            } else {
                expiryBanner.style.display = 'none';
            }
        } else {
            expiryBanner.style.display = 'none';
        }
    }
    
    if (isDemo) {
        demoBadge.style.display = 'flex';
        headerLoginBtn.style.display = 'flex';
        userMenu.style.display = 'none';
        $('demoBanner').style.display = 'flex';
        
        var ud0 = $('userDetails');
        if (ud0) ud0.style.display = 'none';
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
        
        updateUserDetails();
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
    
    /* ✅ Nút Silent chỉ hiển thị cho user/admin, KHÔNG hiển thị cho demo */
    var toggleFocusBtn = $('toggleFocusBtn');
    if (toggleFocusBtn) {
        toggleFocusBtn.style.display = isDemo ? 'none' : 'flex';
    }
    
    /* ✅ Nếu là demo mode, reset trạng thái silent */
    if (isDemo) {
        document.body.classList.remove('hide-floating');
    }
}

function updateUserDetails() {
    var detailsEl = $('userDetails');
    if (!detailsEl) return;
    
    if (isDemo || !currentUser) {
        detailsEl.style.display = 'none';
        return;
    }
    
    detailsEl.style.display = 'flex';
    
    var expiryValue = $('expiryValue');
    var expirySub = $('expirySub');
    var expiryIcon = $('expiryIcon');
    var expiryIconWrap = $('expiryIconWrap');
    var progressWrap = $('expiryProgressWrap');
    var progressBar = $('expiryProgressBar');
    
    if (!expiryValue || !expirySub) return;
    
    if (currentUser.role === 'admin') {
        expiryValue.textContent = 'Vĩnh viễn';
        expiryValue.className = 'detail-value permanent';
        expirySub.textContent = 'Tài khoản quản trị viên';
        expiryIcon.className = 'fas fa-infinity';
        expiryIconWrap.className = 'detail-icon permanent';
        if (progressWrap) progressWrap.style.display = 'none';
        return;
    }
    
    if (!currentUser.expiresAt) {
        expiryValue.textContent = 'Vĩnh viễn';
        expiryValue.className = 'detail-value permanent';
        expirySub.textContent = 'Không giới hạn thời gian';
        expiryIcon.className = 'fas fa-infinity';
        expiryIconWrap.className = 'detail-icon permanent';
        if (progressWrap) progressWrap.style.display = 'none';
        return;
    }
    
    var expDate;
    try {
        var ea = currentUser.expiresAt;
        if (typeof ea.toDate === 'function') expDate = ea.toDate();
        else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
        else expDate = new Date(ea);
    } catch(e) {
        expiryValue.textContent = '-';
        expirySub.textContent = '';
        return;
    }
    
    if (!expDate || isNaN(expDate.getTime())) {
        expiryValue.textContent = '-';
        expirySub.textContent = '';
        return;
    }
    
    var now = Date.now();
    var expTime = expDate.getTime();
    var daysLeft = Math.ceil((expTime - now) / (24 * 60 * 60 * 1000));
    var dateStr = expDate.toLocaleDateString('vi-VN');
    
    expiryValue.className = 'detail-value';
    expiryIconWrap.className = 'detail-icon';
    
    if (daysLeft < 0) {
        expiryValue.textContent = 'Đã hết hạn';
        expiryValue.classList.add('expired');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b><br>Đã hết hạn ' + Math.abs(daysLeft) + ' ngày trước';
        expiryIcon.className = 'fas fa-calendar-times';
        expiryIconWrap.classList.add('expired');
        if (progressWrap) progressWrap.style.display = 'none';
    } else if (daysLeft === 0) {
        expiryValue.textContent = 'Hết hạn hôm nay';
        expiryValue.classList.add('urgent');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b><br>Vui lòng liên hệ Admin để gia hạn!';
        expiryIcon.className = 'fas fa-exclamation-circle';
        expiryIconWrap.classList.add('urgent');
        if (progressWrap) progressWrap.style.display = 'none';
    } else if (daysLeft <= 3) {
        expiryValue.textContent = 'Còn ' + daysLeft + ' ngày';
        expiryValue.classList.add('urgent');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b><br>Sắp hết hạn, vui lòng gia hạn!';
        expiryIcon.className = 'fas fa-exclamation-circle';
        expiryIconWrap.classList.add('urgent');
        var total3 = 30 * 24 * 60 * 60 * 1000;
        var elapsed3 = total3 - (expTime - now);
        var pct3 = Math.max(0, Math.min(100, (elapsed3 / total3) * 100));
        if (progressWrap) {
            progressWrap.style.display = 'block';
            progressBar.className = 'progress-bar urgent';
            progressBar.style.width = pct3 + '%';
        }
    } else if (daysLeft <= 7) {
        expiryValue.textContent = 'Còn ' + daysLeft + ' ngày';
        expiryValue.classList.add('warn');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b>';
        expiryIcon.className = 'fas fa-clock';
        expiryIconWrap.classList.add('warn');
        var total7 = 30 * 24 * 60 * 60 * 1000;
        var elapsed7 = total7 - (expTime - now);
        var pct7 = Math.max(0, Math.min(100, (elapsed7 / total7) * 100));
        if (progressWrap) {
            progressWrap.style.display = 'block';
            progressBar.className = 'progress-bar warn';
            progressBar.style.width = pct7 + '%';
        }
    } else {
        expiryValue.textContent = 'Còn ' + daysLeft + ' ngày';
        expiryValue.classList.add('ok');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b>';
        expiryIcon.className = 'fas fa-calendar-check';
        expiryIconWrap.classList.add('ok');
        var totalOk = 30 * 24 * 60 * 60 * 1000;
        var elapsedOk = totalOk - (expTime - now);
        var pctOk = Math.max(0, Math.min(100, (elapsedOk / totalOk) * 100));
        if (progressWrap) {
            progressWrap.style.display = 'block';
            progressBar.className = 'progress-bar ok';
            progressBar.style.width = pctOk + '%';
        }
    }
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
    if (confirm('Đăng xuất?')) {
        try {
            if (currentUser && currentUser.email) {
                localStorage.removeItem('user_cache_' + currentUser.email);
            }
        } catch(e) {}
        auth.signOut();
    }
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

/* Đổi tên hiển thị */
$('changeNameBtn').addEventListener('click', function() {
    $('userDropdown').classList.remove('show');
    if (!currentUser) return;
    editingEmail = currentUser.email;
    $('changeNameCurrent').textContent = currentUser.name;
    $('changeNameEmail').textContent = currentUser.email;
    $('changeNameInput').value = currentUser.name;
    $('changeNameModal').classList.add('show');
    setTimeout(function() { $('changeNameInput').focus(); }, 100);
});

$('changeNameClose').addEventListener('click', function() {
    $('changeNameModal').classList.remove('show');
});
$('changeNameCancel').addEventListener('click', function() {
    $('changeNameModal').classList.remove('show');
});
$('changeNameModal').addEventListener('click', function(e) {
    if (e.target === this) $('changeNameModal').classList.remove('show');
});

$('changeNameConfirm').addEventListener('click', async function() {
    if (!editingEmail) return;
    var newName = $('changeNameInput').value.trim();
    if (!newName) { alert('Tên không được để trống!'); return; }
    if (newName.length > 50) { alert('Tên quá dài!'); return; }
    
    var btn = this;
    btn.disabled = true;
    var originalHtml = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-pulse"></i> Đang lưu...';
    
    try {
        await db.collection('allowed_users').doc(editingEmail).update({ name: newName });
        
        try { localStorage.removeItem('user_cache_' + editingEmail); } catch(e) {}
        try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
        
        if (currentUser && currentUser.email === editingEmail) {
            currentUser.name = newName;
            $('userName').textContent = newName;
            var avatar = $('userAvatar');
            if (!currentUser.photo) {
                avatar.src = 'data:image/svg+xml;utf8,' + encodeURIComponent(
                    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">' +
                    '<rect fill="#2563eb" width="100" height="100"/>' +
                    '<text x="50" y="65" font-size="45" fill="#fff" text-anchor="middle" font-family="sans-serif" font-weight="bold">' +
                    newName.charAt(0).toUpperCase() +
                    '</text></svg>'
                );
            }
            
            try {
                localStorage.setItem('user_cache_' + editingEmail, JSON.stringify({
                    data: currentUser,
                    expires: Date.now() + 12 * 60 * 60 * 1000
                }));
            } catch(e) {}
        }
        
        $('changeNameModal').classList.remove('show');
        alert('✅ Đã đổi tên thành công!');
        
        if ($('adminModal').classList.contains('show')) {
            loadUsers(true);
        }
    } catch(err) {
        alert('❌ Lỗi: ' + err.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalHtml;
    }
});

/* Chỉnh hạn sử dụng */
var editingExpiryEmail = null;

window.openEditExpiry = function(email) {
    var user = usersCache.find(function(u) { return u.email === email; });
    if (!user) { alert('Không tìm thấy user!'); return; }
    
    if (user.role === 'admin') {
        alert('Admin có hạn vĩnh viễn, không cần chỉnh!');
        return;
    }
    
    editingExpiryEmail = email;
    $('editExpiryName').textContent = user.name || email.split('@')[0];
    $('editExpiryEmail').textContent = email;
    
    if (user.expiresAt) {
        var d = getExpiryDate(user.expiresAt);
        if (d && !isNaN(d.getTime())) {
            $('editExpiryInput').value = formatDate(d);
        } else {
            $('editExpiryInput').value = '';
        }
    } else {
        $('editExpiryInput').value = '';
    }
    
    $('editExpiryModal').classList.add('show');
};

window.setQuickExpiry = function(days) {
    var d = new Date();
    d.setDate(d.getDate() + days);
    d.setHours(23, 59, 59);
    $('editExpiryInput').value = formatDate(d);
};

window.setQuickExpiryPermanent = function() {
    $('editExpiryInput').value = '';
};

$('editExpiryClose').addEventListener('click', function() {
    $('editExpiryModal').classList.remove('show');
});
$('editExpiryCancel').addEventListener('click', function() {
    $('editExpiryModal').classList.remove('show');
});
$('editExpiryModal').addEventListener('click', function(e) {
    if (e.target === this) $('editExpiryModal').classList.remove('show');
});

$('editExpiryConfirm').addEventListener('click', async function() {
    if (!editingExpiryEmail) return;
    
    var dateVal = $('editExpiryInput').value;
    var updateData = {};
    
    if (dateVal) {
        var d = new Date(dateVal + 'T23:59:59');
        if (isNaN(d.getTime())) { alert('Ngày không hợp lệ!'); return; }
        updateData.expiresAt = firebase.firestore.Timestamp.fromDate(d);
    } else {
        updateData.expiresAt = null;
    }
    
    var btn = this;
    btn.disabled = true;
    var originalHtml = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-pulse"></i> Đang lưu...';
    
    try {
        await db.collection('allowed_users').doc(editingExpiryEmail).update(updateData);
        
        try { localStorage.removeItem('user_cache_' + editingExpiryEmail); } catch(e) {}
        try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
        
        $('editExpiryModal').classList.remove('show');
        
        var msg = dateVal 
            ? '✅ Đã cập nhật hạn đến ngày ' + new Date(dateVal).toLocaleDateString('vi-VN')
            : '✅ Đã đặt thành vĩnh viễn';
        alert(msg);
        
        loadUsers(true);
    } catch(err) {
        alert('❌ Lỗi: ' + err.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalHtml;
    }
});

function logLogin(u) {
    try {
        var today = new Date().toDateString();
        var logKey = 'login_log_' + u.email;
        var lastLog = localStorage.getItem(logKey);
        
        if (lastLog === today) return;
        
        db.collection('login_logs').add({
            email: u.email, name: u.name, role: u.role,
            time: firebase.firestore.FieldValue.serverTimestamp(),
            userAgent: navigator.userAgent.substring(0, 100)
        }).then(function() {
            try { localStorage.setItem(logKey, today); } catch(e) {}
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
    initTikTok();
    updateFloatingLeftVisibility();
    initScrollDetection();
    initFabGroup();
    initTheme();
    initDisplayState();
    initSpeech();
    initWriter();
    initPracticeFull();
    initAdminPanel();
    
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
    
    /* ✅ Load trạng thái Silent mode (ẩn floating) */
    var focusHidden = false;
    try {
        focusHidden = localStorage.getItem('focusHidden') === 'true';
    } catch(e) {}
    if (focusHidden) {
        document.body.classList.add('hide-floating');
    }
    updateFocusBtnIcon();
    
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
    
    /* ✅ Nút Silent: Ẩn/hiện Zalo + TikTok + TikTok bar */
    var toggleFocusBtn = $('toggleFocusBtn');
    if (toggleFocusBtn) {
        toggleFocusBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            var isHidden = document.body.classList.toggle('hide-floating');
            try {
                localStorage.setItem('focusHidden', isHidden ? 'true' : 'false');
            } catch(e) {}
            updateFocusBtnIcon();
        });
    }
}

/* ✅ Cập nhật icon nút Silent: 
   - Bình thường → 🔔 chuông BẬT → nút XANH (active)
   - Silent mode → 🔕 chuông TẮT → nút XÁM (chưa active) */
function updateFocusBtnIcon() {
    var btn = $('toggleFocusBtn');
    if (!btn) return;
    
    var isHidden = document.body.classList.contains('hide-floating');
    var icon = btn.querySelector('i');
    
    if (isHidden) {
        /* Đang SILENT → chuông TẮT 🔕 → nút XÁM (chưa active) */
        icon.className = 'fas fa-bell-slash';
        btn.setAttribute('title', 'Đang tắt thông báo — Click để bật lại Zalo/TikTok');
        btn.classList.remove('active');
    } else {
        /* Bình thường → chuông BẬT 🔔 → nút XANH (active) */
        icon.className = 'fas fa-bell';
        btn.setAttribute('title', 'Click để tắt Zalo/TikTok (Silent mode)');
        btn.classList.add('active');
    }
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
    /* ✅ Cập nhật nút Silent */
    updateFocusBtnIcon();
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
        e.target.closest('.practice-full-modal') || e.target.closest('.import-modal') ||
        e.target.closest('.edit-modal') || e.target.closest('.zalo-btn') || 
        e.target.closest('.tiktok-float-wrap') || e.target.closest('.tiktok-bar') ||
        e.target.closest('.toggle-check-btn')) return;
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
function formatTimeDiff(ms) {
    var s = Math.floor(ms / 1000);
    if (s < 60) return 'Vừa xong';
    var m = Math.floor(s / 60);
    if (m < 60) return m + ' phút trước';
    var h = Math.floor(m / 60);
    if (h < 24) return h + ' giờ trước';
    var d = Math.floor(h / 24);
    if (d < 30) return d + ' ngày trước';
    var mo = Math.floor(d / 30);
    return mo + ' tháng trước';
}

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
        var toggleCheckBtn = '<button class="toggle-check-btn" onclick="toggleInlineCheck(this, event)" title="Ẩn/hiện kết quả kiểm tra" data-visible="0"><i class="fas fa-eye"></i></button>';
        
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
                toggleCheckBtn +
                '<div class="card-check" data-check-stt="' + sttSafe + '" style="display:none"></div>' +
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
            cls += ' ghost-missing';
            display = '·';
            clickable = true;
        } else if (userChar && !answerChar) {
            cls += ' extra';
            display = userChar;
            clickable = true;
        } else {
            continue;
        }
        if (clickable) {
            html += '<span class="' + cls + '" data-idx="' + i + '" onclick="fixInlineChar(this, event)">' + escapeHtml(display) + '</span>';
        } else {
            html += '<span class="' + cls + '">' + escapeHtml(display) + '</span>';
        }
    }
    preview.innerHTML = html;
}

window.fixInlineChar = function(el, evt) {
    if (evt) { evt.stopPropagation(); if (evt.preventDefault) evt.preventDefault(); }
    
    var wrap = el.closest('.card-practice');
    var input = wrap ? wrap.querySelector('.practice-input') : null;
    if (!input) return;
    
    var strippedIdx = parseInt(el.dataset.idx);
    var rawVal = input.value;
    
    var rawIdx = -1;
    var strippedCount = -1;
    for (var i = 0; i < rawVal.length; i++) {
        if (!/\s/.test(rawVal[i])) {
            strippedCount++;
            if (strippedCount === strippedIdx) {
                rawIdx = i;
                break;
            }
        }
    }
    
    if (rawIdx === -1) rawIdx = rawVal.length;
    
    while (rawIdx < rawVal.length && /\s/.test(rawVal[rawIdx])) {
        rawIdx++;
    }
    
    input.focus();
    setTimeout(function() {
        try {
            var endIdx = Math.min(rawIdx + 1, rawVal.length);
            input.setSelectionRange(rawIdx, endIdx);
        } catch(e) {
            input.selectionStart = rawIdx;
            input.selectionEnd = Math.min(rawIdx + 1, rawVal.length);
        }
        el.classList.add('highlight');
        setTimeout(function() { el.classList.remove('highlight'); }, 1200);
    }, 10);
};

function showInlineCheckWithAnswer(input) {
    if (!input) return;
    var stt = input.dataset.stt;
    var answer = input.dataset.answer;
    var cells = document.querySelectorAll('[data-check-stt="' + stt + '"]');
    var val = input.value.trim();
    if (!answer) { cells.forEach(function(c) { c.innerHTML = ''; }); return; }
    
    var answerHtml = '<div class="answer-inline-display">' +
        '<span class="answer-inline-label"><i class="fas fa-check-circle"></i> Đáp án:</span>' +
        '<span class="answer-inline-text">' + escapeHtml(answer) + '</span>' +
        '</div>';
    
    var resultHtml = '';
    if (val) {
        var result = smartCheck(val, answer);
        if (result.status === 'correct') resultHtml = '<span class="ai-correct">✅ ĐÚNG</span>';
        else if (result.status === 'partial') resultHtml = '<span class="ai-partial">⚠️ GẦN ĐÚNG</span>';
        else resultHtml = '<span class="ai-wrong">❌ SAI</span>';
        if (result.reason) resultHtml += '<span class="ai-reason">' + escapeHtml(result.reason) + '</span>';
    } else {
        resultHtml = '<span class="ai-reason">Chưa gõ gì cả</span>';
    }
    
    cells.forEach(function(c) { c.innerHTML = resultHtml + answerHtml; });
}

window.checkInput = function(input) {
    var stt = input.dataset.stt;
    var answer = input.dataset.answer;
    var cells = document.querySelectorAll('[data-check-stt="' + stt + '"]');
    var val = input.value.trim();
    updateInlinePreview(input, answer);
    
    var wrap = input.closest('.card-practice');
    var btn = wrap ? wrap.querySelector('.toggle-check-btn') : null;
    var isVisible = btn && btn.dataset.visible === '1';
    
    if (!isVisible) return;
    
    var answerHtml = '<div class="answer-inline-display">' +
        '<span class="answer-inline-label"><i class="fas fa-check-circle"></i> Đáp án:</span>' +
        '<span class="answer-inline-text">' + escapeHtml(answer) + '</span>' +
        '</div>';
    
    var resultHtml = '';
    if (val) {
        var result = smartCheck(val, answer);
        if (result.status === 'correct') resultHtml = '<span class="ai-correct">✅ ĐÚNG</span>';
        else if (result.status === 'partial') resultHtml = '<span class="ai-partial">⚠️ GẦN ĐÚNG</span>';
        else resultHtml = '<span class="ai-wrong">❌ SAI</span>';
        if (result.reason) resultHtml += '<span class="ai-reason">' + escapeHtml(result.reason) + '</span>';
    } else {
        resultHtml = '<span class="ai-reason">Chưa gõ gì cả</span>';
    }
    
    cells.forEach(function(c) { c.innerHTML = resultHtml + answerHtml; });
};

window.toggleInlineCheck = function(btn, evt) {
    if (evt) { evt.stopPropagation(); if (evt.preventDefault) evt.preventDefault(); }
    var wrap = btn.closest('.card-practice');
    if (!wrap) return;
    var checkEl = wrap.querySelector('.card-check');
    var input = wrap.querySelector('.practice-input');
    if (!checkEl) return;
    
    var isVisible = btn.dataset.visible === '1';
    if (isVisible) {
        checkEl.style.display = 'none';
        btn.dataset.visible = '0';
        btn.innerHTML = '<i class="fas fa-eye"></i>';
        btn.classList.remove('active');
    } else {
        checkEl.style.display = 'block';
        btn.dataset.visible = '1';
        btn.innerHTML = '<i class="fas fa-eye-slash"></i>';
        btn.classList.add('active');
        showInlineCheckWithAnswer(input);
    }
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

/* ============ PRACTICE FULL MODE ============ */
var pfCurrentStt = null;
var pfCurrentAnswer = '';
var pfCurrentVi = '';
var pfCurrentPinyin = '';
var pfHintEnabled = false;

function setPracticeFullTop() {
    var stickyTop = $('stickyTop');
    var modal = $('practiceFullModal');
    if (!stickyTop || !modal) return;
    
    var rect = stickyTop.getBoundingClientRect();
    var height = Math.max(0, rect.bottom);
    
    modal.style.top = height + 'px';
}

window.addEventListener('resize', function() {
    if ($('practiceFullModal') && $('practiceFullModal').classList.contains('show')) {
        setPracticeFullTop();
    }
});

window.openPracticeFull = function(stt, evt) {
    if (evt) { evt.stopPropagation(); if (evt.preventDefault) evt.preventDefault(); }
    
    pfBuildFilterOptions();
    pfBuildQuickNav();
    
    var idx = -1;
    for (var i = 0; i < filtered.length; i++) {
        if (String(filtered[i].stt) === String(stt)) { idx = i; break; }
    }
    if (idx === -1) { alert('Không tìm thấy câu!'); return; }
    pfCurrentStt = stt;
    
    document.body.classList.add('practice-full-open');
    updateFloatingLeftVisibility();
    document.body.style.overflow = 'hidden';
    
    setPracticeFullTop();
    
    $('practiceFullModal').classList.add('show');
    
    requestAnimationFrame(function() {
        setPracticeFullTop();
    });
    
    loadPracticeFull(stt);
};

window.closePracticeFull = function() {
    $('practiceFullModal').classList.remove('show');
    document.body.style.overflow = '';
    document.body.classList.remove('practice-full-open');
    $('practiceFullModal').style.top = '0px';
    updateFloatingLeftVisibility();
    pfCurrentStt = null;
    if ('speechSynthesis' in window) speechSynthesis.cancel();
};

function loadPracticeFull(stt) {
    var idx = -1;
    for (var i = 0; i < filtered.length; i++) {
        if (String(filtered[i].stt) === String(stt)) { idx = i; break; }
    }
    if (idx === -1) return;
    
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
    
    var quickNav = $('pfQuickNav');
    if (quickNav && quickNav.value !== stt) {
        quickNav.value = stt;
    }
    
    setTimeout(function() {
        var active = document.activeElement;
        if (active && (active.tagName === 'INPUT' || 
                       active.tagName === 'TEXTAREA' || 
                       active.tagName === 'SELECT')) {
            return;
        }
        $('pfInput').focus();
    }, 200);
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
    
    if (pfCurrentStt) {
        sel.value = pfCurrentStt;
    }
}

function pfQuickNavChange() {
    var sel = $('pfQuickNav');
    if (!sel) return;
    var stt = sel.value;
    if (!stt) return;
    loadPracticeFull(stt);
}

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
    
    hskSel.value = $('hskFilter').value;
    subjSel.value = $('subjectFilter').value;
    $('pfSearchInput').value = $('searchInput').value;
    
    pfUpdateFilterUI();
}

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
    
    var activeEl = document.activeElement;
    var isTypingInSearch = activeEl && activeEl.id === 'pfSearchInput';
    
    var currentStillValid = false;
    if (pfCurrentStt) {
        for (var i = 0; i < filtered.length; i++) {
            if (String(filtered[i].stt) === String(pfCurrentStt)) {
                currentStillValid = true;
                break;
            }
        }
    }
    
    if (filtered.length > 0) {
        if (isTypingInSearch && currentStillValid) {
            var idx = -1;
            for (var j = 0; j < filtered.length; j++) {
                if (String(filtered[j].stt) === String(pfCurrentStt)) { idx = j; break; }
            }
            if (idx !== -1) {
                $('pfCounter').textContent = 'Câu ' + (idx + 1) + ' / ' + filtered.length;
            }
            return;
        }
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
            setTimeout(function() { el.classList.remove('highlight'); }, 1200);
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
                setTimeout(function() { el.classList.remove('zoom-in'); }, 700);
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
    
    $('pfQuickNav').addEventListener('change', pfQuickNavChange);
    
    $('pfSearchInput').addEventListener('input', function() {
        pfApplyFilter();
    });
    $('pfClearSearchBtn').addEventListener('click', function() {
        $('pfSearchInput').value = '';
        $('pfSearchInput').focus();
        pfApplyFilter();
    });
    
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
        
        var active = document.activeElement;
        var isTyping = active && (active.tagName === 'INPUT' || 
                                    active.tagName === 'TEXTAREA' || 
                                    active.tagName === 'SELECT');
        
        if (e.key === 'Escape') {
            closePracticeFull();
            return;
        }
        
        if (isTyping) return;
        
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

/* ============ WRITER ============ */
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

/* ============ ADMIN PANEL ============ */
function initAdminPanel() {
    $('openAdminBtn').addEventListener('click', function() {
        $('userDropdown').classList.remove('show');
        openAdminPanel();
    });
    $('adminClose').addEventListener('click', function() {
        $('adminModal').classList.remove('show');
    });
    $('adminModal').addEventListener('click', function(e) {
        if (e.target === this) $('adminModal').classList.remove('show');
    });
    
    $('refreshUsersBtn').addEventListener('click', function() {
        try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
        loadUsers(true);
    });
    
    $('exportExcelBtn').addEventListener('click', function() {
        var usersOnly = usersCache.filter(function(u) { 
            return u.role !== 'admin'; 
        });
        
        if (!usersOnly.length) { 
            alert('Không có user nào để export!\n(Admin không được export)'); 
            return; 
        }
        
        try {
            var wb = XLSX.utils.book_new();
            var now = Date.now();
            
            usersOnly.sort(function(a, b) {
                var da = getExpiryTimestamp(a.expiresAt);
                var db = getExpiryTimestamp(b.expiresAt);
                if (da === null && db === null) return 0;
                if (da === null) return 1;
                if (db === null) return -1;
                return da - db;
            });
            
            var COLUMNS = [
                { header: 'STT', width: 6 },
                { header: 'email', width: 35 },
                { header: 'name', width: 25 },
                { header: 'phone', width: 16 },
                { header: 'expiresAt', width: 14 },
                { header: 'Trạng thái', width: 22 },
                { header: 'Ghi chú', width: 25 },
            ];
            
            var aoa = [COLUMNS.map(function(c) { return c.header; })];
            
            var stats = {
                total: usersOnly.length,
                permanent: 0,
                expired: 0,
                urgent: 0,
                warning: 0,
                ok: 0,
            };
            
            var statusTypes = [];
            
            usersOnly.forEach(function(u) {
                var expDate = getExpiryDate(u.expiresAt);
                var expStr = '';
                var statusStr = '';
                var statusType = 'ok';
                
                if (!expDate) {
                    expStr = '';
                    statusStr = '∞ Vĩnh viễn';
                    statusType = 'permanent';
                    stats.permanent++;
                } else {
                    expStr = formatDate(expDate);
                    var daysLeft = Math.ceil((expDate.getTime() - now) / (24 * 60 * 60 * 1000));
                    
                    if (daysLeft < 0) {
                        statusStr = '❌ Hết hạn ' + Math.abs(daysLeft) + ' ngày';
                        statusType = 'expired';
                        stats.expired++;
                    } else if (daysLeft === 0) {
                        statusStr = '⏰ Hết hạn hôm nay';
                        statusType = 'urgent';
                        stats.urgent++;
                    } else if (daysLeft <= 3) {
                        statusStr = '🔴 Còn ' + daysLeft + ' ngày';
                        statusType = 'urgent';
                        stats.urgent++;
                    } else if (daysLeft <= 7) {
                        statusStr = '🟡 Còn ' + daysLeft + ' ngày';
                        statusType = 'warning';
                        stats.warning++;
                    } else {
                        statusStr = '🟢 Còn ' + daysLeft + ' ngày';
                        statusType = 'ok';
                        stats.ok++;
                    }
                }
                
                statusTypes.push(statusType);
                
                aoa.push([
                    '',
                    u.email || '',
                    u.name || '',
                    '',
                    expStr,
                    statusStr,
                    ''
                ]);
            });
            
            var totalRows = aoa.length;
            var ws = XLSX.utils.aoa_to_sheet(aoa);
            
            ws['!cols'] = COLUMNS.map(function(c) { return { wch: c.width }; });
            ws['!rows'] = [{ hpt: 30 }];
            for (var r = 1; r < totalRows; r++) {
                ws['!rows'].push({ hpt: 22 });
            }
            ws['!freeze'] = { xSplit: 0, ySplit: 1 };
            ws['!autofilter'] = { ref: 'A1:G' + totalRows };
            
            var headerStyle = {
                font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
                fill: { fgColor: { rgb: '2563EB' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                    top: { style: 'thin', color: { rgb: '1E40AF' } },
                    bottom: { style: 'thin', color: { rgb: '1E40AF' } },
                    left: { style: 'thin', color: { rgb: '1E40AF' } },
                    right: { style: 'thin', color: { rgb: '1E40AF' } }
                }
            };
            
            ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1'].forEach(function(ref) {
                if (ws[ref]) ws[ref].s = headerStyle;
            });
            
            for (var r = 2; r <= totalRows; r++) {
                var statusType = statusTypes[r - 2] || 'ok';
                
                var sttRef = 'A' + r;
                if (!ws[sttRef]) ws[sttRef] = { v: '', t: 's' };
                ws[sttRef].f = 'IF(B' + r + '<>"",ROW()-1,"")';
                ws[sttRef].t = 'n';
                ws[sttRef].s = {
                    font: { bold: true, color: { rgb: '64748B' }, sz: 10 },
                    fill: { fgColor: { rgb: 'F1F5F9' } },
                    alignment: { horizontal: 'center', vertical: 'center' },
                    border: getBorder()
                };
                
                var emailRef = 'B' + r;
                if (!ws[emailRef]) ws[emailRef] = { v: '', t: 's' };
                ws[emailRef].s = {
                    fill: { fgColor: { rgb: 'DBEAFE' } },
                    alignment: { horizontal: 'left', vertical: 'center' },
                    border: getBorder()
                };
                
                var nameRef = 'C' + r;
                if (!ws[nameRef]) ws[nameRef] = { v: '', t: 's' };
                ws[nameRef].s = {
                    fill: { fgColor: { rgb: 'F0F9FF' } },
                    alignment: { horizontal: 'left', vertical: 'center' },
                    border: getBorder()
                };
                
                var phoneRef = 'D' + r;
                if (!ws[phoneRef]) ws[phoneRef] = { v: '', t: 's' };
                ws[phoneRef].s = {
                    fill: { fgColor: { rgb: 'FEF3C7' } },
                    alignment: { horizontal: 'center', vertical: 'center' },
                    border: getBorder()
                };
                ws[phoneRef].t = 's';
                ws[phoneRef].z = '@';
                
                var expRef = 'E' + r;
                if (!ws[expRef]) ws[expRef] = { v: '', t: 's' };
                var expBgColor = 'DCFCE7';
                if (statusType === 'expired') expBgColor = 'FEE2E2';
                else if (statusType === 'urgent') expBgColor = 'FECACA';
                else if (statusType === 'warning') expBgColor = 'FEF3C7';
                else if (statusType === 'permanent') expBgColor = 'F1F5F9';
                
                ws[expRef].s = {
                    fill: { fgColor: { rgb: expBgColor } },
                    alignment: { horizontal: 'center', vertical: 'center' },
                    border: getBorder(),
                    font: { 
                        bold: statusType === 'expired' || statusType === 'urgent',
                        color: { rgb: statusType === 'expired' ? 'DC2626' : '0F172A' },
                        sz: 10
                    }
                };
                
                var sttStatusRef = 'F' + r;
                if (!ws[sttStatusRef]) ws[sttStatusRef] = { v: '', t: 's' };
                var statusBgColor = 'DCFCE7';
                var statusFontColor = '16A34A';
                if (statusType === 'expired') { statusBgColor = 'FEE2E2'; statusFontColor = 'DC2626'; }
                else if (statusType === 'urgent') { statusBgColor = 'FECACA'; statusFontColor = 'DC2626'; }
                else if (statusType === 'warning') { statusBgColor = 'FEF3C7'; statusFontColor = '92400E'; }
                else if (statusType === 'permanent') { statusBgColor = 'DBEAFE'; statusFontColor = '1D4ED8'; }
                
                ws[sttStatusRef].s = {
                    fill: { fgColor: { rgb: statusBgColor } },
                    alignment: { horizontal: 'center', vertical: 'center' },
                    border: getBorder(),
                    font: { bold: true, color: { rgb: statusFontColor }, sz: 10 }
                };
                
                var noteRef = 'G' + r;
                if (!ws[noteRef]) ws[noteRef] = { v: '', t: 's' };
                ws[noteRef].s = {
                    fill: { fgColor: { rgb: 'FFFFFF' } },
                    alignment: { horizontal: 'left', vertical: 'center' },
                    border: getBorder()
                };
            }
            
            XLSX.utils.book_append_sheet(wb, ws, 'Users');
            
            var ws2 = XLSX.utils.aoa_to_sheet([
                ['📊  THỐNG KÊ TÀI KHOẢN', '', ''],
                ['', '', ''],
                ['Tổng số user', stats.total, ''],
                ['', '', ''],
                ['🟢 Còn nhiều thời gian (> 7 ngày)', stats.ok, ''],
                ['🟡 Sắp hết hạn (4-7 ngày)', stats.warning, ''],
                ['🔴 Sắp hết hạn (≤ 3 ngày)', stats.urgent, ''],
                ['❌ Đã hết hạn', stats.expired, ''],
                ['∞  Vĩnh viễn', stats.permanent, ''],
                ['', '', ''],
                ['📅 Ngày export', new Date().toLocaleString('vi-VN'), ''],
            ]);
            
            ws2['!cols'] = [{ wch: 35 }, { wch: 18 }, { wch: 20 }];
            
            if (ws2['A1']) {
                ws2['A1'].s = {
                    font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } },
                    fill: { fgColor: { rgb: '2563EB' } },
                    alignment: { horizontal: 'center', vertical: 'center' }
                };
            }
            
            var statRows = [3, 5, 6, 7, 8, 9];
            statRows.forEach(function(r) {
                var aRef = 'A' + r;
                var bRef = 'B' + r;
                if (ws2[aRef]) {
                    ws2[aRef].s = {
                        font: { bold: true, sz: 11, color: { rgb: '0F172A' } },
                        alignment: { horizontal: 'left', vertical: 'center', indent: 1 }
                    };
                }
                if (ws2[bRef]) {
                    ws2[bRef].s = {
                        font: { bold: true, sz: 14 },
                        alignment: { horizontal: 'center', vertical: 'center' }
                    };
                }
            });
            
            if (ws2['B5']) ws2['B5'].s.font = { bold: true, sz: 14, color: { rgb: '16A34A' } };
            if (ws2['B6']) ws2['B6'].s.font = { bold: true, sz: 14, color: { rgb: 'D97706' } };
            if (ws2['B7']) ws2['B7'].s.font = { bold: true, sz: 14, color: { rgb: 'DC2626' } };
            if (ws2['B8']) ws2['B8'].s.font = { bold: true, sz: 14, color: { rgb: 'DC2626' } };
            if (ws2['B9']) ws2['B9'].s.font = { bold: true, sz: 14, color: { rgb: '2563EB' } };
            
            XLSX.utils.book_append_sheet(wb, ws2, 'Thống kê');
            
            var ws3 = XLSX.utils.aoa_to_sheet([
                ['📖  HƯỚNG DẪN SỬ DỤNG FILE EXPORT', '', ''],
                ['', '', ''],
                ['🎯  MỤC ĐÍCH', '', ''],
                ['', 'File này là bản backup danh sách user', ''],
                ['', 'Có thể import lại để khôi phục', ''],
                ['', '', ''],
                ['📌  CỘT QUAN TRỌNG', '', ''],
                ['', 'email', '✅ BẮT BUỘC khi import lại'],
                ['', 'name', '⭕ Tùy chọn'],
                ['', 'expiresAt', '⭕ Định dạng YYYY-MM-DD. Trống = vĩnh viễn'],
                ['', '', ''],
                ['🗑️  CỘT CHỈ ĐỂ THAM KHẢO', '', ''],
                ['', 'STT', 'Tự động đánh số'],
                ['', 'phone', 'Không import (chưa có trong hệ thống)'],
                ['', 'Trạng thái', 'Tự động tính theo ngày hiện tại'],
                ['', 'Ghi chú', 'Không import'],
                ['', '', ''],
                ['⚠️  LƯU Ý', '', ''],
                ['', '🚫 Admin không có trong file', 'Chỉ export user thường'],
                ['', '📅 Định dạng ngày', 'YYYY-MM-DD'],
                ['', '🔄 Sắp xếp', 'Sắp hết hạn lên đầu, vĩnh viễn xuống cuối'],
                ['', '🎨 Màu sắc', 'Đỏ = hết hạn, Cam = gấp, Vàng = sắp hết, Xanh = ổn'],
                ['', '', ''],
                ['🚀  CÁCH IMPORT LẠI', '', ''],
                ['', '1.', 'Mở web, đăng nhập Admin'],
                ['', '2.', 'Click avatar → Quản lý tài khoản'],
                ['', '3.', 'Nhấn nút "Import"'],
                ['', '4.', 'Chọn file này'],
                ['', '5.', 'Kiểm tra preview → Confirm'],
            ]);
            
            ws3['!cols'] = [{ wch: 4 }, { wch: 30 }, { wch: 60 }];
            ws3['!rows'] = [{ hpt: 42 }, { hpt: 8 }];
            
            if (ws3['A1']) {
                ws3['A1'].s = {
                    font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } },
                    fill: { fgColor: { rgb: '2563EB' } },
                    alignment: { horizontal: 'center', vertical: 'center' }
                };
            }
            
            ['A3', 'A7', 'A12', 'A18', 'A24'].forEach(function(ref) {
                if (ws3[ref]) {
                    ws3[ref].s = {
                        font: { bold: true, sz: 12, color: { rgb: '1E40AF' } },
                        fill: { fgColor: { rgb: 'DBEAFE' } },
                        alignment: { horizontal: 'left', vertical: 'center', indent: 1 }
                    };
                }
            });
            
            for (var r = 4; r <= 30; r++) {
                var bRef = 'B' + r;
                var cRef = 'C' + r;
                if (ws3[bRef] && ws3[bRef].v) {
                    ws3[bRef].s = {
                        font: { bold: true, sz: 10, color: { rgb: '0F172A' } },
                        alignment: { horizontal: 'left', vertical: 'center' }
                    };
                }
                if (ws3[cRef] && ws3[cRef].v) {
                    ws3[cRef].s = {
                        font: { sz: 10, color: { rgb: '475569' } },
                        alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
                    };
                }
            }
            
            XLSX.utils.book_append_sheet(wb, ws3, 'Hướng dẫn');
            
            var today = new Date();
            var dateStr = today.getFullYear() +
                          String(today.getMonth() + 1).padStart(2, '0') +
                          String(today.getDate()).padStart(2, '0') + '_' +
                          String(today.getHours()).padStart(2, '0') +
                          String(today.getMinutes()).padStart(2, '0');
            
            var fname = 'users_export_' + dateStr + '.xlsx';
            
            XLSX.writeFile(wb, fname, { 
                bookType: 'xlsx',
                cellStyles: true
            });
        } catch(err) {
            console.error('Lỗi export:', err);
            alert('❌ Lỗi export: ' + err.message);
        }
    });
    
    $('importExcelBtn').addEventListener('click', function() {
        $('importFileInput').click();
    });
    
    $('importFileInput').addEventListener('change', function(e) {
        var file = e.target.files[0];
        if (!file) return;
        
        var reader = new FileReader();
        reader.onload = function(evt) {
            try {
                var data = new Uint8Array(evt.target.result);
                var workbook = XLSX.read(data, { type: 'array' });
                var sheet = workbook.Sheets[workbook.SheetNames[0]];
                var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
                processImport(rows);
            } catch(err) {
                alert('❌ Lỗi đọc file: ' + err.message);
            }
            e.target.value = '';
        };
        reader.readAsArrayBuffer(file);
    });
    
    $('importClose').addEventListener('click', function() {
        $('importModal').classList.remove('show');
        importRows = [];
    });
    $('importCancelBtn').addEventListener('click', function() {
        $('importModal').classList.remove('show');
        importRows = [];
    });
    $('importModal').addEventListener('click', function(e) {
        if (e.target === this) {
            $('importModal').classList.remove('show');
            importRows = [];
        }
    });
    
    $('importConfirmBtn').addEventListener('click', doImport);
    
    $('showAddUserBtn').addEventListener('click', function() {
        $('addUserForm').classList.toggle('show');
        if ($('addUserForm').classList.contains('show')) $('newUserEmail').focus();
    });
    $('cancelAddUser').addEventListener('click', function() {
        $('addUserForm').classList.remove('show');
        $('newUserEmail').value = '';
        $('newUserName').value = '';
        $('newUserRole').value = 'user';
        $('newUserExpires').value = '';
    });
    $('confirmAddUser').addEventListener('click', async function() {
        var email = $('newUserEmail').value.trim().toLowerCase();
        var name = $('newUserName').value.trim();
        var role = $('newUserRole').value;
        var expiresVal = $('newUserExpires').value;
        
        if (!email || !email.includes('@')) { alert('Email không hợp lệ'); return; }
        if (!name) name = email.split('@')[0];
        
        if (role === 'admin' && !isSuperAdmin()) {
            alert('⚠️ Chỉ Super Admin mới có quyền thêm admin!');
            return;
        }
        
        try {
            var docRef = db.collection('allowed_users').doc(email);
            var doc = await docRef.get();
            if (doc.exists) { alert('Email này đã tồn tại!'); return; }
            
            var setData = {
                name: name, role: role,
                addedAt: firebase.firestore.FieldValue.serverTimestamp(),
                addedBy: currentUser.email
            };
            
            if (expiresVal && role !== 'admin') {
                var d = new Date(expiresVal + 'T23:59:59');
                if (!isNaN(d.getTime())) {
                    setData.expiresAt = firebase.firestore.Timestamp.fromDate(d);
                }
            }
            
            await docRef.set(setData);
            $('addUserForm').classList.remove('show');
            $('newUserEmail').value = '';
            $('newUserName').value = '';
            $('newUserRole').value = 'user';
            $('newUserExpires').value = '';
            try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
            loadUsers(true);
        } catch(e) { alert('Lỗi: ' + e.message); }
    });
}

function getBorder() {
    return {
        top:    { style: 'thin', color: { rgb: 'CBD5E1' } },
        bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
        left:   { style: 'thin', color: { rgb: 'CBD5E1' } },
        right:  { style: 'thin', color: { rgb: 'CBD5E1' } }
    };
}

function getExpiryDate(expiresAt) {
    if (!expiresAt) return null;
    try {
        var ea = expiresAt;
        if (typeof ea.toDate === 'function') return ea.toDate();
        if (ea.seconds) return new Date(ea.seconds * 1000);
        return new Date(ea);
    } catch(e) {
        return null;
    }
}

function getExpiryTimestamp(expiresAt) {
    var d = getExpiryDate(expiresAt);
    return d ? d.getTime() : null;
}

function formatDate(d) {
    if (!d) return '';
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, '0');
    var day = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
}

function openAdminPanel() {
    if (!currentUser || currentUser.role !== 'admin') return;
    $('adminModal').classList.add('show');
    loadUsers(false);
    loadLogs();
}

function loadUsers(forceRefresh) {
    var cacheKey = 'admin_users_cache';
    
    if (forceRefresh) {
        try { localStorage.removeItem(cacheKey); } catch(e) {}
    }
    
    if (!forceRefresh) {
        try {
            var cached = JSON.parse(localStorage.getItem(cacheKey) || 'null');
            if (cached && cached.expires > Date.now() && cached.data && Array.isArray(cached.data)) {
                usersCache = cached.data;
                renderUsers(usersCache);
                renderAdminStats();
                loadLastLoginMap();
                return;
            }
        } catch(e) {
            try { localStorage.removeItem(cacheKey); } catch(e2) {}
        }
    }
    
    $('userList').innerHTML = '<div class="no-data"><i class="fas fa-spinner fa-pulse"></i>Đang tải...</div>';
    
    db.collection('allowed_users').get()
        .then(function(snapshot) {
            usersCache = [];
            snapshot.forEach(function(doc) {
                var data = doc.data() || {};
                usersCache.push({ 
                    email: doc.id, 
                    name: data.name || '',
                    role: data.role || 'user',
                    expiresAt: data.expiresAt || null
                });
            });
            usersCache.sort(function(a, b) { return (a.email || '').localeCompare(b.email || ''); });
            
            try {
                localStorage.setItem(cacheKey, JSON.stringify({
                    data: usersCache,
                    expires: Date.now() + 5 * 60 * 1000
                }));
            } catch(e) {}
            
            renderUsers(usersCache);
            renderAdminStats();
            return loadLastLoginMap();
        })
        .catch(function(err) {
            $('userList').innerHTML = '<div class="no-data" style="color:#dc2626;padding:1rem"><i class="fas fa-exclamation-triangle"></i>Lỗi: ' + err.message + '</div>';
        });
}

function loadLastLoginMap() {
    var hidden = isHiddenAdmin();
    var myEmail = (currentUser && currentUser.email ? currentUser.email.toLowerCase() : '');
    
    return db.collection('login_logs').orderBy('time', 'desc').limit(500).get()
        .then(function(snapshot) {
            lastLoginMap = {};
            snapshot.forEach(function(doc) {
                var d = doc.data();
                var email = (d.email || '').toLowerCase();
                
                if (hidden && d.role === 'admin' && email !== myEmail) {
                    return;
                }
                
                if (!lastLoginMap[email] && d.time) {
                    lastLoginMap[email] = d.time.toDate();
                }
            });
            renderUsers(usersCache);
        })
        .catch(function() {});
}

function renderAdminStats() {
    var hidden = isHiddenAdmin();
    
    var usersCacheOnly = usersCache.filter(function(u) { 
        return u.role !== 'admin'; 
    });
    var users = usersCacheOnly.length;
    var admins = usersCache.length - users;
    
    var now = Date.now();
    var day1 = 24 * 60 * 60 * 1000;
    var day7 = 7 * 24 * 60 * 60 * 1000;
    
    var active = 0;
    var online = 0;
    var never = 0;
    var expired = 0;
    var expiring = 0;
    
    usersCacheOnly.forEach(function(u) {
        var last = lastLoginMap[(u.email || '').toLowerCase()];
        if (last) {
            var diff = now - last.getTime();
            if (diff <= day7) active++;
            if (diff <= day1) online++;
        } else {
            never++;
        }
        
        if (u.expiresAt) {
            var d = getExpiryDate(u.expiresAt);
            if (d && !isNaN(d.getTime())) {
                var daysLeft = Math.ceil((d.getTime() - now) / (24 * 60 * 60 * 1000));
                if (daysLeft < 0) expired++;
                else if (daysLeft <= 7) expiring++;
            }
        }
    });
    
    if (hidden) {
        $('adminStats').innerHTML = 
            '<div class="stat-card">' +
                '<div class="num">' + users + '</div>' +
                '<div class="label">Tổng User</div>' +
            '</div>' +
            '<div class="stat-card">' +
                '<div class="num" style="color:#16a34a">' + online + '</div>' +
                '<div class="label">Đang hoạt động</div>' +
            '</div>' +
            '<div class="stat-card">' +
                '<div class="num" style="color:#f59e0b">' + expiring + '</div>' +
                '<div class="label">Sắp hết hạn</div>' +
            '</div>' +
            '<div class="stat-card">' +
                '<div class="num" style="color:#dc2626">' + expired + '</div>' +
                '<div class="label">Hết hạn</div>' +
            '</div>';
        
        $('adminUserCount').textContent = users;
        return;
    }
    
    $('adminStats').innerHTML =
        '<div class="stat-card"><div class="num">' + users + '</div><div class="label">Tổng User</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#16a34a">' + online + '</div><div class="label">Đang hoạt động</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#3b82f6">' + active + '</div><div class="label">Active 7d</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#f59e0b">' + admins + '</div><div class="label">Admin</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#94a3b8">' + never + '</div><div class="label">Chưa login</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#f59e0b">' + expiring + '</div><div class="label">Sắp hết hạn</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#dc2626">' + expired + '</div><div class="label">Hết hạn</div></div>';
    
    $('adminUserCount').textContent = users;
}

function renderUsers(items) {
    var list = $('userList');
    var hidden = isHiddenAdmin();
    var superAdmin = isSuperAdmin();
    var myEmail = (currentUser && currentUser.email ? currentUser.email.toLowerCase() : '');
    
    var displayItems = items;
    if (hidden) {
        displayItems = items.filter(function(u) {
            var uEmail = (u.email || '').toLowerCase();
            if (u.role === 'admin' && uEmail !== myEmail) return false;
            return true;
        });
    }
    
    if (!displayItems.length) {
        list.innerHTML = '<div class="no-data" style="padding:1.5rem;font-size:.85rem"><i class="fas fa-search"></i>Không có user nào</div>';
        return;
    }
    
    var adminCount = usersCache.filter(function(u) { return u.role === 'admin'; }).length;
    var now = Date.now();
    var day7 = 7 * 24 * 60 * 60 * 1000;
    var day30 = 30 * 24 * 60 * 60 * 1000;
    
    list.innerHTML = displayItems.map(function(u) {
        var isMe = u.email === currentUser.email;
        var isAdmin = u.role === 'admin';
        var targetIsSuper = (u.email || '').toLowerCase() === SUPER_ADMIN.toLowerCase();
        
        var canModifyAdmin = superAdmin && isAdmin && !isMe && !targetIsSuper;
        
        var roleBtn = '';
        if (isAdmin) {
            if (canModifyAdmin) {
                roleBtn = '<button class="u-btn" onclick="changeRole(\'' + escapeJs(u.email) + '\', \'user\')" title="Hạ xuống User"><i class="fas fa-user"></i></button>';
            } else {
                var reason = isMe ? 'Không thể tự hạ quyền chính mình' : (targetIsSuper ? 'Không thể hạ quyền Super Admin' : 'Chỉ Super Admin mới hạ quyền được');
                roleBtn = '<button class="u-btn" disabled title="' + escapeHtml(reason) + '"><i class="fas fa-user"></i></button>';
            }
        } else {
            if (superAdmin) {
                roleBtn = '<button class="u-btn" onclick="changeRole(\'' + escapeJs(u.email) + '\', \'admin\')" title="Nâng lên Admin"><i class="fas fa-shield-alt"></i></button>';
            } else {
                roleBtn = '<button class="u-btn" disabled title="Chỉ Super Admin mới nâng quyền được"><i class="fas fa-shield-alt"></i></button>';
            }
        }
        
        var deleteBtn = '';
        if (isMe) {
            deleteBtn = '<button class="u-btn danger" disabled title="Không thể tự xóa chính mình"><i class="fas fa-trash"></i></button>';
        } else if (targetIsSuper) {
            deleteBtn = '<button class="u-btn danger" disabled title="Không thể xóa Super Admin"><i class="fas fa-trash"></i></button>';
        } else if (isAdmin) {
            if (canModifyAdmin) {
                deleteBtn = '<button class="u-btn danger" onclick="deleteUser(\'' + escapeJs(u.email) + '\')" title="Xóa admin"><i class="fas fa-trash"></i></button>';
            } else {
                deleteBtn = '<button class="u-btn danger" disabled title="Chỉ Super Admin mới xóa được admin"><i class="fas fa-trash"></i></button>';
            }
        } else {
            deleteBtn = '<button class="u-btn danger" onclick="deleteUser(\'' + escapeJs(u.email) + '\')" title="Xóa"><i class="fas fa-trash"></i></button>';
        }
        
        var expiryBtn = '';
        if (!isAdmin) {
            var btnCls = 'u-btn expiry';
            var tooltip = 'Chỉnh hạn sử dụng';
            
            if (u.expiresAt) {
                var d = getExpiryDate(u.expiresAt);
                if (d && !isNaN(d.getTime())) {
                    var daysLeftExp = Math.ceil((d.getTime() - Date.now()) / (24 * 60 * 60 * 1000));
                    tooltip = 'Chỉnh hạn (hiện tại: ' + d.toLocaleDateString('vi-VN') + ', còn ' + Math.max(0, daysLeftExp) + ' ngày)';
                    if (daysLeftExp <= 7) btnCls += ' urgent';
                }
            } else {
                tooltip = 'Chỉnh hạn (hiện tại: Vĩnh viễn)';
            }
            
            expiryBtn = '<button class="' + btnCls + '" onclick="openEditExpiry(\'' + escapeJs(u.email) + '\')" title="' + escapeHtml(tooltip) + '"><i class="fas fa-calendar-alt"></i></button>';
        }
        
        var lastLoginHtml = '';
        var last = lastLoginMap[(u.email || '').toLowerCase()];
        if (last) {
            var diff = now - last.getTime();
            var cls = diff <= day7 ? 'active' : (diff <= day30 ? 'recent' : '');
            lastLoginHtml = '<div class="u-last-login ' + cls + '"><i class="fas fa-clock"></i> ' + formatTimeDiff(diff) + '</div>';
        } else {
            lastLoginHtml = '<div class="u-last-login"><i class="fas fa-times-circle"></i> Chưa đăng nhập</div>';
        }
        
        var expiryHtml = '';
        if (isAdmin) {
            expiryHtml = '<div class="u-expiry permanent"><i class="fas fa-infinity"></i> Vĩnh viễn</div>';
        } else if (!u.expiresAt) {
            expiryHtml = '<div class="u-expiry permanent" onclick="openEditExpiry(\'' + escapeJs(u.email) + '\')" title="Click để chỉnh"><i class="fas fa-infinity"></i> Vĩnh viễn</div>';
        } else {
            var expDate = getExpiryDate(u.expiresAt);
            
            if (expDate && !isNaN(expDate.getTime())) {
                var daysLeft = Math.ceil((expDate.getTime() - now) / (24 * 60 * 60 * 1000));
                var expCls = 'ok', expIcon = 'fa-calendar-check', expText = '';
                
                if (daysLeft < 0) {
                    expCls = 'expired'; expIcon = 'fa-calendar-times';
                    expText = 'Hết hạn ' + Math.abs(daysLeft) + ' ngày trước';
                } else if (daysLeft === 0) {
                    expCls = 'urgent'; expIcon = 'fa-exclamation-circle';
                    expText = 'Hết hạn hôm nay';
                } else if (daysLeft <= 3) {
                    expCls = 'urgent'; expIcon = 'fa-exclamation-circle';
                    expText = 'Còn ' + daysLeft + ' ngày';
                } else if (daysLeft <= 7) {
                    expCls = 'warn'; expIcon = 'fa-clock';
                    expText = 'Còn ' + daysLeft + ' ngày';
                } else {
                    expText = 'Còn ' + daysLeft + ' ngày';
                }
                
                expiryHtml = '<div class="u-expiry ' + expCls + '" onclick="openEditExpiry(\'' + escapeJs(u.email) + '\')" title="Click để chỉnh"><i class="fas ' + expIcon + '"></i> ' + 
                             escapeHtml(expText) + ' • ' + expDate.toLocaleDateString('vi-VN') + '</div>';
            }
        }
        
        var roleBadge = isAdmin 
            ? '<span class="u-role ' + (targetIsSuper ? 'super' : 'admin') + '">' + (targetIsSuper ? '👑 super' : 'admin') + '</span>'
            : '<span class="u-role user">user</span>';
        
        return '<div class="user-row" data-email="' + escapeHtml(u.email) + '">' +
            '<div class="u-info">' +
                '<div class="u-name">' + escapeHtml(u.name || u.email.split('@')[0]) + (isMe ? ' <span style="color:#94a3b8;font-size:.7rem">(bạn)</span>' : '') + '</div>' +
                '<div class="u-email">' + escapeHtml(u.email) + '</div>' +
                lastLoginHtml +
                expiryHtml +
            '</div>' +
            roleBadge +
            '<div class="u-actions">' + expiryBtn + roleBtn + deleteBtn + '</div>' +
        '</div>';
    }).join('');
}

window.changeRole = async function(email, newRole) {
    var target = usersCache.find(function(u) { return u.email === email; });
    if (!target) { alert('Không tìm thấy user!'); return; }
    
    var isMe = email === currentUser.email;
    var isAdmin = target.role === 'admin';
    var superAdmin = isSuperAdmin();
    var targetIsSuper = (email || '').toLowerCase() === SUPER_ADMIN.toLowerCase();
    
    if (isMe && newRole === 'user') { 
        alert('⚠️ Không thể tự hạ quyền admin của chính mình!'); 
        return; 
    }
    
    if (targetIsSuper) {
        alert('⚠️ Không thể thay đổi quyền của Super Admin!');
        return;
    }
    
    if (isAdmin && newRole === 'user') {
        if (!superAdmin) {
            alert('⚠️ Chỉ Super Admin mới có quyền hạ cấp admin khác!');
            return;
        }
    }
    
    if (!isAdmin && newRole === 'admin') {
        if (!superAdmin) {
            alert('⚠️ Chỉ Super Admin mới có quyền nâng cấp lên admin!');
            return;
        }
    }
    
    var action;
    if (newRole === 'admin') action = 'NÂNG LÊN ADMIN';
    else action = 'HẠ XUỐNG USER';
    
    if (!confirm(action + ' cho tài khoản:\n\n' + email + '\n\nBạn có chắc không?')) return;
    
    try {
        await db.collection('allowed_users').doc(email).update({ role: newRole });
        try { localStorage.removeItem('user_cache_' + email); } catch(e) {}
        try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
        loadUsers(true);
    }
    catch(e) { alert('Lỗi: ' + e.message); }
};

window.deleteUser = async function(email) {
    var target = usersCache.find(function(u) { return u.email === email; });
    if (!target) { alert('Không tìm thấy user!'); return; }
    
    var isMe = email === currentUser.email;
    var isAdmin = target.role === 'admin';
    var superAdmin = isSuperAdmin();
    var targetIsSuper = (email || '').toLowerCase() === SUPER_ADMIN.toLowerCase();
    
    if (isMe) { 
        alert('⚠️ Không thể tự xóa tài khoản của chính mình!'); 
        return; 
    }
    
    if (targetIsSuper) {
        alert('⚠️ Không thể xóa Super Admin!');
        return;
    }
    
    if (isAdmin && !superAdmin) {
        alert('⚠️ Chỉ Super Admin mới có quyền xóa admin khác!');
        return;
    }
    
    var confirmMsg = isAdmin 
        ? '⚠️ XÓA ADMIN\n\n' + email + '\n\nNgười này sẽ mất quyền quản trị và không đăng nhập được nữa.\n\nBạn có chắc không?'
        : '⚠️ XÓA TÀI KHOẢN\n\n' + email + '\n\nNgười này sẽ không đăng nhập được nữa.\n\nBạn có chắc không?';
    
    if (!confirm(confirmMsg)) return;
    
    try {
        await db.collection('allowed_users').doc(email).delete();
        try { localStorage.removeItem('user_cache_' + email); } catch(e) {}
        try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
        loadUsers(true);
    }
    catch(e) { alert('Lỗi: ' + e.message); }
};

function loadLogs() {
    var hidden = isHiddenAdmin();
    var logTitleEl = $('logsTitle');
    
    if (hidden) {
        if (logTitleEl) logTitleEl.style.display = 'none';
        $('logsList').style.display = 'none';
        $('logsList').innerHTML = '';
        return;
    }
    
    if (logTitleEl) logTitleEl.style.display = 'flex';
    $('logsList').style.display = 'block';
    
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

function processImport(rows) {
    if (!rows || rows.length < 2) {
        alert('❌ File rỗng hoặc thiếu header!');
        return;
    }
    
    var headerRowIdx = -1;
    for (var i = 0; i < Math.min(5, rows.length); i++) {
        var r = rows[i].map(function(c) { return String(c || '').toLowerCase().trim(); });
        if (r.indexOf('email') !== -1) { headerRowIdx = i; break; }
    }
    if (headerRowIdx === -1) {
        alert('❌ Không tìm thấy cột "email"!\n\nFile Excel cần có ít nhất cột "email".');
        return;
    }
    
    var header = rows[headerRowIdx].map(function(c) { 
        return String(c || '').toLowerCase().trim(); 
    });
    var emailCol = header.indexOf('email');
    var nameCol = header.indexOf('name');
    var expCol = header.indexOf('expiresat');
    
    var warnings = [];
    if (nameCol === -1) warnings.push('Thiếu cột "name" → sẽ dùng phần trước @ của email làm tên');
    if (expCol === -1) warnings.push('Thiếu cột "expiresAt" → tài khoản sẽ là vĩnh viễn');
    
    var ignoredCols = [];
    header.forEach(function(h, idx) {
        if (idx !== emailCol && idx !== nameCol && idx !== expCol && h) {
            ignoredCols.push(h);
        }
    });
    
    var existingUserMap = {};
    var existingAdminSet = {};
    usersCache.forEach(function(u) { 
        var email = (u.email || '').toLowerCase();
        if (u.role === 'admin') {
            existingAdminSet[email] = true;
        } else {
            existingUserMap[email] = u; 
        }
    });
    
    importRows = [];
    var stats = { total: 0, newUser: 0, update: 0, invalid: 0, skippedAdmin: 0 };
    var seenInFile = {};
    
    for (var i = headerRowIdx + 1; i < rows.length; i++) {
        var row = rows[i];
        if (!row || row.length === 0) continue;
        
        var email = emailCol >= 0 ? String(row[emailCol] || '').trim().toLowerCase() : '';
        var name = nameCol >= 0 ? String(row[nameCol] || '').trim() : '';
        var expRaw = expCol >= 0 ? row[expCol] : '';
        
        if (!email && !name) continue;
        if (email.indexOf('←') === 0 || email.indexOf('•') === 0 || 
            email.indexOf('xóa dòng') !== -1 || email.indexOf('#') === 0 ||
            email.indexOf('⚠') === 0 || email.indexOf('ví dụ') === 0) continue;
        
        if (existingAdminSet[email]) {
            stats.skippedAdmin++;
            continue;
        }
        
        var status = 'ok';
        var reason = '';
        var isUpdate = !!existingUserMap[email];
        
        if (!email) {
            status = 'error'; reason = 'Thiếu email'; stats.invalid++;
        } else if (!email.includes('@') || !email.includes('.')) {
            status = 'error'; reason = 'Email không hợp lệ'; stats.invalid++;
        } else if (seenInFile[email]) {
            status = 'error'; reason = 'Trùng trong file'; stats.invalid++;
        } else if (isUpdate) {
            stats.update++;
            seenInFile[email] = true;
        } else {
            stats.newUser++;
            seenInFile[email] = true;
        }
        
        if (!name && email.indexOf('@') > 0) {
            name = email.split('@')[0];
        }
        
        var expDate = null;
        var expStr = '';
        if (expRaw) {
            var raw = expRaw;
            if (typeof raw === 'number' && raw > 25569) {
                var d = new Date((raw - 25569) * 86400 * 1000);
                if (!isNaN(d.getTime())) {
                    expDate = d;
                    var yy = d.getFullYear();
                    var mm = String(d.getMonth() + 1).padStart(2, '0');
                    var dd = String(d.getDate()).padStart(2, '0');
                    expStr = yy + '-' + mm + '-' + dd;
                }
            } else {
                var s = String(raw).trim();
                if (s) {
                    var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
                    if (m) {
                        var d2 = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]), 23, 59, 59);
                        if (!isNaN(d2.getTime())) {
                            expDate = d2;
                            expStr = m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0');
                        }
                    } else {
                        var d3 = new Date(s);
                        if (!isNaN(d3.getTime())) {
                            expDate = d3;
                            var y3 = d3.getFullYear();
                            var mo3 = String(d3.getMonth() + 1).padStart(2, '0');
                            var da3 = String(d3.getDate()).padStart(2, '0');
                            expStr = y3 + '-' + mo3 + '-' + da3;
                        } else {
                            if (status === 'ok') {
                                status = 'warn';
                                reason = 'Ngày không hợp lệ (bỏ qua hạn)';
                            }
                            expDate = null;
                        }
                    }
                }
            }
        }
        
        importRows.push({
            rowNum: i + 1,
            email: email,
            name: name,
            role: 'user',
            expDate: expDate,
            expStr: expStr,
            status: status,
            reason: reason,
            isUpdate: isUpdate
        });
    }
    
    if (importRows.length === 0) {
        var msg = '❌ Không có dòng dữ liệu hợp lệ!';
        if (stats.skippedAdmin > 0) {
            msg += '\n\n(Bỏ qua ' + stats.skippedAdmin + ' admin - không import admin qua Excel)';
        }
        alert(msg);
        return;
    }
    
    renderImportPreview();
    $('importModal').classList.add('show');
}

function renderImportPreview() {
    var tbody = $('importTableBody');
    if (!tbody) return;
    
    var html = '';
    var countOk = 0, countUpdate = 0, countWarn = 0, countErr = 0;
    
    importRows.forEach(function(r) {
        var rowCls = '';
        var statusHtml = '';
        
        if (r.status === 'ok' && r.isUpdate) {
            rowCls = 'row-update';
            statusHtml = '<span class="status-badge update"><i class="fas fa-sync-alt"></i> Cập nhật</span>';
            countUpdate++;
        } else if (r.status === 'ok') {
            rowCls = 'row-new';
            statusHtml = '<span class="status-badge ok"><i class="fas fa-plus"></i> Thêm mới</span>';
            countOk++;
        } else if (r.status === 'warn') {
            rowCls = 'row-warn';
            statusHtml = '<span class="status-badge warn"><i class="fas fa-exclamation-triangle"></i> ' + escapeHtml(r.reason) + '</span>';
            countWarn++;
        } else {
            rowCls = 'row-error';
            statusHtml = '<span class="status-badge err"><i class="fas fa-times"></i> ' + escapeHtml(r.reason) + '</span>';
            countErr++;
        }
        
        var expDisplay = '—';
        if (r.expStr) {
            var d = new Date(r.expStr + 'T23:59:59');
            var daysLeft = Math.ceil((d.getTime() - Date.now()) / (24 * 60 * 60 * 1000));
            var expColor = daysLeft < 0 ? '#dc2626' : (daysLeft <= 7 ? '#f59e0b' : '#16a34a');
            expDisplay = '<span style="color:' + expColor + ';font-size:.7rem;font-weight:600;">' + r.expStr + '</span>';
        } else {
            expDisplay = '<span style="color:#94a3b8;font-size:.7rem;">Vĩnh viễn</span>';
        }
        
        html += '<tr class="' + rowCls + '">' +
            '<td>' + r.rowNum + '</td>' +
            '<td><b>' + escapeHtml(r.email) + '</b></td>' +
            '<td>' + escapeHtml(r.name) + '</td>' +
            '<td><span class="role-badge user">user</span></td>' +
            '<td>' + expDisplay + '</td>' +
            '<td>' + statusHtml + '</td>' +
        '</tr>';
    });
    
    tbody.innerHTML = html;
    
    var summaryEl = $('importSummary');
    if (summaryEl) {
        summaryEl.innerHTML =
            '<div class="import-stat"><div class="num">' + importRows.length + '</div><div class="label">Tổng</div></div>' +
            '<div class="import-stat ok"><div class="num">' + countOk + '</div><div class="label">Thêm mới</div></div>' +
            '<div class="import-stat update"><div class="num">' + countUpdate + '</div><div class="label">Cập nhật</div></div>' +
            '<div class="import-stat warn"><div class="num">' + countWarn + '</div><div class="label">Cảnh báo</div></div>' +
            '<div class="import-stat err"><div class="num">' + countErr + '</div><div class="label">Lỗi</div></div>';
    }
    
    var totalImportable = countOk + countUpdate + countWarn;
    
    var countEl = $('importCount');
    if (countEl) {
        countEl.textContent = totalImportable;
    }
    
    var confirmBtn = $('importConfirmBtn');
    if (confirmBtn) {
        confirmBtn.disabled = totalImportable === 0;
    }
}

async function doImport() {
    var skipDuplicates = $('importSkipDuplicates').checked;
    var skipInvalid = $('importSkipInvalid').checked;
    
    var toImport = importRows.filter(function(r) {
        if (r.status === 'error') return false;
        if (r.status === 'warn' && skipInvalid) return false;
        if (r.isUpdate && skipDuplicates) return false;
        return true;
    });
    
    if (toImport.length === 0) {
        alert('⚠️ Không có user nào để import!');
        return;
    }
    
    var totalNew = toImport.filter(function(r) { return !r.isUpdate; }).length;
    var totalUpdate = toImport.filter(function(r) { return r.isUpdate; }).length;
    
    if (!confirm('📥 IMPORT ' + toImport.length + ' TÀI KHOẢN?\n\n' +
                 '• Thêm mới: ' + totalNew + '\n' +
                 '• Cập nhật: ' + totalUpdate + '\n\n' +
                 '(Chỉ import user, KHÔNG import admin)\n\n' +
                 'Bạn có chắc không?')) return;
    
    var btn = $('importConfirmBtn');
    if (!btn) return;
    
    var originalHTML = btn.innerHTML;
    var originalDisabled = btn.disabled;
    
    btn.disabled = true;
    
    var icon = btn.querySelector('i');
    if (icon) icon.className = 'fas fa-spinner fa-pulse';
    
    var success = 0;
    var failed = 0;
    var errors = [];
    
    var adminEmails = {};
    usersCache.forEach(function(u) {
        if (u.role === 'admin') adminEmails[(u.email || '').toLowerCase()] = true;
    });
    
    var BATCH_SIZE = 400;
    for (var i = 0; i < toImport.length; i += BATCH_SIZE) {
        var chunk = toImport.slice(i, i + BATCH_SIZE);
        
        chunk = chunk.filter(function(r) {
            if (adminEmails[r.email]) {
                console.warn('Skip admin:', r.email);
                return false;
            }
            return true;
        });
        
        if (chunk.length === 0) continue;
        
        var batch = db.batch();
        
        chunk.forEach(function(r) {
            var ref = db.collection('allowed_users').doc(r.email);
            var data = {
                name: r.name,
                role: 'user',
                addedBy: currentUser.email
            };
            
            if (!r.isUpdate) {
                data.addedAt = firebase.firestore.FieldValue.serverTimestamp();
                data.importedFromExcel = true;
            } else {
                data.updatedAt = firebase.firestore.FieldValue.serverTimestamp();
            }
            
            if (r.expDate) {
                try {
                    data.expiresAt = firebase.firestore.Timestamp.fromDate(r.expDate);
                } catch(e) {
                    data.expiresAt = null;
                }
            } else {
                data.expiresAt = null;
            }
            
            batch.set(ref, data, { merge: true });
        });
        
        try {
            await batch.commit();
            success += chunk.length;
        } catch(err) {
            failed += chunk.length;
            errors.push(err.message);
            console.error('Batch error:', err);
        }
        
        var countEl = btn.querySelector('#importCount');
        if (countEl) {
            countEl.textContent = success + '/' + toImport.length;
        }
    }
    
    btn.disabled = originalDisabled;
    btn.innerHTML = originalHTML;
    
    importRows = [];
    
    try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
    toImport.forEach(function(r) {
        try { localStorage.removeItem('user_cache_' + r.email); } catch(e) {}
    });
    
    var msg = '✅ Import hoàn tất!\n\n' +
              '✓ Thành công: ' + success + '\n' +
              (failed ? '✗ Thất bại: ' + failed + '\n' : '') +
              (errors.length ? '\nLỗi:\n' + errors.slice(0, 3).join('\n') : '');
    alert(msg);
    
    $('importModal').classList.remove('show');
    
    loadUsers(true);
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
    .replace("__TIKTOK_USERNAME__", TIKTOK_USERNAME)
    .replace("__TIKTOK_NICKNAME__", TIKTOK_NICKNAME)
    .replace("__TIKTOK_AVATAR__", TIKTOK_AVATAR)
    .replace("__TIKTOK_URL__", TIKTOK_URL)
    .replace("__DEMO_LIMIT__", str(DEMO_LIMIT))
    .replace("__DEMO_DAILY_LIMIT__", str(DEMO_DAILY_LIMIT))
    .replace("__DEMO_HSK_MAX__", str(DEMO_HSK_MAX))
    .replace("__TARGET_ADMINS__", str(TARGET_ADMINS))
    .replace("__SUPER_ADMIN__", SUPER_ADMIN)
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
print(f"👑 Super admin: {SUPER_ADMIN}")
print(f"🎵 TikTok: @{TIKTOK_USERNAME} ({TIKTOK_NICKNAME})")
print(f"   → Header: 'Học tiếng Trung' + 'Văn phòng & Công xưởng' (lùi phải)")
print(f"   → TikTok info bar dưới header (giữ nguyên Thảo nói 中文)")
print(f"   → Chế độ thường: Zalo + TikTok (floating left)")
print(f"   → Chế độ luyện tập: CHỈ TikTok")
print(f"   → PC hover TikTok: card info (avatar + nickname + stats)")
print(f"   → Mobile 1 tap TikTok: card info; 2 tap: mở TikTok")
print(f"   → Mobile TikTok button: CHỈ icon tròn nhỏ gọn")
print(f"   → Avatar TikTok: {'Có' if TIKTOK_AVATAR else 'Fallback SVG'}")
print(f"🚫 FIX: Đã sửa lỗi TRÙNG LẶP 2 ô tìm kiếm / 2 ô HSK / 2 ô chủ đề")
print(f"      → Ẩn search bar + filters của header khi mở modal luyện tập")
print(f"      → Modal set top 1 lần theo chiều cao header, không cập nhật liên tục")
print(f"✅ Header KHÔNG biến mất khi bật chế độ luyện tập")
print(f"🔔🔕 Nút FAB 'Silent mode' - ICON CHUÔNG BẬT / CHUÔNG TẮT:")
print(f"      → Bình thường: 🔔 fa-bell + nút XANH primary (active)")
print(f"      → Silent mode: 🔕 fa-bell-slash + nút XÁM (chưa active)")
print(f"      → Màu GIỐNG các nút FAB khác")
print(f"      → Ẩn/hiện Zalo + TikTok + TikTok bar")
print(f"      → CHỈ hiển thị cho user/admin (KHÔNG hiển thị demo)")
print(f"      → Lưu trạng thái vào localStorage (focusHidden)")
print(f"      → Đã bỏ tooltip ::after → không còn chấm đen khi hover/nhấn")
