# -*- coding: utf-8 -*-
"""Load + validate config.json. KHÔNG cần sửa khi đổi cấu trúc Excel/UI."""
import json
import os
import sys

CONFIG_FILE = "scripts/config.json"

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
DEFAULT_FILLERS = ["了", "的", "吗", "呢", "吧", "啊", "呀", "哦", "嘛", "哈", "哪", "着", "过"]


def load_config():
    if not os.path.exists(CONFIG_FILE):
        print(f"❌ Không tìm thấy file cấu hình: {CONFIG_FILE}")
        sys.exit(1)

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    if not cfg.get("firebase_config", {}).get("apiKey"):
        print(f"❌ Firebase config chưa được cấu hình trong {CONFIG_FILE}")
        sys.exit(1)

    # Defaults
    cfg.setdefault("excel_file", "data/input.xlsx")
    cfg.setdefault("output_html", "index.html")
    cfg.setdefault("sheet_index", 0)
    cfg.setdefault("demo_limit", 50)
    cfg.setdefault("demo_daily_limit", 100)
    cfg.setdefault("demo_hsk_max", 3)
    cfg.setdefault("target_admins", 2)
    cfg.setdefault("super_admin", "hoanginvest@gmail.com")
    cfg.setdefault("zalo_phone", "")
    cfg.setdefault("zalo_name", "Hỗ trợ")
    cfg.setdefault("tiktok_username", "thaonoizhongwen")
    cfg.setdefault("tiktok_nickname", "Thảo nói 中文")
    cfg.setdefault("tiktok_avatar", "")
    cfg.setdefault("tiktok_url",
                   f"https://www.tiktok.com/@{cfg['tiktok_username']}")
    cfg.setdefault("synonyms", DEFAULT_SYNONYMS)
    cfg.setdefault("filler_words", DEFAULT_FILLERS)

    return cfg


def print_banner(CONFIG):
    print(f"⚙️  Đã đọc cấu hình từ: {CONFIG_FILE}")
    print(f"   📞 Zalo: {CONFIG['zalo_phone']} ({CONFIG['zalo_name']})")
    print(f"   🎵 TikTok: @{CONFIG['tiktok_username']} ({CONFIG['tiktok_nickname']})")
    print(f"   🖼️  TikTok Avatar: {'Có' if CONFIG['tiktok_avatar'] else 'Không (dùng fallback)'}")
    print(f"   🎁 Demo: {CONFIG['demo_limit']} câu + HSK1-{CONFIG['demo_hsk_max']} + {CONFIG['demo_daily_limit']} lượt")
    print(f"   👑 Super admin: {CONFIG['super_admin']}")
