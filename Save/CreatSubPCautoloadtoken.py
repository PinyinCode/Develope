# -*- coding: utf-8 -*-
import os
import re
import sys
import threading
import json
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog, simpledialog
from pypinyin import pinyin, Style
import yt_dlp
from deep_translator import GoogleTranslator
import whisper

def get_ffmpeg_path():
    """Tự động tìm đường dẫn thư mục bin của FFmpeg"""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
        bundled_ffmpeg = os.path.join(base_path, 'ffmpeg', 'bin')
        if os.path.exists(bundled_ffmpeg):
            return bundled_ffmpeg
    default_path = r'C:\ffmpeg\bin'
    return default_path

def extract_video_id(url):
    """Trích xuất Video ID từ đường dẫn YouTube"""
    match = re.search(r'(?:v=|\/)([0-9A-Za-z_-]{11}).*', url)
    return match.group(1) if match else "unknown_video"

class SubtitleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Phần Mềm Tạo Phụ Đề AI - iOS Style")
        self.root.geometry("580x720")
        self.root.minsize(520, 680)
        
        self.config_file = 'config.json'

        # --- CẤU HÌNH STYLE ---
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        BG_COLOR = "#000000"
        CARD_BG = "#121214"
        CARD_ELEMENT = "#1C1C1E"
        TEXT_PRIMARY = "#FFFFFF"
        TEXT_SECONDARY = "#8E8E93"
        
        self.root.configure(bg=BG_COLOR)
        
        self.style.configure('TFrame', background=BG_COLOR)
        self.style.configure('Card.TFrame', background=CARD_BG, relief='flat', borderwidth=0)
        self.style.configure('TLabel', background=BG_COLOR, foreground=TEXT_PRIMARY, font=('Segoe UI', 9))
        self.style.configure('Card.TLabel', background=CARD_BG, foreground=TEXT_PRIMARY, font=('Segoe UI', 9))
        self.style.configure('Title.TLabel', background=BG_COLOR, font=('Segoe UI', 16, 'bold'), foreground=TEXT_PRIMARY, anchor='center')
        self.style.configure('SectionTitle.TLabel', background=CARD_BG, font=('Segoe UI', 9, 'bold'), foreground=TEXT_SECONDARY)
        self.style.configure('Status.TLabel', background=CARD_BG, font=('Segoe UI', 9), foreground=TEXT_SECONDARY)
        
        self.style.configure('TCheckbutton', background=CARD_BG, foreground=TEXT_PRIMARY, font=('Segoe UI', 9))
        self.style.map('TCheckbutton', background=[('active', CARD_BG)], indicatorcolor=[('selected', '#0A84FF')])
        
        self.style.configure('TButton', font=('Segoe UI', 9, 'bold'), borderwidth=0, relief='flat')
        self.style.configure('Primary.TButton', background='#FFFFFF', foreground='#000000', padding=(12, 6))
        self.style.map('Primary.TButton',
            background=[('active', '#D1D1D6'), ('disabled', '#2C2C2E')],
            foreground=[('active', '#000000'), ('disabled', TEXT_SECONDARY)]
        )
        
        self.style.configure('Danger.TButton', background='#2C2C2E', foreground='#FF453A', padding=(12, 6))
        self.style.map('Danger.TButton',
            background=[('active', '#48484A'), ('disabled', '#1C1C1E')],
            foreground=[('active', '#FF6961'), ('disabled', TEXT_SECONDARY)]
        )

        self.style.configure('Square.TButton', background='#2C2C2E', foreground='#FFFFFF', padding=(8, 4))
        self.style.map('Square.TButton',
            background=[('active', '#3A3A3C'), ('disabled', '#1C1C1E')],
            foreground=[('active', '#FFFFFF'), ('disabled', TEXT_SECONDARY)]
        )

        self.style.configure('TEntry', fieldbackground=CARD_ELEMENT, foreground=TEXT_PRIMARY, insertcolor=TEXT_PRIMARY, borderwidth=0, relief='flat')
        self.style.configure('TProgressbar', background='#0A84FF', troughcolor=CARD_ELEMENT, borderwidth=0, thickness=6)

        self.cancel_event = threading.Event()
        self.current_output_file = None
        self.is_processing = False

        # --- FOOTER ---
        footer_frame = tk.Frame(root, bg=BG_COLOR, height=32)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=24, pady=(0, 10))
        dev_label = tk.Label(footer_frame, text="Nhà phát triển: +84528484668  |  hoanginvest199@gmail.com", font=("Segoe UI", 8, "bold"), fg=TEXT_SECONDARY, bg=BG_COLOR)
        dev_label.pack(expand=True)

        # --- MAIN CONTAINER ---
        main_container = ttk.Frame(root, style='TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=24, pady=(20, 5))

        title_label = ttk.Label(main_container, text="AI Pinyin Subtitle", style='Title.TLabel', anchor='center')
        title_label.pack(fill=tk.X, pady=(0, 16))

        # --- CARD 1: NHẬP LINK ---
        link_card = ttk.Frame(main_container, style='Card.TFrame', padding=16)
        link_card.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(link_card, text="ĐƯỜNG DẪN YOUTUBE", style='SectionTitle.TLabel').pack(anchor=tk.W, pady=(0, 4))
        ttk.Label(link_card, text="Tự động nhận diện và dịch sang Trung - Pinyin - Việt", style='Status.TLabel').pack(anchor=tk.W, pady=(0, 8))
        
        url_row = ttk.Frame(link_card, style='Card.TFrame')
        url_row.pack(fill=tk.X, pady=(0, 10))

        self.url_entry = ttk.Entry(url_row, font=('Segoe UI', 10))
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6, ipadx=6)

        self.btn_clear_url = ttk.Button(url_row, text="✕", style='Square.TButton', command=self.clear_url_entry)
        self.btn_clear_url.pack(side=tk.RIGHT, padx=(6, 0))

        action_row = ttk.Frame(link_card, style='Card.TFrame')
        action_row.pack(fill=tk.X)

        self.btn_action = ttk.Button(action_row, text="TẠO PHỤ ĐỀ AI", style='Primary.TButton', command=self.handle_action_button)
        self.btn_action.pack(side=tk.LEFT)

        # --- CARD 2: CÀI ĐẶT GITHUB ---
        save_card = ttk.Frame(main_container, style='Card.TFrame', padding=16)
        save_card.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(save_card, text="THƯ MỤC LƯU FILE", style='SectionTitle.TLabel').pack(anchor=tk.W, pady=(0, 8))
        
        self.save_dir_var = tk.StringVar(value=os.getcwd())
        
        save_row = ttk.Frame(save_card, style='Card.TFrame')
        save_row.pack(fill=tk.X)
        
        self.save_entry = ttk.Entry(save_row, textvariable=self.save_dir_var, font=('Segoe UI', 9), state="readonly")
        self.save_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6, ipadx=6)
        
        self.btn_browse = ttk.Button(save_row, text="Duyệt...", style='Square.TButton', command=self.browse_save_dir)
        self.btn_browse.pack(side=tk.RIGHT, padx=(6, 0))

        # === Ô NHẬP TOKEN (TỰ ĐỘNG CẬP NHẬT) ===
        ttk.Label(save_card, text="GITHUB TOKEN", style='SectionTitle.TLabel').pack(anchor=tk.W, pady=(12, 4))
        self.token_var = tk.StringVar()
        self.token_entry = ttk.Entry(save_card, textvariable=self.token_var, font=('Segoe UI', 9), show='*')
        self.token_entry.pack(fill=tk.X, ipady=4, ipadx=4)
        self.token_entry.bind('<FocusOut>', self.on_token_changed)

        # Checkbox GitHub
        self.upload_gh_var = tk.BooleanVar(value=False)
        self.chk_github = ttk.Checkbutton(save_card, text="☁️ Đồng bộ bản sao lên GitHub Gist", variable=self.upload_gh_var, style='TCheckbutton')
        self.chk_github.pack(anchor=tk.W, pady=(8, 0))

        # Load token đã lưu
        self.load_token_to_ui()

        # --- CARD 3: TIẾN TRÌNH ---
        progress_card = ttk.Frame(main_container, style='Card.TFrame', padding=16)
        progress_card.pack(fill=tk.X, pady=(0, 10))

        self.progress_label = ttk.Label(progress_card, text="Trạng thái: Sẵn sàng hoạt động", style='Status.TLabel')
        self.progress_label.pack(anchor=tk.W, pady=(0, 8))

        self.progress_bar = ttk.Progressbar(progress_card, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X)

        # --- CARD 4: NHẬT KÝ ---
        log_card = ttk.Frame(main_container, style='Card.TFrame', padding=16)
        log_card.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
        
        ttk.Label(log_card, text="NHẬT KÝ HỆ THỐNG", style='SectionTitle.TLabel').pack(anchor=tk.W, pady=(0, 8))
        
        log_inner = tk.Frame(log_card, bg=CARD_ELEMENT, bd=0)
        log_inner.pack(fill=tk.BOTH, expand=True)
        
        self.log_area = scrolledtext.ScrolledText(
            log_inner, font=("Consolas", 8), bg=CARD_ELEMENT, fg=TEXT_PRIMARY, insertbackground=TEXT_PRIMARY,
            relief='flat', borderwidth=8, highlightthickness=0, wrap=tk.WORD, selectbackground="#3A3A3C", selectforeground="#FFFFFF"
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)

    # ===== TOKEN MANAGEMENT =====
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            except:
                pass
        return {}

    def save_config(self, data):
        try:
            with open(self.config_file, 'w') as f:
                json.dump(data, f)
        except Exception as e:
            self.log(f"Không thể lưu cấu hình: {e}")

    def on_token_changed(self, event=None):
        """Tự động lưu token khi thay đổi"""
        new_token = self.token_var.get().strip()
        config = self.load_config()
        old_token = config.get('gh_pat_token', '')
        
        if new_token != old_token:
            config['gh_pat_token'] = new_token
            self.save_config(config)
            if new_token:
                self.log("Đã cập nhật Token GitHub mới ✓")
            else:
                self.log("Đã xóa Token GitHub")

    def load_token_to_ui(self):
        """Hiển thị token đã lưu lên giao diện"""
        config = self.load_config()
        token = config.get('gh_pat_token', '')
        self.token_var.set(token)

    def clear_url_entry(self):
        self.url_entry.delete(0, tk.END)

    def browse_save_dir(self):
        dir_path = filedialog.askdirectory(title="Chọn thư mục lưu file phụ đề", initialdir=self.save_dir_var.get())
        if dir_path:
            self.save_dir_var.set(dir_path)

    def log(self, message):
        def _append():
            self.log_area.insert(tk.END, message + "\n")
            self.log_area.see(tk.END)
        self.root.after(0, _append)

    def set_progress(self, val, text=""):
        def _update():
            self.progress_bar['value'] = val
            if text:
                self.progress_label.config(text=text)
        self.root.after(0, _update)

    def check_cancel(self):
        if self.cancel_event.is_set():
            raise Exception("TÁC VỤ ĐÃ BỊ HỦY BỞI NGƯỜI DÙNG!")

    def ytdl_progress_hook(self, d):
        if self.cancel_event.is_set():
            raise yt_dlp.utils.DownloadError("Đã hủy tải theo yêu cầu người dùng.")
        if d['status'] == 'downloading':
            downloaded = d.get('downloaded_bytes', 0)
            total = d.get('total_bytes') or d.get('total_bytes_estimate', 0)
            if total > 0:
                percent = (downloaded / total) * 100
                speed = d.get('_speed_str', '').strip()
                self.set_progress(percent, f"Đang tải âm thanh: {percent:.1f}% ({speed})")

    def handle_action_button(self):
        if not self.is_processing:
            url = self.url_entry.get().strip()
            if not url:
                messagebox.showwarning("Cảnh báo", "Vui lòng dán link YouTube vào ô trống!")
                return
            
            if self.upload_gh_var.get():
                # Lưu token hiện tại từ ô nhập
                self.on_token_changed()
                config = self.load_config()
                token = config.get('gh_pat_token', '')
                if not token:
                    messagebox.showwarning("Cảnh báo", "Vui lòng nhập GitHub Token vào ô bên trên!")
                    return
            
            self.cancel_event.clear()
            self.current_output_file = None
            self.is_processing = True
            
            self.btn_action.config(text="HỦY BỎ", style='Danger.TButton')
            self.log_area.delete(1.0, tk.END)
            self.set_progress(0, "Đang khởi tạo...")
            
            threading.Thread(target=self.process_video_with_ai, args=(url,), daemon=True).start()
        else:
            if not self.cancel_event.is_set():
                self.cancel_event.set()
                self.log("Đang tiến hành hủy tác vụ và dọn dẹp...")
                self.set_progress(0, "Đang dừng...")
                self.btn_action.config(state=tk.DISABLED)

    def cleanup_files(self):
        for file in os.listdir('.'):
            if file.startswith('temp_audio') or file.startswith('temp_sub'):
                try:
                    os.remove(file)
                except: pass
        if self.cancel_event.is_set() and self.current_output_file and os.path.exists(self.current_output_file):
            try:
                os.remove(self.current_output_file)
                self.log(f"Đã xóa file chưa hoàn tất: {self.current_output_file}")
            except: pass

    def upload_to_github_gist(self, file_path, video_id):
        self.log("Kiểm tra và đồng bộ phụ đề lên GitHub Gist...")
        self.set_progress(95, "Đang đồng bộ GitHub...")
        config = self.load_config()
        token = config.get('gh_pat_token')
        
        if not token:
            self.log("Lỗi: Không tìm thấy Token cấu hình.")
            return

        target_filename = f"{video_id}.vtt"

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            existing_gist_id = None
            page = 1
            while True:
                req_list = urllib.request.Request(f"https://api.github.com/gists?page={page}&per_page=100")
                req_list.add_header("Authorization", f"token {token}")
                try:
                    with urllib.request.urlopen(req_list) as resp:
                        gists = json.loads(resp.read().decode('utf-8'))
                        if not gists:
                            break
                        found = False
                        for g in gists:
                            if target_filename in g.get('files', {}):
                                existing_gist_id = g['id']
                                found = True
                                break
                        if found:
                            break
                        page += 1
                except:
                    break

            if existing_gist_id:
                self.log(f"Phát hiện Gist cũ cho video [{video_id}]. Đang tiến hành ghi đè cập nhật...")
                payload = {
                    "description": f"Pinyin AI Subtitle - {video_id}",
                    "files": {
                        target_filename: {
                            "content": content
                        }
                    }
                }
                req = urllib.request.Request(f"https://api.github.com/gists/{existing_gist_id}", method="PATCH")
            else:
                self.log(f"Không tìm thấy Gist cũ. Đang tạo mới...")
                payload = {
                    "description": f"Pinyin AI Subtitle - {video_id}",
                    "public": True,
                    "files": {
                        target_filename: {
                            "content": content
                        }
                    }
                }
                req = urllib.request.Request("https://api.github.com/gists", method="POST")

            req.add_header("Authorization", f"token {token}")
            req.add_header("Content-Type", "application/json")
            
            data = json.dumps(payload).encode('utf-8')
            response = urllib.request.urlopen(req, data=data)
            res_data = json.loads(response.read().decode('utf-8'))
            
            raw_url = res_data['files'][target_filename]['raw_url']
            self.log(f"✅ Đã đồng bộ thành công lên GitHub Gist!\n🔗 Raw URL: {raw_url}")
            
        except urllib.error.HTTPError as e:
            if e.code == 401:
                self.log("❌ Lỗi: Token GitHub không hợp lệ hoặc đã hết hạn. Vui lòng nhập token mới.")
                config.pop('gh_pat_token', None)
                self.save_config(config)
                self.token_var.set('')
            else:
                self.log(f"❌ Lỗi HTTP khi đẩy lên GitHub: Mã {e.code}")
        except Exception as e:
            self.log(f"❌ Lỗi không xác định khi đẩy lên GitHub: {str(e)}")

    def format_time(self, seconds):
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        milliseconds = int(int((seconds - int(seconds)) * 1000))
        return f"{hours:02d}:{minutes:02d}:{secs:02d}.{milliseconds:03d}"

    def clean_filename(self, name):
        return re.sub(r'[\\/*?:"<>|]', "", name).strip()

    def align_texts_by_meaning(self, zh_text, translator):
        if not zh_text.strip(): return "", ""
        p_list = pinyin(zh_text, style=Style.TONE, heteronym=False)
        pinyin_text = " ".join([item[0] for item in p_list])
        try:
            vi_text = translator.translate(zh_text)
            return pinyin_text, vi_text if vi_text else ""
        except:
            return pinyin_text, ""

    def is_valid_vtt_structure(self, sub_file):
        try:
            with open(sub_file, 'r', encoding='utf-8', errors='ignore') as f: content = f.read()
            if not content.startswith("WEBVTT"): return False
            blocks = re.split(r'\n\s*\n', content)
            valid_block_count = 0
            for block in blocks:
                block = block.strip()
                if not block or 'WEBVTT' in block or 'Kind:' in block or 'Language:' in block: continue
                lines = block.split('\n')
                if lines[0].isdigit(): lines = lines[1:]
                if not any('-->' in line for line in lines): return False
                text_lines = [l for l in lines if '-->' not in l and not l.isdigit()]
                if len(text_lines) != 3: return False
                valid_block_count += 1
            return valid_block_count > 0
        except: return False

    def process_video_with_ai(self, url):
        ffmpeg_path = get_ffmpeg_path()
        save_directory = self.save_dir_var.get()
        video_id = extract_video_id(url)
        
        try:
            self.check_cancel()
            self.log("Kết nối YouTube để lấy thông tin video...")
            self.set_progress(5, "Đang lấy thông tin video...")
            
            ydl_info_opts = {'ffmpeg_location': ffmpeg_path}
            with yt_dlp.YoutubeDL(ydl_info_opts) as ydl:
                info_dict = ydl.extract_info(url, download=False)
                video_title = info_dict.get('title', 'video_sub')
                safe_title = self.clean_filename(video_title)
                self.current_output_file = os.path.join(save_directory, f"{safe_title}.vtt")

            self.check_cancel()
            self.log("Kiểm tra phụ đề sẵn có trên YouTube...")
            self.set_progress(10, "Đang kiểm tra phụ đề...")
            
            sub_file = None
            sub_lang_found = "zh"
            try:
                ydl_opts_sub = {
                    'skip_download': True, 'writesubtitles': True, 'writeautomaticsub': True,
                    'subtitleslangs': ['zh', 'zh-Hans', 'zh-Hant', 'zh-CN', 'zh-TW', 'en', 'ja', 'ko'],
                    'outtmpl': 'temp_sub', 'ffmpeg_location': ffmpeg_path,
                }
                with yt_dlp.YoutubeDL(ydl_opts_sub) as ydl:
                    info_sub = ydl.extract_info(url, download=True)
                    sub_lang_found = list(info_sub.get('requested_subtitles', {}).keys())[0] if info_sub.get('requested_subtitles') else "zh"
                    
                for file in os.listdir('.'):
                    if ('temp_sub' in file) and (file.endswith('.vtt') or file.endswith('.srt')):
                        sub_file = file
                        break
            except:
                self.log("Không lấy được phụ đề YouTube. Chuyển sang dùng AI Whisper.")

            if sub_file:
                if self.is_valid_vtt_structure(sub_file):
                    self.check_cancel()
                    self.log(f"Phụ đề hợp lệ được tìm thấy [{sub_lang_found.upper()}]. Đang xử lý chuyển đổi...")
                    self.convert_existing_sub(sub_file, self.current_output_file, sub_lang_found, video_id)
                    self.cleanup_files()
                    return
                else:
                    self.log("Phụ đề không đạt chuẩn 3 dòng. Chuyển sang chế độ AI Whisper.")
                    self.cleanup_files()

            self.check_cancel()
            self.log("Tải file âm thanh để xử lý bằng AI Whisper...")
            self.set_progress(15, "Chuẩn bị tải âm thanh...")
            
            ydl_opts_audio = {
                'format': 'bestaudio/best', 'outtmpl': 'temp_audio',
                'postprocessors': [{'key': 'FFmpegExtractAudio', 'preferredcodec': 'mp3', 'preferredquality': '192'}],
                'ffmpeg_location': ffmpeg_path, 'progress_hooks': [self.ytdl_progress_hook]
            }
            
            with yt_dlp.YoutubeDL(ydl_opts_audio) as ydl: ydl.download([url])

            audio_file = next((f for f in os.listdir('.') if f.startswith('temp_audio') and f.endswith('.mp3')), None)

            self.check_cancel()
            self.log("AI đang phân tích và nhận diện giọng nói (Whisper)...")
            self.set_progress(40, "AI đang nhận diện âm thanh...")
            
            model = whisper.load_model("base")
            result = model.transcribe(audio_file)
            detected_lang = result.get("language", "zh")
            self.log(f"Ngôn ngữ phát hiện: [{detected_lang.upper()}]")

            self.check_cancel()
            self.log("Đang dịch thuật và tạo Pinyin song song...")
            
            to_zh_translator = GoogleTranslator(source=detected_lang, target='zh-CN') if not detected_lang.startswith('zh') else None
            to_vi_translator = GoogleTranslator(source='zh-CN', target='vi')
            
            segments = result.get("segments", [])
            total_segments = len(segments)
            new_lines = ["WEBVTT\nKind: captions\nLanguage: zh-TW\n\n"]

            for i, segment in enumerate(segments, start=1):
                self.check_cancel()
                start_time = self.format_time(segment["start"])
                end_time = self.format_time(segment["end"])
                raw_text = segment["text"].strip()
                
                if not raw_text: continue

                zh_text = raw_text if detected_lang.startswith('zh') else (to_zh_translator.translate(raw_text) or raw_text)
                pinyin_text, vi_text = self.align_texts_by_meaning(zh_text, to_vi_translator)
                
                new_lines.append(f"{start_time} --> {end_time}\n{zh_text}\n{pinyin_text}\n{vi_text}\n\n")
                if total_segments > 0:
                    pct = 40 + (i / total_segments) * 55
                    self.set_progress(pct, f"Đang dịch & tạo Pinyin: {i}/{total_segments} câu ({int(pct)}%)")

            self.check_cancel()
            with open(self.current_output_file, 'w', encoding='utf-8') as f:
                f.writelines(new_lines)

            if self.upload_gh_var.get():
                self.upload_to_github_gist(self.current_output_file, video_id)

            self.cleanup_files()
            self.set_progress(100, "Hoàn thành!")
            self.log(f"\nThành công!\nFile phụ đề lưu tại: {self.current_output_file}")
            messagebox.showinfo("Thành công", f"Đã tạo xong phụ đề:\n{self.current_output_file}")

        except Exception as e:
            self.cleanup_files()
            if self.cancel_event.is_set():
                self.log("\nĐã hủy tác vụ thành công.")
                self.set_progress(0, "Đã hủy.")
            else:
                self.log(f"Lỗi: {str(e)}")
                self.set_progress(0, "Lỗi xảy ra!")
                messagebox.showerror("Lỗi", str(e))
            
        finally:
            self.root.after(0, self.reset_ui)

    def convert_existing_sub(self, sub_file, output_file, source_lang, video_id):
        with open(sub_file, 'r', encoding='utf-8', errors='ignore') as f: content = f.read()
            
        blocks = re.split(r'\n\s*\n', content)
        new_lines = ["WEBVTT\nKind: captions\nLanguage: zh-TW\n\n"]
        
        is_chinese = source_lang.startswith('zh')
        to_zh_translator = GoogleTranslator(source=source_lang, target='zh-CN') if not is_chinese else None
        to_vi_translator = GoogleTranslator(source='zh-CN', target='vi')
        
        total_valid = len(blocks)
        processed_count = 0

        for block in blocks:
            self.check_cancel()
            if not block.strip() or 'WEBVTT' in block or 'Kind:' in block or 'Language:' in block: continue
                
            block_lines = block.strip().split('\n')
            time_line, text_lines = "", []
            
            for line in block_lines:
                if '-->' in line: time_line = line.replace(',', '.') 
                elif not line.isdigit() and line.strip(): text_lines.append(line.strip())
            
            if time_line and text_lines:
                clean_text = re.sub(r'<[^>]+>', '', " ".join(text_lines))
                if clean_text:
                    zh_text = clean_text if is_chinese else (to_zh_translator.translate(clean_text) or clean_text)
                    pinyin_text, vi_text = self.align_texts_by_meaning(zh_text, to_vi_translator)
                    new_lines.append(f"{time_line}\n{zh_text}\n{pinyin_text}\n{vi_text}\n\n")

            processed_count += 1
            if total_valid > 0:
                pct = (processed_count / total_valid) * 100
                self.set_progress(pct, f"Đang đồng bộ chuẩn 3 dòng: {processed_count}/{total_valid} câu ({int(pct)}%)")

        self.check_cancel()
        with open(output_file, 'w', encoding='utf-8', errors='ignore') as f:
            f.writelines(new_lines)

        if self.upload_gh_var.get():
            self.upload_to_github_gist(output_file, video_id)

        self.set_progress(100, "Hoàn thành!")
        self.log(f"\nThành công!\nFile phụ đề lưu tại: {output_file}")
        messagebox.showinfo("Thành công", f"Đã tạo xong file phụ đề:\n{output_file}")

    def reset_ui(self):
        self.is_processing = False
        self.btn_action.config(text="TẠO PHỤ ĐỀ AI", style='Primary.TButton', state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = SubtitleApp(root)
    root.mainloop()
