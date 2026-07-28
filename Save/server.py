#!/usr/bin/env python3
"""Server tải audio xong hoàn toàn rồi mới gửi"""
import os, sys, tempfile, time, subprocess
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

PORT = 8888

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        p = urlparse(self.path)
        q = parse_qs(p.query)
        url = q.get('url', [None])[0]
        
        if p.path == '/health':
            self.send_response(200)
            self.send_header('Content-Type', 'text/plain')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(b'OK')
            return
        
        if p.path == '/download' and url:
            tmp_file = None
            try:
                print(f"\n=== Downloading: {url[:60]}...")
                tmp_file = os.path.join(tempfile.gettempdir(), f'audio_{int(time.time())}.m4a')
                
                # Tải audio - đợi hoàn toàn xong
                result = subprocess.run([
                    sys.executable, '-m', 'yt_dlp',
                    '-f', 'worstaudio[ext=m4a]/worstaudio',
                    '-o', tmp_file,
                    '--quiet', '--no-warnings',
                    '--extractor-args', 'youtube:js_runtimes=deno',
                    '--no-playlist',
                    url
                ], capture_output=True, text=True, timeout=300)
                
                if result.returncode != 0:
                    raise Exception(f"yt-dlp error: {result.stderr[:200]}")
                
                if not os.path.exists(tmp_file) or os.path.getsize(tmp_file) == 0:
                    raise Exception("File not created or empty")
                
                size = os.path.getsize(tmp_file)
                print(f"Downloaded: {size/1024:.0f}KB - Sending...")
                
                # Đọc toàn bộ file vào bộ nhớ rồi gửi 1 lần
                with open(tmp_file, 'rb') as f:
                    file_data = f.read()
                
                self.send_response(200)
                self.send_header('Content-Type', 'audio/mp4')
                self.send_header('Content-Length', str(len(file_data)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Connection', 'close')
                self.end_headers()
                self.wfile.write(file_data)
                
                print(f"Sent: {len(file_data)/1024:.0f}KB")
                
            except Exception as e:
                print(f"Error: {e}")
                self.send_response(500)
                self.send_header('Content-Type', 'text/plain')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(f"Error: {str(e)}".encode())
            finally:
                if tmp_file and os.path.exists(tmp_file):
                    try: os.remove(tmp_file)
                    except: pass
            return
        
        self.send_response(404)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
    
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', '*')
        self.end_headers()

if __name__ == '__main__':
    print(f"Server: http://localhost:{PORT}")
    HTTPServer(('0.0.0.0', PORT), Handler).serve_forever()
