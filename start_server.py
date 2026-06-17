#!/usr/bin/env python3
"""
Lokalny serwer dla Szukajki Ofert
Uruchom: python start_server.py
Potem otwórz: http://localhost:8080/price-finder.html
"""

import http.server
import socketserver
import webbrowser
import os
import threading

PORT = 8080
FILE = "price-finder.html"

class CORSHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, x-api-key, anthropic-version, anthropic-dangerous-direct-browser-access")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def log_message(self, format, *args):
        pass  # wycisz logi

def open_browser():
    import time
    time.sleep(0.8)
    webbrowser.open(f"http://localhost:{PORT}/{FILE}")

os.chdir(os.path.dirname(os.path.abspath(__file__)))

print("=" * 50)
print("  💰 Szukajka Ofert — Lokalny Serwer")
print("=" * 50)
print(f"\n  ✅ Serwer działa na: http://localhost:{PORT}")
print(f"  🌐 Otwórz:          http://localhost:{PORT}/{FILE}")
print(f"\n  Aby zatrzymać: naciśnij Ctrl+C\n")

threading.Thread(target=open_browser, daemon=True).start()

with socketserver.TCPServer(("", PORT), CORSHandler) as httpd:
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n  🛑 Serwer zatrzymany.")
