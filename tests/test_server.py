"""خادم اختبار محلي: يقدّم ملفات المشروع ويستقبل ملفات Excel المولدة من صفحة الاختبار (POST /save?name=...)."""
import http.server, os, sys, urllib.parse

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'tests', 'output')

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=ROOT, **kw)

    def end_headers(self):
        # منع التخزين المؤقت حتى تُختبر أحدث نسخة من الملفات دائمًا
        self.send_header('Cache-Control', 'no-store')
        super().end_headers()

    def do_POST(self):
        q = urllib.parse.urlparse(self.path)
        if q.path != '/save':
            self.send_error(404); return
        name = urllib.parse.parse_qs(q.query).get('name', ['out.bin'])[0]
        name = os.path.basename(name)
        data = self.rfile.read(int(self.headers.get('Content-Length', 0)))
        os.makedirs(OUT, exist_ok=True)
        with open(os.path.join(OUT, name), 'wb') as f:
            f.write(data)
        self.send_response(200); self.end_headers(); self.wfile.write(b'ok')

if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    http.server.ThreadingHTTPServer(('127.0.0.1', port), Handler).serve_forever()
