"""回归测试：流式 SSE 中文内容含 0x85 字节时不能被切行（llm_parser._stream）。"""
import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

from tabularedge import llm_parser


def test_stream_utf8_nel_byte():
    text = '{"summary":"删除括号内容，并用公式按部门汇总全部金额","lang":"zh","steps":[]}'
    assert b"\x85" in text.encode("utf-8")  # “…” = e2 80 a6，“删”等字含 0x85

    class H(BaseHTTPRequestHandler):
        def do_POST(self):
            self.rfile.read(int(self.headers["Content-Length"]))
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream")  # 故意不带 charset
            self.end_headers()
            for ch in (text[:20], text[20:]):
                self.wfile.write(b"data: " + json.dumps({"choices": [{"delta": {"content": ch}}]}, ensure_ascii=False).encode() + b"\n\n")
            self.wfile.write(b"data: [DONE]\n\n")

        def log_message(self, *a):
            pass
    srv = HTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    got = []
    buf, info = llm_parser._stream([], on_summary=got.append, url=f"http://127.0.0.1:{srv.server_port}")
    srv.shutdown()
    assert buf == text
    assert got == ["删除括号内容，并用公式按部门汇总全部金额"]
