#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
本地接收服务器：等待浏览器端 JS 推送 Homes 数据(JSON)，落盘后调用 fill_template.py 生成 Excel。
用法: python homes_server.py --date YYYYMMDD [--port 8765] [--timeout 240]
      [--template 模板路径] [--outdir 输出目录]
输出: 打印一行 JSON 结果 {status, out, checks...}，进程退出码 0=成功。
"""
import http.server, json, argparse, os, sys, io, subprocess, time, threading

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
BASE = os.path.dirname(os.path.abspath(__file__))
PYTHON = sys.executable
DEFAULT_TEMPLATE = r"C:\Users\泇辰461427\Desktop\盒马Homes自动生成MTD营收对比表.xlsx"

PAYLOAD = {'data': None}

class Handler(http.server.BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def do_POST(self):
        ln = int(self.headers.get('Content-Length', 0))
        PAYLOAD['data'] = self.rfile.read(ln)
        self.send_response(200)
        self._cors()
        self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self.end_headers()
        self.wfile.write('RECEIVED'.encode())

    def log_message(self, *a):
        pass

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--date', required=True)
    ap.add_argument('--port', type=int, default=8765)
    ap.add_argument('--timeout', type=int, default=240)
    ap.add_argument('--template', default=DEFAULT_TEMPLATE)
    ap.add_argument('--outdir', default=os.path.join(BASE, 'out'))
    args = ap.parse_args()

    srv = http.server.HTTPServer(('127.0.0.1', args.port), Handler)
    th = threading.Thread(target=srv.serve_forever, daemon=True)
    th.start()

    deadline = time.time() + args.timeout
    while PAYLOAD['data'] is None and time.time() < deadline:
        time.sleep(0.5)
    srv.shutdown()

    if PAYLOAD['data'] is None:
        print(json.dumps({'status': 'timeout', 'error': '等待浏览器推送数据超时'}, ensure_ascii=False))
        sys.exit(2)

    try:
        raw = json.loads(PAYLOAD['data'].decode('utf-8'))
    except Exception as e:
        print(json.dumps({'status': 'bad_payload', 'error': str(e)}, ensure_ascii=False))
        sys.exit(3)
    if str(raw.get('date')) != args.date:
        print(json.dumps({'status': 'date_mismatch', 'expect': args.date, 'got': raw.get('date')}, ensure_ascii=False))
        sys.exit(4)

    period = str(raw.get('period', 'mtd'))
    label = '单日' if period == '1d' else 'MTD'

    raw_dir = os.path.join(BASE, 'raw')
    os.makedirs(raw_dir, exist_ok=True)
    raw_path = os.path.join(raw_dir, 'homes_%s_%s.json' % (period, args.date))
    with open(raw_path, 'w', encoding='utf-8') as f:
        f.write(json.dumps(raw, ensure_ascii=False))

    os.makedirs(args.outdir, exist_ok=True)
    out_path = os.path.join(args.outdir, '盒马Homes自动生成%s营收对比表_%s.xlsx' % (label, args.date))

    r = subprocess.run([PYTHON, os.path.join(BASE, 'fill_template.py'), raw_path, args.template, out_path, period],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        print(json.dumps({'status': 'fill_error', 'error': (r.stderr or r.stdout)[-2000:]}, ensure_ascii=False))
        sys.exit(5)
    print(r.stdout.strip())

if __name__ == '__main__':
    main()
