#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
xlsx → 长图 一条龙：渲染 HTML → Chrome headless 截图 → 按内容底边裁切。
用法: python xlsx_to_image.py <in.xlsx> <out.png>
说明:
  - 截图必须先落 C:\\Windows\\Temp（headless 直写工作区会被拒 0x5），裁好再拷到目标路径。
  - 内容底边按页面实际背景色 ±4 容差自底向上扫描，避免裁掉页脚或留大段空白。
"""
import sys, io, os, subprocess, glob
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from render_report_image import render

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

TMP = r"C:\Windows\Temp"
CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]

def find_chrome():
    for c in CHROME_CANDIDATES:
        if os.path.exists(c):
            return c
    raise RuntimeError("未找到 chrome.exe")

def main():
    in_xlsx, out_png = sys.argv[1], sys.argv[2]
    html_path = os.path.join(TMP, "homes_report.html")
    raw_png = os.path.join(TMP, "homes_report_raw.png")

    # 1) 渲染 HTML
    render(in_xlsx, html_path)

    # 2) Chrome headless 截图（2x 高清；窗口宽度贴合 1120 容器 + 边距，高度给足）
    chrome = find_chrome()
    file_url = "file:///" + html_path.replace("\\", "/")
    cmd = [chrome, "--headless", "--disable-gpu", "--no-sandbox", "--hide-scrollbars",
           "--force-device-scale-factor=2", "--virtual-time-budget=5000",
           "--window-size=1160,3600", f"--screenshot={raw_png}", file_url]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if not os.path.exists(raw_png) or os.path.getsize(raw_png) < 10000:
        raise RuntimeError("截图失败: " + (r.stderr or "")[-500:])

    # 3) 检测内容底边并裁切
    from PIL import Image
    im = Image.open(raw_png)
    w, h = im.size
    px = im.convert("RGB").load()
    bg = px[5, 5]
    def is_bg(y):
        return all(abs(px[x, y][c] - bg[c]) <= 4 for c in range(3) for x in range(0, w, 40))
    bottom = next((y for y in range(h - 1, 0, -1) if not is_bg(y)), h)
    os.makedirs(os.path.dirname(os.path.abspath(out_png)), exist_ok=True)
    im.crop((0, 0, w, min(bottom + 30, h))).save(out_png)
    print(f"IMAGE_OK {out_png} size={w}x{min(bottom+30,h)} bytes={os.path.getsize(out_png)}")

if __name__ == "__main__":
    main()
