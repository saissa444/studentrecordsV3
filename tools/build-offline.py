"""產生單機版：把所有程式與 Excel 元件包進一個 HTML 檔，雙擊即可離線使用。
用法：python3 tools/build-offline.py  → 產生 dist/學生加扣分獎懲_單機版.html"""
import base64, os, re
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
read = lambda p: open(os.path.join(ROOT, p), encoding="utf-8").read()
html = read("index.html")

def inline(match):
    src = match.group(1).split("?")[0].lstrip("./")
    code = read(src)
    assert "</script" not in code.lower(), src
    return "<script>\n" + code + "\n</script>"

html = re.sub(r'<script defer src="([^"]+)"></script>', inline, html)
# 單機版不需要 PWA 相關設定
html = re.sub(r'<script>if \("serviceWorker".*?</script>\n?', "", html)
html = re.sub(r'<link rel="manifest"[^>]*>\n?', "", html)
icon = base64.b64encode(open(os.path.join(ROOT, "icons/icon-180.png"), "rb").read()).decode()
html = re.sub(r'<link rel="apple-touch-icon"[^>]*>', '<link rel="icon" href="data:image/png;base64,' + icon + '">', html)
html = html.replace("<title>學生加扣分・獎懲登錄</title>", "<title>學生加扣分・獎懲登錄（單機版）</title>")
html = html.replace('版本 2026.10.18', '單機版｜版本 2026.10.18')
assert "src=\"./" not in html and "href=\"./" not in html, "仍有外部檔案參照"
os.makedirs(os.path.join(ROOT, "dist"), exist_ok=True)
out = os.path.join(ROOT, "dist", "學生加扣分獎懲_單機版.html")
open(out, "w", encoding="utf-8").write(html)
print(out, round(len(html.encode()) / 1024), "KB")
