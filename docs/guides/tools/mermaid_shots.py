"""Renders the mermaid diagrams of docs/TECHNICAL.md to PNG (Edge, headless, mermaid from jsDelivr)."""
import os
import re
import sys
from playwright.sync_api import sync_playwright

MD, OUT = sys.argv[1], sys.argv[2]
os.makedirs(OUT, exist_ok=True)
blocks = re.findall(r"```mermaid\n(.*?)```", open(MD, encoding="utf-8").read(), re.S)
PAGE = """<!doctype html><html><head><meta charset="utf-8">
<script src="https://cdn.jsdelivr.net/npm/mermaid@10.9.1/dist/mermaid.min.js"></script>
<style>body{margin:0;padding:16px;background:#fff;font-family:Calibri,Segoe UI,sans-serif}</style></head>
<body><pre class="mermaid" id="m"></pre></body></html>"""
with sync_playwright() as p:
    browser = p.chromium.launch(channel="msedge", headless=True)
    page = browser.new_page(viewport={"width": 1400, "height": 900}, device_scale_factor=2)
    for n, code in enumerate(blocks, 1):
        code = code.replace("flowchart LR", "flowchart TB")   # taller, so it reads at page width
        page.set_content(PAGE)
        page.wait_for_function("window.mermaid !== undefined", timeout=30000)
        page.evaluate("""async ([code, n]) => {
            mermaid.initialize({ startOnLoad: false, theme: 'neutral', fontFamily: 'Calibri, Segoe UI, sans-serif',
              themeVariables: { primaryColor: '#FDF1E9', primaryBorderColor: '#C2571A', lineColor: '#1F3A5F', fontSize: '15px' } });
            const out = await mermaid.render('d' + n, code);
            document.getElementById('m').innerHTML = out.svg;
        }""", [code, n])
        page.wait_for_timeout(800)
        page.locator("#m svg").screenshot(path=os.path.join(OUT, "diagram%d.png" % n))
        print("saved diagram%d" % n)
    browser.close()
