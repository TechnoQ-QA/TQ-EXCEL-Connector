"""Screenshots of the TQ Connector panel and setup page for the guides (Edge, headless).
Real pages from the installed connector; the sandbox account, the user's name and the company are replaced
with neutral example values before each picture, because the guides go to customers."""
import os
import sys
from playwright.sync_api import sync_playwright

OUT = sys.argv[1]
WHICH = sys.argv[2] if len(sys.argv) > 2 else "install"
BASE = "https://localhost:3000"
os.makedirs(OUT, exist_ok=True)

SANITIZE = r"""
() => {
  const swap = (t) => t
    .replace(/65987[_-]SB1/gi, '1234567')
    .replace(/65987/g, '1234567')
    .replace(/Zubair Salamat/g, 'Finance User')
    .replace(/Techno Q \(sandbox\)/g, 'Your Company')
    .replace(/TQ-1GF4774/g, 'FIN-PC-01')
    .replace(/C:\\Users\\[^\\]+\\/g, 'C:\\Users\\finance.user\\');
  const walk = (n) => {
    if (n.nodeType === 3) { n.nodeValue = swap(n.nodeValue); return; }
    n.childNodes.forEach(walk);
  };
  walk(document.body);
  document.querySelectorAll('input, textarea').forEach((i) => {
    if (i.type === 'password' || i.type === 'file' || i.type === 'checkbox' || i.type === 'radio') return;
    let v = swap(i.value || '');
    if (i.id === 'clientId' && v) v = '3f9c1a7e5b2d48c6a0e4f1b2c3d4e5f6a7b8c9d0e1f2a3b4c5d6e7f8a9b0c1d2';
    if (i.id === 'restletUrl' && v) v = 'https://1234567.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=101&deploy=1';
    i.value = v;
  });
  document.querySelectorAll('[title]').forEach((e) => { e.title = swap(e.title); });
  if (document.activeElement && document.activeElement.blur) document.activeElement.blur();   // no keyboard focus ring in pictures
  const logo = document.getElementById('brandLogo');
  if (logo) { logo.remove(); const m = document.getElementById('brandMark'); if (m) m.hidden = false; }
  if (!window.__tqWatch) {   // text filled in later (status, welcome checks) is swapped too
    window.__tqWatch = new MutationObserver((list) => list.forEach((r) => r.addedNodes.forEach(walk)));
    window.__tqWatch.observe(document.body, { childList: true, subtree: true, characterData: true });
  }
}
"""


def shot(page, name, full=False, clip=None):
    page.evaluate(SANITIZE)
    page.wait_for_timeout(1200)
    page.evaluate(SANITIZE)
    page.screenshot(path=os.path.join(OUT, name + ".png"), full_page=full, clip=clip)
    print("saved", name)


with sync_playwright() as p:
    browser = p.chromium.launch(channel="msedge", headless=True)
    ctx = browser.new_context(viewport={"width": 360, "height": 780}, device_scale_factor=2, ignore_https_errors=True,
                              color_scheme="light")
    page = ctx.new_page()

    def panel(screen=None, welcomed=True, wait=6000):
        page.goto(BASE + "/taskpane.html")
        page.evaluate("(w) => { try { if (w) localStorage.setItem('tq.welcomed','1'); else localStorage.removeItem('tq.welcomed'); } catch (e) {} }", welcomed)
        page.reload()
        page.wait_for_timeout(wait)
        if screen:
            page.evaluate("(s) => TQNav.go(s)", screen)
            page.wait_for_timeout(900)

    def setup(section, wait=2000):
        if "/setup.html" not in page.url:
            page.goto(BASE + "/setup.html")
            page.wait_for_timeout(2000)
        page.evaluate("(s) => document.querySelector('#snav button[data-section=\"' + s + '\"]').click()", section)
        page.wait_for_timeout(wait)

    if WHICH == "install":
        panel(welcomed=False)
        shot(page, "panel-welcome")
        panel()
        shot(page, "panel-home")
        setup("overview")
        shot(page, "setup-overview", full=True)
        setup("connection")
        page.evaluate("() => { const d = document.getElementById('advanced'); if (d) d.open = true; const b = document.getElementById('savebar'); if (b) { b.style.position = 'static'; } }")
        shot(page, "setup-connection", full=True)
        setup("licence")
        shot(page, "setup-licence")
        setup("seats", wait=3500)
        shot(page, "setup-seats")
        setup("updates")
        shot(page, "setup-updates")
        panel("transfer")
        shot(page, "panel-transfer", full=True)
        panel("keys")
        shot(page, "panel-keys")
    if WHICH == "user":
        def tall(h):
            page.set_viewport_size({"width": 360, "height": h})

        def js(code, wait=700):
            page.evaluate(code)
            page.wait_for_timeout(wait)

        # Home: search and help
        panel()
        js("() => { const i = document.getElementById('qsInput'); i.value = 'aging'; i.dispatchEvent(new Event('input')); }")
        shot(page, "u-search")
        js("() => { const i = document.getElementById('qsInput'); i.value = ''; i.dispatchEvent(new Event('input')); document.getElementById('helpButton').click(); }")
        shot(page, "u-help")
        js("() => document.getElementById('helpButton').click()")
        js("() => document.getElementById('menuButton').click()")
        shot(page, "u-menu")
        js("() => document.getElementById('menuButton').click()")

        # New report wizard
        tall(1500)
        js("() => TQHome.openWizard()", 1500)
        shot(page, "u-wizard-kinds")
        tall(1000)
        js("() => TQHome.openKind('monthytd')", 1500)
        shot(page, "u-wizard-month")
        js("() => [...document.querySelectorAll('#wizard button')].find(b => b.textContent === 'Next').click()", 2500)
        page.wait_for_timeout(2500)
        js("""() => { const boxes = [...document.querySelectorAll('#wizard .mpick')];
          const dep = boxes.find((m) => /Department/.test(m.parentNode.parentNode.textContent)) || boxes[0];
          if (!dep) return; dep.querySelector('.mpick-head').click();
          const b = dep.querySelectorAll('.pick input'); if (b[1]) b[1].click(); if (b[3]) b[3].click(); }""", 600)
        shot(page, "u-wizard-filters")
        js("() => { const n = () => [...document.querySelectorAll('#wizard button')].find(b => b.textContent === 'Next'); n().click(); }", 600)
        shot(page, "u-wizard-sheets")
        js("() => [...document.querySelectorAll('#wizard button')].find(b => b.textContent === 'Next').click()", 600)
        shot(page, "u-wizard-check")

        # This report (staged: outside Excel the screen has no sheet to read)
        tall(900)
        js("""() => {
          TQNav.go('report');
          document.getElementById('rpTitle').textContent = 'Monthly P&L';
          document.getElementById('rpAbout').textContent = 'Sheet "Monthly P&L", from layout "TQ Layout Monthly P&L". Changes there rebuild it in place.';
          const st = document.getElementById('rpStamp'); st.hidden = false; st.className = 'stamp'; st.textContent = 'Refreshed today at 10:42';
          document.getElementById('rpMonthBox').hidden = false;
          document.getElementById('hMonth').innerHTML = '<option>Oct 2026</option>';
          const box = document.getElementById('fCurrent'); box.textContent = '';
          [['Subsidiary', 'Hotel Riyadh'], ['Department', '22 Food & Beverage*']].forEach(([f, v]) => {
            const it = document.createElement('div'); it.className = 'item';
            const s = document.createElement('span'); s.textContent = f + ': '; const sm = document.createElement('small'); sm.textContent = v; s.appendChild(sm);
            const b = document.createElement('button'); b.textContent = '\u2715'; it.appendChild(s); it.appendChild(b); box.appendChild(it); });
          document.getElementById('rpRestoreRow').hidden = false;
          document.getElementById('rpLayoutRow').hidden = false;
        }""")
        shot(page, "u-report")
        tall(1100)
        js("() => TQNav.go('filter')", 2500)
        shot(page, "u-filter")

        # Layout screen with a real layout from the wizard template (built by the connector, not written anywhere)
        tall(1500)
        js("() => { TQNav.go('layout'); document.getElementById('dTemplate').value = 'monthytd'; document.getElementById('dNewName').value = 'Monthly P&L'; document.getElementById('dNew').click(); }", 4000)
        js("() => { document.getElementById('dMessage').textContent = ''; const f = document.getElementById('libFold'); f.open = true; }", 2500)
        shot(page, "u-library")

        tall(900)
        for screen in ("dash", "send", "upload", "formulas", "data", "keys", "chat"):
            js("() => TQNav.go('%s')" % screen, 2500)
            shot(page, "u-" + screen)

        # Drill down with real figures, then the large window
        js("() => TQPanel.breakdown({ account: 'type:Income', fromPeriod: 'Jan 2025', toPeriod: 'Dec 2025' })", 6000)
        shot(page, "u-drill")
        page.set_viewport_size({"width": 1100, "height": 300})
        page.goto(BASE + "/drill.html#" + __import__('urllib.parse').parse.quote('{"kind":"breakdown","request":{"account":"type:Income","fromPeriod":"Jan 2025","toPeriod":"Dec 2025"}}'))
        page.wait_for_timeout(5000)
        shot(page, "u-drill-large")

        # Setup page: logo
        page.set_viewport_size({"width": 360, "height": 780})
        setup("logo")
        shot(page, "u-setup-logo")
    browser.close()
