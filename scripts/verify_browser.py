#!/usr/bin/env python3
"""
scripts/verify_browser.py
Browser regression suite for dist/ (Playwright, 375x812 phone viewport).

Expectations are derived from dist/itinerary_data.json, so the suite keeps
working when itinerary.xlsx changes. Read-only checks run concurrently; the
check that edits itinerary_data.json gets its own server and copy of dist/.

Setup (once):   python3 -m venv .venv && .venv/bin/pip install playwright
Run:            .venv/bin/python scripts/verify_browser.py [--only name,name] [--shots DIR]
Uses the installed Microsoft Edge / Chrome; falls back to Playwright's Chromium.
"""
import argparse
import asyncio
import json
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path

from playwright.async_api import async_playwright

DIST = Path(__file__).resolve().parent.parent / "dist"
VIEWPORT = {"width": 375, "height": 812}
TAIWAN_LEG = re.compile(r"桃園|^入境$|返家")  # mirrors TAIWAN_LEG in index.html


def to_min(hhmm):
    m = re.match(r"^(\d{1,2}):(\d{2})", hhmm or "")
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def fmt(minutes):
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def timed_items(day):
    return [it for it in day["items"] if to_min(it["time_start"]) is not None and not TAIWAN_LEG.search(it["spot"])]


@contextmanager
def serve(root):
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    proc = subprocess.Popen([sys.executable, "-m", "http.server", str(port), "--bind", "127.0.0.1", "--directory", str(root)],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        for _ in range(50):
            try:
                socket.create_connection(("127.0.0.1", port), timeout=0.1).close()
                break
            except OSError:
                time.sleep(0.1)
        yield f"http://127.0.0.1:{port}"
    finally:
        proc.terminate()
        proc.wait()


async def launch(p):
    for kwargs in ({"channel": "msedge"}, {"channel": "chrome"}, {}):
        try:
            return await p.chromium.launch(headless=True, **kwargs)
        except Exception:
            continue
    raise SystemExit("No browser found: install Edge/Chrome or run `.venv/bin/playwright install chromium`")


async def sw_ready(page):
    await page.evaluate("navigator.serviceWorker.ready")
    await page.wait_for_function("navigator.serviceWorker.controller !== null", timeout=10000)


class Suite:
    def __init__(self, browser, url, data, shots):
        self.browser, self.url, self.data, self.shots = browser, url, data, shots

    async def page(self, path="index.html", **ctx_opts):
        ctx = await self.browser.new_context(viewport=VIEWPORT, **ctx_opts)
        page = await ctx.new_page()
        if path is not None:
            await page.goto(f"{self.url}/{path}")
            await page.wait_for_selector(".tl-item, .empty-state")
        return ctx, page

    async def shot(self, page, name):
        if self.shots:
            await page.screenshot(path=str(Path(self.shots) / f"{name}.png"), full_page=True)

    # ---- checks: each returns (ok, detail) ----

    async def now_focus(self):
        """A stop in progress is marked is-now and scrolled into view on its trip day."""
        best = None
        for day in self.data["schedule"]:
            timed = timed_items(day)
            for i, it in enumerate(timed):
                start = to_min(it["time_start"])
                nxt = to_min(timed[i + 1]["time_start"]) if i + 1 < len(timed) else None
                until = to_min(it["time_end"]) or nxt or start + 60
                if until - start >= 15 and (nxt is None or nxt > start + 5):
                    pos = day["items"].index(it)
                    if best is None or pos > best[2]:
                        best = (day, it, pos)
        day, it, _ = best
        now = f"{day['date']}T{fmt(to_min(it['time_start']) + 5)}"
        ctx, page = await self.page(f"index.html?now={now}")
        await page.wait_for_timeout(1500)  # smooth scroll
        state = await page.evaluate("""id => {
            const el = document.getElementById(id);
            const r = el && el.getBoundingClientRect();
            return { isNow: !!el && el.classList.contains('is-now'),
                     inView: !!r && r.bottom > 0 && r.top < innerHeight,
                     active: document.querySelector('.day-chip.active')?.dataset.dayIndex };
        }""", it["id"])
        await self.shot(page, "now_focus")
        await ctx.close()
        ok = state["isNow"] and state["inView"] and state["active"] == str(self.data["schedule"].index(day))
        return ok, f"?now={now} → {it['spot']} {state}"

    async def tomorrow_reminder(self):
        """After 23:00 the now-panel names tomorrow's first stop; not during the day."""
        days = self.data["schedule"]
        first = timed_items(days[1])[0]
        ctx, page = await self.page(f"index.html?now={days[0]['date']}T23:10")
        night = await page.locator("#nowPanel").inner_text()
        await ctx.close()
        ctx, page = await self.page(f"index.html?now={days[1]['date']}T10:00")
        day = await page.locator("#nowPanel").inner_text()
        await ctx.close()
        ok = "⏰ 明天" in night and f"第一站 {first['time_start']}" in night and "⏰ 明天" not in day
        return ok, f"23:10 → {night.splitlines()[-1]!r}; 10:00 has reminder: {'⏰ 明天' in day}"

    async def offline_bar(self):
        """Offline bar toggles with the connection and never covers the sticky day bar."""
        ctx, page = await self.page()
        hidden_online = await page.locator("#offlineBar").is_hidden()
        await ctx.set_offline(True)
        await page.wait_for_timeout(300)
        await page.evaluate("scrollTo(0, 600)")
        await page.wait_for_timeout(300)
        r = await page.evaluate("""() => ({ bar: document.getElementById('offlineBar').getBoundingClientRect().bottom,
                                            day: document.querySelector('.day-bar').getBoundingClientRect().top })""")
        shown = await page.locator("#offlineBar").is_visible()
        await self.shot(page, "offline_bar")
        await ctx.set_offline(False)
        await page.wait_for_timeout(300)
        hidden_again = await page.locator("#offlineBar").is_hidden()
        await ctx.close()
        ok = hidden_online and shown and hidden_again and r["day"] >= r["bar"] - 1
        return ok, f"hidden/shown/hidden={hidden_online}/{shown}/{hidden_again}, day-bar.top={r['day']:.0f} offline.bottom={r['bar']:.0f}"

    async def offline_reload(self):
        """After one online visit, the page reloads offline from the service worker cache."""
        ctx, page = await self.page()
        await sw_ready(page)
        await page.wait_for_timeout(500)
        await ctx.set_offline(True)
        await page.reload()
        await page.wait_for_selector(".tl-item")
        count = await page.locator(".tl-item").count()
        await ctx.close()
        return count > 0, f"{count} timeline cards offline"

    async def single_shell_cache(self):
        """Every visit (/, ?now=…) shares one cached app shell; no query-string entries."""
        ctx, page = await self.page("")
        await sw_ready(page)
        await page.goto(f"{self.url}/?now={self.data['schedule'][1]['date']}T10:00")
        await page.wait_for_selector(".tl-item")
        await page.wait_for_timeout(800)
        urls = await page.evaluate("""async () => {
            const out = [];
            for (const n of await caches.keys()) for (const r of await (await caches.open(n)).keys()) out.push(r.url);
            return out; }""")
        await ctx.close()
        pages = [u for u in urls if u.endswith("/") or u.endswith(".html") or "?" in u]
        ok = pages == [f"{self.url}/index.html"]
        return ok, f"page entries: {pages}"

    async def manifest(self):
        """Manifest parses and every icon it lists is served."""
        ctx, page = await self.page(None)
        res = await page.request.get(f"{self.url}/manifest.webmanifest")
        icons = {i["src"]: (await page.request.get(f"{self.url}/{i['src']}")).status for i in (await res.json())["icons"]}
        await ctx.close()
        return res.status == 200 and all(s == 200 for s in icons.values()), f"manifest {res.status}, icons {icons}"

    async def theme_toggle(self):
        """Light by default; toggle flips theme + status-bar colour; dark survives reload before first paint."""
        ctx, page = await self.page()
        await page.add_init_script("""document.addEventListener('DOMContentLoaded',
            () => { window.__themeAtDCL = document.documentElement.dataset.theme; }, { once: true });""")
        meta = "document.querySelector('meta[name=theme-color]').content"
        light = await page.evaluate(f"[document.documentElement.dataset.theme, {meta}, document.getElementById('themeToggle').getAttribute('aria-pressed')]")
        await page.click("#themeToggle")
        dark = await page.evaluate(f"[document.documentElement.dataset.theme, {meta}, document.getElementById('themeToggle').getAttribute('aria-pressed')]")
        await page.reload()
        await page.wait_for_selector(".tl-item")
        at_dcl = await page.evaluate("window.__themeAtDCL")
        await page.click("#themeToggle")
        back = await page.evaluate(f"[document.documentElement.dataset.theme, {meta}]")
        await ctx.close()
        ok = (light[0] != "dark" and light[2] == "false" and dark == ["dark", "#0f172a", "true"]
              and at_dcl == "dark" and back[0] == "light" and back[1] == light[1])
        return ok, f"light={light} dark={dark} themeAtDOMContentLoaded={at_dcl} back={back}"

    async def dark_contrast(self):
        """Dark mode never makes text harder to read than light mode, below WCAG AA (4.5, or 3 for large text)."""
        audit = """() => {
            const lum = c => { const [r, g, b] = c.map(v => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; });
                               return 0.2126 * r + 0.7152 * g + 0.0722 * b; };
            const rgb = s => (s.match(/[\\d.]+/g) || []).map(Number);
            const bgOf = el => { for (; el; el = el.parentElement) { const cs = getComputedStyle(el);
                const m = cs.backgroundImage.match(/rgba?\\([^)]*\\)|#[0-9a-f]{3,8}/i);
                if (m && m[0].startsWith('rgb')) return rgb(m[0]);
                const c = rgb(cs.backgroundColor); if (c.length === 3 || c[3] > 0.5) return c; }
                return rgb(getComputedStyle(document.body).backgroundColor); };
            const out = {};
            document.querySelectorAll('#app *').forEach(el => {
                const text = [...el.childNodes].filter(n => n.nodeType === 3).map(n => n.textContent.trim()).join('');
                if (!text || !el.offsetParent || /^[\\p{Extended_Pictographic}\\s️]+$/u.test(text)) return;
                const cs = getComputedStyle(el);
                let alpha = 1;  // faded ancestors (e.g. past stops) blend text into the background
                for (let a = el; a; a = a.parentElement) alpha *= Number(getComputedStyle(a).opacity);
                if (alpha === 0) return;
                const bg = bgOf(el), fg = rgb(cs.color).slice(0, 3).map((c, i) => c * alpha + bg[i] * (1 - alpha));
                const [l1, l2] = [lum(fg), lum(bg)].sort((a, b) => b - a);
                const size = parseFloat(cs.fontSize), bold = Number(cs.fontWeight) >= 700;
                const need = size >= 24 || (size >= 18.66 && bold) ? 3 : 4.5;
                out[el.tagName + '.' + el.className + '|' + text.slice(0, 20)] = { ratio: +((l1 + 0.05) / (l2 + 0.05)).toFixed(2), need };
            });
            return out; }"""
        days = self.data["schedule"]
        results = {}
        for theme in ("light", "dark"):
            ctx, page = await self.page(None)
            await page.add_init_script(f"localStorage.setItem('japan2026.theme', '{theme}')")
            await page.goto(f"{self.url}/index.html?now={days[len(days) // 2]['date']}T16:10")
            await page.wait_for_selector(".tl-item")
            await page.evaluate("document.querySelectorAll('details').forEach(d => d.open = true)")
            # Tab fade-in animations would otherwise be caught mid-way and skew the measured opacity
            await page.add_style_tag(content="*, *::before, *::after { animation: none !important; transition: none !important; }")
            merged = {}
            for tab in ("schedule", "flights", "coupons", "practical"):
                await page.click(f'.nav-item[data-tab="{tab}"]')
                await page.wait_for_timeout(250)
                merged.update({f"{tab}:{k}": v for k, v in (await page.evaluate(audit)).items()})
                await self.shot(page, f"{theme}_{tab}")
            results[theme] = merged
            await ctx.close()
        # Fail when dark is below AA and either light passed or dark is clearly worse than light
        worse = []
        for k, v in results["dark"].items():
            light = results["light"].get(k, {"ratio": 99})["ratio"]
            if v["ratio"] < v["need"] and (light >= v["need"] or v["ratio"] < light - 0.3):
                worse.append((k, light, v["ratio"]))
        detail = f"{len(results['dark'])} text elements; dark-only failures: {len(worse)}"
        if worse:
            detail += "\n      " + "\n      ".join(f"{k}  light {l} → dark {d}" for k, l, d in worse[:15])
        return not worse, detail

    async def data_update(self):
        """A changed itinerary_data.json re-renders with a toast and keeps the day the user picked."""
        days = self.data["schedule"]
        picked = len(days) - 2
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "dist"
            shutil.copytree(DIST, root)
            with serve(root) as url:
                ctx = await self.browser.new_context(viewport=VIEWPORT)
                page = await ctx.new_page()
                await page.goto(f"{url}/index.html?now={days[1]['date']}T10:00")
                await page.wait_for_selector(".tl-item")
                await sw_ready(page)
                await page.click(f'.day-chip[data-day-index="{picked}"]')
                changed = json.loads(self.data_text)
                marker = "【回歸測試更新】"
                changed["schedule"][picked]["items"][0]["spot"] = marker + changed["schedule"][picked]["items"][0]["spot"]
                (root / "itinerary_data.json").write_text(json.dumps(changed, ensure_ascii=False), encoding="utf-8")
                await page.evaluate("fetch('itinerary_data.json', { cache: 'no-cache' })")
                try:
                    await page.wait_for_function("document.getElementById('toast').textContent.includes('行程已更新')", timeout=8000)
                    toast = True
                except Exception:
                    toast = False
                state = await page.evaluate("""m => ({ active: document.querySelector('.day-chip.active')?.dataset.dayIndex,
                                                       rendered: document.getElementById('timeline').textContent.includes(m) })""", marker)
                await ctx.close()
        ok = toast and state["active"] == str(picked) and state["rendered"]
        return ok, f"toast={toast} {state} (picked day index {picked})"


CHECKS = ["now_focus", "tomorrow_reminder", "offline_bar", "offline_reload", "single_shell_cache",
          "manifest", "theme_toggle", "dark_contrast", "data_update"]


async def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--only", help="comma-separated check names: " + ", ".join(CHECKS))
    ap.add_argument("--shots", help="directory for screenshots (skipped when omitted)")
    args = ap.parse_args()
    names = args.only.split(",") if args.only else CHECKS
    if args.shots:
        Path(args.shots).mkdir(parents=True, exist_ok=True)

    data_text = (DIST / "itinerary_data.json").read_text(encoding="utf-8")
    async with async_playwright() as p:
        browser = await launch(p)
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "dist"
            shutil.copytree(DIST, root)
            with serve(root) as url:
                suite = Suite(browser, url, json.loads(data_text), args.shots)
                suite.data_text = data_text

                async def run(name):
                    try:
                        return name, *(await getattr(suite, name)())
                    except Exception as e:  # report and keep going
                        return name, False, f"error: {e!r}"

                results = await asyncio.gather(*(run(n) for n in names))
        await browser.close()

    for name, ok, detail in results:
        print(f"{'✓' if ok else '✗'} {name}: {detail}")
    failed = [n for n, ok, _ in results if not ok]
    print(f"\n[{'PASSED' if not failed else 'FAILED'}] browser regression: {len(results) - len(failed)}/{len(results)} checks passed"
          + (f" (failed: {', '.join(failed)})" if failed else ""))
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    asyncio.run(main())
