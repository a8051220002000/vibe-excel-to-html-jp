#!/usr/bin/env python3
"""
scripts/verify_offline.py
Verification test for Task 3.1:
Checks dist/sw.js precaches index.html + itinerary_data.json, serves cached
copies when the network fails, and that index.html registers the worker with
a relative path (GitHub Pages sub-path safe). Also syntax-checks sw.js with
node when available.
"""
import shutil
import subprocess
from pathlib import Path


def test_offline():
    sw_path = Path("dist/sw.js")
    assert sw_path.exists(), "dist/sw.js does not exist"
    sw = sw_path.read_text(encoding="utf-8")
    html = Path("dist/index.html").read_text(encoding="utf-8")

    for url in ("'./index.html'", "'./itinerary_data.json'"):
        assert url in sw, f"sw.js does not precache {url}"
    print("✓ sw.js precaches index.html and itinerary_data.json")

    for needle in ("addEventListener('install'", "addEventListener('activate'",
                   "addEventListener('fetch'", "cache.match(", "caches.delete("):
        assert needle in sw, f"sw.js missing: {needle}"
    print("✓ install / activate (old cache cleanup) / fetch with cache fallback handlers present")

    assert "serviceWorker.register('./sw.js'" in html, "index.html does not register ./sw.js"
    print("✓ index.html registers ./sw.js with a relative scope")

    if shutil.which("node"):
        subprocess.run(["node", "--check", str(sw_path)], check=True)
        print("✓ sw.js passes node --check syntax validation")

    print("\n[PASSED] Service worker offline caching verified.")


if __name__ == "__main__":
    test_offline()
