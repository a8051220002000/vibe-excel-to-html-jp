#!/usr/bin/env python3
"""
scripts/verify_cards.py
Verification test for Task 2.4:
Checks flight cards (outbound / inbound), coupon grid, copy-to-clipboard
utility (Clipboard API + legacy fallback) and the "已複製" Toast wiring,
plus the data the cards depend on.
"""
import json
import re
from pathlib import Path


def test_cards():
    html = Path("dist/index.html").read_text(encoding="utf-8")
    data = json.loads(Path("dist/itinerary_data.json").read_text(encoding="utf-8"))

    for fn in ("function renderFlights", "function renderCoupons", "function renderPractical"):
        assert fn in html, f"Missing renderer: {fn}"
    assert "flight-card" in html and "coupon-grid" in html, "Card markup classes missing"
    assert re.search(r"\.coupon-grid\s*\{[^}]*display:\s*grid", html), "Coupon grid is not a CSS grid"
    print("✓ Flight cards, coupon grid and practical renderers present")

    assert "navigator.clipboard" in html and "execCommand('copy')" in html, "Copy utility needs Clipboard API + fallback"
    assert "closest('[data-copy]')" in html, "Delegated copy click handler missing"
    assert "已複製" in html and "function showToast" in html, "Toast '已複製' wiring missing"
    print("✓ Copy-to-clipboard utility with fallback and '已複製' Toast wired")

    flights = data.get("flights", [])
    types = " ".join(f.get("type", "") for f in flights)
    assert "去程" in types and "回程" in types, "Need both outbound (去程) and inbound (回程) flights"
    print(f"✓ Flights: {len(flights)} (去程 + 回程)")

    coupons = data.get("coupons", [])
    assert coupons and all(c.get("title") for c in coupons), "Coupons missing or untitled"
    print(f"✓ Coupons: {len(coupons)}")

    p = data.get("practical", {})
    addrs = [a.get("address") for a in p.get("accommodations", []) if a.get("address")]
    tels = [c.get("tel") for c in p.get("emergency_contacts", []) if c.get("tel")]
    assert addrs and tels, "Need at least one address and one phone number to copy"
    print(f"✓ Copyable data: {len(addrs)} address(es), {len(tels)} phone number(s)")

    print("\n[PASSED] Flight cards, coupon grid and copy utilities verified.")


if __name__ == "__main__":
    test_cards()
