#!/usr/bin/env python3
"""
scripts/verify_timeline.py
Verification test for Task 2.3:
Checks the timeline view wiring in dist/index.html (sticky day bar, day chips,
day filtering, card fields, expandable remarks) and that the JSON data provides
what the timeline consumes. Simulates the per-day filter against the data.
"""
import json
import re
from pathlib import Path


def test_timeline():
    html = Path("dist/index.html").read_text(encoding="utf-8")
    data = json.loads(Path("dist/itinerary_data.json").read_text(encoding="utf-8"))

    # Markup / CSS hooks
    assert 'id="dayBar"' in html and 'id="timeline"' in html, "Day bar / timeline containers missing"
    assert re.search(r"\.day-bar\s*\{[^}]*position:\s*sticky", html), "Day bar is not sticky"
    assert re.search(r"\.day-bar\s*\{[^}]*overflow-x:\s*auto", html), "Day bar is not horizontally scrollable"
    print("✓ Sticky, horizontally scrollable day bar present")

    # JS behaviour hooks
    for needle in ("function renderDayBar", "function selectDay", "function renderTimelineItem",
                   "fetch('itinerary_data.json'", "data-day-index", "it.day === day.day"):
        assert needle in html, f"Timeline JS missing: {needle}"
    for field in ("time_start", "time_end", "spot", "transport", "remarks"):
        assert f"it.{field}" in html, f"Timeline card does not render field: {field}"
    assert '<details class="tl-remarks">' in html, "Remarks are not expandable"
    print("✓ Day filter, card fields (time / spot / transport) and expandable remarks wired")

    # Data contract + filter simulation
    schedule = data["schedule"]
    assert len(schedule) > 0, "Schedule empty"
    for day in schedule:
        filtered = [it for it in day["items"] if it["day"] == day["day"]]
        assert filtered, f"{day['day']} has no items after filtering"
        assert len(filtered) == len(day["items"]), f"{day['day']} contains items from other days"
        print(f"✓ {day['day']} ({day['date']}): {len(filtered)} items")

    print("\n[PASSED] Interactive timeline with day filtering verified.")


if __name__ == "__main__":
    test_timeline()
