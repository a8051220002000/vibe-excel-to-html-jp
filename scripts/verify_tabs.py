#!/usr/bin/env python3
"""
scripts/verify_tabs.py
Verification test for Task 2.2:
Checks that 4 Tabs (schedule, flights, coupons, practical) exist,
their corresponding tab-view containers exist,
and tab switching mechanics are correctly implemented.
"""
import sys
import re
from pathlib import Path

def test_tabs():
    html_path = Path("dist/index.html")
    assert html_path.exists(), "dist/index.html does not exist"
    html = html_path.read_text(encoding="utf-8")

    expected_tabs = [
        ("schedule", "每日行程"),
        ("flights", "航班資訊"),
        ("coupons", "優惠券"),
        ("practical", "實用資訊")
    ]

    # Verify buttons
    for tab_id, tab_label in expected_tabs:
        pattern = rf'data-tab=["\']{tab_id}["\'][^>]*>[\s\S]*?{tab_label}'
        assert re.search(pattern, html), f"Bottom nav button for {tab_id} ({tab_label}) not found"
        print(f"✓ Found bottom nav tab button: {tab_label} (data-tab='{tab_id}')")

    # Verify tab views
    for tab_id, _ in expected_tabs:
        view_id = f"view-{tab_id}"
        assert f'id="{view_id}"' in html, f"Tab view #{view_id} not found in DOM"
        print(f"✓ Found corresponding view container: #{view_id}")

    # Verify switchTab function and click listener logic
    assert "function switchTab(tabId)" in html or "switchTab =" in html, "switchTab function not defined"
    assert "addEventListener('click'" in html, "Click listener for navigation tabs not found"
    assert "classList.add('active')" in html, "Active state toggling not found"

    print("\n[PASSED] Bottom nav bar with 4 tabs and instant switching verified successfully.")

if __name__ == "__main__":
    test_tabs()
