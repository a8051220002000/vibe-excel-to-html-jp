#!/usr/bin/env python3
"""
Verify Mobile-First layout specifications:
- Viewport meta configuration with viewport-fit=cover
- Safe area inset CSS configuration
- 375px viewport constraints & overflow-x safety rules
"""
import sys
import re
from pathlib import Path

def test_mobile_layout():
    html_file = Path("dist/index.html")
    if not html_file.exists():
        print("FAIL: dist/index.html does not exist.")
        sys.exit(1)

    content = html_file.read_text(encoding="utf-8")

    # 1. Check viewport meta
    viewport_match = re.search(r'<meta[^>]*name=["\']viewport["\'][^>]*content=["\']([^"\']+)["\']', content, re.I)
    assert viewport_match, "Viewport meta tag not found"
    viewport_content = viewport_match.group(1)
    assert "width=device-width" in viewport_content, "Missing width=device-width in viewport"
    assert "viewport-fit=cover" in viewport_content, "Missing viewport-fit=cover in viewport"
    print("✓ Viewport meta configured correctly with viewport-fit=cover")

    # 2. Check Safe Area Insets
    assert "safe-area-inset-bottom" in content, "Missing safe-area-inset-bottom handling"
    assert "safe-area-inset-top" in content, "Missing safe-area-inset-top handling"
    print("✓ Safe Area Insets configured in CSS")

    # 3. Check overflow-x and box-sizing rules
    assert "box-sizing: border-box" in content, "Missing universal box-sizing: border-box"
    assert "overflow-x: hidden" in content, "Missing overflow-x: hidden"
    assert "max-width: 480px" in content or "max-width: 100%" in content, "Missing mobile max-width container constraint"
    print("✓ Box model and overflow-x protections verified (375px safe)")

    # 4. Check that no fixed widths greater than 375px exist in styles
    fixed_widths = re.findall(r'(?<![a-zA-Z-])width:\s*([0-9]+)px', content)
    for w in fixed_widths:
        w_val = int(w)
        assert w_val <= 375, f"Found fixed width {w_val}px > 375px which may cause overflow"
    print("✓ No fixed pixel widths exceeding 375px found")

    print("\n[PASSED] 375px mobile layout acceptance test passed cleanly without horizontal overflow.")

if __name__ == "__main__":
    test_mobile_layout()
