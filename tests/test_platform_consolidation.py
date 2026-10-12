import sys
import os

_TEST_DIR = os.path.dirname(os.path.abspath(__file__))
_PARENT_DIR = os.path.abspath(os.path.join(_TEST_DIR, ".."))
for _p in [_TEST_DIR, _PARENT_DIR]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

from collections import Counter
from PIL import Image

# Import functions from lambda_function
from lambda_function import extract_device_platform, normalize_platform, render_donut_chart

def run_tests():
    print("=== TEST 1: Platform Extraction & Consolidation ===")
    tests = [
        ({"platform": "Win32"}, "Windows PC"),
        ({"platform": "Workstation"}, "Windows PC"),
        ({"platform": "Windows"}, "Windows PC"),
        ({"platform": "Desktop"}, "Windows PC"),
        ({"userAgent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}, "Windows PC"),
        ({"platform": "", "rawPlatform": "", "userAgent": ""}, "Windows PC"),
        # Linux must remain distinct!
        ({"platform": "Linux", "userAgent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64)"}, "Linux"),
        ({"platform": "linux armv8l", "userAgent": "curl/7.68.0"}, "Linux"),
        # Other devices
        ({"platform": "Android", "userAgent": "Mozilla/5.0 (Linux; Android 13; Pixel 7)"}, "Android"),
        ({"platform": "iPhone", "userAgent": "Mozilla/5.0 (iPhone; CPU iPhone OS 16_5 like Mac OS X)"}, "iPhone / iOS"),
        ({"platform": "MacIntel", "userAgent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"}, "MacBook / macOS"),
    ]

    for meta, expected in tests:
        actual = extract_device_platform(meta)
        assert actual == expected, f"Failed for {meta}: expected {expected}, got {actual}"
        print(f"  [PASS] {meta} -> '{actual}'")

    print("\n=== TEST 2: Platform Normalization ===")
    norm_tests = [
        ("Workstation", "Windows PC"),
        ("Windows PC", "Windows PC"),
        ("Windows", "Windows PC"),
        ("Win32", "Windows PC"),
        ("PC", "Windows PC"),
        ("Linux", "Linux"),
        ("Linux Workstation", "Linux"),
        ("Android", "Android"),
        ("iPhone / iOS", "iPhone / iOS"),
        ("MacBook / macOS", "MacBook / macOS"),
    ]
    for inp, expected in norm_tests:
        actual = normalize_platform(inp)
        assert actual == expected, f"Failed for '{inp}': expected {expected}, got {actual}"
        print(f"  [PASS] normalize('{inp}') -> '{actual}'")

    print("\n=== TEST 3: User Scenario Consolidation (8 Attendees) ===")
    # Exact scenario from user's screenshot:
    # 3 Android, 3 Workstation, 1 iPhone / iOS, 1 Windows PC
    raw_participants = [
        {"name": "P1", "platform": "Android"},
        {"name": "P2", "platform": "Android"},
        {"name": "P3", "platform": "Android"},
        {"name": "P4", "platform": "Workstation"},
        {"name": "P5", "platform": "Workstation"},
        {"name": "P6", "platform": "Workstation"},
        {"name": "P7", "platform": "iPhone / iOS"},
        {"name": "P8", "platform": "Windows PC"},
    ]

    normalized_counts = Counter(normalize_platform(p["platform"]) for p in raw_participants)
    print("Normalized Counts:", dict(normalized_counts))
    assert "Workstation" not in normalized_counts, "Workstation should not appear in counts!"
    assert normalized_counts["Windows PC"] == 4, f"Expected 4 Windows PC, got {normalized_counts['Windows PC']}"
    assert normalized_counts["Android"] == 3
    assert normalized_counts["iPhone / iOS"] == 1

    print("\n=== TEST 4: Linux Distinctness ===")
    raw_with_linux = raw_participants + [{"name": "P9", "platform": "Linux"}]
    linux_counts = Counter(normalize_platform(p["platform"]) for p in raw_with_linux)
    print("Counts with Linux:", dict(linux_counts))
    assert linux_counts["Linux"] == 1, "Linux should be distinct!"
    assert linux_counts["Windows PC"] == 4, "Windows PC should remain 4!"

    print("\n=== TEST 5: Render Donut Chart for Consolidated Scenario ===")
    ranked = normalized_counts.most_common(6)
    cats, vals = [c for c, _ in ranked], [v for _, v in ranked]
    donut_bytes = render_donut_chart(cats, vals, "Operating System Share", is_dark=False, primary_color="#E40000", card_bg="#FFFFFF")
    out_img = "test_consolidated_os_donut.png"
    with open(out_img, "wb") as f:
        f.write(donut_bytes)
    print(f"Saved consolidated donut chart: {out_img} ({os.path.getsize(out_img)} bytes)")

    print("\nALL PLATFORM CONSOLIDATION TESTS PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    run_tests()
