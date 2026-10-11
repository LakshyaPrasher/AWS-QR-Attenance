import os
import sys
import json
import tempfile

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from lambda_function import (
    TEMPLATE_SPECS,
    LocationValidator,
    UniversalTemplateBuilder,
    verify_deck,
    verify_slide_content,
    event_config,
    detect_event_source,
    extract_device_platform,
    extract_screen_info,
    S3DataLayer
)
from export_slides import export_presentation

print("=" * 70)
print("RUNNING FULL SUITE TEST FOR OMEGA 5.0 (WHITE BACKGROUND DEFAULT)")
print("=" * 70)

# 1. Test S3 & HTTP Event Source Detection
s3_test_event = {"Event": "s3:TestEvent"}
assert detect_event_source(s3_test_event) == "S3_TEST"

s3_put_event = {
    "Records": [{
        "eventSource": "aws:s3",
        "s3": {
            "bucket": {"name": "attendance-data-lakshyaprasher"},
            "object": {"key": "images/2026-09-22/Joseph+Banda.jpg"}
        }
    }]
}
assert detect_event_source(s3_put_event) == "S3"
date_extracted = S3DataLayer.extract_date_from_key("images/2026-09-22/Joseph+Banda.jpg")
assert date_extracted == "2026-09-22"
print("✓ S3 Event detection & key parsing passed.")

# 2. Test Platform Extraction (Never 'Unknown')
assert extract_device_platform({"platform": "", "userAgent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"}) == "MacBook / macOS"
assert extract_device_platform({"platform": "iPhone", "userAgent": ""}) == "iPhone / iOS"
assert extract_device_platform({"platform": "", "userAgent": "Mozilla/5.0 (Linux; Android 13; Pixel 7)"}) == "Android"
assert extract_device_platform({}) in ["Mobile Device", "Windows PC"]
print("✓ Device platform parser passed (guaranteed zero 'Unknown').")

# 3. Test Location Validation Engine
test_meta = {
    "name": "Joseph Banda",
    "location": {
        "latitude": -15.4167,
        "longitude": 28.2833
    }
}
loc_res = LocationValidator.resolve_and_enrich(test_meta, None, None, None)
print(f"✓ Location resolved: '{loc_res['city_country']}'")
assert "Lusaka" in loc_res["city_country"]
assert "Zambia" in loc_res["city_country"]
assert "°" not in loc_res["city_country"]
assert "S" not in loc_res["city_country"].split(",")[0]  # No latitude orientation tag
print("✓ Zero-Coordinates Policy strictly verified on location string.")

# 4. Test White Background Presentation Generation (Default 2_swiss)
config = event_config({"client": "Airtel Zambia", "workshop": "Cyber Crisis Simulation", "target_date": "2026-09-22"})
template_spec = TEMPLATE_SPECS[config["template"]]
print(f"✓ Loaded template: {template_spec['name']}")
assert template_spec["is_dark"] == False, "Default template must be White / Light background!"

builder = UniversalTemplateBuilder(template_spec, config)

participants = [
    {
        "name": "Joseph Banda",
        "platform": "Android",
        "screen_display": "1080 × 2400 @2.75x",
        "screen_sort_value": 1080 * 2400 * 2.75 * 2.75,
        "camera_permission": True,
        "location_permission": True,
        "city": "Lusaka",
        "country": "Zambia",
        "city_country": "Lusaka, Zambia",
        "lat": -15.4167,
        "lon": 28.2833,
        "image": None,
        "has_image": True
    },
    {
        "name": "Chileshe Mulenga",
        "platform": "iPhone / iOS",
        "screen_display": "1170 × 2532 @3x",
        "screen_sort_value": 1170 * 2532 * 9,
        "camera_permission": True,
        "location_permission": True,
        "city": "Kitwe",
        "country": "Zambia",
        "city_country": "Kitwe, Zambia",
        "lat": -12.8000,
        "lon": 28.2000,
        "image": None,
        "has_image": True
    },
    {
        "name": "Mwamba Phiri",
        "platform": "Windows PC",
        "screen_display": "1920 × 1080 Full HD",
        "screen_sort_value": 1920 * 1080,
        "camera_permission": True,
        "location_permission": True,
        "city": "Lusaka",
        "country": "Zambia",
        "city_country": "Lusaka, Zambia",
        "lat": -15.4167,
        "lon": 28.2833,
        "image": None,
        "has_image": True
    },
    {
        "name": "Mutale Lubinda",
        "platform": "MacBook / macOS",
        "screen_display": "1440 × 900 @2x",
        "screen_sort_value": 1440 * 900 * 4,
        "camera_permission": True,
        "location_permission": True,
        "city": "Lusaka",
        "country": "Zambia",
        "city_country": "Lusaka, Zambia",
        "lat": -15.4167,
        "lon": 28.2833,
        "image": None,
        "has_image": True
    },
    {
        "name": "Kondwani Tembo",
        "platform": "Android",
        "screen_display": "1080 × 2340 @2.6x",
        "screen_sort_value": 1080 * 2340 * 2.6 * 2.6,
        "camera_permission": True,
        "location_permission": False,
        "city": "Ndola",
        "country": "Zambia",
        "city_country": "Ndola, Zambia",
        "lat": -12.9667,
        "lon": 28.6333,
        "image": None,
        "has_image": True
    }
]

builder.build_cover_slide(len(participants))
builder.build_permission_slide(participants)
builder.build_gallery_slides(participants)
builder.build_location_slide(participants)
builder.build_screen_slide(participants)
builder.build_platform_slide(participants)
builder.build_summary_slide(participants)

out_pptx = os.path.join(os.getcwd(), "Swiss_White_Deck.pptx")
builder.prs.save(out_pptx)
print(f"✓ Saved presentation to: {out_pptx}")

verif = verify_deck(out_pptx)
print("✓ Verification report:", json.dumps(verif, indent=2))
assert verif["ok"] == True
assert verif["slide_count"] == 7
assert len(verif["issues"]) == 0, f"Issues detected: {verif['issues']}"

content = verify_slide_content(builder.prs)
print("✓ Content verification:", json.dumps(content, indent=2))
assert content["ok"] == True, f"Content verification failed: {content['warnings']}"
for s in content["slides"]:
    assert s["shapes"] >= 3, f"Slide {s['name']} has too few shapes: {s['shapes']}"
print("✓ All slides have sufficient content (bulletproof check passed).")

# Export and verify PNG rendering
export_dir = os.path.join(os.getcwd(), "swiss_white_preview")
os.makedirs(export_dir, exist_ok=True)
exported = export_presentation(out_pptx, export_dir)
assert exported == True
print("✓ Successfully exported all 7 slides to PNG!")

print("\n" + "=" * 70)
print("ALL TESTS PASSED SUCCESSFULLY! DEPLOYMENT READY!")
print("=" * 70)
