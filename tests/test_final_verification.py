import os
import sys
import json
import math
from io import BytesIO
from PIL import Image

_TEST_DIR = os.path.dirname(os.path.abspath(__file__))
_PARENT_DIR = os.path.abspath(os.path.join(_TEST_DIR, ".."))
for _p in [_TEST_DIR, _PARENT_DIR]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from lambda_function import (
    LocationValidator,
    UniversalTemplateBuilder,
    TEMPLATE_SPECS,
    event_config,
    verify_deck,
    verify_slide_content
)
from export_slides import export_presentation

print("=" * 70)
print("RUNNING COMPREHENSIVE FINAL VERIFICATION")
print("=" * 70)

# 1. Test Coordinate Clustering (e.g. 13.4005 vs 13.4001)
c1 = LocationValidator.cluster_coords(13.4005, 28.2801)
c2 = LocationValidator.cluster_coords(13.4001, 28.2805)
print(f"Cluster 1: {c1}, Cluster 2: {c2}")
assert c1 == c2, f"Clustering failed: {c1} != {c2}"
print("✓ Coordinate clustering passed: 13.4005 and 13.4001 map to exact same (13.4, 28.28) cluster!")

# 2. Test Coordinate Preservation (never overwritten with default)
meta1 = {"name": "Test User 1", "location": {"latitude": 13.4005, "longitude": 28.2801}}
meta2 = {"name": "Test User 2", "location": {"latitude": 13.4001, "longitude": 28.2805}}

res1 = LocationValidator.resolve_and_enrich(meta1)
res2 = LocationValidator.resolve_and_enrich(meta2)

assert res1["lat"] == 13.4005, f"Latitude was overwritten: {res1['lat']}"
assert res2["lat"] == 13.4001, f"Latitude was overwritten: {res2['lat']}"
assert res1["city_country"] == res2["city_country"], f"City country mismatch: {res1['city_country']} vs {res2['city_country']}"
print(f"✓ Coordinate preservation passed: {res1['lat']} and {res2['lat']} preserved; resolved to '{res1['city_country']}'!")

# 3. Test 12 Participants on ONE Gallery Slide
config = event_config({"client": "Airtel Zambia", "workshop": "Cyber Crisis Simulation", "target_date": "2026-09-22"})
template_spec = TEMPLATE_SPECS[config["template"]]
builder = UniversalTemplateBuilder(template_spec, config)

participants_12 = []
# Create 12 realistic participants matching user's workshop data
names = [
    "Kapa Kaumba", "Anil Dwivedi", "Shrikant Vyas", "Rakesh Kumar Agarwal",
    "Soft Mulizwa", "Keith Banda", "Chanda Mwewa", "Bwalya Chilufya",
    "Nchimunya Haimbe", "Lubinda Mwiinga", "Taonga Phiri", "Zondani Sakala"
]
for i, name in enumerate(names):
    # Slight coordinate variation within the same venue (e.g. -15.4167 +/- 0.002)
    lat_var = -15.4160 + (i * 0.0001)
    lon_var = 28.2830 + (i * 0.0001)
    participants_12.append({
        "name": name,
        "platform": ["Android", "iPhone / iOS", "Windows PC", "MacBook / macOS"][i % 4],
        "screen_display": "1080 × 2400 @2.75x",
        "screen_sort_value": 1080 * 2400 * 2.75 * 2.75,
        "camera_permission": True,
        "location_permission": True if i < 9 else False,
        "city": "Lusaka",
        "country": "Zambia",
        "city_country": "Lusaka, Zambia",
        "lat": lat_var,
        "lon": lon_var,
        "has_coords": True if i < 9 else False,
        "image": None,
        "has_image": True
    })

builder.build_cover_slide(len(participants_12))
builder.build_permission_slide(participants_12)
builder.build_gallery_slides(participants_12)
builder.build_location_slide(participants_12)
builder.build_screen_slide(participants_12)
builder.build_platform_slide(participants_12)
builder.build_summary_slide(participants_12)

out_pptx = os.path.join(os.getcwd(), "Airtel_Final_Deck.pptx")
builder.prs.save(out_pptx)
print(f"✓ Saved 12-attendee presentation to: {out_pptx}")

# Verify Slide Count is exactly 7 (all 12 on 1 gallery slide)
verif = verify_deck(out_pptx)
print(f"✓ Verification: slide_count={verif['slide_count']}, ok={verif['ok']}")
assert verif["ok"] == True
assert verif["slide_count"] == 7, f"Expected 7 slides, got {verif['slide_count']}"

# Verify Content
content = verify_slide_content(builder.prs)
assert content["ok"] == True
gallery_slide = content["slides"][2]
assert gallery_slide["name"] == "Gallery"
print(f"✓ Gallery slide verified: shapes={gallery_slide['shapes']}, pictures={gallery_slide['pictures']}")
assert gallery_slide["pictures"] == 12, f"Expected 12 pictures on gallery slide, got {gallery_slide['pictures']}"

# 4. Verify Picture Borders are completely gone
for s_idx, slide in enumerate(builder.prs.slides):
    for shape in slide.shapes:
        if shape.shape_type == 13:  # Picture
            ln = shape._element.spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ln')
            if ln is not None:
                no_fill = ln.find('{http://schemas.openxmlformats.org/drawingml/2006/main}noFill')
                assert no_fill is not None, f"Slide {s_idx+1} picture {shape.name} has outline fill!"
print("✓ All pictures verified to have <a:noFill/> (zero square borders)!")

# 5. Export slides to PNG
export_dir = os.path.join(os.getcwd(), "final_preview")
os.makedirs(export_dir, exist_ok=True)
exported = export_presentation(out_pptx, export_dir)
assert exported == True
print("✓ Successfully exported all 7 slides to final_preview/")
print("=" * 70)
print("ALL TESTS PASSED WITH 100% PERFECTION!")
print("=" * 70)
