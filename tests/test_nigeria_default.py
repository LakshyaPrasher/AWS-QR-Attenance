"""Production-scenario test: default config (Airtel Nigeria) + Lagos participants.

Additive only - does not touch or depend on export_slides (no PowerPoint needed).
Checks: client default, Nigeria resolution, 7-slide deck, zero raw coordinates in
slide text, and the cover shows Airtel Nigeria (not Zambia).
"""
import os
import re
import sys
import tempfile

if sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from pptx import Presentation
from lambda_function import (
    TEMPLATE_SPECS, LocationValidator, UniversalTemplateBuilder,
    event_config, verify_deck, verify_slide_content,
)

print("=" * 70)
print("NIGERIA DEFAULT-CONFIG SCENARIO")
print("=" * 70)

# 1. Defaults
config = event_config({"client": "Airtel Nigeria"})
assert config["client"] == "Airtel Nigeria", config["client"]
assert config["workshop"] == "Cyber Crisis Simulation"
assert TEMPLATE_SPECS[config["template"]]["is_dark"] is False
print("✓ event_config with Airtel Nigeria / Cyber Crisis Simulation / light template")

# 2. Banana Island GPS resolves to Nigeria with no coordinates in the label
res = LocationValidator.resolve_and_enrich({"name": "Ada", "location": {"latitude": 6.4682, "longitude": 3.4472}}, None, None, None)
label = res["city_country"]
assert "Nigeria" in label and "Zambia" not in label, label
assert not re.search(r"\d+\.\d+", label) and "°" not in label, label
print("✓ Banana Island GPS resolves to: %r (no coordinates)" % label)

# 3. Build the 7-slide deck for Lagos participants
def p(name, plat, disp, sv, lat, lon):
    return {"name": name, "platform": plat, "screen_display": disp, "screen_sort_value": sv,
            "camera_permission": True, "location_permission": True,
            "city": "Lagos", "country": "Nigeria", "city_country": "Lagos, Nigeria",
            "lat": lat, "lon": lon, "image": None, "has_image": True}

people = [
    p("Ada Nwosu", "Windows PC", "1920 × 1080 @1x", 1920 * 1080, 6.4682, 3.4472),
    p("Tunde Bello", "Android", "1080 × 2400 @2.75x", 1080 * 2400 * 7.5625, 6.4684, 3.4474),
    p("Chioma Eze", "iPhone / iOS", "1170 × 2532 @3x", 1170 * 2532 * 9, 6.4679, 3.4471),
    p("Ibrahim Musa", "Linux", "2560 × 1440 @1x", 2560 * 1440, 6.4681, 3.4469),
]
builder = UniversalTemplateBuilder(TEMPLATE_SPECS[config["template"]], config)
builder.build_cover_slide(len(people))
builder.build_permission_slide(people)
builder.build_gallery_slides(people)
builder.build_location_slide(people)
builder.build_screen_slide(people)
builder.build_platform_slide(people)
builder.build_summary_slide(people)

out = os.path.join(tempfile.gettempdir(), "nigeria_default_check.pptx")
builder.prs.save(out)
v = verify_deck(out)
assert v["ok"] and v["slide_count"] == 7 and not v["issues"], v
assert verify_slide_content(builder.prs)["ok"]
print("✓ 7-slide deck built and verified")

# 4. Text checks on the saved deck
prs = Presentation(out)
all_text = []
for i, slide in enumerate(prs.slides, 1):
    for sh in slide.shapes:
        if sh.has_text_frame:
            all_text.append((i, sh.text_frame.text))
joined = "\n".join(t for _, t in all_text)
assert "AIRTEL NIGERIA" in joined.upper(), "cover must show Airtel Nigeria"
assert "ZAMBIA" not in joined.upper() and "LUSAKA" not in joined.upper(), "Zambia leaked into Nigeria deck"
print("✓ Deck says Airtel Nigeria; no Zambia/Lusaka text")

coord = re.compile(r"-?\d{1,3}\.\d{3,}")
for i, t in all_text:
    assert not coord.search(t), "raw coordinate on slide %d: %r" % (i, t)
print("✓ Zero raw coordinates in any slide text box")

# 5. Platform slide carries the Pillow donut chart (labels live in the image;
#    OS consolidation itself is covered by test_platform_consolidation.py)
plat_slide = prs.slides[5]
assert any(sh.shape_type == 13 for sh in plat_slide.shapes), "donut chart picture missing on slide 6"
print("✓ Platform slide has the donut chart picture")

print("=" * 70)
print("ALL NIGERIA DEFAULT-CONFIG TESTS PASSED!")
print("=" * 70)
