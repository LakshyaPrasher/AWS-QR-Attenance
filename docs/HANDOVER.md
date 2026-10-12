# 📋 HANDOVER DOCUMENT FOR CLAUDE
## Airtel Cyber Crisis Simulation & Executive Presentation Engine (Omega 5.0)

> **Document Version:** 1.0.0  
> **Target Recipient:** Claude (AI Pair Programmer / Systems Engineer)  
> **Project Root Directory:** `.  
> **Target OS:** Windows (PowerShell environment; invoke Python using `py <script>`, NOT `python`)  
> **Status:** Production-Ready & Verified (All test suites passing 100%)

---

## 🛑 PRIME DIRECTIVE: KEEP IT SIMPLE — DO NOT OVER-ENGINEER

> [!IMPORTANT]
> ### THE GOLDEN RULE FOR CLAUDE
> **WE DO NOT WANT TO OVER-ENGINEER ANYTHING. WE JUST WANT TO PASS.**
> 
> 1. **Surgical, minimal changes only:** Fix or implement exactly what is asked. Do not rewrite working modules or "clean up" unrelated code.
> 2. **Respect working code:** The presentation engine (`lambda_function.py`), attendance handler (`attendancehandler.py`), frontend (`app.js`, `index.html`), and all test scripts are currently passing 100%. Do not refactor them unless specifically instructed.
> 3. **Pragmatic execution:** Always prefer the simplest solution that satisfies requirements and passes tests. No unnecessary abstractions, no speculative features, no architectural over-hauls.
> 4. **Run existing verification scripts:** After making any change, run `py <test_script.py>` and verify everything exits with code 0 before concluding.

---

## 1. Executive System Overview & Architecture

The system is an end-to-end event-driven cyber simulation reporting platform built on AWS:

```
[Participant Browser] (index.html + app.js)
        │
        │ HTTP POST (multipart/form-data: metadata JSON + webcam JPEG)
        ▼
[AWS API Gateway] 
  URL: https://bag6d60ugc.execute-api.ap-south-1.amazonaws.com/prod/attendance
        │
        ▼
[Lambda 1: Attendance Handler] (attendancehandler.py)
  ├─ Parses multipart payload
  ├─ Consolidates platform ("Windows PC", "Linux", etc.)
  ├─ 4-tier location cascade: GPS -> CloudFront Headers -> IP Geolocation -> Timezone
  ├─ Enriches via Nominatim reverse geocoder
  └─ Writes to S3 Bucket: `attendance-data-lakshyaprasher`
        ├─ raw-text/{YYYY-MM-DD}/{Name}.txt
        ├─ records/{YYYY-MM-DD}/{Name}.json
        └─ images/{YYYY-MM-DD}/{Name}.jpg
        │
        │ S3 Object Created Event (or HTTP trigger)
        ▼
[Lambda 2: PPT Generator] (lambda_function.py — Omega 5.0)
  ├─ PASS 1: Telemetry extraction, face detection (OpenCV/PIL), coordinate clustering
  ├─ PASS 2: Session venue majority vote (Lagos, Nigeria hubs prioritized over legacy)
  ├─ Cartographic Map Engine (PIL tile stitching, water cleaning, test_var_a floating card)
  ├─ 6 Executive Design Templates (Default: `2_swiss` Swiss Precision Editorial Light)
  ├─ Strict Zero-Coordinates Policy (lat/long never displayed raw)
  └─ Uploads final .pptx to S3 Bucket: `final-ppt-report` -> reports/{slug}_{date}.pptx
```

---

## 2. File Map & Project Layout

All files reside in: `.

| File | Type | Lines | Role & Key Responsibilities |
|---|---|---|---|
| **`app.js`** | Frontend JS | ~340 | Client-side attendance logic. Permissions API camera pre-check (silent for returning users), non-blocking geolocation, typing speed tracking, FormData assembly. |
| **`index.html`** | Frontend HTML | 1055 | Responsive 100dvh zero-scroll single-page application. Authentic Airtel Nigeria branding, animated SVG telecom transmission tower, executive success view with confetti. |
| **`attendancehandler.py`** | AWS Lambda | ~540 | Ingests attendance submissions, parses multipart/form-data, normalizes OS, resolves 4-tier location, writes `.jpg`, `.json`, `.txt` to S3. Zero external dependencies (urllib fallback). |
| **`lambda_function.py`** | AWS Lambda | ~2410 | Core Omega 5.0 PPTX generator. OpenCV face detection, cartographic map generation (test_var_a), coordinate clustering, 7-slide deck builder, 6 template themes. |
| **`test_location_fallbacks.py`** | Test Script | ~50 | Unit tests for attendancehandler 4-tier location cascade (GPS, CloudFront, Timezone). |
| **`test_platform_consolidation.py`** | Test Script | ~130 | Verifies Windows PC consolidation (Workstation/Windows/Win32 -> Windows PC, Linux distinct). |
| **`test_final_verification.py`** | Test Script | ~170 | Comprehensive slide deck generation test with coordinate clustering and borderless pictures. |
| **`test_full_pipeline.py`** | Test Script | ~230 | End-to-end Omega 5.0 pipeline test, template validation, zero coordinates verification. |
| **`unpack_files.py`** | Utility | ~60 | Historical extraction script from transcript archives. |
| **`clean_index_html.py`** | Utility | ~15 | Cleaned trailing script lines in index.html. |

---

## 3. Strict Business Rules & Non-Negotiables

### A. Platform Consolidation Rule
- **Rule:** Workstation, Windows, Win32, Desktop, and PC must **always** be consolidated and displayed as **`"Windows PC"`**.
- **Exception:** **`"Linux"`** must **always remain distinct**.
- **Mobile/Mac:** `"Android"`, `"iPhone / iOS"`, and `"MacBook / macOS"` are kept distinct.
- **Implemented in:**
  - `app.js` -> `getPlatformGroup()`
  - `attendancehandler.py` -> `normalize_platform()`
  - `lambda_function.py` -> `normalize_platform()` and `extract_device_platform()`

### B. Zero-Coordinates Policy
- **Rule:** Raw latitude/longitude coordinates must **NEVER** appear as numbers on any slide text box, table, KPI card, or chart.
- **Enforcement:** All coordinates are clustered to 2-decimal precision (`~1.1km`) and reverse-geocoded to **`"City, Country"`** (e.g., `"Lagos, Nigeria"`).
- **Map Visual Exception:** The cartographic map on Slide 4 includes a stylized minimalist coordinates pill in the floating metadata card (`test_var_a` design), which is the only authorized place coordinates may appear.

### C. Client & Geographic Context
- **Primary Client:** **`"Airtel Nigeria"`** (Cyber Crisis Simulation).
- **Primary Venue:** **`"Banana Island, Ikoyi, Lagos, Nigeria"`** (Coordinates: ~6.4682° N, 3.4472° E).
- **Default Theme:** **`2_swiss`** (Swiss Precision Editorial — international typographic white background).

### D. Map Visual Specifications (`test_var_a`)
- **Design:** Minimalist white floating card anchored on the bottom-left of the map.
- **Elements:** Red pulsing status dot, `"AIRTEL CRISIS EPICENTER"` title, address subtitle, coordinates pill.
- **Strictly Removed:** 
  - ❌ Zoom +/- navigation controls
  - ❌ Red "PRIMARY SIMULATION EPICENTER" header text
  - ❌ "Verified via OpenStreetMap" footer text
  - ❌ Blue dashed ferry/water routes on Lagos Lagoon (cleaned via numpy/Pillow color masks)
  - ❌ Excessively cluttered small street labels (major regions enlarged: BANANA ISLAND, IKOYI, LEKKI PHASE 1).

---

## 4. Camera & Location Mechanics (Recent Major Upgrade)

### Camera Capture:
1. **Returning Users (Silent & Invisible):**
   - `app.js` runs `queryCameraPermissionState()` using the Permissions API (`navigator.permissions.query({name: "camera"})`).
   - If state is `"granted"`: `getUserMedia()` activates **silently with zero browser prompts**. A hidden `<video>` plays, a 500ms delay elapses, and a hidden `<canvas>` grabs a 640×480 frame. Tracks are stopped immediately. User never sees a viewfinder or shutter button.
2. **First-Time Users:**
   - Status displays `"Verifying participant identity…"`. Browser displays the native one-time "Allow" dialog. Once allowed, frame is captured and browser caches the permission for all future visits.
3. **Denied Users:**
   - If state is `"denied"`, the function returns immediately without hanging or throwing user-facing errors.

### Location Capture (4-Tier Zero-Prompt Cascade):
Implemented in `attendancehandler.py` (`normalize_location`):
1. **Tier 1 — Browser GPS:** Uses exact GPS if granted by user via `navigator.geolocation`.
2. **Tier 2 — AWS CloudFront Viewer Headers (Prompt-Free):**
   - `cloudfront-viewer-latitude`, `cloudfront-viewer-longitude`
   - `cloudfront-viewer-city`, `cloudfront-viewer-country`
   - `cloudfront-viewer-country-region`
   - Zero browser permission needed. Edge CDN injects headers directly at the HTTP layer.
3. **Tier 3 — Server-Side IP Geolocation:**
   - Calls `https://ipapi.co/{ip}/json/` via single-lookup cache `fetch_ip_info(ip)`.
   - Resolves public IP to City, Region, Country, and approximate coordinates.
4. **Tier 4 — Client Timezone Inference:**
   - Reads client timezone from metadata (e.g., `"Africa/Lagos"` -> `"Lagos, Nigeria"`).
5. **Presentation Engine Recognition:**
   - `lambda_function.py`'s `permission_granted(meta, "location")` recognizes `"network_inferred"` and `"timezone_inferred"`, ensuring participants resolved without browser prompts are still counted as verified location attendees.

---

## 5. Slide Deck Blueprint (Omega 5.0 — 7 Slides)

| Slide # | Name | Description & Key Visual Elements |
|---|---|---|
| **Slide 1** | **Executive Cover** | Crimson brand typography, client name, workshop title, date, editorial layout rules. |
| **Slide 2** | **Permission & Telemetry Audit** | Metric cards: Camera Permission Rate, Location Permission Rate, Typing Speed KPI, Security Awareness Scorecard. |
| **Slide 3** | **Participant Gallery** | Grid of participant webcam photos with zero-border formatting (`<a:noFill/>`), names, and timestamps. Black screens and face-less captures filtered. |
| **Slide 4** | **Cartographic Venue Map** | High-definition map of Banana Island / Lagos with `test_var_a` floating card, regional breakdown table. |
| **Slide 5** | **Display & Hardware Telemetry** | Screen resolution breakdown, aspect ratios, DPI scales. |
| **Slide 6** | **Platform Ecosystem** | Pillow-rendered Donut Chart showing OS breakdown (Windows PC consolidated, Linux distinct, Android, iOS). |
| **Slide 7** | **Executive Summary & Action Plan** | Summary KPIs, security vulnerability insights, participant totals, final sign-off. |

---

## 6. Known Gotchas & Pitfalls to Avoid

1. **Windows Python Execution:**
   - On this machine, use **`py <script.py>`**. Do NOT use `python <script.py>`.
2. **PowerShell Quote Escaping:**
   - Running multiline Python inline via `py -c "..."` often fails in PowerShell due to quote escaping and syntax errors. Always write a temporary `.py` script and run `py <script.py>` instead.
3. **Python-pptx Shape Deletion Bug:**
   - Never call `sp.getparent().remove(sp)` directly in `python-pptx` without rebuilding relationships. It corrupts PowerPoint XML and causes COM error `-2147352567`. To make shapes invisible, use transparent formatting or zero-fill XML elements (`<a:noFill/>`).
4. **Standard Library Fallbacks in Lambda:**
   - `attendancehandler.py` includes `http_get_json()` with `urllib.request` fallback so it runs cleanly even if `requests` is not installed in the environment.
5. **Tile Caching:**
   - Map tiles are cached in `%TEMP%\osm_tiles\` (Windows) or `/tmp/osm_tiles/` (Lambda). Master pre-rendered HD map asset is `airtel_banana_island_minimalist_map_hd.png`.
6. **Past "Lusaka Bug":**
   - Ensure `event_config()` default client remains `"Airtel Nigeria"` (not Zambia), and `NIGERIA_HUBS` are prioritized in `LocationValidator`.
7. **Python 3.13 Upgrade & Zero-Billing AWS Protocol:**
   - Dockerfile base image: `public.ecr.aws/lambda/python:3.13`.
   - Managed Lambda runtimes: `python3.13` (AL2023).
   - Dependencies in `requirements.txt` are pinned to Python 3.13 compatible baselines (`numpy>=2.1.0`, `Pillow>=11.0.0`, etc.).
   - Never attach Lambda to a VPC NAT Gateway or enable Provisioned Concurrency ($0.00 free tier guarantee).

---

## 7. How to Verify Everything (Test Suite Guide)

Run the master test runner from `.

```powershell
py run_all_tests.py
# Runs all 6 test suites across all 3 Lambda services (100% pass expected):
# 1. test_location_fallbacks.py       -> Attendance Handler location cascade
# 2. test_platform_consolidation.py   -> Platform consolidation & distinct Linux
# 3. test_final_verification.py       -> 7-slide deck & coordinate clustering
# 4. test_full_pipeline.py            -> Full Omega 5.0 pipeline & Swiss white theme
# 5. test_nigeria_default.py          -> Production default config & Banana Island GPS
# 6. test_excel_compiler.py           -> Attendance Excel compiler & S3 workbook generation
```

---

## 8. Final Checklist Before Answering the User

When picking up any future task from the user:
- [ ] Did you keep the solution as simple as possible without over-engineering?
- [ ] Did you preserve existing platform consolidation rules ("Windows PC", Linux distinct)?
- [ ] Did you preserve the Zero-Coordinates policy?
- [ ] Did you run the relevant test scripts (`py <test>.py`) to confirm zero regressions?
- [ ] Did you verify that no syntax errors were introduced?
