# Executive System Architecture & Engineering Specifications
## Event-Driven Cyber Crisis Simulation Reporting Engine (Omega 5.0)

### 1. High-Level System Architecture

The reporting engine is an event-driven AWS serverless architecture designed for high-concurrency simulation workshops:

```
[Participant Browser] (index.html + app.js)
        │
        │ HTTP POST (multipart/form-data: metadata JSON + webcam JPEG)
        ▼
[AWS API Gateway] 
  Endpoint: /prod/attendance
        │
        ▼
[Lambda 1: Attendance Handler] (attendancehandler.py)
  ├─ Multipart Parser & Image Extractor
  ├─ OS & Device Normalization ("Windows PC", "Linux", etc.)
  ├─ 4-Tier Location Cascade: GPS -> CloudFront Edge -> IP Geolocation -> Timezone
  ├─ Nominatim Reverse Geocoder & Administrative Cleaner
  └─ Amazon S3 Data Sink: `attendance-data`
        ├─ raw-text/{YYYY-MM-DD}/{Name}.txt
        ├─ records/{YYYY-MM-DD}/{Name}.json
        └─ images/{YYYY-MM-DD}/{Name}.jpg
        │
        │ S3 Object Created Event Trigger
        ▼
[Lambda 2: Presentation Generator] (lambda_function.py — Omega 5.0)
  ├─ PASS 1: Telemetry parsing, OpenCV face validation, coordinate clustering
  ├─ PASS 2: Session venue majority vote algorithm
  ├─ Cartographic Engine: Tile stitching, water polygon cleansing, HUD card
  ├─ 6 Executive Design Templates (Default: Swiss Precision Editorial White)
  ├─ Strict Zero-Coordinates Policy (lat/long never displayed raw)
  └─ Amazon S3 Report Sink: `final-ppt-report` -> reports/{slug}_{date}.pptx
```

---

### 2. Core Operational Pillars

#### A. Zero-Coordinates Policy
* **Requirement**: Raw GPS coordinates (latitude/longitude numerals) must **never** appear on slide textboxes, participant badges, tables, or charts.
* **Resolution**: All coordinates are clustered within ~1.1 km radius (2 decimal places) and reverse-geocoded to administrative `City, Country` labels.
* **Map Visual Anchor**: The only approved appearance of coordinates is within a stylized minimalist metadata pill inside the floating HUD card (`test_var_a`).

#### B. Platform Ecosystem Normalization
* **Unified Desktop Tier**: "Win32", "Windows", "Workstation", "Desktop", and "PC" are strictly consolidated into **`Windows PC`**.
* **Distinct Unix Tier**: **`Linux`** remains distinct from all other desktop categories.
* **Mobile & Apple Tiers**: `Android`, `iPhone / iOS`, and `MacBook / macOS` are preserved with distinct telemetry.

#### C. 4-Tier Location Cascade (Zero-Prompt Fallback)
1. **Tier 1 (Browser Geolocation API)**: Direct GPS coordinates if granted by user.
2. **Tier 2 (CloudFront Edge Headers)**: `cloudfront-viewer-latitude`, `cloudfront-viewer-longitude`, `cloudfront-viewer-city`. Zero browser permission prompt required.
3. **Tier 3 (Server-Side IP Geolocation)**: Automatic reverse lookup against IP geolocation services with in-memory caching.
4. **Tier 4 (Timezone Inference)**: Client timezone resolution (e.g., `Africa/Lagos` → `Lagos, Nigeria`).

---

### 3. Multi-Template Executive Themes

1. **`2_swiss` (Default)**: Swiss Precision Editorial (International Typographic Light). High-contrast white canvas, Georgia serif headers, black body, scarlet hairline rules.
2. **`5_prestige`**: Monochrome Prestige (Ultra-Luxury Minimalist Light).
3. **`1_obsidian`**: Obsidian Executive (Boardroom Stealth Dark).
4. **`3_cyber`**: Cyber Incident Command (Tactical War Room Dark).
5. **`4_linear`**: Linear Modern Tech (Silicon Valley Keynote).
6. **`6_flagship`**: Flagship Brand Crimson (Immersive Brand Glow).

---

### 4. Runtime Architecture & Zero-Billing Specifications
* **Lambda Runtimes**: Upgraded to **Python 3.13** across all services.
  * `lambda_function.py`: OCI Container Image using `public.ecr.aws/lambda/python:3.13` (AL2023 base) to bundle headless OpenCV and cartographic dependencies within ECR.
  * `attendancehandler.py`: Managed `python3.13` runtime (Amazon Linux 2023) with standard library `urllib` fallback.
  * `compile_attendance_excel.py`: Managed `python3.13` runtime with `openpyxl`.
* **Zero-Billing Guarantees**:
  * Function invocations are On-Demand, operating within the AWS Always-Free Tier (1M invocations / 3.2M compute seconds/month).
  * Direct internet & S3 Gateway VPC Endpoints are utilized — no billable VPC NAT Gateways are provisioned.
  * ECR repository retains `:latest` tag overwrites to prevent image storage accrual.

