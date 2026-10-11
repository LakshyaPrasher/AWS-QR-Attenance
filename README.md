# Intelligent Workshop Report Engine — Omega 5.0 (Executive Edition)

Event-driven AWS Lambda engine that automatically generates an executive-grade PowerPoint report whenever a participant photograph is uploaded to Amazon S3.

Designed with an international editorial design aesthetic, publication-grade cartographic vector mapping, guaranteed non-empty data states, and strict Zero-Coordinates enforcement.

---

## 1. Core Architecture (Model B — Live Deterministic Overwrite)

```
Participant → QR Attendance Page → Attendance Lambda
                                        ↓
                    S3: attendance-data-lakshyaprasher
                    ├── images/{date}/participant.jpg
                    └── raw-text/{date}/participant.txt
                                        ↓
                              S3 PUT Event Trigger
                                        ↓
                        PPT Generator Lambda (Omega 5.0)
                                        ↓
                    S3: final-ppt-report/reports/{slug}_{date}.pptx
```

Every participant check-in regenerates a single deterministic report file per date:
```
reports/airtel-zambia_2026-09-22.pptx  ← Always reflects the complete accumulated attendance
```

---

## 2. Key Enhancements in Omega 5.0

1. **White Background Executive Design (Default)**:
   - **Swiss Precision Editorial (`2_swiss`)**: Clean white/light canvas (`#F9F9FB` / `#FFFFFF`), Georgia serif typography, stark black type, scarlet hairline rules, and `01/02/03` section index numerals.
   - Includes 5 alternative templates selectable via `BRAND_TEMPLATE` or payload:
     - `5_prestige`: Monochrome Prestige (Ultra-Luxury Minimalist Light)
     - `1_obsidian`: Obsidian Executive (Boardroom Stealth Dark)
     - `3_cyber`: Cyber Incident Command (Tactical War Room Dark)
     - `4_linear`: Linear Modern Tech (Silicon Valley Keynote)
     - `6_flagship`: Airtel Flagship Crimson (Immersive Brand Glow)

2. **Aesthetic Cartographic Vector Map**:
   - Visualizes the **Top 1 Location Epicenter** with concentric radar rings and a pulsing beacon.
   - High-contrast floating glassmorphic HUD telemetry card displaying city, country, participant count, and hub share.
   - **100% Offline & Lambda Stable**: Zero external tile servers, zero Mapbox/Google API keys, zero network failure modes.

3. **Guaranteed Zero "No Data" States**:
   - **Device Parser**: Resolves `platform`, `rawPlatform`, and `userAgent` regex. Never outputs "Unknown".
   - **Location Resolver**: 3-tier resolution (existing metadata → OpenStreetMap administrative hierarchy → Zambian coordinate bounding boxes → timezone inference). Slide 4 (Map) and Slide 6 (Platform) are always fully populated.

4. **Strict Zero-Coordinates Policy**:
   - Latitude and longitude numbers are **never displayed** anywhere on slides or under photos.
   - Photos display strictly `City, Country` (e.g. `Lusaka, Zambia`).
   - The map HUD card identifies the epicenter cluster by administrative title and regional timezone.

5. **Integrated Dynamic Macro Geocoder (Standalone Lambda Deleted)**:
   - All logic from `validate-locations` is ported into `lambda_function.py`.
   - Reverse geocodes coordinates via OpenStreetMap Nominatim with rate-limiting and in-memory caching.
   - Automatically cleans administrative clutter (`Tehsil`, `District`, `Subdivision`, `Municipal Corporation`).
   - Writes enriched location data back to S3 `raw-text/{date}/{participant}.txt`.

---

## 3. 7-Slide Presentation Deck Architecture

| Slide | Title | Purpose & Visual Elements |
|---|---|---|
| **01** | **Cover Hero** | Client name in Georgia serif, workshop title, date badge, verified attendee count, telemetry sync status, and Model B live sync indicator. |
| **02** | **Telemetry & Permission Audit** | Dual cards with Camera and Location authorization donut charts, participant active counts, and verified location telemetry. |
| **03** | **Participant Directory** | High-res circular attendee photos/avatars with drop shadows, attendee names, and strictly `City, Country` subtitle. |
| **04** | **Geographic Operations & Epicenter Map** | Left: Regional cartographic vector map with concentric radar beacon. Right: Primary hub share KPI + Regional demographic distribution bar chart. |
| **05** | **Hardware & Display Recognition** | "Largest Display in the Room" trophy award, winner attendee name, viewport resolution with DPR multiplier (`@3x`), and platform pill. |
| **06** | **Operating System & Device Ecosystem** | Donut chart breaking down platform share across Android, iPhone / iOS, Windows PC, MacBook, and Linux. |
| **07** | **Executive Summary & Audit** | 4 primary KPI tiles (Total Attended, Photos Captured, Top Platform, Primary Hub) + 5-point data quality & telemetry verification audit. |

---

## 4. Environment Variables

| Variable | Default | Purpose |
|---|---|---|
| `BRAND_TEMPLATE` | `2_swiss` | Presentation template (`2_swiss`, `5_prestige`, `1_obsidian`, `3_cyber`, `4_linear`, `6_flagship`) |
| `BRAND_CLIENT` | `Airtel Zambia` | Client organization title |
| `BRAND_WORKSHOP` | `Cyber Crisis Simulation` | Simulation / event subtitle |
| `SOURCE_BUCKET` | `attendance-data-lakshyaprasher` | Source S3 bucket with images and raw-text |
| `REPORT_BUCKET` | `final-ppt-report` | Destination bucket for generated `.pptx` |
| `URL_EXPIRY` | `86400` | Presigned download URL expiry (24 hours) |
| `FACE_DETECTION_ENABLED` | `1` | OpenCV face detection toggle (`1` = on, `0` = off) |

---

## 5. Deployment

```bash
# Build container image
docker build -t ppt-generator .

# Authenticate & push to ECR
aws ecr get-login-password --region ap-south-1 | docker login --username AWS --password-stdin <account>.dkr.ecr.ap-south-1.amazonaws.com
docker tag ppt-generator:latest <account>.dkr.ecr.ap-south-1.amazonaws.com/ppt-generator:latest
docker push <account>.dkr.ecr.ap-south-1.amazonaws.com/ppt-generator:latest

# Update Lambda function code
aws lambda update-function-code \
  --function-name ppt-generator \
  --image-uri <account>.dkr.ecr.ap-south-1.amazonaws.com/ppt-generator:latest
```
