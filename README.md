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

## 5. Deployment & Python Upgrade (Zero-Billing AWS Procedure)

### Python Runtime Specifications
- **PPT Generator Lambda (`lambda_function.py`)**: AWS Lambda Container Image (`public.ecr.aws/lambda/python:3.13`).
- **Attendance Handler Lambda (`attendancehandler.py`)**: AWS Lambda Managed Runtime (`python3.13`).
- **Excel Compiler Lambda (`compile_attendance_excel.py`)**: AWS Lambda Managed Runtime (`python3.13`).

### Exact Zero-Cost Upgrade Steps

#### A. PPT Generator (ECR Container Lambda)
1. **Build Container with Python 3.13 base**:
   ```bash
   docker build -t ppt-generator:latest .
   ```
2. **Authenticate & Push to existing ECR repository (overwrite `:latest` tag to prevent storage accumulation)**:
   ```bash
   aws ecr get-login-password --region ap-south-1 | docker login --username AWS --password-stdin <account>.dkr.ecr.ap-south-1.amazonaws.com
   docker tag ppt-generator:latest <account>.dkr.ecr.ap-south-1.amazonaws.com/ppt-generator:latest
   docker push <account>.dkr.ecr.ap-south-1.amazonaws.com/ppt-generator:latest
   ```
3. **Update Lambda function code**:
   ```bash
   aws lambda update-function-code \
     --function-name ppt-generator \
     --image-uri <account>.dkr.ecr.ap-south-1.amazonaws.com/ppt-generator:latest
   ```

#### B. Attendance Ingestion & Excel Lambdas (Managed Runtime)
Update the function runtime in-place (instant, $0.00 cost):
```bash
aws lambda update-function-configuration \
  --function-name attendance-handler \
  --runtime python3.13

aws lambda update-function-configuration \
  --function-name compile-attendance-excel \
  --runtime python3.13
```

#### C. Zero-Billing Guarantees & Cost Safety
- **No NAT Gateways**: Lambda functions operate with direct internet or S3 Gateway Endpoints (100% free). Never attach a VPC NAT Gateway ($32.40/month baseline cost).
- **No Provisioned Concurrency**: Functions run on On-Demand concurrency covered under the AWS Free Tier (1,000,000 invocations & 3.2M compute seconds free/month).
- **ECR Storage Zero-Creep**: Pushing with `:latest` replaces the active tag. Untagged images can be cleaned with an ECR Lifecycle rule (`expire untagged images > 1 day`).

#### D. Running the Full Verification Test Suite
Run the master test runner from the repository root:
```bash
py run_all_tests.py
```
This executes all 6 verification test suites covering all 3 AWS Lambda functions:
1. `test_location_fallbacks.py`: Verifies `attendancehandler.py` 4-tier location cascade.
2. `test_platform_consolidation.py`: Verifies `lambda_function.py` platform consolidation and distinct Linux.
3. `test_final_verification.py`: Verifies `lambda_function.py` 7-slide deck generation and coordinate clustering.
4. `test_full_pipeline.py`: Verifies `lambda_function.py` full S3 event pipeline and Swiss white theme.
5. `test_nigeria_default.py`: Verifies `lambda_function.py` default config and Banana Island venue GPS.
6. `test_excel_compiler.py`: Verifies `compile_attendance_excel.py` S3 ingestion, aggregation, and Excel generation.


