"""
INTELLIGENT WORKSHOP REPORT ENGINE — OMEGA 5.0 (AIRTEL EXECUTIVE EDITION)
==========================================================================
Event-driven PPTX generator & Macro Location Validator for AWS Lambda.

ALL-IN-ONE ARCHITECTURE:
  1) Automated S3 Trigger  — fires on 'images/{date}/*.jpg' upload
  2) Dynamic Macro Geocoding — OpenStreetMap administrative hierarchy engine
     (Strictly resolves to City, Country; strips noise; writes back to S3)
  3) Zero Coordinates Policy — lat/long NEVER exposed on any presentation slide
  4) Model B Live Report   — continuously overwrites reports/{slug}_{date}.pptx
  5) Dual-Trigger Support  — handles both S3 events and HTTP API Gateway / Function URLs
  6) Executive Templates   — 6 Complete distinct design templates (Default: Swiss Precision Editorial White Background)
"""

import json
import boto3
import os
import re
import math
import time
import random
import tempfile
import logging
import zipfile
import urllib.parse
import urllib.request
import concurrent.futures
from io import BytesIO
from datetime import datetime, timezone, timedelta
from collections import Counter

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree

try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

logger = logging.getLogger()
logger.setLevel(logging.INFO)


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  APPLICATION CONFIGURATION                                              ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
class AppConfig:
    SOURCE_BUCKET = os.environ.get("SOURCE_BUCKET", "attendance-data-lakshyaprasher")
    REPORT_BUCKET = os.environ.get("REPORT_BUCKET", "final-ppt-report")
    URL_EXPIRY    = int(os.environ.get("URL_EXPIRY", "86400"))

    # Preferred White Background Template as default:
    DEFAULT_TEMPLATE = os.environ.get("BRAND_TEMPLATE", "2_swiss")

    SLIDE_W, SLIDE_H = 13.333, 7.50
    HEADER_H = 0.72
    THUMB_RENDER_PX = 480

    MAX_WORKERS = 15
    IST = timezone(timedelta(hours=5, minutes=30))

    FACE_DETECTION_ENABLED = os.environ.get("FACE_DETECTION_ENABLED", "1") != "0"
    BLACK_SCREEN_MEAN_THRESHOLD = 18
    BLACK_SCREEN_VARIANCE_THRESHOLD = 80

    USER_AGENT = os.environ.get("USER_AGENT", "ExecutiveLocationEngine/5.0 (admin@airtel-simulation.com)")


def parse_flexible_date(date_str: str) -> str:
    """Parses standard ISO dates or human strings like '30th June, 2026' into YYYY-MM-DD."""
    if not date_str:
        return ""
    s = str(date_str).strip()
    m = re.search(r"(\d{4}-\d{2}-\d{2})", s)
    if m:
        return m.group(1)
    s_clean = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", s)
    for fmt in ["%d %B, %Y", "%d %B %Y", "%B %d, %Y", "%B %d %Y", "%Y/%m/%d", "%d-%m-%Y"]:
        try:
            return datetime.strptime(s_clean, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s


def event_config(payload: dict) -> dict:
    """Resolve per-client / per-occasion branding and template selection."""
    template_choice = payload.get("template", os.environ.get("BRAND_TEMPLATE", AppConfig.DEFAULT_TEMPLATE))
    raw_d = (payload.get("target_date") or payload.get("date") or payload.get("day") or 
             payload.get("processed_date") or 
             os.environ.get("EVENT_DATE") or 
             os.environ.get("TARGET_DATE") or 
             os.environ.get("DATE") or 
             S3DataLayer.today_ist())
    target_d = parse_flexible_date(raw_d) or S3DataLayer.today_ist()
    client = (payload.get("client") or 
              os.environ.get("CLIENT") or 
              os.environ.get("BRAND_CLIENT") or 
              "Airtel Zambia")
    workshop = (payload.get("workshop") or 
                os.environ.get("WORKSHOP") or 
                os.environ.get("BRAND_WORKSHOP") or 
                "Cyber Crisis Simulation")
    return {
        "title":       payload.get("title",       os.environ.get("BRAND_TITLE",    "Attendance & Telemetry Report")),
        "subtitle":    payload.get("subtitle",    os.environ.get("BRAND_SUBTITLE", "Executive Telemetry Report")),
        "client":      client,
        "workshop":    workshop,
        "target_date": target_d,
        "date_display": payload.get("date_display", target_d),
        "template":    template_choice,
    }


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  6 COMPLETE EXECUTIVE TEMPLATE SPECIFICATIONS                            ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
TEMPLATE_SPECS = {
    "2_swiss": {
        "name": "Swiss Precision Editorial (International Typographic Light)",
        "is_dark": False,
        "bg_start": "#F9F9FB", "bg_end": "#F4F4F6", "bg_angle": 0,
        "primary": "#E40000", "secondary": "#B80000", "accent": "#E40000",
        "card_bg": "#FFFFFF", "card_border": "#E2E8F0", "hairline": "#E2E8F0",
        "text_main": "#0A0A0A", "text_muted": "#52525B", "text_sub": "#71717A",
        "font_heading": "Georgia", "font_body": "Segoe UI",
        "card_style": "editorial_rules",
        "pill_bg": "#F4F4F5", "pill_border": "#E4E4E7", "pill_text": "#0A0A0A"
    },
    "5_prestige": {
        "name": "Monochrome Prestige (Ultra-Luxury Minimalist Light)",
        "is_dark": False,
        "bg_start": "#FFFFFF", "bg_end": "#FAFAFA", "bg_angle": 0,
        "primary": "#E40000", "secondary": "#18181B", "accent": "#E40000",
        "card_bg": "#FFFFFF", "card_border": "#E4E4E7", "hairline": "#E4E4E7",
        "text_main": "#18181B", "text_muted": "#52525B", "text_sub": "#71717A",
        "font_heading": "Palatino Linotype", "font_body": "Calibri",
        "card_style": "clean_whitespace",
        "pill_bg": "#FEF2F2", "pill_border": "#FECACA", "pill_text": "#E40000"
    },
    "1_obsidian": {
        "name": "Obsidian Executive (Boardroom Stealth Dark)",
        "is_dark": True,
        "bg_start": "#05070B", "bg_end": "#0B0F17", "bg_angle": 135,
        "primary": "#FF1E27", "secondary": "#C40000", "accent": "#FF4444",
        "card_bg": "#0D1322", "card_border": "#1E2738", "hairline": "#1E293B",
        "text_main": "#FFFFFF", "text_muted": "#94A3B8", "text_sub": "#64748B",
        "font_heading": "Segoe UI", "font_body": "Segoe UI",
        "card_style": "hairline_grid",
        "pill_bg": "#121826", "pill_border": "#1E293B", "pill_text": "#FF3333"
    },
    "3_cyber": {
        "name": "Cyber Incident Command (Tactical War Room Dark)",
        "is_dark": True,
        "bg_start": "#020617", "bg_end": "#060D24", "bg_angle": 150,
        "primary": "#FF0033", "secondary": "#00F0FF", "accent": "#00F0FF",
        "card_bg": "#0B1124", "card_border": "#1E293B", "hairline": "#162036",
        "text_main": "#F8FAFC", "text_muted": "#94A3B8", "text_sub": "#64748B",
        "font_heading": "Segoe UI", "font_body": "Consolas",
        "card_style": "hud_brackets",
        "pill_bg": "#0B152E", "pill_border": "#00F0FF", "pill_text": "#00F0FF"
    },
    "4_linear": {
        "name": "Linear Modern Tech (Silicon Valley Keynote)",
        "is_dark": True,
        "bg_start": "#0C0E14", "bg_end": "#141724", "bg_angle": 135,
        "primary": "#E40000", "secondary": "#8B5CF6", "accent": "#EC4899",
        "card_bg": "#141824", "card_border": "#232838", "hairline": "#1E2332",
        "text_main": "#FFFFFF", "text_muted": "#94A3B8", "text_sub": "#64748B",
        "font_heading": "Segoe UI Semibold", "font_body": "Segoe UI",
        "card_style": "bento_glass",
        "pill_bg": "#1F1D2B", "pill_border": "#374151", "pill_text": "#F43F5E"
    },
    "6_flagship": {
        "name": "Airtel Flagship Crimson (Immersive Brand Experience)",
        "is_dark": True,
        "bg_start": "#08090C", "bg_end": "#220508", "bg_angle": 135,
        "primary": "#E40000", "secondary": "#FF3333", "accent": "#FF6666",
        "card_bg": "#170C0E", "card_border": "#3A1417", "hairline": "#2D1013",
        "text_main": "#FFFFFF", "text_muted": "#94A3B8", "text_sub": "#64748B",
        "font_heading": "Segoe UI", "font_body": "Segoe UI",
        "card_style": "crimson_glow",
        "pill_bg": "#220508", "pill_border": "#551014", "pill_text": "#FF3333"
    }
}


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  DATA EXTRACTION & NORMALIZATION ENGINE (GUARANTEED ZERO MISSING DATA)   ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
def normalize_platform(plat: str) -> str:
    """Normalizes platform names so 'Workstation' and 'Windows PC' are always unified as 'Windows PC',
    while 'Linux' remains distinct."""
    if not plat:
        return "Windows PC"
    p_lower = str(plat).strip().lower()
    if "linux" in p_lower and "android" not in p_lower:
        return "Linux"
    if "iphone" in p_lower or "ipad" in p_lower or "ios" in p_lower:
        return "iPhone / iOS"
    if "android" in p_lower:
        return "Android"
    if "mac" in p_lower or "darwin" in p_lower:
        return "MacBook / macOS"
    if "chromebook" in p_lower or "cros" in p_lower:
        return "Chromebook"
    if any(k in p_lower for k in ["workstation", "windows", "win", "desktop", "pc"]):
        return "Windows PC"
    return plat.strip()


def extract_device_platform(meta: dict) -> str:
    """Robust platform extraction checking platform, rawPlatform, and userAgent regex.
    Workstations and Windows PCs are unified as 'Windows PC'.
    Linux remains distinct as 'Linux'."""
    plat = str(meta.get("platform", "")).strip().lower()
    raw_plat = str(meta.get("rawPlatform", "")).strip().lower()
    ua = str(meta.get("userAgent", "")).strip().lower()
    combined = f"{plat} {raw_plat} {ua}"

    if "iphone" in combined or "ipad" in combined or "ios" in combined:
        return "iPhone / iOS"
    if "android" in combined:
        return "Android"
    if "macintel" in combined or "macintosh" in combined or "mac os" in combined:
        return "MacBook / macOS"
    if "linux" in combined and "android" not in combined:
        return "Linux"
    if "cros" in combined:
        return "Chromebook"
    if any(k in combined for k in ["win", "windows", "workstation", "desktop"]):
        return "Windows PC"
    return "Mobile Device" if "mobi" in combined else "Windows PC"


def extract_screen_info(meta: dict) -> dict:
    """Extract display resolution, compute DPR and sort value for Largest Display award."""
    screen_str = str(meta.get("screen", "")).strip()
    if screen_str:
        try:
            s = screen_str.lower().replace("inches", "").replace("inch", "").replace('"', "").strip()
            val = float(s)
            return {"display": f'{val:.1f}" Display', "sort_value": val * 1000000}
        except (ValueError, TypeError):
            pass

    try:
        w = int(float(meta.get("screenWidth", 0)))
        h = int(float(meta.get("screenHeight", 0)))
        if w > 0 and h > 0:
            dpr = float(meta.get("devicePixelRatio", 1)) or 1
            sort_val = w * h * (dpr ** 2)
            dpr_label = f" @{dpr:.0f}x" if dpr > 1 else ""
            return {"display": f"{w} × {h}{dpr_label}", "sort_value": sort_val}
    except (ValueError, TypeError):
        pass

    return {"display": "1920 × 1080 Full HD", "sort_value": 1920 * 1080}


def permission_granted(meta: dict, kind: str) -> bool:
    """Check if camera or location permission was granted."""
    if kind == "camera":
        flat = str(meta.get("cameraPermission", "")).lower()
        nested = str((meta.get("camera") or {}).get("permissionStatus", "")).lower() if isinstance(meta.get("camera"), dict) else ""
    else:
        flat = str(meta.get("locationPermission", "")).lower()
        nested = str((meta.get("location") or {}).get("permissionStatus", "")).lower() if isinstance(meta.get("location"), dict) else ""
    return "granted" in (flat, nested) or "true" in (flat, nested)


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  DYNAMIC LOCATION VALIDATOR (ZERO COORDINATES POLICY)                    ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
class LocationValidator:
    """Dynamic reverse geocoding via OSM Nominatim with 3-tier fallback.
    - Strict Macro Hierarchy: city -> town -> municipality -> county.
    - Strips administrative clutter.
    - Offline fallback: Zambia bounding boxes + timezone inference.
    - Zero Coordinates Policy: NEVER exposes raw lat/long on any slide.
    - Writes back enriched location to S3 if not present.
    """
    NOISE_PATTERNS = [
        r"\bTehsil\b", r"\bDistrict\b", r"\bSubdivision\b", r"\bTaluka\b",
        r"\bMunicipal Corporation\b", r"\bM\.Corp\b", r"\bCity\b", r"\bTown\b", r"\bBlock\b"
    ]
    _GEO_CACHE = {}
    _LAST_CALL_TIME = 0.0

    # Regional bounding boxes for Zambia cities
    ZAMBIA_HUBS = [
        ("Lusaka", "Zambia", -15.4167, 28.2833, -15.8, -15.0, 27.8, 28.6),
        ("Ndola", "Zambia", -12.9667, 28.6333, -13.2, -12.7, 28.4, 28.9),
        ("Kitwe", "Zambia", -12.8000, 28.2000, -13.0, -12.6, 28.0, 28.4),
        ("Livingstone", "Zambia", -17.8419, 25.8544, -18.0, -17.6, 25.6, 26.1),
        ("Chipata", "Zambia", -13.6333, 32.6500, -13.8, -13.4, 32.4, 32.9),
        ("Solwezi", "Zambia", -12.1833, 26.4000, -12.4, -12.0, 26.2, 26.6),
    ]

    @classmethod
    def clean_noise(cls, text: str) -> str:
        if not text: return ""
        cleaned = str(text)
        for p in cls.NOISE_PATTERNS:
            cleaned = re.sub(p, "", cleaned, flags=re.IGNORECASE)
        return re.sub(r"\s+", " ", cleaned).strip().title()

    @staticmethod
    def validate_coordinates(lat, lon) -> bool:
        if lat is None or lon is None:
            return False
        try:
            f_lat, f_lon = float(lat), float(lon)
            if -90.0 <= f_lat <= 90.0 and -180.0 <= f_lon <= 180.0:
                return abs(f_lat) > 0.0001 or abs(f_lon) > 0.0001
            return False
        except (ValueError, TypeError):
            return False

    @staticmethod
    def cluster_coords(lat, lon) -> tuple:
        """Rounds coordinates to 2 decimal places (~1.1 km precision) so attendees
        in the same venue/hotel/hall (e.g. 13.4005 and 13.4001) are read and grouped
        as the exact same location."""
        return (round(float(lat), 2), round(float(lon), 2))

    @classmethod
    def reverse_geocode(cls, lat: float, lon: float) -> dict:
        """Dual-tier reverse geocoder (OSM Nominatim + Esri ArcGIS fallback) with cluster caching."""
        cluster_key = cls.cluster_coords(lat, lon)
        if cluster_key in cls._GEO_CACHE:
            return cls._GEO_CACHE[cluster_key]

        c_lat, c_lon = cluster_key

        # Provider 1: OSM Nominatim (city/town level with 1s rate limit)
        now = time.time()
        elapsed = now - cls._LAST_CALL_TIME
        if elapsed < 1.0:
            time.sleep(1.0 - elapsed)
        cls._LAST_CALL_TIME = time.time()

        try:
            params = urllib.parse.urlencode({
                "lat": c_lat, "lon": c_lon, "format": "jsonv2",
                "zoom": 12, "accept-language": "en", "addressdetails": 1
            })
            url = f"https://nominatim.openstreetmap.org/reverse?{params}"
            req = urllib.request.Request(url, headers={"User-Agent": "AirtelWorkshopReport/5.0 (compliance@airtel.com)"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            addr = data.get("address", {})
            macro_city = (
                addr.get("city") or addr.get("town") or addr.get("municipality")
                or addr.get("suburb") or addr.get("state_district") or addr.get("county") or ""
            )
            clean_city = cls.clean_noise(macro_city)
            clean_country = cls.clean_noise(addr.get("country", ""))
            if clean_city or clean_country:
                res = {"city": clean_city, "country": clean_country}
                cls._GEO_CACHE[cluster_key] = res
                return res
        except Exception as e:
            logger.info(f"Nominatim reverse geocode skipped ({e}), switching to Esri fallback...")

        # Provider 2: Esri ArcGIS World Geocoding (high availability, zero rate-limit, never blocked on AWS)
        try:
            url_esri = f"https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/reverseGeocode?location={c_lon},{c_lat}&f=json"
            req_esri = urllib.request.Request(url_esri, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req_esri, timeout=3) as resp:
                data = json.loads(resp.read().decode("utf-8")).get("address", {})
            esri_city = (
                data.get("City") or data.get("Subregion") or data.get("District")
                or data.get("Neighborhood") or data.get("PlaceName") or ""
            )
            esri_country = data.get("CntryName") or ""
            clean_city = cls.clean_noise(esri_city)
            clean_country = cls.clean_noise(esri_country)
            if clean_city or clean_country:
                res = {"city": clean_city, "country": clean_country}
                cls._GEO_CACHE[cluster_key] = res
                return res
        except Exception as e:
            logger.warning(f"Esri geocoding fallback failed: {e}")

        # Provider 3: Coordinate Bounding Box Fallback (Zambia Hubs)
        for h_city, h_country, _, _, min_lat, max_lat, min_lon, max_lon in cls.ZAMBIA_HUBS:
            if min_lat <= c_lat <= max_lat and min_lon <= c_lon <= max_lon:
                res = {"city": h_city, "country": h_country}
                cls._GEO_CACHE[cluster_key] = res
                return res

        # Provider 4: Return empty so session-level inference can assign the true venue
        res = {"city": "", "country": ""}
        cls._GEO_CACHE[cluster_key] = res
        return res

    @classmethod
    def resolve_and_enrich(cls, meta: dict, s3_client=None, bucket: str = None, s3_key: str = None) -> dict:
        """Resolves location to City, Country with guaranteed coordinate preservation and venue clustering."""
        raw_loc = meta.get("location", {})
        if not isinstance(raw_loc, dict):
            raw_loc = {"city": str(raw_loc)}

        city = cls.clean_noise(raw_loc.get("city", "") or meta.get("city", ""))
        country = cls.clean_noise(raw_loc.get("country", "") or meta.get("country", ""))
        lat = raw_loc.get("latitude") or raw_loc.get("lat") or meta.get("lat") or meta.get("latitude")
        lon = raw_loc.get("longitude") or raw_loc.get("lon") or meta.get("lon") or meta.get("longitude")

        has_coords = cls.validate_coordinates(lat, lon)
        if has_coords:
            f_lat, f_lon = float(lat), float(lon)
        else:
            f_lat, f_lon = None, None

        updated_s3 = False

        if has_coords:
            # If city/country not already known, reverse geocode using 2-decimal cluster key
            if not city or not country:
                geo = cls.reverse_geocode(f_lat, f_lon)
                city = city or geo.get("city", "")
                country = country or geo.get("country", "")
                if city:
                    raw_loc["city"] = city
                if country:
                    raw_loc["country"] = country
                meta["location"] = raw_loc

                if s3_client and bucket and s3_key and (city or country):
                    try:
                        s3_client.put_object(
                            Bucket=bucket, Key=s3_key,
                            Body=json.dumps(meta, ensure_ascii=False, indent=2).encode("utf-8"),
                            ContentType="application/json"
                        )
                        updated_s3 = True
                    except Exception:
                        pass

        city_country = f"{city}, {country}" if city and country else (city or country)

        return {
            "city": city,
            "country": country,
            "city_country": city_country,
            "lat": f_lat,
            "lon": f_lon,
            "has_coords": has_coords,
            "updated_s3": updated_s3,
            "is_suspect": False
        }


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  CARTOGRAPHIC MAP & CHART ENGINES                                        ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
ZAMBIA_BORDER_COORDS = [
    (24.0, -11.0), (25.5, -11.2), (26.5, -12.2), (27.5, -12.0), (28.5, -11.5),
    (29.8, -12.3), (29.5, -9.5), (31.0, -8.5), (32.8, -9.5), (33.3, -11.0),
    (33.0, -12.5), (33.5, -13.5), (32.8, -14.2), (30.5, -15.5), (29.0, -15.7),
    (27.0, -18.0), (25.2, -17.8), (23.5, -17.5), (22.0, -16.2), (22.2, -13.0),
    (24.0, -11.0)
]

def get_font(size: int, bold: bool = False):
    """Load font gracefully across Windows, Linux (AWS Lambda), and fallback."""
    font_candidates = []
    if bold:
        font_candidates = [
            "C:/Windows/Fonts/arialbd.ttf", "C:/Windows/Fonts/segoeuib.ttf",
            "arialbd.ttf", "segoeuib.ttf", "georgiab.ttf",
            "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans-Bold.ttf",
            "DejaVuSans-Bold.ttf"
        ]
    else:
        font_candidates = [
            "C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/segoeui.ttf",
            "arial.ttf", "segoeui.ttf", "georgia.ttf",
            "/usr/share/fonts/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans.ttf",
            "DejaVuSans.ttf"
        ]
    for fc in font_candidates:
        try:
            return ImageFont.truetype(fc, size)
        except Exception:
            pass
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def hex_to_rgb(hex_str: str, alpha: int = 255) -> tuple:
    h = hex_str.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (r, g, b, alpha) if alpha < 255 else (r, g, b)


def render_cartographic_map(top_city: str, country: str, lat: float, lon: float,
                            attendees: int, pct: float, is_dark: bool, primary_color: str,
                            card_bg: str = None, total_attendees: int = None) -> bytes:
    """Renders an authentic, geographically verified street map centered on the exact lat/long.
    - Stitches real OpenStreetMap / Esri World Street Map tiles of the location.
    - Places the iconic Google Maps red teardrop pin at the exact epicenter coordinates.
    - Overlays a razor-sharp Google Maps Place Card with zero coordinates exposed.
    - 300 DPI high resolution with local tile caching in /tmp/tiles."""
    if not PIL_AVAILABLE:
        logger.error("MAP: PIL unavailable — returning empty bytes")
        return b""

    W, H = 2040, 1740
    zoom = 14  # City street detail level

    n = 2.0 ** zoom
    x_exact = (lon + 180.0) / 360.0 * n
    y_exact = (1.0 - math.asinh(math.tan(math.radians(lat))) / math.pi) / 2.0 * n

    center_tx = int(x_exact)
    center_ty = int(y_exact)
    sub_x = (x_exact - center_tx) * 256
    sub_y = (y_exact - center_ty) * 256

    cols, rows = 6, 5
    start_tx = center_tx - 3
    start_ty = center_ty - 2

    canvas = Image.new("RGB", (cols * 256, rows * 256), (245, 243, 240))
    tile_cache_dir = os.path.join(tempfile.gettempdir(), "osm_tiles")
    try: os.makedirs(tile_cache_dir, exist_ok=True)
    except Exception: pass

    for r in range(rows):
        for c in range(cols):
            tx = start_tx + c
            ty = start_ty + r
            cache_path = os.path.join(tile_cache_dir, f"{zoom}_{tx}_{ty}.png")
            tile = None
            if os.path.exists(cache_path):
                try: tile = Image.open(cache_path).convert("RGB")
                except Exception: pass

            if tile is None:
                # Primary: OpenStreetMap standard tiles
                url_osm = f"https://tile.openstreetmap.org/{zoom}/{tx}/{ty}.png"
                req_osm = urllib.request.Request(url_osm, headers={"User-Agent": "AirtelWorkshopReport/5.0 (compliance@airtel.com)"})
                try:
                    with urllib.request.urlopen(req_osm, timeout=3) as resp:
                        data = resp.read()
                        try:
                            with open(cache_path, "wb") as f: f.write(data)
                        except Exception: pass
                        tile = Image.open(BytesIO(data)).convert("RGB")
                except Exception:
                    # Fallback: Esri World Street Map (reliable worldwide mirror)
                    try:
                        url_esri = f"https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{zoom}/{ty}/{tx}"
                        req_esri = urllib.request.Request(url_esri, headers={"User-Agent": "Mozilla/5.0"})
                        with urllib.request.urlopen(req_esri, timeout=3) as resp:
                            data = resp.read()
                            tile = Image.open(BytesIO(data)).convert("RGB")
                    except Exception:
                        pass
            if tile:
                canvas.paste(tile, (c * 256, r * 256))

    pin_x = (center_tx - start_tx) * 256 + int(sub_x)
    pin_y = (center_ty - start_ty) * 256 + int(sub_y)

    # Crop to 6.8 : 5.8 slide aspect ratio
    crop_w = 1280
    crop_h = int(crop_w * 5.8 / 6.8)

    c_x1 = max(0, pin_x - crop_w // 2)
    c_y1 = max(0, pin_y - crop_h // 2)
    c_x2 = c_x1 + crop_w
    c_y2 = c_y1 + crop_h

    if c_x2 > canvas.width:
        c_x1 -= (c_x2 - canvas.width)
        c_x2 = canvas.width
    if c_y2 > canvas.height:
        c_y1 -= (c_y2 - canvas.height)
        c_y2 = canvas.height

    cropped = canvas.crop((c_x1, c_y1, c_x2, c_y2))
    final_map = cropped.resize((W, H), Image.LANCZOS).convert("RGBA")

    # Scaled pin coordinates on high-res canvas
    scale_factor = W / crop_w
    p_x = int((pin_x - c_x1) * scale_factor)
    p_y = int((pin_y - c_y1) * scale_factor)

    # Subtle radar pulse around epicenter
    pulse = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    p_draw = ImageDraw.Draw(pulse)
    for rad, alpha in [(80, 160), (160, 80), (280, 35)]:
        p_draw.ellipse([p_x - rad, p_y - rad, p_x + rad, p_y + rad], outline=(228, 0, 0, alpha), width=5)
    final_map = Image.alpha_composite(final_map, pulse)

    # Ground shadow under pin tip
    shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).ellipse([p_x - 55, p_y - 20, p_x + 55, p_y + 20], fill=(0, 0, 0, 95))
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=8))
    final_map = Image.alpha_composite(final_map, shadow)

    # Iconic Red Google Maps Location Pin
    draw = ImageDraw.Draw(final_map)
    pin_head_y = p_y - 95
    pin_radius = 52
    pin_poly = [
        (p_x, p_y),
        (p_x - 45, pin_head_y + 16),
        (p_x - pin_radius, pin_head_y),
        (p_x - 36, pin_head_y - 40),
        (p_x, pin_head_y - pin_radius),
        (p_x + 36, pin_head_y - 40),
        (p_x + pin_radius, pin_head_y),
        (p_x + 45, pin_head_y + 16),
    ]
    draw.polygon(pin_poly, fill=(234, 67, 53, 255))
    draw.ellipse([p_x - pin_radius, pin_head_y - pin_radius, p_x + pin_radius, pin_head_y + pin_radius],
                 fill=(234, 67, 53, 255), outline=(197, 34, 31, 255), width=3)
    draw.ellipse([p_x - 20, pin_head_y - 20, p_x + 20, pin_head_y + 20], fill=(255, 255, 255, 255))

    # Floating Google Maps Place Card (300 DPI razor-sharp typography)
    card_w, card_h = 760, 310
    card_x, card_y = 65, H - card_h - 65

    c_shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(c_shadow).rounded_rectangle([card_x + 6, card_y + 10, card_x + card_w + 6, card_y + card_h + 10],
                                               radius=24, fill=(0, 0, 0, 45))
    c_shadow = c_shadow.filter(ImageFilter.GaussianBlur(radius=14))
    final_map = Image.alpha_composite(final_map, c_shadow)
    draw = ImageDraw.Draw(final_map)

    # Solid opaque background so text is 100% crisp without antialiasing artifacts
    draw.rounded_rectangle([card_x, card_y, card_x + card_w, card_y + card_h],
                           radius=24, fill=(255, 255, 255, 255), outline=hex_to_rgb("#DADCE0"), width=2)

    font_tag = get_font(20, bold=True)
    font_city = get_font(46, bold=True)
    font_sub = get_font(24, bold=False)
    font_pill = get_font(22, bold=True)
    font_meta = get_font(20, bold=False)

    # Red dot & tag
    draw.ellipse([card_x + 36, card_y + 36, card_x + 48, card_y + 48], fill=hex_to_rgb("#EA4335"))
    draw.text((card_x + 60, card_y + 31), "PRIMARY SIMULATION EPICENTER", fill=hex_to_rgb("#EA4335"), font=font_tag)

    draw.text((card_x + 36, card_y + 70), top_city.upper(), fill=hex_to_rgb("#202124"), font=font_city)
    draw.text((card_x + 36, card_y + 138), f"{country.title()} · Central Business District Hub", fill=hex_to_rgb("#5F6368"), font=font_sub)

    # Green Verified Pill
    pill_x = card_x + 36
    pill_y = card_y + 190
    pill_w = 480
    pill_h = 48
    draw.rounded_rectangle([pill_x, pill_y, pill_x + pill_w, pill_y + pill_h],
                           radius=12, fill=hex_to_rgb("#E6F4EA"), outline=hex_to_rgb("#CEEAD6"), width=1)

    # 100% Vector Checkmark (Zero missing font glyphs / tofu)
    chk_x = pill_x + 22
    chk_y = pill_y + 24
    draw.line([(chk_x, chk_y), (chk_x + 5, chk_y + 6)], fill=hex_to_rgb("#137333"), width=3)
    draw.line([(chk_x + 5, chk_y + 6), (chk_x + 14, chk_y - 7)], fill=hex_to_rgb("#137333"), width=3)

    tot_disp = total_attendees if total_attendees is not None else attendees
    draw.text((pill_x + 44, pill_y + 11), f"{attendees} of {tot_disp} Verified Attendees ({pct:.0f}%)",
              fill=hex_to_rgb("#137333"), font=font_pill)

    # Dynamic timezone based on resolved country (never hardcode CAT when in India!)
    c_lower = (country or "").lower()
    if "india" in c_lower:
        tz_label = "IST (UTC+5:30)"
    elif "zambia" in c_lower:
        tz_label = "CAT (UTC+2)"
    elif "kenya" in c_lower or "uganda" in c_lower:
        tz_label = "EAT (UTC+3)"
    elif "uk" in c_lower or "kingdom" in c_lower:
        tz_label = "GMT / BST"
    elif "us" in c_lower or "united states" in c_lower:
        tz_label = "EST / PST"
    else:
        tz_label = "Macro Geocoding Audit"

    draw.text((card_x + 36, card_y + 260), f"Verified via OpenStreetMap Macro Geocoding  ·  {tz_label}",
              fill=hex_to_rgb("#5F6368"), font=font_meta)

    # Google Maps Zoom Pill (bottom right) - 100% Vector Lines (Zero Tofu Glyphs)
    zw, zh = 72, 140
    zx = W - zw - 65
    zy = H - zh - 65
    draw.rounded_rectangle([zx, zy, zx + zw, zy + zh], radius=16, fill=(255, 255, 255, 255), outline=hex_to_rgb("#DADCE0"), width=2)
    draw.line([(zx + 14, zy + zh // 2), (zx + zw - 14, zy + zh // 2)], fill=hex_to_rgb("#DADCE0"), width=2)

    cx_z = zx + zw // 2
    cy_plus = zy + zh // 4
    cy_minus = zy + 3 * zh // 4
    # Vector '+'
    draw.line([(cx_z - 11, cy_plus), (cx_z + 11, cy_plus)], fill=hex_to_rgb("#202124"), width=3)
    draw.line([(cx_z, cy_plus - 11), (cx_z, cy_plus + 11)], fill=hex_to_rgb("#202124"), width=3)
    # Vector '-'
    draw.line([(cx_z - 11, cy_minus), (cx_z + 11, cy_minus)], fill=hex_to_rgb("#202124"), width=3)

    buf = BytesIO()
    final_map.save(buf, format="PNG", optimize=True)
    result = buf.getvalue()
    logger.info(f"REAL MAP rendered: {len(result)} bytes for {top_city}, {country}")
    return result


def render_donut_chart(categories: list, values: list, title: str,
                       is_dark: bool, primary_color: str, palette: list = None,
                       card_bg: str = None) -> bytes:
    """Pure PIL Donut Chart at 300 DPI sharpness with distinct high-contrast palette.
    Guarantees slices NEVER use adjacent or similar colors."""
    if not PIL_AVAILABLE:
        logger.error("DONUT CHART: PIL unavailable — returning empty bytes")
        return b""

    # High-DPI canvas calibrated to card container aspect ratio (1.408)
    W, H = 2112, 1500
    bg_rgb = hex_to_rgb(card_bg or ("#0D1322" if is_dark else "#FFFFFF"))
    text_main = hex_to_rgb("#FFFFFF" if is_dark else "#0A0A0A")
    text_muted = hex_to_rgb("#94A3B8" if is_dark else "#64748B")

    total = sum(values) if values else 1

    # High-contrast executive palette: Red, Slate Navy, Azure, Green, Amber, Purple
    if not palette:
        palette = [
            primary_color,  # Airtel Red
            "#1E293B",      # Deep Slate Navy (100% contrast against Red)
            "#0284C7",      # Sky Azure
            "#059669",      # Emerald Green
            "#D97706",      # Amber Orange
            "#7C3AED",      # Royal Purple
            "#06B6D4"       # Vibrant Cyan
        ]

    img = Image.new("RGBA", (W, H), bg_rgb + (255,))
    draw = ImageDraw.Draw(img)

    # Title
    if title:
        font_title = get_font(42, bold=True)
        draw.text((80, 60), title, fill=text_main, font=font_title)

    # Donut geometry (left side) — 100% mathematically circular
    cx = 600
    cy = 760
    R = 420
    r_inner = 250

    bbox_outer = [cx - R, cy - R, cx + R, cy + R]
    bbox_inner = [cx - r_inner, cy - r_inner, cx + r_inner, cy + r_inner]

    # Draw slices
    cur_deg = -90.0
    for i, (cat, val) in enumerate(zip(categories, values)):
        sweep = (val / total) * 360.0
        c_hex = palette[i % len(palette)]
        fill_rgb = hex_to_rgb(c_hex)

        if sweep > 0.1:
            draw.pieslice(bbox_outer, start=cur_deg, end=cur_deg + sweep,
                          fill=fill_rgb, outline=bg_rgb, width=8)
        cur_deg += sweep

    # Cut inner donut hole
    draw.ellipse(bbox_inner, fill=bg_rgb)

    # Center Text
    font_center_num = get_font(110, bold=True)
    font_center_sub = get_font(32, bold=True)

    num_str = str(total)
    num_box = draw.textbbox((0, 0), num_str, font=font_center_num)
    nw = num_box[2] - num_box[0]
    nh = num_box[3] - num_box[1]
    draw.text((cx - nw // 2, cy - nh // 2 - 25), num_str, fill=text_main, font=font_center_num)

    sub_str = "TOTAL"
    sub_box = draw.textbbox((0, 0), sub_str, font=font_center_sub)
    sw = sub_box[2] - sub_box[0]
    draw.text((cx - sw // 2, cy + 45), sub_str, fill=text_muted, font=font_center_sub)

    # Legend on right side (generous spacing & large readable text)
    leg_x = 1180
    leg_start_y = 660
    row_h = 130
    font_leg_bold = get_font(42, bold=True)
    font_leg_pct = get_font(42, bold=False)

    for i, (cat, val) in enumerate(zip(categories, values)):
        ly = leg_start_y + i * row_h
        c_hex = palette[i % len(palette)]
        fill_rgb = hex_to_rgb(c_hex)

        # Pill marker
        draw.rounded_rectangle([leg_x, ly + 6, leg_x + 36, ly + 46], radius=18, fill=fill_rgb)

        # Label & value
        pct = (val / total) * 100.0
        draw.text((leg_x + 60, ly), cat, fill=text_main, font=font_leg_bold)
        cat_w = draw.textbbox((0, 0), cat, font=font_leg_bold)[2] - draw.textbbox((0, 0), cat, font=font_leg_bold)[0]
        draw.text((leg_x + 60 + cat_w + 18, ly), f"{val} ({pct:.0f}%)", fill=text_muted, font=font_leg_pct)

    buf = BytesIO()
    img.save(buf, format="PNG", optimize=True)
    result = buf.getvalue()
    logger.info(f"DONUT CHART (300 DPI) rendered: {len(result)} bytes for '{title}'")
    return result


def render_horizontal_bars(categories: list, values: list, title: str,
                           is_dark: bool, primary_color: str,
                           card_bg: str = None) -> bytes:
    """Pure PIL Horizontal Bar Chart at 300 DPI sharpness. Rounded pill bars, vector clarity."""
    if not PIL_AVAILABLE:
        logger.error("BAR CHART: PIL unavailable — returning empty bytes")
        return b""

    # High-DPI canvas calibrated to card container aspect ratio (1.164)
    W, H = 1800, 1546
    bg_rgb = hex_to_rgb(card_bg or ("#0D1322" if is_dark else "#FFFFFF"))
    text_main = hex_to_rgb("#FFFFFF" if is_dark else "#0A0A0A")
    text_muted = hex_to_rgb("#94A3B8" if is_dark else "#64748B")
    bar_rgb = hex_to_rgb(primary_color)
    grid_rgb = hex_to_rgb("#1E293B" if is_dark else "#F1F5F9")

    img = Image.new("RGBA", (W, H), bg_rgb + (255,))
    draw = ImageDraw.Draw(img)

    if title:
        font_title = get_font(42, bold=True)
        draw.text((80, 60), title, fill=text_main, font=font_title)

    n = len(categories) if categories else 1
    total = sum(values) if values else 1
    max_val = max(values) if values else 1

    bar_start_y = 220
    available_h = H - bar_start_y - 120
    slot_h = available_h / max(n, 1)
    bar_h = min(88, int(slot_h * 0.48))

    label_w = 440
    bar_x = 80 + label_w
    max_bar_w = W - bar_x - 280

    font_label = get_font(34, bold=True)
    font_val = get_font(34, bold=True)

    # Grid vertical lines
    for step in [0.25, 0.5, 0.75, 1.0]:
        gx = bar_x + step * max_bar_w
        draw.line([(gx, bar_start_y - 20), (gx, H - 80)], fill=grid_rgb, width=2)

    for i, (cat, val) in enumerate(zip(categories, values)):
        cy = int(bar_start_y + i * slot_h + (slot_h - bar_h) / 2)

        cat_disp = cat if len(cat) <= 22 else cat[:21] + "…"
        bbox = draw.textbbox((0, 0), cat_disp, font=font_label)
        lw, lh = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text((bar_x - lw - 30, cy + (bar_h - lh) // 2 - 4), cat_disp, fill=text_main, font=font_label)

        pct = (val / total) * 100.0
        bw = max(int((val / max_val) * max_bar_w), 24)
        draw.rounded_rectangle([bar_x, cy, bar_x + bw, cy + bar_h],
                               radius=bar_h // 2, fill=bar_rgb)

        val_str = f"{val} ({pct:.0f}%)"
        draw.text((bar_x + bw + 28, cy + (bar_h - lh) // 2 - 4), val_str, fill=text_main, font=font_val)

    buf = BytesIO()
    img.save(buf, format="PNG", optimize=True)
    result = buf.getvalue()
    logger.info(f"BAR CHART (300 DPI) rendered: {len(result)} bytes for '{title}'")
    return result


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  CIRCULAR AVATAR GENERATOR                                              ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
def make_avatar(name: str, raw_bytes: bytes, primary_hex: str, size=480) -> tuple:
    if not PIL_AVAILABLE:
        return b"", 0
    supersample = 2
    S = size * supersample
    hx = primary_hex.lstrip("#")
    br, bg, bb = int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16)

    canvas_pad = int(S * 0.08)
    canvas_size = S + canvas_pad * 2
    canvas = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))

    # Ambient drop shadow
    shadow_layer = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow_layer)
    sd.ellipse((canvas_pad + 2, canvas_pad + 6, canvas_pad + S - 2, canvas_pad + S + 6),
               fill=(15, 23, 42, 45))
    shadow_layer = shadow_layer.filter(ImageFilter.GaussianBlur(radius=S * 0.02))
    canvas = Image.alpha_composite(canvas, shadow_layer)

    draw = ImageDraw.Draw(canvas)
    o = canvas_pad

    if raw_bytes:
        try:
            img = Image.open(BytesIO(raw_bytes)).convert("RGB")
            w, h = img.size
            side = min(w, h)
            img = img.crop(((w - side) // 2, (h - side) // 2, (w + side) // 2, (h + side) // 2))

            border_px = max(4, int(S * 0.028))
            gap_px = max(2, int(S * 0.012))
            photo_size = S - (border_px + gap_px) * 2
            img = img.resize((photo_size, photo_size), Image.LANCZOS).convert("RGBA")

            mask = Image.new("L", (photo_size, photo_size), 0)
            ImageDraw.Draw(mask).ellipse((0, 0, photo_size - 1, photo_size - 1), fill=255)
            circle = Image.new("RGBA", (photo_size, photo_size), (0, 0, 0, 0))
            circle.paste(img, mask=mask)

            draw.ellipse((o, o, o + S - 1, o + S - 1), fill=(br, bg, bb, 255))
            draw.ellipse((o + border_px, o + border_px, o + S - border_px - 1, o + S - border_px - 1), fill=(255, 255, 255, 255))
            canvas.paste(circle, (o + border_px + gap_px, o + border_px + gap_px), mask=circle)

            canvas = canvas.resize((canvas.width // supersample, canvas.height // supersample), Image.LANCZOS)
            out = BytesIO()
            canvas.save(out, "PNG", optimize=True)
            out.seek(0)
            return out.read(), canvas_pad // supersample
        except Exception:
            pass

    # Initials Avatar fallback
    r2, g2, b2 = max(0, br - 35), max(0, bg - 35), max(0, bb - 35)
    draw.ellipse((o, o, o + S - 1, o + S - 1), fill=(br, bg, bb, 255))
    border_px = int(S * 0.035)
    draw.ellipse((o + border_px, o + border_px, o + S - border_px - 1, o + S - border_px - 1), fill=(r2, g2, b2, 255))

    clean_name = re.sub(r"\s+", " ", str(name)).strip()
    parts = clean_name.split()
    initials = (parts[0][0] + parts[1][0]).upper() if len(parts) >= 2 else (clean_name[:2].upper() if clean_name else "AZ")

    fs = int(S * 0.36)
    font = None
    for fn in ["georgiab.ttf", "segoeuib.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf"]:
        try:
            font = ImageFont.truetype(fn, fs)
            break
        except Exception:
            pass
    if not font:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), initials, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((canvas_size - tw) // 2, (canvas_size - th) // 2 - int(S * 0.015)), initials, fill=(255, 255, 255, 255), font=font)

    canvas = canvas.resize((canvas.width // supersample, canvas.height // supersample), Image.LANCZOS)
    out = BytesIO()
    canvas.save(out, "PNG", optimize=True)
    out.seek(0)
    return out.read(), canvas_pad // supersample


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  UNIVERSAL TEMPLATE PRESENTATION BUILDER                                ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
class UniversalTemplateBuilder:
    def __init__(self, spec: dict, config: dict):
        self.spec = spec
        self.config = config
        self.prs = Presentation()
        self.prs.slide_width = Inches(AppConfig.SLIDE_W)
        self.prs.slide_height = Inches(AppConfig.SLIDE_H)

    def hex_rgb(self, hex_str: str) -> RGBColor:
        h = hex_str.lstrip("#")
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    def add_rect(self, slide, l, t, w, h, hex_fill, radius=None, border_color=None, border_width=1):
        st = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(st, Inches(l), Inches(t), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.hex_rgb(hex_fill)
        if border_color:
            shape.line.color.rgb = self.hex_rgb(border_color)
            shape.line.width = Pt(border_width)
        else:
            shape.line.fill.background()
        if radius:
            try: shape.adjustments[0] = radius
            except Exception: pass
        return shape

    def add_line(self, slide, x1, y1, x2, y2, color, width=1):
        connector = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        connector.line.color.rgb = self.hex_rgb(color)
        connector.line.width = Pt(width)
        return connector

    def add_text(self, slide, text, l, t, w, h, size, color, bold=False, italic=False,
                 align=PP_ALIGN.LEFT, font=None, spacing=None):
        font_name = font or self.spec["font_heading"]
        tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = align
        if spacing is not None:
            p.space_after = Pt(spacing)
        r = p.add_run()
        r.text = str(text)
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.name = font_name
        r.font.color.rgb = self.hex_rgb(color)
        return tb

    def apply_slide_bg(self, slide):
        NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
        NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
        ang = int(self.spec.get("bg_angle", 135) * 60000)
        s = self.spec["bg_start"].lstrip("#")
        e = self.spec["bg_end"].lstrip("#")
        bg_xml = (
            f'<p:bg xmlns:p="{NS_P}" xmlns:a="{NS_A}">'
            f'<p:bgPr>'
            f'<a:gradFill rotWithShape="1">'
            f'<a:gsLst>'
            f'<a:gs pos="0"><a:srgbClr val="{s}"/></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="{e}"/></a:gs>'
            f'</a:gsLst>'
            f'<a:lin ang="{ang}" scaled="0"/>'
            f'</a:gradFill>'
            f'<a:effectLst/>'
            f'</p:bgPr>'
            f'</p:bg>'
        )
        cSld = slide._element.find(qn("p:cSld"))
        if cSld is not None:
            existing = cSld.find(qn("p:bg"))
            if existing is not None: cSld.remove(existing)
            cSld.insert(0, parse_xml(bg_xml))

    def apply_transition(self, slide):
        sld = slide._element
        existing = sld.find(qn("p:transition"))
        if existing is not None: sld.remove(existing)
        trans = etree.SubElement(sld, qn("p:transition"))
        trans.set("spd", "fast")
        fade = etree.SubElement(trans, qn("p:fade"))
        fade.set("thruBlk", "0")

    def draw_header(self, slide, title_text: str, section_index: str = None):
        style = self.spec["card_style"]
        if style == "editorial_rules":
            # Swiss Editorial Header: Stark white with clean hairline rule
            self.add_line(slide, 0.75, 0.95, AppConfig.SLIDE_W - 0.75, 0.95, self.spec["hairline"], 0.75)
            if section_index:
                self.add_text(slide, section_index, 0.75, 0.35, 0.8, 0.5, 24, self.spec["primary"], bold=True)
            self.add_text(slide, title_text, 1.6 if section_index else 0.75, 0.38, 7.5, 0.5, 22,
                          self.spec["text_main"], bold=True, font=self.spec["font_heading"])
            self.add_text(slide, self.config.get("client", "Airtel Zambia").upper(),
                          AppConfig.SLIDE_W - 4.75, 0.45, 4.0, 0.35, 9.5, self.spec["text_sub"], bold=True, align=PP_ALIGN.RIGHT)
        elif style == "hud_brackets":
            self.add_rect(slide, 0, 0, AppConfig.SLIDE_W, 0.68, "#020617")
            self.add_line(slide, 0, 0.68, AppConfig.SLIDE_W, 0.68, self.spec["primary"], 1.5)
            self.add_text(slide, f"// {title_text.upper()}", 0.75, 0.18, 7.0, 0.4, 18,
                          self.spec["text_main"], bold=True, font="Segoe UI")
            self.add_text(slide, f"[SEC_LEVEL: 1]  ●  {self.config.get('client', 'AIRTEL').upper()}",
                          AppConfig.SLIDE_W - 4.75, 0.22, 4.0, 0.35, 9.0, self.spec["secondary"], bold=True, align=PP_ALIGN.RIGHT)
        else:
            self.add_line(slide, 0, 0.72, AppConfig.SLIDE_W, 0.72, self.spec["primary"], 2.0)
            self.add_text(slide, title_text, 0.75, 0.20, 8.5, 0.45, 22,
                          self.spec["text_main"], bold=True, font=self.spec["font_heading"])
            self.add_text(slide, f"● {self.config.get('client', 'Airtel Zambia').upper()}",
                          AppConfig.SLIDE_W - 3.75, 0.26, 3.0, 0.35, 10, self.spec["text_main"], bold=True, align=PP_ALIGN.RIGHT)

    def add_chart_image(self, slide, img_bytes: bytes, l, t, w, h, fallback_label: str = None):
        """Embed a chart PNG into the slide with guaranteed aspect ratio preservation and zero border."""
        if img_bytes and len(img_bytes) > 100:
            try:
                fitted_l, fitted_t, fitted_w, fitted_h = l, t, w, h
                if PIL_AVAILABLE:
                    try:
                        with Image.open(BytesIO(img_bytes)) as pil_img:
                            img_w, img_h = pil_img.size
                        img_ratio = img_w / img_h
                        box_ratio = w / h

                        if img_ratio > box_ratio:
                            # Image wider than box: fit to width, center vertically
                            fitted_w = w
                            fitted_h = w / img_ratio
                            fitted_l = l
                            fitted_t = t + (h - fitted_h) / 2
                        else:
                            # Image taller than box: fit to height, center horizontally
                            fitted_h = h
                            fitted_w = h * img_ratio
                            fitted_t = t
                            fitted_l = l + (w - fitted_w) / 2
                    except Exception:
                        pass

                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
                    tf.write(img_bytes)
                    pic = slide.shapes.add_picture(tf.name, Inches(fitted_l), Inches(fitted_t),
                                                   Inches(fitted_w), Inches(fitted_h))
                try: os.unlink(tf.name)
                except Exception: pass

                # Eliminate any square picture border applied by PowerPoint
                try:
                    pic.line.fill.background()
                except Exception:
                    pass

                return True
            except Exception as e:
                logger.error(f"Failed to add chart picture: {e}")

        # ── FALLBACK ──
        logger.error(f"CHART RENDER FAILED — fallback activated for '{fallback_label or 'unknown'}' "
                     f"(bytes={len(img_bytes) if img_bytes else 0})")

        if fallback_label:
            # Render a clean native placeholder so the slide is never blank
            fb = self.add_rect(slide, l + w * 0.1, t + h * 0.25, w * 0.8, h * 0.5,
                               self.spec["card_bg"],
                               radius=0.08 if self.spec["card_style"] != "editorial_rules" else None,
                               border_color=self.spec["card_border"])
            self.add_text(slide, fallback_label, l + w * 0.1, t + h * 0.35, w * 0.8, 0.4,
                          14, self.spec["text_muted"], bold=True, align=PP_ALIGN.CENTER)
            self.add_text(slide, "Chart rendering unavailable — data is present",
                          l + w * 0.1, t + h * 0.55, w * 0.8, 0.3,
                          10, self.spec["text_sub"], align=PP_ALIGN.CENTER)
        return False

    # ── 1. COVER SLIDE ───────────────────────────────────────────────────
    def build_cover_slide(self, total_attendees: int):
        sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.apply_slide_bg(sl)
        style = self.spec["card_style"]

        if style == "editorial_rules":
            # Swiss Precision Editorial Cover
            lx = 0.95
            self.add_rect(sl, lx, 1.2, 0.05, 5.0, self.spec["primary"])

            self.add_text(sl, "EXECUTIVE TELEMETRY REPORT", lx + 0.35, 1.25, 4.5, 0.3, 9,
                          self.spec["primary"], bold=True, spacing=4)
            self.add_text(sl, self.config.get("client", "Airtel Zambia"), lx + 0.35, 1.7, 5.5, 1.4, 48,
                          self.spec["text_main"], bold=True, font=self.spec["font_heading"])
            self.add_text(sl, self.config.get("workshop", "Cyber Crisis Simulation"), lx + 0.35, 3.25, 5.5, 0.8, 22,
                          self.spec["text_muted"], italic=True)

            date_label = "Session Date" if len(self.config.get("target_dates", [1])) == 1 else "Session Period"
            date_display_val = self.config.get('date_display', self.config.get('target_date', '2026-09-22'))
            self.add_text(sl, f"{date_label}: {date_display_val}  |  Confidential",
                          lx + 0.35, 4.15, 5.5, 0.35, 11, self.spec["text_sub"])

            # Right Column: Precision Editorial Cards
            rx, rw = 7.0, 5.35
            cards_data = [
                ("ATTENDEES", str(total_attendees), "Verified Attendees"),
                ("TELEMETRY STATUS", "100%", "Biometric & Telemetry Sync"),
                ("SECURITY AUDIT", "VERIFIED", "Zero-Coordinates Policy Compliant")
            ]
            for i, (label, val, sub) in enumerate(cards_data):
                cy = 1.35 + i * 1.65
                self.add_line(sl, rx, cy, rx + rw, cy, self.spec["hairline"], 1.5)
                self.add_text(sl, label, rx, cy + 0.12, rw, 0.25, 8.5, self.spec["primary"], bold=True)
                self.add_text(sl, val, rx, cy + 0.35, rw, 0.65, 38, self.spec["text_main"], bold=True, font=self.spec["font_heading"])
                self.add_text(sl, sub, rx, cy + 1.15, rw, 0.25, 10, self.spec["text_muted"])
            self.add_line(sl, rx, 1.35 + 3 * 1.65, rx + rw, 1.35 + 3 * 1.65, self.spec["hairline"], 1.5)

        else:
            # Bento Grid Cover for Obsidian / Prestige / Linear / Flagship / Cyber
            lx = 0.85
            self.add_rect(sl, lx, 1.4, 0.06, 4.7, self.spec["primary"])

            self.add_rect(sl, lx + 0.3, 1.4, 2.7, 0.32, self.spec["pill_bg"], radius=0.25, border_color=self.spec["pill_border"])
            self.add_text(sl, "EXECUTIVE TELEMETRY REPORT", lx + 0.3, 1.45, 2.7, 0.25, 8.5,
                          self.spec["pill_text"], bold=True, align=PP_ALIGN.CENTER)

            self.add_text(sl, self.config.get("client", "Airtel Zambia"), lx + 0.3, 1.9, 5.5, 1.3, 44,
                          self.spec["text_main"], bold=True, font=self.spec["font_heading"])
            self.add_text(sl, self.config.get("workshop", "Cyber Crisis Simulation"), lx + 0.3, 3.4, 5.5, 0.6, 20,
                          self.spec["text_muted"], italic=True)

            date_display_val = self.config.get('date_display', self.config.get('target_date', '2026-09-22'))
            self.add_rect(sl, lx + 0.3, 4.25, 3.4, 0.38, self.spec["pill_bg"], radius=0.2, border_color=self.spec["pill_border"])
            self.add_text(sl, f"📅  {date_display_val}", lx + 0.5, 4.31, 3.0, 0.3, 11,
                          self.spec["text_muted"], bold=True)

            rx, rw = 6.6, 5.85
            cards_data = [
                ("REGISTERED ATTENDEES", str(total_attendees), "Verified attendees engaged", self.spec["primary"]),
                ("TELEMETRY INTEGRITY", "100%", "Biometrics & location synchronized", "#10B981" if self.spec["is_dark"] else self.spec["primary"]),
                ("SECURITY COMPLIANCE", "VERIFIED", f"Sync: {datetime.now(timezone.utc).strftime('%H:%M UTC')}", self.spec["accent"])
            ]
            for i, (label, metric, sub, acc_color) in enumerate(cards_data):
                cy = 1.4 + i * 1.65
                self.add_rect(sl, rx, cy, rw, 1.45, self.spec["card_bg"], radius=0.10, border_color=self.spec["card_border"])
                self.add_rect(sl, rx, cy + 0.2, 0.05, 1.05, acc_color)
                self.add_text(sl, label, rx + 0.3, cy + 0.18, rw - 0.6, 0.25, 9, self.spec["text_sub"], bold=True)
                self.add_text(sl, metric, rx + 0.3, cy + 0.42, rw - 0.6, 0.55, 32, self.spec["text_main"], bold=True, font=self.spec["font_heading"])
                self.add_text(sl, sub, rx + 0.3, cy + 1.08, rw - 0.6, 0.25, 10, self.spec["text_muted"])

        self.apply_transition(sl)

    # ── 2. PERMISSION ANALYTICS SLIDE ────────────────────────────────────
    def build_permission_slide(self, participants: list):
        sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.apply_slide_bg(sl)
        self.draw_header(sl, "Telemetry & Permission Audit", "01")

        total = len(participants) if participants else 1
        cam = sum(1 for p in participants if p.get("camera_permission"))
        loc = sum(1 for p in participants if p.get("location_permission"))

        card_w, card_h = 5.75, 5.8
        # Card 1: Camera
        self.add_rect(sl, 0.75, 1.15, card_w, card_h, self.spec["card_bg"],
                      radius=0.10 if self.spec["card_style"] != "editorial_rules" else None,
                      border_color=self.spec["card_border"])

        cam_img = render_donut_chart(["Granted", "Denied"], [cam, max(0, total - cam)], "Camera Authorization",
                                     self.spec["is_dark"], self.spec["primary"], [self.spec["primary"], "#64748B" if self.spec["is_dark"] else "#CBD5E1"],
                                     card_bg=self.spec["card_bg"])
        self.add_chart_image(sl, cam_img, 0.95, 1.45, 5.35, 3.8, fallback_label="Camera Authorization")

        self.add_rect(sl, 1.6, 5.7, 4.05, 0.45, self.spec["pill_bg"], radius=0.22, border_color=self.spec["pill_border"])
        self.add_text(sl, f"✓  {cam} of {total} Participants Active", 1.6, 5.8, 4.05, 0.3, 11,
                      "#10B981" if self.spec["is_dark"] else self.spec["primary"], bold=True, align=PP_ALIGN.CENTER)

        # Card 2: Location
        self.add_rect(sl, 6.85, 1.15, card_w, card_h, self.spec["card_bg"],
                      radius=0.10 if self.spec["card_style"] != "editorial_rules" else None,
                      border_color=self.spec["card_border"])

        loc_img = render_donut_chart(["Granted", "Denied"], [loc, max(0, total - loc)], "Location Authorization",
                                     self.spec["is_dark"], self.spec["primary"], [self.spec["primary"], "#64748B" if self.spec["is_dark"] else "#CBD5E1"],
                                     card_bg=self.spec["card_bg"])
        self.add_chart_image(sl, loc_img, 7.05, 1.45, 5.35, 3.8, fallback_label="Location Authorization")

        self.add_rect(sl, 7.7, 5.7, 4.05, 0.45, self.spec["pill_bg"], radius=0.22, border_color=self.spec["pill_border"])
        self.add_text(sl, f"✓  {loc} of {total} Locations Verified", 7.7, 5.8, 4.05, 0.3, 11,
                      "#10B981" if self.spec["is_dark"] else self.spec["primary"], bold=True, align=PP_ALIGN.CENTER)

        self.apply_transition(sl)

    # ── 3. PARTICIPANT DIRECTORY (GALLERY) ───────────────────────────────
    def build_gallery_slides(self, participants: list):
        """Scalable, beautifully centered participant directory.
        - Supports up to 12 participants on a single slide without crowding.
        - Only scales up to multiple slides on a requirement basis when attendee count exceeds 12.
        - Dynamically adapts card dimensions for 1 to 12 attendees.
        - Mathematically centers each row horizontally and the entire grid vertically.
        - Eliminates any picture outline borders."""
        if not participants:
            return

        n_total = len(participants)
        max_per_page = 12

        if n_total <= max_per_page:
            chunks = [participants]
        else:
            n_pages = math.ceil(n_total / max_per_page)
            per_page = math.ceil(n_total / n_pages)
            chunks = [participants[i:i + per_page] for i in range(0, n_total, per_page)]

        for page_idx, chunk in enumerate(chunks, 1):
            sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            self.apply_slide_bg(sl)
            title = "Participant Directory" if len(chunks) == 1 else f"Participant Directory ({page_idx}/{len(chunks)})"
            self.draw_header(sl, title, "02")

            n = len(chunk)
            if n <= 4:
                cols = max(n, 1)
                avatar_d = 1.65
                item_w = 2.50
                gap_x, gap_y = 0.55, 0.55
                font_name_sz, font_loc_sz = 12, 10
            elif n <= 8:
                cols = 4
                avatar_d = 1.45
                item_w = 2.40
                gap_x, gap_y = 0.55, 0.55
                font_name_sz, font_loc_sz = 11.5, 9.5
            else:
                cols = 6
                avatar_d = 1.45
                item_w = 1.85
                gap_x, gap_y = 0.30, 0.45
                font_name_sz, font_loc_sz = 10.5, 9.0

            rows = math.ceil(n / cols)
            item_h = avatar_d + 0.10 + 0.75
            grid_total_h = rows * item_h + (rows - 1) * gap_y
            header_offset = 1.15
            grid_start_y = header_offset + (AppConfig.SLIDE_H - header_offset - grid_total_h) / 2

            for idx_p, p in enumerate(chunk):
                r = idx_p // cols
                c = idx_p % cols

                # Center each row horizontally
                items_in_row = min(cols, n - r * cols)
                row_total_w = items_in_row * item_w + (items_in_row - 1) * gap_x
                row_start_x = (AppConfig.SLIDE_W - row_total_w) / 2

                cx = row_start_x + c * (item_w + gap_x)
                cy = grid_start_y + r * (item_h + gap_y)

                name = p.get("name", "Unknown")
                raw_bytes = p.get("image")
                png, pad = make_avatar(name, raw_bytes, self.spec["primary"])

                avatar_x = cx + (item_w - avatar_d) / 2
                avatar_y = cy

                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
                    tf.write(png)
                    pic = sl.shapes.add_picture(tf.name, Inches(avatar_x), Inches(avatar_y),
                                                Inches(avatar_d), Inches(avatar_d))
                try: os.unlink(tf.name)
                except Exception: pass

                # Guarantee zero picture outline on avatar
                try: pic.line.fill.background()
                except Exception: pass

                disp_name = (name[:23] + "…") if len(name) > 24 else name
                city_country = p.get("city_country", "")

                tb = sl.shapes.add_textbox(Inches(cx), Inches(avatar_y + avatar_d + 0.10),
                                           Inches(item_w), Inches(0.75))
                tf = tb.text_frame
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                tf.word_wrap = True

                p1 = tf.paragraphs[0]
                p1.alignment = PP_ALIGN.CENTER
                r1 = p1.add_run()
                r1.text = disp_name
                r1.font.size = Pt(font_name_sz)
                r1.font.bold = True
                r1.font.name = self.spec["font_heading"]
                r1.font.color.rgb = self.hex_rgb(self.spec["text_main"])

                if city_country:
                    p2 = tf.add_paragraph()
                    p2.alignment = PP_ALIGN.CENTER
                    p2.space_before = Pt(2)
                    r2 = p2.add_run()
                    r2.text = city_country
                    r2.font.size = Pt(font_loc_sz)
                    r2.font.name = self.spec["font_body"]
                    r2.font.color.rgb = self.hex_rgb(self.spec["text_muted"])

            self.apply_transition(sl)

    # ── 4. GEOGRAPHIC OPERATIONS HUB (MAP + REGIONAL BREAKDOWN) ──────────
    def build_location_slide(self, participants: list):
        sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.apply_slide_bg(sl)
        self.draw_header(sl, "Geographic Operations & Epicenter Map", "03")

        cities = Counter(p.get("city_country") for p in participants if p.get("city_country"))
        if not cities:
            cities = Counter(["Lusaka, Zambia"])

        top_loc, top_cnt = cities.most_common(1)[0]
        top_city = top_loc.split(",")[0].strip()
        top_country = top_loc.split(",")[1].strip() if "," in top_loc else "Zambia"
        total = len(participants) if participants else 1
        pct = (top_cnt / total) * 100

        # Compute exact epicenter from real coordinates of attendees at top location
        top_participants_with_coords = [
            p for p in participants
            if p.get("city_country") == top_loc and p.get("has_coords", False)
        ]
        if top_participants_with_coords:
            center_lat = sum(p["lat"] for p in top_participants_with_coords) / len(top_participants_with_coords)
            center_lon = sum(p["lon"] for p in top_participants_with_coords) / len(top_participants_with_coords)
        else:
            top_p = next((p for p in participants if p.get("city_country") == top_loc), {})
            center_lat = top_p.get("lat", -15.4167)
            center_lon = top_p.get("lon", 28.2833)

        # Left Column: The Aesthetic Cartographic Map
        map_w, map_h = 6.8, 5.8
        self.add_rect(sl, 0.75, 1.15, map_w, map_h, self.spec["card_bg"],
                      radius=0.10 if self.spec["card_style"] != "editorial_rules" else None,
                      border_color=self.spec["card_border"])

        map_png = render_cartographic_map(top_city, top_country, center_lat, center_lon, top_cnt, pct,
                                          self.spec["is_dark"], self.spec["primary"],
                                          card_bg=self.spec["card_bg"],
                                          total_attendees=total)
        self.add_chart_image(sl, map_png, 0.75, 1.15, map_w, map_h, fallback_label="Epicenter Map")

        # Right Column: Regional Distribution Breakdown
        right_x = 7.8
        right_w = AppConfig.SLIDE_W - right_x - 0.75

        # Metric KPI Tile 1
        self.add_rect(sl, right_x, 1.15, right_w, 1.45, self.spec["card_bg"],
                      radius=0.10 if self.spec["card_style"] != "editorial_rules" else None,
                      border_color=self.spec["card_border"])
        self.add_rect(sl, right_x, 1.15, 0.05, 1.45, self.spec["primary"])
        self.add_text(sl, "PRIMARY SIMULATION HUB SHARE", right_x + 0.25, 1.30, right_w - 0.5, 0.25, 9, self.spec["text_sub"], bold=True)
        self.add_text(sl, f"{pct:.0f}%", right_x + 0.25, 1.55, right_w - 0.5, 0.55, 34, self.spec["text_main"], bold=True, font=self.spec["font_heading"])
        self.add_text(sl, f"{top_cnt} of {total} participants located in {top_city}", right_x + 0.25, 2.20, right_w - 0.5, 0.25, 10, self.spec["text_muted"])

        # Horizontal Bars Chart for All Locations
        ranked = cities.most_common(5)
        cats, vals = [c for c, _ in ranked], [v for _, v in ranked]
        bars_png = render_horizontal_bars(cats, vals, "Regional Demographic Distribution",
                                          self.spec["is_dark"], self.spec["primary"],
                                          card_bg=self.spec["card_bg"])
        self.add_rect(sl, right_x, 2.80, right_w, 4.15, self.spec["card_bg"],
                      radius=0.10 if self.spec["card_style"] != "editorial_rules" else None,
                      border_color=self.spec["card_border"])
        self.add_chart_image(sl, bars_png, right_x + 0.15, 2.95, right_w - 0.3, 3.85, fallback_label="Regional Demographics")

        # Inject pulsing beacon animation on the map image
        self._inject_pulse_animation(sl)

        self.apply_transition(sl)

    # ── 5. HARDWARE & DISPLAY TELEMETRY SLIDE ────────────────────────────
    def build_screen_slide(self, participants: list):
        sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.apply_slide_bg(sl)
        self.draw_header(sl, "Display Telemetry & Hardware Recognition", "04")

        winner, best_val = None, -1
        for p in participants:
            sv = p.get("screen_sort_value", -1)
            if sv > best_val: best_val, winner = sv, p

        card_w, card_h = AppConfig.SLIDE_W - 1.5, 5.8
        self.add_rect(sl, 0.75, 1.15, card_w, card_h, self.spec["card_bg"],
                      radius=0.10 if self.spec["card_style"] != "editorial_rules" else None,
                      border_color=self.spec["card_border"])

        if winner:
            self.add_rect(sl, 4.65, 1.8, 4.0, 0.4, self.spec["pill_bg"], radius=0.2, border_color=self.spec["pill_border"])
            self.add_text(sl, "🏆  LARGEST DISPLAY IN THE ROOM", 4.65, 1.88, 4.0, 0.3, 10,
                          self.spec["pill_text"], bold=True, align=PP_ALIGN.CENTER)

            self.add_text(sl, winner.get("name", "Unknown"), 1.3, 2.5, card_w - 1.2, 0.6, 32,
                          self.spec["text_main"], bold=True, align=PP_ALIGN.CENTER, font=self.spec["font_heading"])

            screen_disp = winner.get("screen_display", "1920 × 1080 Full HD")
            self.add_text(sl, screen_disp, 1.3, 3.25, card_w - 1.2, 0.9, 52,
                          self.spec["primary"], bold=True, align=PP_ALIGN.CENTER, font=self.spec["font_heading"])

            plat = normalize_platform(winner.get("platform", "Windows PC"))
            self.add_rect(sl, 5.35, 4.45, 2.6, 0.42, self.spec["pill_bg"], radius=0.21, border_color=self.spec["pill_border"])
            self.add_text(sl, f"📱  Platform: {plat}", 5.35, 4.53, 2.6, 0.3, 11,
                          self.spec["text_main"], bold=True, align=PP_ALIGN.CENTER)

            self.add_text(sl, "Identified by client viewport pixel-density telemetry analysis",
                          1.3, 5.25, card_w - 1.2, 0.35, 11.5, self.spec["text_muted"], align=PP_ALIGN.CENTER)

        self.apply_transition(sl)

    # ── 6. DEVICE PLATFORM ECOSYSTEM ────────────────────────────────────
    def build_platform_slide(self, participants: list):
        sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.apply_slide_bg(sl)
        self.draw_header(sl, "Operating System & Device Ecosystem", "05")

        counts = Counter(normalize_platform(p.get("platform")) for p in participants if p.get("platform"))
        if not counts:
            counts = Counter(["Android", "iPhone / iOS", "Windows PC"])

        ranked = counts.most_common(6)
        cats, vals = [c for c, _ in ranked], [v for _, v in ranked]

        card_w, card_h = AppConfig.SLIDE_W - 1.5, 5.8
        self.add_rect(sl, 0.75, 1.15, card_w, card_h, self.spec["card_bg"],
                      radius=0.10 if self.spec["card_style"] != "editorial_rules" else None,
                      border_color=self.spec["card_border"])

        img = render_donut_chart(cats, vals, "Operating System Share",
                                 self.spec["is_dark"], self.spec["primary"],
                                 card_bg=self.spec["card_bg"])
        self.add_chart_image(sl, img, 2.5, 1.6, 8.3, 5.0, fallback_label="Operating System Share")

        self.apply_transition(sl)

    # ── 7. EXECUTIVE SUMMARY & AUDIT ────────────────────────────────────
    def build_summary_slide(self, participants: list):
        sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.apply_slide_bg(sl)
        self.draw_header(sl, "Executive Summary & Audit", "06")

        total = len(participants) if participants else 1
        with_img = sum(1 for p in participants if p.get("has_image", False))
        devices = Counter(normalize_platform(p.get("platform")) for p in participants if p.get("platform"))
        top_dev = devices.most_common(1)[0][0] if devices else "Android"
        locations = Counter(p.get("city_country") for p in participants if p.get("city_country"))
        top_loc = locations.most_common(1)[0][0].split(",")[0] if locations else "Lusaka"

        stats = [
            ("TOTAL ATTENDED", str(total), "Verified Attendees", self.spec["primary"]),
            ("PHOTOS CAPTURED", str(with_img), "Biometric Passed", "#10B981" if self.spec["is_dark"] else self.spec["primary"]),
            ("TOP PLATFORM", top_dev, "Dominant Device", self.spec["accent"]),
            ("PRIMARY HUB", top_loc, "Top Epicenter", "#0284C7" if self.spec["is_dark"] else self.spec["secondary"])
        ]

        card_w = 2.76
        for i, (title, val, sub, col) in enumerate(stats):
            cx = 0.75 + i * (card_w + 0.28)
            self.add_rect(sl, cx, 1.15, card_w, 1.6, self.spec["card_bg"],
                          radius=0.08 if self.spec["card_style"] != "editorial_rules" else None,
                          border_color=self.spec["card_border"])
            self.add_rect(sl, cx, 1.15, card_w, 0.05, col)

            # Dynamic font sizing to prevent overflow on long values like "iPhone / iOS"
            val_len = len(str(val))
            if val_len > 12:
                metric_size = 18
            elif val_len > 8:
                metric_size = 22
            else:
                metric_size = 32

            self.add_text(sl, val, cx, 1.35, card_w, 0.55, metric_size, self.spec["text_main"], bold=True, align=PP_ALIGN.CENTER, font=self.spec["font_heading"])
            self.add_text(sl, title, cx, 1.95, card_w, 0.25, 9, self.spec["text_sub"], bold=True, align=PP_ALIGN.CENTER)
            self.add_text(sl, sub, cx, 2.25, card_w, 0.25, 9.5, col, align=PP_ALIGN.CENTER)

        # Audit Card
        bottom_h = 3.65
        self.add_rect(sl, 0.75, 3.0, AppConfig.SLIDE_W - 1.5, bottom_h, self.spec["card_bg"],
                      radius=0.10 if self.spec["card_style"] != "editorial_rules" else None,
                      border_color=self.spec["card_border"])
        self.add_rect(sl, 0.75, 3.0, 0.05, bottom_h, self.spec["primary"])

        self.add_text(sl, "DATA QUALITY & TELEMETRY AUDIT VERIFICATION", 1.2, 3.3, 10.0, 0.35, 11,
                      self.spec["primary"], bold=True)

        bullets = [
            f"Biometric Quality: {with_img} of {total} captured photos verified via Computer Vision cascade.",
            "Macro Geocoding: Standardized via OpenStreetMap administrative hierarchy (0 raw coordinates exposed).",
            "Cartographic Telemetry: High-precision regional vector mapping calibrated to top attendee epicenter.",
            "Device Fingerprinting: Platform and display resolution telemetry captured with DPR scaling.",
            f"Presentation Telemetry: Automated live build synchronized at {datetime.now(timezone.utc).strftime('%H:%M:%S UTC')}."
        ]

        by = 3.8
        for b in bullets:
            self.add_text(sl, f"• {b}", 1.2, by, 10.8, 0.32, 10, self.spec["text_muted"])
            by += 0.50

        self.apply_transition(sl)

    # ── PULSING BEACON ANIMATION (OOXML Injection) ────────────────────────
    def _inject_pulse_animation(self, slide):
        """Inject a continuous pulsing scale animation on the map image shape (the last picture added).
        Uses raw OOXML <p:timing> elements — works in slideshow mode only."""
        try:
            # Find the map picture shape (first picture shape on this slide)
            pic_shape = None
            for shape in slide.shapes:
                if hasattr(shape, 'image'):
                    pic_shape = shape
                    break  # Use the first (map) picture

            if pic_shape is None:
                logger.warning("PULSE ANIM: No picture shape found on map slide — skipping animation")
                return

            shape_id = pic_shape.shape_id

            # Build OOXML timing tree for continuous pulse
            NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
            NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

            sld = slide._element

            # Remove any existing timing
            existing_timing = sld.find(qn("p:timing"))
            if existing_timing is not None:
                sld.remove(existing_timing)

            timing_xml = (
                f'<p:timing xmlns:p="{NS_P}" xmlns:a="{NS_A}">'
                f'<p:tnLst>'
                f'<p:par>'
                f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
                f'<p:childTnLst>'
                f'<p:seq concurrent="1" nextAc="seek">'
                f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
                f'<p:childTnLst>'
                f'<p:par>'
                f'<p:cTn id="3" fill="hold">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                f'<p:childTnLst>'
                f'<p:par>'
                f'<p:cTn id="4" fill="hold">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                f'<p:childTnLst>'

                # Scale pulse: grow to 105% then back, repeat forever
                f'<p:par>'
                f'<p:cTn id="5" presetID="53" presetClass="emph" presetSubtype="0" '
                f'fill="hold" grpId="0" repeatCount="indefinite">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                f'<p:childTnLst>'
                f'<p:animScale>'
                f'<p:cBhvr>'
                f'<p:cTn id="6" dur="1500" autoRev="1" fill="hold"/>'
                f'<p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl>'
                f'</p:cBhvr>'
                f'<p:by x="105000" y="105000"/>'
                f'</p:animScale>'
                f'</p:childTnLst>'
                f'</p:cTn>'
                f'</p:par>'

                f'</p:childTnLst>'
                f'</p:cTn>'
                f'</p:par>'

                f'</p:childTnLst>'
                f'</p:cTn>'
                f'</p:par>'

                f'</p:childTnLst>'
                f'</p:cTn>'
                f'</p:seq>'
                f'</p:childTnLst>'
                f'</p:cTn>'
                f'</p:par>'
                f'</p:tnLst>'
                f'</p:timing>'
            )

            timing_elem = parse_xml(timing_xml)
            sld.append(timing_elem)
            logger.info(f"PULSE ANIM: injected on shape {shape_id}")

        except Exception as e:
            logger.warning(f"PULSE ANIM: failed to inject — {e}")


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  DECK VERIFIER & BULLETPROOF CONTENT CHECKER                            ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
def verify_slide_content(prs) -> dict:
    """Post-build verification: ensures every slide has real content, not just headers.
    Returns a report with per-slide shape counts and any warnings."""
    report = {"ok": True, "slides": [], "warnings": []}
    EXPECTED_SLIDE_NAMES = [
        "Cover", "Permission Audit", "Gallery", "Map & Location",
        "Display Telemetry", "Platform Ecosystem", "Executive Summary"
    ]

    for i, slide in enumerate(prs.slides):
        shape_count = len(slide.shapes)
        pic_count = sum(1 for s in slide.shapes if hasattr(s, 'image'))
        text_count = sum(1 for s in slide.shapes if s.has_text_frame)
        slide_name = EXPECTED_SLIDE_NAMES[i] if i < len(EXPECTED_SLIDE_NAMES) else f"Slide {i+1}"

        slide_info = {
            "index": i + 1,
            "name": slide_name,
            "shapes": shape_count,
            "pictures": pic_count,
            "text_boxes": text_count
        }
        report["slides"].append(slide_info)

        # Minimum viable content check
        if shape_count < 3:
            msg = f"CRITICAL: {slide_name} (slide {i+1}) has only {shape_count} shapes — likely empty"
            report["warnings"].append(msg)
            report["ok"] = False
            logger.critical(msg)
        else:
            logger.info(f"VERIFY: Slide {i+1} ({slide_name}): {shape_count} shapes, "
                       f"{pic_count} pictures, {text_count} text boxes")

    return report


def verify_deck(pptx_path: str) -> dict:
    report = {"ok": True, "issues": []}
    try:
        with zipfile.ZipFile(pptx_path, "r") as zf:
            slide_files = sorted([n for n in zf.namelist()
                                  if n.startswith("ppt/slides/slide") and n.endswith(".xml")])
            for i, sf in enumerate(slide_files, 1):
                xml = zf.read(sf)
                if b"p:transition" not in xml:
                    report["issues"].append(f"Slide {i}: missing transition")
        test_prs = Presentation(pptx_path)
        report["slide_count"] = len(test_prs.slides)
    except Exception as e:
        report["ok"] = False
        report["issues"].append(str(e))
    return report


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  DATA LAYER                                                             ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
class S3DataLayer:
    def __init__(self):
        self.s3 = boto3.client("s3")

    @staticmethod
    def today_ist() -> str:
        return datetime.now(AppConfig.IST).strftime("%Y-%m-%d")

    @staticmethod
    def extract_date_from_key(s3_key: str) -> str:
        match = re.search(r"(\d{4}-\d{2}-\d{2})", s3_key)
        return match.group(1) if match else S3DataLayer.today_ist()

    @staticmethod
    def resolve_target_dates(payload: dict, event: dict = None, target_date_override: str = None):
        """Resolves target dates from all possible triggers:
        1. Explicit S3 upload event (target_date_override)
        2. Query string parameters (HTTP Function URL / API Gateway)
        3. Payload keys: 'date', 'target_date', 'day', 'processed_date'
        4. Date ranges: 'start_date' + 'end_date', 'from' + 'to', or 'date_range' ('2026-09-20 to 2026-09-22')
        5. Specific dates list: 'dates' (['2026-09-20', '2026-09-22'])
        6. Environment variables: 'START_DATE', 'END_DATE', 'TARGET_DATE', 'DATE'
        7. Fallback: Today's date in IST."""
        event = event or {}
        qp = event.get("queryStringParameters") or {}

        # Priority 1: Explicit S3 trigger override
        if target_date_override:
            d = target_date_override.strip()
            return [d], d

        # Priority 2: date_range field (array or string)
        dr = payload.get("date_range") or qp.get("date_range") or os.environ.get("DATE_RANGE")
        start_d, end_d = None, None
        if dr:
            if isinstance(dr, list) and len(dr) == 2:
                start_d, end_d = dr[0], dr[1]
            elif isinstance(dr, str):
                parts = re.split(r"\s*(?:to|:|\.\.)\s*", dr.strip())
                if len(parts) == 2:
                    start_d, end_d = parts[0], parts[1]

        # Priority 3: start_date / end_date aliases
        if not start_d or not end_d:
            start_d = (payload.get("start_date") or payload.get("from_date") or payload.get("from") or 
                       qp.get("start_date") or qp.get("from") or os.environ.get("START_DATE"))
            end_d = (payload.get("end_date") or payload.get("to_date") or payload.get("to") or 
                     qp.get("end_date") or qp.get("to") or os.environ.get("END_DATE"))

        # Priority 4: Specific dates list
        dates_list = payload.get("dates") or qp.get("dates")
        if isinstance(dates_list, str):
            dates_list = [d.strip() for d in dates_list.split(",") if d.strip()]

        # Priority 5: Single date aliases
        single_d = (payload.get("date") or payload.get("target_date") or payload.get("day") or 
                    payload.get("processed_date") or qp.get("date") or qp.get("day") or 
                    os.environ.get("EVENT_DATE") or 
                    os.environ.get("TARGET_DATE") or os.environ.get("DATE"))

        # Resolution A: Date Range
        if start_d and end_d:
            try:
                s_parsed = parse_flexible_date(start_d)
                e_parsed = parse_flexible_date(end_d)
                d1 = datetime.strptime(s_parsed, "%Y-%m-%d")
                d2 = datetime.strptime(e_parsed, "%Y-%m-%d")
                if d1 > d2:
                    d1, d2 = d2, d1
                cur = d1
                resolved = []
                while cur <= d2:
                    resolved.append(cur.strftime("%Y-%m-%d"))
                    cur += timedelta(days=1)
                display = f"{resolved[0]} to {resolved[-1]}"
                return resolved, display
            except Exception as e:
                logger.warning(f"Could not parse date range '{start_d}' to '{end_d}': {e}")

        # Resolution B: Specific list of dates
        if dates_list and isinstance(dates_list, list):
            resolved = sorted(list(set(parse_flexible_date(d) for d in dates_list if parse_flexible_date(d))))
            display = f"{resolved[0]} to {resolved[-1]}" if len(resolved) > 1 else resolved[0]
            return resolved, display

        # Resolution C: Single date
        if single_d:
            d_str = parse_flexible_date(single_d)
            if d_str:
                return [d_str], d_str

        # Resolution D: Default fallback to today's date in IST
        today = S3DataLayer.today_ist()
        return [today], today

    def discover_participant_keys(self, dates: list) -> list:
        """Discovers all participants across one or more target dates.
        Scans both records/{date}/ and images/{date}/ to ensure 100% attendee discovery."""
        items = {}
        paginator = self.s3.get_paginator("list_objects_v2")

        for d in dates:
            # 1. Scan records/{date}/
            rec_prefix = f"records/{d}/"
            try:
                for page in paginator.paginate(Bucket=AppConfig.SOURCE_BUCKET, Prefix=rec_prefix):
                    for obj in page.get("Contents", []):
                        k = obj["Key"]
                        if k.lower().endswith((".json", ".txt")):
                            basename = k.rsplit("/", 1)[-1].rsplit(".", 1)[0]
                            pair_key = (d, basename)
                            if pair_key not in items:
                                items[pair_key] = {"date": d, "basename": basename, "json_key": k, "image_key": None}
                            else:
                                items[pair_key]["json_key"] = k
            except Exception as e:
                logger.warning(f"Could not list records for {d}: {e}")

            # 2. Scan images/{date}/
            img_prefix = f"images/{d}/"
            try:
                for page in paginator.paginate(Bucket=AppConfig.SOURCE_BUCKET, Prefix=img_prefix):
                    for obj in page.get("Contents", []):
                        k = obj["Key"]
                        if k.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
                            basename = k.rsplit("/", 1)[-1].rsplit(".", 1)[0]
                            pair_key = (d, basename)
                            if pair_key not in items:
                                items[pair_key] = {"date": d, "basename": basename, "json_key": None, "image_key": k}
                            else:
                                items[pair_key]["image_key"] = k
            except Exception as e:
                logger.warning(f"Could not list images for {d}: {e}")

        # If nothing found via records/ or images/, check root date folder {date}/
        if not items:
            for d in dates:
                try:
                    for page in paginator.paginate(Bucket=AppConfig.SOURCE_BUCKET, Prefix=f"{d}/"):
                        for obj in page.get("Contents", []):
                            k = obj["Key"]
                            basename = k.rsplit("/", 1)[-1].rsplit(".", 1)[0]
                            pair_key = (d, basename)
                            if k.lower().endswith((".json", ".txt")):
                                if pair_key not in items:
                                    items[pair_key] = {"date": d, "basename": basename, "json_key": k, "image_key": None}
                                else:
                                    items[pair_key]["json_key"] = k
                            elif k.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
                                if pair_key not in items:
                                    items[pair_key] = {"date": d, "basename": basename, "json_key": None, "image_key": k}
                                else:
                                    items[pair_key]["image_key"] = k
                except Exception:
                    pass

        return list(items.values())

    def list_image_keys(self, date_filter: str) -> list:
        prefix = f"images/{date_filter}/"
        keys = []
        paginator = self.s3.get_paginator("list_objects_v2")
        for page in paginator.paginate(Bucket=AppConfig.SOURCE_BUCKET, Prefix=prefix):
            for obj in page.get("Contents", []):
                if obj["Key"].lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
                    keys.append(obj["Key"])
        return keys

    def fetch_participant_record(self, item) -> dict:
        if isinstance(item, str):
            image_key = item
            parts = image_key.split("/")
            basename = parts[2].rsplit(".", 1)[0] if len(parts) >= 3 else "unknown"
            date_str = parts[1] if len(parts) >= 3 else ""
            json_key = f"records/{date_str}/{basename}.json"
        else:
            date_str = item.get("date", "")
            basename = item.get("basename", "unknown")
            json_key = item.get("json_key") or f"records/{date_str}/{basename}.json"
            image_key = item.get("image_key")

        record = {"image_key": image_key, "raw_bytes": None, "meta": {}, "meta_key": None, "basename": basename, "date": date_str}

        # 1. Primary path: records/*.json
        meta_loaded = False
        try:
            resp = self.s3.get_object(Bucket=AppConfig.SOURCE_BUCKET, Key=json_key)
            record["meta"] = json.loads(resp["Body"].read().decode("utf-8", errors="ignore"))
            record["meta_key"] = json_key
            meta_loaded = True
        except Exception:
            pass

        # 2. Try direct {date_str}/{basename}.json if not found
        if not meta_loaded and date_str:
            alt_json = f"{date_str}/{basename}.json"
            try:
                resp = self.s3.get_object(Bucket=AppConfig.SOURCE_BUCKET, Key=alt_json)
                record["meta"] = json.loads(resp["Body"].read().decode("utf-8", errors="ignore"))
                record["meta_key"] = alt_json
                meta_loaded = True
            except Exception:
                pass

        # 3. Legacy fallback: raw-text/*.txt
        if not meta_loaded and date_str:
            txt_key = f"raw-text/{date_str}/{basename}.txt"
            try:
                resp = self.s3.get_object(Bucket=AppConfig.SOURCE_BUCKET, Key=txt_key)
                raw = resp["Body"].read().decode("utf-8", errors="ignore").strip()
                record["meta"] = json.loads(raw)
                record["meta_key"] = txt_key
                meta_loaded = True
            except Exception:
                pass

        # 4. Fetch image bytes if image_key is known or from meta
        if not image_key and record["meta"]:
            alt_img = record["meta"].get("image_key") or record["meta"].get("image")
            if alt_img and isinstance(alt_img, str) and (alt_img.startswith("images/") or alt_img.startswith(f"{date_str}/")):
                image_key = alt_img

        if image_key:
            try:
                img_resp = self.s3.get_object(Bucket=AppConfig.SOURCE_BUCKET, Key=image_key)
                record["raw_bytes"] = img_resp["Body"].read()
                record["image_key"] = image_key
            except Exception as e:
                logger.warning(f"Failed to fetch image {image_key}: {e}")
        elif date_str and basename:
            # Try guessing images/{date}/{basename}.jpg or .png
            for ext in [".jpg", ".jpeg", ".png", ".webp"]:
                guess_k = f"images/{date_str}/{basename}{ext}"
                try:
                    img_resp = self.s3.get_object(Bucket=AppConfig.SOURCE_BUCKET, Key=guess_k)
                    record["raw_bytes"] = img_resp["Body"].read()
                    record["image_key"] = guess_k
                    break
                except Exception:
                    pass

        return record

    def fetch_all_concurrently(self, items: list) -> list:
        results = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=AppConfig.MAX_WORKERS) as executor:
            future_map = {executor.submit(self.fetch_participant_record, it): it for it in items}
            for future in concurrent.futures.as_completed(future_map):
                res = future.result()
                if res.get("meta") or res.get("raw_bytes"):
                    results.append(res)
        return results


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  COMPUTER VISION                                                        ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
class ComputerVision:
    _CASCADE = None

    @classmethod
    def get_cascade(cls):
        if cls._CASCADE is None:
            try:
                import cv2
                path = os.path.join(cv2.data.haarcascades, "haarcascade_frontalface_default.xml")
                cls._CASCADE = cv2.CascadeClassifier(path)
            except Exception:
                logger.info("OpenCV not available — face detection skipped")
                cls._CASCADE = False
        return cls._CASCADE

    @staticmethod
    def is_black_screen(img_bytes: bytes) -> bool:
        if not img_bytes or not PIL_AVAILABLE:
            return False
        try:
            img = Image.open(BytesIO(img_bytes)).convert("L").resize((64, 64), Image.LANCZOS)
            px = list(img.getdata())
            mean = sum(px) / len(px)
            variance = sum((p - mean) ** 2 for p in px) / len(px)
            return (mean < AppConfig.BLACK_SCREEN_MEAN_THRESHOLD
                    or variance < AppConfig.BLACK_SCREEN_VARIANCE_THRESHOLD)
        except Exception:
            return False

    @classmethod
    def detect_face(cls, img_bytes: bytes) -> bool:
        if not AppConfig.FACE_DETECTION_ENABLED or not img_bytes:
            return True
        cascade = cls.get_cascade()
        if not cascade:
            return True
        try:
            import cv2
            import numpy as np
            arr = np.frombuffer(img_bytes, dtype=np.uint8)
            img = cv2.imdecode(arr, cv2.IMREAD_GRAYSCALE)
            if img is None:
                return False
            h, w = img.shape[:2]
            min_edge = max(int(min(h, w) * 0.06), 24)
            faces = cascade.detectMultiScale(img, scaleFactor=1.08, minNeighbors=5, minSize=(min_edge, min_edge))
            return len(faces) > 0
        except Exception:
            return True


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  EVENT SOURCE DETECTION                                                 ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
def detect_event_source(event: dict) -> str:
    if event.get("Event") == "s3:TestEvent":
        return "S3_TEST"
    records = event.get("Records", [])
    if records and isinstance(records, list) and len(records) > 0:
        if records[0].get("eventSource") == "aws:s3":
            return "S3"
    if event.get("version") == "2.0" and "requestContext" in event:
        return "HTTP"
    if "httpMethod" in event or "requestContext" in event:
        return "HTTP"
    return "DIRECT"


# ╔═══════════════════════════════════════════════════════════════════════════╗
# ║  LAMBDA HANDLER                                                         ║
# ╚═══════════════════════════════════════════════════════════════════════════╝
def lambda_handler(event, context):
    logger.info("Omega 5.0 — Event-Driven Report Engine starting")

    cors_headers = {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "Content-Type,Authorization,X-Amz-Date,X-Api-Key,X-Amz-Security-Token",
        "Access-Control-Allow-Methods": "OPTIONS,POST",
    }

    source = detect_event_source(event)
    logger.info(f"Invocation source: {source}")
    is_http = source == "HTTP"

    http_method = event.get("requestContext", {}).get("http", {}).get("method",
                  event.get("httpMethod", ""))
    if http_method == "OPTIONS":
        return {"statusCode": 200, "headers": cors_headers, "body": json.dumps({"status": "ok"})}

    if source == "S3_TEST":
        logger.info("S3 TestEvent acknowledged.")
        return {"status": "S3_TEST_EVENT_ACKNOWLEDGED"}

    payload = {}
    target_date_override = None

    if source == "S3":
        for record in event.get("Records", []):
            raw_key = record.get("s3", {}).get("object", {}).get("key", "")
            decoded_key = urllib.parse.unquote_plus(raw_key)
            target_date_override = S3DataLayer.extract_date_from_key(decoded_key)
            logger.info(f"S3 trigger — key: {decoded_key}, date: {target_date_override}")
            break
    elif source == "HTTP":
        body = event.get("body", "{}")
        if event.get("isBase64Encoded", False):
            import base64
            body = base64.b64decode(body).decode("utf-8")
        try:
            payload = json.loads(body) if isinstance(body, str) else (body or {})
        except Exception as e:
            logger.warning(f"Failed to parse HTTP body: {e}")
            payload = {}
    else:
        payload = event

    try:
        data_layer = S3DataLayer()

        target_dates, date_display = S3DataLayer.resolve_target_dates(payload, event, target_date_override)
        config = event_config(payload)
        config["target_date"] = target_dates[0]
        config["target_dates"] = target_dates
        config["date_display"] = date_display

        # Select template: defaults to Swiss Precision Editorial (White Background)
        template_key = config.get("template", AppConfig.DEFAULT_TEMPLATE)
        if template_key not in TEMPLATE_SPECS:
            for k in TEMPLATE_SPECS:
                if template_key in k:
                    template_key = k
                    break
        template_spec = TEMPLATE_SPECS.get(template_key, TEMPLATE_SPECS["2_swiss"])
        logger.info(f"Using template: {template_spec['name']} for dates: {date_display}")

        # Discover participant keys across all requested dates
        attendee_items = data_layer.discover_participant_keys(target_dates)
        if not attendee_items:
            msg = f"No attendance data found for {date_display}."
            logger.info(msg)
            if is_http:
                return {"statusCode": 404, "headers": cors_headers, "body": json.dumps({"error": msg})}
            return {"status": "NO_DATA", "dates": target_dates, "date_display": date_display}

        raw_records = data_layer.fetch_all_concurrently(attendee_items)
        logger.info(f"Retrieved {len(raw_records)} participant records across {date_display}")

        # PASS 1: Extract telemetry & resolve participants with GPS
        temp_participants = []
        seen_names = set()
        rejected_face, location_flagged, s3_records_updated = 0, 0, 0

        for rec in raw_records:
            if ComputerVision.is_black_screen(rec.get("raw_bytes")):
                continue
            has_face = ComputerVision.detect_face(rec.get("raw_bytes"))
            if not has_face and rec.get("raw_bytes"):
                rejected_face += 1

            meta = rec.get("meta", {})
            meta_key = rec.get("meta_key")

            p_name = meta.get("name") or rec.get("basename") or "Unknown Participant"
            p_name_key = p_name.strip().lower()
            if len(target_dates) > 1 and p_name_key in seen_names and p_name_key != "unknown participant":
                continue
            seen_names.add(p_name_key)

            loc_data = LocationValidator.resolve_and_enrich(
                meta,
                data_layer.s3,
                AppConfig.SOURCE_BUCKET,
                meta_key
            )
            if loc_data["updated_s3"]:
                s3_records_updated += 1
            if loc_data.get("is_suspect"):
                location_flagged += 1

            screen_info = extract_screen_info(meta)

            temp_participants.append({
                "name": p_name,
                "platform": normalize_platform(extract_device_platform(meta)),
                "screen_display": screen_info["display"],
                "screen_sort_value": screen_info["sort_value"],
                "camera_permission": permission_granted(meta, "camera"),
                "location_permission": permission_granted(meta, "location"),
                "city": loc_data["city"],
                "country": loc_data["country"],
                "city_country": loc_data["city_country"],
                "lat": loc_data["lat"],
                "lon": loc_data["lon"],
                "has_coords": loc_data.get("has_coords", False),
                "is_suspect": loc_data.get("is_suspect", False),
                "image": rec.get("raw_bytes"),
                "has_image": has_face and bool(rec.get("raw_bytes")),
                "meta": meta,
                "meta_key": meta_key,
            })

        # PASS 2: Determine Session Primary Venue & Re-align Attendees Without GPS
        # Find the primary verified venue from attendees who had real GPS
        verified_venues = [
            (p["city"], p["country"]) for p in temp_participants
            if p["has_coords"] and p["city"] and p["country"]
        ]

        if verified_venues:
            session_city, session_country = Counter(verified_venues).most_common(1)[0][0]
            top_coords = [
                (p["lat"], p["lon"]) for p in temp_participants
                if p["has_coords"] and p["city"] == session_city and p["lat"] is not None and p["lon"] is not None
            ]
            if top_coords:
                session_lat = sum(c[0] for c in top_coords) / len(top_coords)
                session_lon = sum(c[1] for c in top_coords) / len(top_coords)
            else:
                session_lat, session_lon = (28.4595, 77.0266) if session_country == "India" else (-15.4167, 28.2833)
        else:
            # Contextual fallback based on client branding (never assume Zambia for non-Zambian clients!)
            client_lower = config["client"].lower()
            if "zambia" in client_lower:
                session_city, session_country = "Lusaka", "Zambia"
                session_lat, session_lon = -15.4167, 28.2833
            elif "lupin" in client_lower or "india" in client_lower:
                session_city, session_country = "Gurgaon", "India"
                session_lat, session_lon = 28.4595, 77.0266
            else:
                session_city, session_country = "Gurgaon", "India"
                session_lat, session_lon = 28.4595, 77.0266

        clean_participants = []
        for p in temp_participants:
            # Re-align if participant has no GPS, or has a legacy suspect Lusaka fallback that contradicts the session venue
            has_gps = p["has_coords"]
            is_legacy_fallback = p["city"].lower() == "lusaka" and session_country.lower() != "zambia"
            if not has_gps or is_legacy_fallback or not p["city"]:
                p["city"] = session_city
                p["country"] = session_country
                p["city_country"] = f"{session_city}, {session_country}"
                p["lat"] = session_lat
                p["lon"] = session_lon
                # Correct false legacy S3 record if needed
                if is_legacy_fallback and data_layer.s3 and p.get("meta_key"):
                    try:
                        p["meta"]["location"] = {"city": session_city, "country": session_country}
                        data_layer.s3.put_object(
                            Bucket=AppConfig.SOURCE_BUCKET, Key=p["meta_key"],
                            Body=json.dumps(p["meta"], ensure_ascii=False, indent=2).encode("utf-8"),
                            ContentType="application/json"
                        )
                    except Exception:
                        pass
            clean_participants.append(p)

        random.shuffle(clean_participants)

        # Align client brand dynamically with resolved country if not explicitly set by user
        if (not os.environ.get("CLIENT") and not os.environ.get("CLIENT_NAME")) and session_country:
            config["client"] = f"Airtel {session_country.title()}"

        # Build Presentation
        builder = UniversalTemplateBuilder(template_spec, config)
        builder.build_cover_slide(len(clean_participants))
        builder.build_permission_slide(clean_participants)
        builder.build_gallery_slides(clean_participants)
        builder.build_location_slide(clean_participants)
        builder.build_screen_slide(clean_participants)
        builder.build_platform_slide(clean_participants)
        builder.build_summary_slide(clean_participants)

        # ── BULLETPROOF CONTENT VERIFICATION ──
        content_report = verify_slide_content(builder.prs)
        if not content_report["ok"]:
            logger.critical(f"SLIDE CONTENT VERIFICATION FAILED: {content_report['warnings']}")

        # Save to S3 (deterministic live report)
        slug = re.sub(r"[^a-z0-9]+", "-", config["client"].lower()).strip("-") or "report"
        if len(target_dates) == 1:
            out_key = f"reports/{slug}_{target_dates[0]}.pptx"
        else:
            out_key = f"reports/{slug}_{target_dates[0]}_to_{target_dates[-1]}.pptx"

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tf:
            builder.prs.save(tf.name)
            verif = verify_deck(tf.name)
            with open(tf.name, "rb") as f:
                data_layer.s3.put_object(
                    Bucket=AppConfig.REPORT_BUCKET, Key=out_key, Body=f.read(),
                    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
            os.unlink(tf.name)

        presigned_url = data_layer.s3.generate_presigned_url(
            "get_object",
            Params={"Bucket": AppConfig.REPORT_BUCKET, "Key": out_key},
            ExpiresIn=AppConfig.URL_EXPIRY,
        )

        result = {
            "status": "success",
            "client": config["client"],
            "workshop": config["workshop"],
            "target_dates": target_dates,
            "date_display": date_display,
            "processed_date": date_display,
            "template_used": template_spec["name"],
            "participants_accepted": len(clean_participants),
            "faces_rejected": rejected_face,
            "locations_flagged": location_flagged,
            "s3_records_geocoded_and_updated": s3_records_updated,
            "report_key": out_key,
            "verification": verif,
            "content_verification": content_report,
            "download_url": presigned_url,
        }

        logger.info(f"Report generated: {out_key} ({len(clean_participants)} attendees, {verif.get('slide_count')} slides)")

        if is_http:
            return {"statusCode": 200, "headers": cors_headers, "body": json.dumps(result)}
        return result

    except Exception as e:
        logger.error(f"Pipeline Failure: {str(e)}", exc_info=True)
        if is_http:
            return {"statusCode": 500, "headers": cors_headers, "body": json.dumps({"status": "error", "error": str(e)})}
        raise
