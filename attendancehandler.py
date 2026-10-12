import json
import base64
import boto3
import datetime
import re
import math
try:
    import requests
except ImportError:
    requests = None
import urllib.request
import urllib.parse
from email.parser import BytesParser
from email.policy import default

def http_get_json(url, params=None, headers=None, timeout=5):
    if requests is not None:
        try:
            return requests.get(url, params=params, headers=headers, timeout=timeout).json()
        except Exception:
            pass
    try:
        full_url = url
        if params:
            full_url = f"{url}?{urllib.parse.urlencode(params)}"
        req = urllib.request.Request(full_url, headers=headers or {"User-Agent": "AttendanceHandler/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return {}

# =========================
# CONFIG
# =========================
DATA_BUCKET = "attendance-data-lakshyaprasher"
IMAGE_PREFIX = "images"
TEXT_PREFIX = "raw-text"
JSON_PREFIX = "records"

s3 = boto3.client("s3")

# =========================
# HELPERS
# =========================
def response(status_code, body):
    return {
        "statusCode": status_code,
        "headers": {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Headers": "*",
            "Access-Control-Allow-Methods": "OPTIONS,POST,GET"
        },
        "body": json.dumps(body)
    }

def safe_name(name):
    name = name or "unknown"
    name = re.sub(r"[^a-zA-Z0-9\s_-]", "", name)
    name = re.sub(r"\s+", "_", name.strip().lower())
    return name or "unknown"

def normalize_platform(metadata):
    platform = (metadata.get("platform") or "").strip()
    raw_platform = (metadata.get("rawPlatform") or "").lower()
    ua = (metadata.get("userAgent") or "").lower()

    if platform in ["Android", "IOS", "Windows", "Linux"]:
        return platform

    if "android" in ua:
        return "Android"

    if any(x in ua for x in ["iphone", "ipad", "ipod"]):
        return "IOS"

    if "win" in raw_platform or "windows" in ua:
        return "Windows"

    if "linux" in raw_platform or "linux" in ua:
        return "Linux"

    return ""

def determine_screen(metadata):
    try:
        sw = int(metadata.get("screenWidth") or 0)
        sh = int(metadata.get("screenHeight") or 0)
        dpr = float(metadata.get("devicePixelRatio") or 1)
    except:
        return "Not Determined"

    if sw <= 0 or sh <= 0:
        return "Not Determined"

    ua = (metadata.get("userAgent") or "").lower()
    diag_px = math.sqrt((sw * dpr) ** 2 + (sh * dpr) ** 2)

    if any(x in ua for x in ["iphone", "android"]):
        ppi = 460 if dpr >= 3 else 400 if dpr >= 2 else 320
        return f"{round(diag_px / ppi, 1)} inches"

    ppi_candidates = [96, 100, 110]
    inches = [round(diag_px / p, 1) for p in ppi_candidates]
    best = min(inches, key=lambda x: abs(x - 15))
    return f"{best} inches"

def determine_network(metadata, ip):
    base = "wifi"
    eff = (metadata.get("effectiveType") or "").lower()

    if eff in ["2g", "3g", "4g", "5g"]:
        base = "data"

    try:
        if ip:
            g = http_get_json(f"https://ipapi.co/{ip}/json/", timeout=2)
            org = (g.get("org") or "").lower()

            if any(k in org for k in ["vpn", "hosting", "cloud", "aws", "azure", "google"]):
                return f"vpn with {base}"
    except:
        pass

    return base

def normalize_camera(metadata):
    camera = metadata.get("camera", {})
    if not isinstance(camera, dict):
        camera = {}

    status = camera.get("permissionStatus", "")
    if status != "granted":
        status = ""

    return status

def normalize_location(metadata):
    location = metadata.get("location", {})
    if not isinstance(location, dict):
        location = {}

    permission_status = location.get("permissionStatus", "")

    if permission_status != "granted":
        return {
            "permissionStatus": "",
            "latitude": "",
            "longitude": "",
            "accuracy": "",
            "altitude": "",
            "altitudeAccuracy": "",
            "heading": "",
            "speed": "",
            "timestamp": "",
            "state": "",
            "city": "",
            "streetName": ""
        }

    return {
        "permissionStatus": "granted",
        "latitude": location.get("latitude", ""),
        "longitude": location.get("longitude", ""),
        "accuracy": location.get("accuracy", ""),
        "altitude": location.get("altitude", ""),
        "altitudeAccuracy": location.get("altitudeAccuracy", ""),
        "heading": location.get("heading", ""),
        "speed": location.get("speed", ""),
        "timestamp": location.get("timestamp", ""),
        "state": "",
        "city": "",
        "streetName": ""
    }

def reverse_geocode_location(location):
    lat = location.get("latitude", "")
    lon = location.get("longitude", "")

    if location.get("permissionStatus") != "granted":
        return location

    if lat in (None, "") or lon in (None, ""):
        return location

    try:
        url = "https://nominatim.openstreetmap.org/reverse"

        params = {
            "format": "jsonv2",
            "lat": str(lat),
            "lon": str(lon),
            "addressdetails": 1
        }

        headers = {
            "User-Agent": "LupinAttendanceApp/1.0 contact-lakshyaprasher"
        }

        result = http_get_json(
            url,
            params=params,
            headers=headers,
            timeout=5
        )

        address = result.get("address", {}) or {}

        state = (
            address.get("state")
            or address.get("region")
            or address.get("state_district")
            or ""
        )

        city = (
            address.get("city")
            or address.get("town")
            or address.get("village")
            or address.get("municipality")
            or address.get("county")
            or address.get("state_district")
            or ""
        )

        street_name = (
            address.get("road")
            or address.get("pedestrian")
            or address.get("footway")
            or address.get("residential")
            or address.get("neighbourhood")
            or address.get("suburb")
            or ""
        )

        location["state"] = state
        location["city"] = city
        location["streetName"] = street_name

    except Exception:
        location["state"] = ""
        location["city"] = ""
        location["streetName"] = ""

    return location

def build_text(r):
    loc = r["location"]

    return "\n".join([
        "ATTENDANCE RECORD",
        "-----------------",
        f"Name: {r['name']}",
        f"Time (UTC): {r['timestamp']}",
        f"Typing Time (s): {r['typingTimeSeconds']}",
        "",
        "DEVICE",
        "------",
        f"Platform: {r['platform']}",
        f"Language: {r['language']}",
        f"Screen: {r['screen']}",
        f"Timezone: {r['timezone']}",
        "",
        "NETWORK",
        "-------",
        f"Connection: {r['network']}",
        f"IP: {r['ip']}",
        "",
        "CAMERA",
        "------",
        f"Camera Permission Status: {r['cameraPermission']}",
        "",
        "LOCATION",
        "--------",
        f"Permission Status: {loc['permissionStatus']}",
        f"Latitude: {loc['latitude']}",
        f"Longitude: {loc['longitude']}",
        f"Accuracy Meters: {loc['accuracy']}",
        f"State: {loc['state']}",
        f"City: {loc['city']}",
        f"Street Name: {loc['streetName']}",
        f"Altitude: {loc['altitude']}",
        f"Altitude Accuracy: {loc['altitudeAccuracy']}",
        f"Heading: {loc['heading']}",
        f"Speed: {loc['speed']}",
        f"Location Timestamp: {loc['timestamp']}"
    ])

# =========================
# HANDLER
# =========================
def lambda_handler(event, context):

    method = (
        event.get("requestContext", {}).get("http", {}).get("method")
        or event.get("httpMethod")
        or ""
    )

    if method == "OPTIONS":
        return response(200, {"success": True})

    if "body" not in event or event.get("body") in (None, ""):
        return response(400, {
            "success": False,
            "message": "Missing request body. Expected multipart/form-data."
        })

    headers = {k.lower(): v for k, v in (event.get("headers") or {}).items()}
    content_type = headers.get("content-type", "")

    if not content_type:
        return response(400, {
            "success": False,
            "message": "Missing Content-Type header. Expected multipart/form-data."
        })

    try:
        if event.get("isBase64Encoded"):
            body = base64.b64decode(event["body"])
        else:
            body = event["body"].encode()
    except Exception as e:
        return response(400, {
            "success": False,
            "message": "Unable to decode request body.",
            "error": str(e)
        })

    try:
        msg = BytesParser(policy=default).parsebytes(
            b"Content-Type: " + content_type.encode() + b"\r\n\r\n" + body
        )
    except Exception as e:
        return response(400, {
            "success": False,
            "message": "Unable to parse multipart form-data.",
            "error": str(e)
        })

    files = {}
    data = {}

    try:
        for part in msg.iter_parts():
            name = part.get_param("name", header="content-disposition")

            if not name:
                continue

            if part.get_content_type().startswith("image"):
                files[name] = part.get_payload(decode=True)
            else:
                payload = part.get_payload(decode=True)
                data[name] = payload.decode("utf-8", errors="ignore") if payload else ""
    except Exception as e:
        return response(400, {
            "success": False,
            "message": "Unable to read multipart parts.",
            "error": str(e)
        })

    try:
        metadata = json.loads(data.get("metadata", "{}"))
    except:
        metadata = {}

    raw_name = metadata.get("name", "unknown")
    person = safe_name(raw_name)

    now = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None).isoformat() + "Z"
    date_str = now[:10]

    ip = headers.get("x-forwarded-for", "").split(",")[0].strip()

    typing_time = metadata.get("typingTimeSeconds")
    if typing_time in (None, "", "null"):
        typing_time = "0.00"

    camera_permission = normalize_camera(metadata)
    location_data = normalize_location(metadata)
    
    record = {
        "name": raw_name,
        "timestamp": now,
        "typingTimeSeconds": typing_time,

        "platform": normalize_platform(metadata),
        "language": metadata.get("language", ""),
        "userAgent": metadata.get("userAgent", ""),
        "screen": determine_screen(metadata),
        "timezone": metadata.get("timezone", ""),

        "network": determine_network(metadata, ip),
        "ip": ip,

        "cameraPermission": camera_permission,
        "locationPermission": location_data["permissionStatus"],

        "location": location_data
    }

    try:
        if "image" in files and files["image"]:
            s3.put_object(
                Bucket=DATA_BUCKET,
                Key=f"{IMAGE_PREFIX}/{date_str}/{person}.jpg",
                Body=files["image"],
                ContentType="image/jpeg"
            )

        s3.put_object(
            Bucket=DATA_BUCKET,
            Key=f"{TEXT_PREFIX}/{date_str}/{person}.txt",
            Body=build_text(record),
            ContentType="text/plain"
        )

        s3.put_object(
            Bucket=DATA_BUCKET,
            Key=f"{JSON_PREFIX}/{date_str}/{person}.json",
            Body=json.dumps(record, indent=2),
            ContentType="application/json"
        )

    except Exception as e:
        return response(500, {
            "success": False,
            "message": "Failed to save attendance data to S3.",
            "error": str(e)
        })

    return response(200, {
        "success": True,
        "message": "Attendance marked successfully.",
        "cameraPermission": record["cameraPermission"],
        "locationPermission": record["locationPermission"],
        "state": record["location"]["state"],
        "city": record["location"]["city"],
        "streetName": record["location"]["streetName"]
    })