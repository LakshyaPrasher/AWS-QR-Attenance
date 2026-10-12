import os
import tempfile
import boto3
import json
import re
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill

# =========================
# CONFIG
# =========================
BUCKET = "attendance-data-lakshyaprasher"
JSON_PREFIX = "records/"
TEXT_PREFIX = "raw-text/"
OUTPUT_KEY = "compiled/attendance.xlsx"

HEADERS = [
    "S No",
    "Name",
    "Time (UTC)",
    "Typing Time (s)",
    "Platform",
    "Language",
    "Screen",
    "Timezone",
    "Network",
    "IP",
    "Camera Permission",
    "Location Permission",
    "Latitude",
    "Longitude",
    "Accuracy Meters",
    "State",
    "City",
    "Street Name",
    "Altitude",
    "Altitude Accuracy",
    "Heading",
    "Speed",
    "Location Timestamp"
]

s3 = boto3.client("s3")

# =========================
# HELPERS
# =========================
def grab(label, text):
    pattern = rf"^{re.escape(label)}:[ \t]*(.*)$"
    m = re.search(pattern, text, flags=re.MULTILINE)
    return m.group(1).strip() if m else ""

def is_granted(value):
    return (value or "").strip().lower() == "granted"

def autosize_and_style(ws):
    header_fill = PatternFill("solid", fgColor="DFF4E7")
    header_font = Font(bold=True)

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(
            vertical="center",
            horizontal="center",
            wrap_text=True
        )

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)

        for cell in col:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))

        ws.column_dimensions[col_letter].width = min(max(max_length + 2, 14), 36)

def normalize_json_record(data):
    location = data.get("location", {})
    if not isinstance(location, dict):
        location = {}

    return {
        "Name": data.get("name", ""),
        "Time": data.get("timestamp", ""),
        "Typing": data.get("typingTimeSeconds", ""),
        "Platform": data.get("platform", ""),
        "Language": data.get("language", ""),
        "Screen": data.get("screen", ""),
        "Timezone": data.get("timezone", ""),
        "Network": data.get("network", ""),
        "IP": data.get("ip", ""),
        "CameraPermission": data.get("cameraPermission", ""),
        "LocationPermission": data.get("locationPermission", ""),
        "Latitude": location.get("latitude", ""),
        "Longitude": location.get("longitude", ""),
        "Accuracy": location.get("accuracy", ""),
        "State": location.get("state", ""),
        "City": location.get("city", ""),
        "StreetName": location.get("streetName", ""),
        "Altitude": location.get("altitude", ""),
        "AltitudeAccuracy": location.get("altitudeAccuracy", ""),
        "Heading": location.get("heading", ""),
        "Speed": location.get("speed", ""),
        "LocationTimestamp": location.get("timestamp", "")
    }

def normalize_text_record(content):
    return {
        "Name": grab("Name", content),
        "Time": grab("Time (UTC)", content),
        "Typing": grab("Typing Time (s)", content),
        "Platform": grab("Platform", content),
        "Language": grab("Language", content),
        "Screen": grab("Screen", content),
        "Timezone": grab("Timezone", content),
        "Network": grab("Connection", content),
        "IP": grab("IP", content),
        "CameraPermission": grab("Camera Permission Status", content),
        "LocationPermission": grab("Permission Status", content),
        "Latitude": grab("Latitude", content),
        "Longitude": grab("Longitude", content),
        "Accuracy": grab("Accuracy Meters", content),
        "State": grab("State", content),
        "City": grab("City", content),
        "StreetName": grab("Street Name", content),
        "Altitude": grab("Altitude", content),
        "AltitudeAccuracy": grab("Altitude Accuracy", content),
        "Heading": grab("Heading", content),
        "Speed": grab("Speed", content),
        "LocationTimestamp": grab("Location Timestamp", content)
    }

def add_record(records, record):
    name = record.get("Name", "")
    time_value = record.get("Time", "")

    if not name:
        return

    unique_key = f"{name.lower()}|{time_value}"
    records[unique_key] = record

# =========================
# HANDLER
# =========================
def lambda_handler(event, context):

    records = {}
    paginator = s3.get_paginator("list_objects_v2")

    # Read JSON records
    for page in paginator.paginate(Bucket=BUCKET, Prefix=JSON_PREFIX):
        for obj in page.get("Contents", []):
            key = obj["Key"]

            if not key.endswith(".json"):
                continue

            try:
                content = s3.get_object(
                    Bucket=BUCKET,
                    Key=key
                )["Body"].read().decode("utf-8", errors="ignore")

                data = json.loads(content)
                record = normalize_json_record(data)
                add_record(records, record)

            except Exception:
                continue

    # Read old TXT records as fallback
    for page in paginator.paginate(Bucket=BUCKET, Prefix=TEXT_PREFIX):
        for obj in page.get("Contents", []):
            key = obj["Key"]

            if not key.endswith(".txt"):
                continue

            try:
                content = s3.get_object(
                    Bucket=BUCKET,
                    Key=key
                )["Body"].read().decode("utf-8", errors="ignore")

                record = normalize_text_record(content)
                add_record(records, record)

            except Exception:
                continue

    sorted_records = sorted(
        records.values(),
        key=lambda r: r.get("Time", "")
    )

    total_attendance = len(sorted_records)

    location_permission_granted = sum(
        1 for r in sorted_records
        if is_granted(r.get("LocationPermission", ""))
    )

    camera_permission_granted = sum(
        1 for r in sorted_records
        if is_granted(r.get("CameraPermission", ""))
    )

    wb = Workbook()

    # =========================
    # ATTENDANCE SHEET
    # =========================
    ws = wb.active
    ws.title = "Attendance"
    ws.append(HEADERS)

    for idx, r in enumerate(sorted_records, start=1):
        ws.append([
            idx,
            r.get("Name", ""),
            r.get("Time", ""),
            r.get("Typing", ""),
            r.get("Platform", ""),
            r.get("Language", ""),
            r.get("Screen", ""),
            r.get("Timezone", ""),
            r.get("Network", ""),
            r.get("IP", ""),
            r.get("CameraPermission", ""),
            r.get("LocationPermission", ""),
            r.get("Latitude", ""),
            r.get("Longitude", ""),
            r.get("Accuracy", ""),
            r.get("State", ""),
            r.get("City", ""),
            r.get("StreetName", ""),
            r.get("Altitude", ""),
            r.get("AltitudeAccuracy", ""),
            r.get("Heading", ""),
            r.get("Speed", ""),
            r.get("LocationTimestamp", "")
        ])

    autosize_and_style(ws)
    ws.freeze_panes = "A2"

    # =========================
    # SUMMARY SHEET
    # =========================
    summary = wb.create_sheet("Summary")

    summary_rows = [
        ["Metric", "Total"],
        ["Total Attendance Marked", total_attendance],
        ["Camera Permission Granted", camera_permission_granted],
        ["Location Permission Granted", location_permission_granted],
        ["Camera Permission Blank/Denied", total_attendance - camera_permission_granted],
        ["Location Permission Blank/Denied", total_attendance - location_permission_granted]
    ]

    for row in summary_rows:
        summary.append(row)

    autosize_and_style(summary)
    summary.freeze_panes = "A2"

    path = os.path.join(tempfile.gettempdir(), f"attendance_{os.getpid()}.xlsx")
    wb.save(path)

    try:
        with open(path, "rb") as f:
            s3.put_object(
                Bucket=BUCKET,
                Key=OUTPUT_KEY,
                Body=f,
                ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    finally:
        if os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass

    return {
        "status": "updated",
        "totalAttendance": total_attendance,
        "cameraPermissionGranted": camera_permission_granted,
        "locationPermissionGranted": location_permission_granted
    }
