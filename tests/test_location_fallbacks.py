import sys
import os

from attendancehandler import normalize_location, reverse_geocode_location

print("--- TEST 1: GPS Granted ---")
meta1 = {
    "location": {
        "permissionStatus": "granted",
        "latitude": 6.4682,
        "longitude": 3.4472
    }
}
loc1 = normalize_location(meta1)
assert loc1["permissionStatus"] == "granted", loc1
assert loc1["latitude"] == 6.4682, loc1
print("Test 1 Passed: Real GPS preserved:", loc1["latitude"], loc1["longitude"])

print("\n--- TEST 2: Location Denied / Blank ---")
meta2 = {
    "location": {
        "permissionStatus": "denied"
    }
}
loc2 = normalize_location(meta2)
assert loc2["permissionStatus"] == "", loc2
assert loc2["latitude"] == "", loc2
print("Test 2 Passed: Denied status properly cleared:", loc2)

print("\n--- TEST 3: Reverse Geocoding with Valid Coords ---")
loc3_input = {
    "permissionStatus": "granted",
    "latitude": 6.4682,
    "longitude": 3.4472
}
loc3 = reverse_geocode_location(loc3_input)
assert loc3["permissionStatus"] == "granted", loc3
print("Test 3 Passed: Reverse geocode processed cleanly:", loc3)

print("\nALL ATTENDANCEHANDLER LOCATION TESTS PASSED!")
