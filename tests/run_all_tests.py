"""Run every verification script and print a one-line summary. Usage: py run_all_tests.py"""
import os
import subprocess
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TESTS = [
    "test_location_fallbacks.py",
    "test_platform_consolidation.py",
    "test_final_verification.py",
    "test_full_pipeline.py",
    "test_nigeria_default.py",
]

print(f"Running verification suite under Python {sys.version.split()[0]}")
print(f"Base directory: {BASE_DIR}\n")

failed = []
for t in TESTS:
    # Support both flat layout and nested tests/ layout
    t_path = os.path.join(BASE_DIR, t)
    if not os.path.isfile(t_path):
        alt_path = os.path.join(BASE_DIR, "tests", t)
        if os.path.isfile(alt_path):
            t_path = alt_path
    cwd = os.path.dirname(t_path) if os.path.isfile(t_path) else BASE_DIR
    code = subprocess.call([sys.executable, t_path], cwd=cwd)
    print("%s  %s" % ("PASS" if code == 0 else "FAIL", t))
    if code != 0:
        failed.append(t)

print("\n" + ("ALL %d TEST SCRIPTS PASSED" % len(TESTS) if not failed else "FAILED: " + ", ".join(failed)))
sys.exit(1 if failed else 0)

