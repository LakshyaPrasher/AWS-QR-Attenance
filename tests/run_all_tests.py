"""Run every verification script and print a one-line summary. Usage: py run_all_tests.py"""
import subprocess
import sys

TESTS = [
    "test_location_fallbacks.py",
    "test_platform_consolidation.py",
    "test_final_verification.py",
    "test_full_pipeline.py",
    "test_nigeria_default.py",
]

failed = []
for t in TESTS:
    code = subprocess.call([sys.executable, t])
    print("%s  %s" % ("PASS" if code == 0 else "FAIL", t))
    if code != 0:
        failed.append(t)

print("\n" + ("ALL %d TEST SCRIPTS PASSED" % len(TESTS) if not failed else "FAILED: " + ", ".join(failed)))
sys.exit(1 if failed else 0)
