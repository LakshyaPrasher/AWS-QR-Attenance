"""Run every verification script and print a one-line summary. Usage: py run_all_tests.py"""
import os
import subprocess
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(BASE_DIR, "..")) if os.path.basename(BASE_DIR) == "tests" else BASE_DIR

TESTS = [
    "test_location_fallbacks.py",
    "test_platform_consolidation.py",
    "test_final_verification.py",
    "test_full_pipeline.py",
    "test_nigeria_default.py",
    "test_excel_compiler.py",
]

print(f"Running verification suite under Python {sys.version.split()[0]}")
print(f"Base directory: {BASE_DIR}")
print(f"Project root:   {ROOT_DIR}\n")

env = os.environ.copy()
env["PYTHONPATH"] = ROOT_DIR + os.pathsep + env.get("PYTHONPATH", "")

failed = []
for t in TESTS:
    # Support both flat layout and nested tests/ layout
    t_path = os.path.join(BASE_DIR, t)
    if not os.path.isfile(t_path):
        alt_path = os.path.join(ROOT_DIR, "tests", t)
        if os.path.isfile(alt_path):
            t_path = alt_path
        elif os.path.isfile(os.path.join(ROOT_DIR, t)):
            t_path = os.path.join(ROOT_DIR, t)
    cwd = os.path.dirname(t_path) if os.path.isfile(t_path) else BASE_DIR
    code = subprocess.call([sys.executable, t_path], cwd=cwd, env=env)
    print("%s  %s" % ("PASS" if code == 0 else "FAIL", t))
    if code != 0:
        failed.append(t)

print("\n" + ("ALL %d TEST SCRIPTS PASSED" % len(TESTS) if not failed else "FAILED: " + ", ".join(failed)))
sys.exit(1 if failed else 0)

