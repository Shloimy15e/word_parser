#!/usr/bin/env python
# Test regenerating a single file with DAF mode
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import subprocess

# Run the main script with DAF mode
cmd = [
    sys.executable,
    "main.py",
    "--book", "תלמוד",
    "--sefer", "ביצה",
    "--docs", r"backup_input\שס 2\ביצה",
    "--out", "test_daf_single",
    "--daf-mode"
]

print("Running:", " ".join(cmd))
print("=" * 80)

result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
print(result.stdout)
if result.stderr:
    print("STDERR:", result.stderr)
print("Exit code:", result.exitcode)

