#!/usr/bin/env python3
"""Oracle reference solution entry (produces /output/result.xlsx).

Launched by solution/solve.sh. Reads the exam input workbook and writes the
canonical answer workbook. The grader (tests/test_outputs.py) imports the same
engine module, so Oracle == golden by construction.
"""
import os
import sys

# Allow running both inside the image (/solution) and locally.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from oracle_engine import compute_from_file  # noqa: E402

INPUT = "/data/init.xlsx"
OUTPUT = "/output/result.xlsx"


def main():
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    compute_from_file(INPUT, OUTPUT)
    print(f"oracle wrote {OUTPUT}")


if __name__ == "__main__":
    main()
