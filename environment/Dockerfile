FROM ubuntu:22.04

# Only dependency needed: python3 + openpyxl (the engine reads/writes xlsx via
# openpyxl and uses the Python stdlib for all arithmetic). pandas is intentionally
# NOT installed -- leaner image, faster build, fewer failure surfaces.
RUN apt-get update && apt-get install -y \
    python3 \
    python3-openpyxl \
    && rm -rf /var/lib/apt/lists/*

RUN mkdir -p /data /output /logs/verifier /solution /tests

# Exam input (dirty loan tape + explicit Params).
COPY environment/sheet/init.xlsx /data/init.xlsx

# Oracle reference solution.
COPY solution/oracle_engine.py /solution/oracle_engine.py
COPY solution/oracle_convert.py /solution/oracle_convert.py
COPY solution/solve.sh /solution/solve.sh
RUN chmod +x /solution/solve.sh

# Verifier.
COPY tests/test_outputs.py /tests/test_outputs.py
COPY tests/test.sh /tests/test.sh
RUN chmod +x /tests/test.sh

CMD ["/bin/bash"]
