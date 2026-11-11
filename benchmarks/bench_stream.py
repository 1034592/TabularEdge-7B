"""Reproducible large-workbook benchmark for the streaming XLSX engine (tabularedge/xlsx_stream.py).

    python test_cases/make_big.py /data/big            # 150 MB synthetic registry, 761,507 rows x 25 cols (~3.3M shared strings)
    python benchmarks/bench_stream.py /data/big         # profile + 2,000-key lookup, prints JSON
    python benchmarks/bench_stream.py /data/big --openpyxl   # also time openpyxl read_only open (needs several GB RAM)

Measures, in one process:
  open_s      : open the zip, read sheet names and <dimension> (what the desktop client shows when a file is selected)
  head_s      : header + first 200 rows with shared strings resolved lazily (what goes into the model prompt)
  lookup_s    : 2,000 keys from a small list matched against the 761k-row table, returning 2 columns
  peak_rss_mb : peak resident memory of the whole process (ru_maxrss), i.e. the entire run, not just one step
and checks correctness: odd-numbered list rows exist in the big table, even-numbered ones do not.
"""
from __future__ import annotations

import argparse
import json
import resource
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from tabularedge.xlsx_stream import XlsxStream, norm_key  # noqa: E402

BIG, SMALL = "单位名录_76万行.xlsx", "商户名单.xlsx"


def rss_mb() -> float:
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024  # Linux: KiB


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dir", type=Path)
    ap.add_argument("--openpyxl", action="store_true", help="also time openpyxl load_workbook(read_only=True) for comparison")
    a = ap.parse_args()
    big, small = a.dir / BIG, a.dir / SMALL
    out = {"file_mb": round(big.stat().st_size / 2**20, 1)}

    t = time.perf_counter()
    xs = XlsxStream(big)
    sheet = xs.sheet_names[0]
    out["dims"] = xs.dims(sheet)
    out["open_s"] = round(time.perf_counter() - t, 3)

    t = time.perf_counter()
    rows = xs.head(sheet, 200)
    out["head_s"] = round(time.perf_counter() - t, 3)
    out["header"] = [str(v) for v in rows[0][:3]] + ["…"]

    from openpyxl import load_workbook
    ws = load_workbook(small, read_only=True).worksheets[0]
    lst = [(r[0], r[2]) for r in ws.iter_rows(min_row=2, values_only=True)]
    keys = {norm_key(c) for _, c in lst}
    t = time.perf_counter()
    hits = xs.lookup(sheet, 1, [2, 15], keys)
    out["lookup_s"] = round(time.perf_counter() - t, 2)
    xs.close()
    found = {k for k in keys if k in hits}
    odd = {norm_key(c) for i, c in lst if i % 2}
    out["keys"] = len(keys)
    out["hits"] = len(found)
    out["correct"] = found == odd
    out["peak_rss_mb"] = round(rss_mb(), 1)

    if a.openpyxl:
        t = time.perf_counter()
        load_workbook(big, read_only=True)
        out["openpyxl_read_only_open_s"] = round(time.perf_counter() - t, 1)
        out["peak_rss_mb_with_openpyxl"] = round(rss_mb(), 1)
    print(json.dumps(out, ensure_ascii=False))


if __name__ == "__main__":
    main()
