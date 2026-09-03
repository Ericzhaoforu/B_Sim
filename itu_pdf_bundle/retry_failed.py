#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
import random
import shutil
import time
from pathlib import Path

import download as dl

FAILED_ROWS = {
    45, 49, 50, 51, 52, 53, 54, 56, 58, 60, 61, 63, 65, 66, 71,
    72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86,
    87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99,
}

OUT_ROOT = Path("out") / "ITU-R建议书1200-1400MHz汇总（第二轮补充）"
LOG_DIR = OUT_ROOT / "清单与日志"

# Use a fresh runner and a sequential pace. If the ITU web-application firewall
# returns an HTML "Request Rejected" page with HTTP 200, retry the same URL after
# warming the main site and backing off.
_original_request_bytes = dl.request_bytes


def robust_request_bytes(url: str, timeout: int = 120, retries: int = 3, headers=None) -> bytes:
    last = None
    max_attempts = max(5, retries)
    for attempt in range(max_attempts):
        try:
            data = _original_request_bytes(url, timeout=timeout, retries=1, headers=headers)
            head = data[:8192]
            if b"Request Rejected" not in head and b"The requested URL was rejected" not in head:
                return data
            last = RuntimeError("ITU WAF returned Request Rejected HTML")
            try:
                _original_request_bytes("https://www.itu.int/", timeout=30, retries=1)
            except Exception:
                pass
        except Exception as exc:
            last = exc
            if "HTTPError 404" in repr(exc) or "HTTPError 410" in repr(exc):
                break
        time.sleep(4 + attempt * 4 + random.random() * 2)
    raise RuntimeError(f"robust download failed: {url}: {last!r}")


dl.request_bytes = robust_request_bytes


def download_item(item):
    work = dict(item)
    # The workbook contains a stale/nonexistent RA.611-1 reference. The official
    # series page identifies RA.611-4 (03/2006) as the current in-force release.
    if work.get("version") == "RA.611-1":
        work["original_version"] = work["version"]
        work["version"] = "RA.611-4"
        work["cn_url"] = "https://www.itu.int/dms_pubrec/itu-r/rec/ra/R-REC-RA.611-4-200603-I!!PDF-C.pdf"
        work["en_url"] = "https://www.itu.int/dms_pubrec/itu-r/rec/ra/R-REC-RA.611-4-200603-I!!PDF-E.pdf"
    result = dl.download_one(work)
    result["listed_version"] = item.get("version")
    result["resolved_version"] = work.get("version")
    result["row"] = item.get("row")
    result["target"] = item.get("target")
    return result


def main() -> int:
    items = [item for item in dl.load_manifest() if int(item.get("row", -1)) in FAILED_ROWS]
    if OUT_ROOT.exists():
        shutil.rmtree(OUT_ROOT)
    dl.OUT_ROOT = OUT_ROOT
    dl.LOG_DIR = LOG_DIR
    dl.PDF_ROOT = OUT_ROOT / "PDF"
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    results = []
    for index, item in enumerate(items, 1):
        result = download_item(item)
        results.append(result)
        print(
            f"[{index:02d}/{len(items)}] {item['version']}: "
            f"{result.get('status')} {result.get('language') or ''} {result.get('bytes') or 0}",
            flush=True,
        )
        time.sleep(1.5 + random.random())

    ok = sum(r.get("status") == "ok" for r in results)
    summary = {
        "requested": len(items),
        "downloaded": ok,
        "failed": len(items) - ok,
        "chinese": sum(r.get("status") == "ok" and r.get("language") == "C" for r in results),
        "english": sum(r.get("status") == "ok" and r.get("language") == "E" for r in results),
        "total_bytes": sum(int(r.get("bytes") or 0) for r in results),
    }
    (LOG_DIR / "第二轮下载结果.json").write_text(
        json.dumps({"summary": summary, "results": results}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    fields = [
        "row", "rec", "listed_version", "resolved_version", "status", "language",
        "source_kind", "bytes", "sha256", "target", "source_url",
        "metadata_source", "metadata_error",
    ]
    with (LOG_DIR / "第二轮下载结果.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(results)
    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
