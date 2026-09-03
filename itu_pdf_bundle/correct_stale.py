#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
import random
import shutil
import time
from pathlib import Path

import retry_failed as rf

CORRECTIONS = {
    45: ("RA.611-4", "200603", "I"),
    50: ("F.759-0", "199203", "S"),
    51: ("BO.789-3", "202603", "I"),
    52: ("BO.1130-5", "202602", "I"),
    53: ("BO.1383-0", "199812", "I"),
    54: ("BS.1547-0", "200111", "I"),
    72: ("P.2108-1", "202109", "I"),
    86: ("M.1179-0", "199510", "I"),
    88: ("M.1372-1", "200306", "I"),
    90: ("M.1634-0", "200306", "I"),
}

OUT_ROOT = Path("out") / "ITU-R建议书1200-1400MHz汇总（版本校正补充）"
LOG_DIR = OUT_ROOT / "清单与日志"


def make_url(version: str, yyyymm: str, stage: str, lang: str) -> str:
    series = version.split(".", 1)[0].lower()
    return (
        f"https://www.itu.int/dms_pubrec/itu-r/rec/{series}/"
        f"R-REC-{version}-{yyyymm}-{stage}!!PDF-{lang}.pdf"
    )


def main() -> int:
    if OUT_ROOT.exists():
        shutil.rmtree(OUT_ROOT)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    rf.dl.OUT_ROOT = OUT_ROOT
    rf.dl.LOG_DIR = LOG_DIR
    rf.dl.PDF_ROOT = OUT_ROOT / "PDF"

    by_row = {int(item["row"]): item for item in rf.dl.load_manifest()}
    results = []
    for index, (row, (version, yyyymm, stage)) in enumerate(CORRECTIONS.items(), 1):
        original = by_row[row]
        item = dict(original)
        item["original_version"] = item["version"]
        item["version"] = version
        item["cn_url"] = make_url(version, yyyymm, stage, "C")
        item["en_url"] = make_url(version, yyyymm, stage, "E")
        result = rf.dl.download_one(item)
        result["row"] = row
        result["listed_version"] = original["version"]
        result["resolved_version"] = version
        result["target"] = original["target"]
        results.append(result)
        print(
            f"[{index:02d}/{len(CORRECTIONS)}] {original['version']} -> {version}: "
            f"{result.get('status')} {result.get('language') or ''} {result.get('bytes') or 0}",
            flush=True,
        )
        time.sleep(1.5 + random.random())

    summary = {
        "requested": len(results),
        "downloaded": sum(r.get("status") == "ok" for r in results),
        "failed": sum(r.get("status") != "ok" for r in results),
        "chinese": sum(r.get("status") == "ok" and r.get("language") == "C" for r in results),
        "english": sum(r.get("status") == "ok" and r.get("language") == "E" for r in results),
        "total_bytes": sum(int(r.get("bytes") or 0) for r in results),
    }
    (LOG_DIR / "版本校正下载结果.json").write_text(
        json.dumps({"summary": summary, "results": results}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    fields = [
        "row", "rec", "listed_version", "resolved_version", "status", "language",
        "source_kind", "bytes", "sha256", "target", "source_url",
    ]
    with (LOG_DIR / "版本校正下载结果.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(results)
    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
