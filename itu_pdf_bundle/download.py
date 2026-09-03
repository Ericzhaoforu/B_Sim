#!/usr/bin/env python3
from __future__ import annotations

import base64
import concurrent.futures
import csv
import hashlib
import html
import json
import os
import random
import re
import shutil
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zlib
from pathlib import Path
from typing import Any

MANIFEST_B64 = "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"
OUT_ROOT = Path("out") / "ITU-R建议书1200-1400MHz汇总（PDF已下载）"
LOG_DIR = OUT_ROOT / "清单与日志"
PDF_ROOT = OUT_ROOT / "PDF"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140 Safari/537.36"
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")


def load_manifest() -> list[dict[str, Any]]:
    data = zlib.decompress(base64.b64decode(MANIFEST_B64))
    return json.loads(data.decode("utf-8"))


def unique(seq):
    seen = set()
    out = []
    for x in seq:
        key = x[1] if isinstance(x, tuple) else x
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(x)
    return out


def request_bytes(url: str, timeout: int = 120, retries: int = 3, headers: dict[str, str] | None = None) -> bytes:
    hdr = {
        "User-Agent": USER_AGENT,
        "Accept": "application/pdf,text/html,application/xhtml+xml,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    }
    if headers:
        hdr.update(headers)
    last: Exception | None = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers=hdr)
            with urllib.request.urlopen(req, timeout=timeout, context=ssl.create_default_context()) as resp:
                return resp.read()
        except urllib.error.HTTPError as exc:
            last = exc
            if exc.code in (404, 410):
                break
            if exc.code not in (408, 425, 429, 500, 502, 503, 504):
                break
        except Exception as exc:
            last = exc
        if attempt + 1 < retries:
            time.sleep((2 ** attempt) + random.random())
    raise RuntimeError(f"download failed: {url}: {last!r}")


def request_text(url: str, timeout: int = 60, retries: int = 2, headers: dict[str, str] | None = None) -> str:
    data = request_bytes(url, timeout=timeout, retries=retries, headers=headers)
    for enc in ("utf-8", "utf-16", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", "replace")


def metadata_urls(version: str) -> tuple[list[str], str | None, str | None]:
    slug = "itu-r-" + version.lower().replace(".", "-") + ".yaml"
    sources = [
        f"https://raw.githubusercontent.com/relaton/relaton-data-itu-r/v2/data/{slug}",
        f"https://raw.githubusercontent.com/relaton/relaton-data-itu-r/master/data/{slug}",
    ]
    errors = []
    for meta_url in sources:
        try:
            text = request_text(meta_url, timeout=45, retries=2)
            urls = re.findall(r"https://www\.itu\.int/dms_pubrec/[^\s'\"]+?PDF-E\.pdf", text)
            urls = [html.unescape(u.rstrip(".,;")) for u in urls]
            if urls:
                return unique(urls), meta_url, None
            errors.append(f"{meta_url}: no PDF-E URL")
        except Exception as exc:
            errors.append(f"{meta_url}: {exc!r}")
    return [], None, " | ".join(errors)


def scrape_page_urls(item: dict[str, Any]) -> list[str]:
    version = item["version"]
    rec = item.get("rec") or version.rsplit("-", 1)[0]
    pages = unique([
        item.get("en_page"),
        item.get("zh_page"),
        f"https://www.itu.int/rec/R-REC-{version}/en",
        f"https://www.itu.int/rec/R-REC-{version}/zh",
        f"https://www.itu.int/rec/R-REC-{rec}/en",
    ])
    found = []
    for page in pages:
        try:
            text = request_text(page, timeout=60, retries=2)
        except Exception:
            continue
        text = html.unescape(text.replace("\\/", "/"))
        for m in re.findall(r"(?:https?://www\.itu\.int)?/dms_pubrec/[^\s'\"<>]+?PDF-[A-Z]\.pdf", text, flags=re.I):
            if m.startswith("/"):
                m = "https://www.itu.int" + m
            found.append(m)
    return unique(found)


def github_search_mirrors(filename: str) -> list[str]:
    if not filename:
        return []
    urls = [f"https://raw.githubusercontent.com/KeyBridge/ITU-R-P.2001-6/main/docs/references/{filename}"]
    if not GITHUB_TOKEN:
        return urls
    q = urllib.parse.quote(f'"{filename}" in:path')
    api = f"https://api.github.com/search/code?q={q}&per_page=10"
    try:
        text = request_text(api, timeout=60, retries=2, headers={
            "Authorization": f"Bearer {GITHUB_TOKEN}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        })
        obj = json.loads(text)
        for it in obj.get("items", []):
            repo = it.get("repository", {}).get("full_name")
            path = it.get("path")
            branch = it.get("repository", {}).get("default_branch", "main")
            if repo and path:
                urls.append(f"https://raw.githubusercontent.com/{repo}/{branch}/{urllib.parse.quote(path)}")
    except Exception:
        pass
    return unique(urls)


def direct_candidate_urls(item: dict[str, Any]) -> tuple[list[tuple[str, str, str]], dict[str, Any]]:
    """Build the fast-path official candidates without scraping any pages."""
    meta, meta_source, meta_error = metadata_urls(item["version"])
    candidates: list[tuple[str, str, str]] = []

    # The Relaton-synchronised ITU metadata is the preferred source because it
    # repairs stale publication dates in workbook URLs. Try Chinese first, then
    # the official English file.
    for u in meta:
        if "PDF-E.pdf" in u:
            candidates.append(("C", u.replace("PDF-E.pdf", "PDF-C.pdf"), "ITU-metadata"))
        candidates.append(("E", u, "ITU-metadata"))

    # Workbook URLs remain a fallback for records absent from the metadata set.
    if item.get("cn_url"):
        candidates.append(("C", item["cn_url"], "ITU-workbook"))
    if item.get("en_url"):
        candidates.append(("E", item["en_url"], "ITU-workbook"))

    return unique(candidates), {
        "metadata_source": meta_source,
        "metadata_error": meta_error,
        "metadata_pdf_urls": meta,
    }


def try_candidates(item: dict[str, Any], result: dict[str, Any], candidates: list[tuple[str, str, str]]) -> bool:
    """Try candidates in order. Return True after a valid PDF is saved."""
    for lang, url, source_kind in unique(candidates):
        attempt = {"language": lang, "source_kind": source_kind, "url": url}
        try:
            referer = item.get("en_page") or f"https://www.itu.int/rec/R-REC-{item['version']}/en"
            data = request_bytes(url, timeout=120, retries=3, headers={"Referer": referer})
            ok, validation = is_pdf(data)
            attempt["bytes"] = len(data)
            attempt["validation"] = validation
            if not ok:
                attempt["status"] = "invalid"
                result["attempts"].append(attempt)
                continue

            dest = OUT_ROOT / item["target"]
            dest.parent.mkdir(parents=True, exist_ok=True)
            tmp = dest.with_suffix(dest.suffix + ".part")
            tmp.write_bytes(data)
            tmp.replace(dest)
            digest = hashlib.sha256(data).hexdigest()
            attempt["status"] = "ok"
            result["attempts"].append(attempt)
            result.update({
                "status": "ok", "language": lang, "source_kind": source_kind,
                "source_url": url, "bytes": len(data), "sha256": digest,
                "validation": validation, "saved_path": str(dest),
            })
            return True
        except Exception as exc:
            attempt["status"] = "error"
            attempt["error"] = repr(exc)
            result["attempts"].append(attempt)
    return False


def is_pdf(data: bytes) -> tuple[bool, str]:
    if not data.startswith(b"%PDF-"):
        return False, f"header={data[:40]!r}"
    if len(data) < 5000:
        return False, f"too small: {len(data)} bytes"
    eof = b"%%EOF" in data[-4096:]
    return True, "ok" if eof else "PDF header valid; EOF marker not found in final 4096 bytes"


def download_one(item: dict[str, Any]) -> dict[str, Any]:
    result = {
        "row": item.get("row"), "category": item.get("category"), "rec": item.get("rec"),
        "version": item.get("version"), "target": item.get("target"), "status": "failed",
        "language": None, "source_kind": None, "source_url": None, "bytes": 0,
        "sha256": None, "validation": None, "attempts": [],
    }

    # Fast path: corrected official URLs from metadata, followed by workbook URLs.
    try:
        direct, meta_info = direct_candidate_urls(item)
        result.update(meta_info)
    except Exception as exc:
        direct = []
        result["candidate_generation_error"] = repr(exc)
    if try_candidates(item, result, direct):
        return result

    # Slow fallback 1: inspect the official recommendation pages only when the
    # direct official candidates failed.
    scraped: list[tuple[str, str, str]] = []
    try:
        for u in scrape_page_urls(item):
            upper = u.upper()
            lang = "C" if "PDF-C.PDF" in upper else ("E" if "PDF-E.PDF" in upper else "O")
            scraped.append((lang, u, "ITU-page"))
    except Exception as exc:
        result["page_scrape_error"] = repr(exc)
    if try_candidates(item, result, scraped):
        return result

    # Slow fallback 2: exact-filename public GitHub mirrors. This is intentionally
    # last so an official ITU copy is always preferred when available.
    mirror_candidates: list[tuple[str, str, str]] = []
    english_urls = [u for lang, u, _ in direct if lang == "E"] + [u for lang, u, _ in scraped if lang == "E"]
    for u in unique(english_urls):
        filename = urllib.parse.unquote(urllib.parse.urlsplit(u).path.rsplit("/", 1)[-1])
        for mirror in github_search_mirrors(filename):
            mirror_candidates.append(("E", mirror, "GitHub-mirror"))
    try_candidates(item, result, mirror_candidates)
    return result


def main() -> int:
    manifest = load_manifest()
    if OUT_ROOT.exists():
        shutil.rmtree(OUT_ROOT)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    PDF_ROOT.mkdir(parents=True, exist_ok=True)
    (LOG_DIR / "原始下载清单.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    results: list[dict[str, Any]] = []
    workers = min(5, max(1, len(manifest)))
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(download_one, item): item for item in manifest}
        for done, fut in enumerate(concurrent.futures.as_completed(futures), 1):
            item = futures[fut]
            try:
                r = fut.result()
            except Exception as exc:
                r = {"row": item.get("row"), "version": item.get("version"), "target": item.get("target"), "status": "failed", "fatal_error": repr(exc)}
            results.append(r)
            print(f"[{done:03d}/{len(manifest)}] {r.get('version')}: {r.get('status')} {r.get('language') or ''} {r.get('bytes') or 0}", flush=True)

    results.sort(key=lambda r: int(r.get("row") or 999999))
    ok = sum(r.get("status") == "ok" for r in results)
    failed = len(results) - ok
    total_bytes = sum(int(r.get("bytes") or 0) for r in results if r.get("status") == "ok")
    summary = {
        "requested": len(manifest), "downloaded": ok, "failed": failed,
        "chinese": sum(r.get("status") == "ok" and r.get("language") == "C" for r in results),
        "english": sum(r.get("status") == "ok" and r.get("language") == "E" for r in results),
        "other_language": sum(r.get("status") == "ok" and r.get("language") not in ("C", "E") for r in results),
        "total_bytes": total_bytes,
    }
    (LOG_DIR / "下载结果.json").write_text(json.dumps({"summary": summary, "results": results}, ensure_ascii=False, indent=2), encoding="utf-8")
    fields = ["row","category","rec","version","status","language","source_kind","bytes","sha256","target","source_url","validation","metadata_source","metadata_error"]
    with (LOG_DIR / "下载结果.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(results)
    with (LOG_DIR / "SHA256校验值.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["建议书版本", "相对路径", "字节数", "SHA256"])
        for r in results:
            if r.get("status") == "ok":
                w.writerow([r.get("version"), r.get("target"), r.get("bytes"), r.get("sha256")])
    readme = (
        "ITU-R建议书1200-1400MHz PDF下载结果\n"
        "========================================\n"
        f"清单条目：{summary['requested']}\n"
        f"成功下载：{summary['downloaded']}\n"
        f"下载失败：{summary['failed']}\n"
        f"中文PDF：{summary['chinese']}\n"
        f"英文PDF：{summary['english']}\n"
        f"总字节数：{summary['total_bytes']}\n\n"
        "下载策略：优先采用Relaton同步的ITU-R官方元数据校正PDF地址，优先中文PDF；中文不可用时回退英文PDF；官方地址仍不可用时再尝试公开GitHub镜像。\n"
        "每个文件均检查%PDF-文件头，并计算SHA256。详细记录见清单与日志/下载结果.json及下载结果.csv。\n"
    )
    (OUT_ROOT / "README_下载说明.txt").write_text(readme, encoding="utf-8-sig")
    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
