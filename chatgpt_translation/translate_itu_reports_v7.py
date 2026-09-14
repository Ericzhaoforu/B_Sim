#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
import random
import re
import time
from pathlib import Path
from typing import Any

import translate_itu_reports_v5 as v5
import translate_itu_reports_v6 as v6

pipeline = v6.pipeline

# Common residual fragments left by the public translator or by PDF line
# extraction. These substitutions are intentionally conservative and target
# only standalone English words that should never remain in the Chinese body.
RESIDUAL_REPLACEMENTS: list[tuple[str, str]] = [
    (r"\bReport ITU-R\s+([A-Z]+\.\d+(?:-\d+)?)\b", r"ITU-R \1号报告"),
    (r"\bRecommendation ITU-R\s+([A-Z]+\.\d+(?:-\d+)?)\b", r"ITU-R \1号建议书"),
    (r"\bResolution\s+(\d+)\b", r"第\1号决议"),
    (r"\bFIGURE\s*(\d*)\b", r"图\1"),
    (r"\bTABLE\s*(\d*)\b", r"表\1"),
    (r"\bNOTE\b", "注"),
    (r"\bAnnex\b", "附件"),
    (r"\bAppendix\b", "附录"),
    (r"\bsatellites\b", "卫星"),
    (r"\bsatellite\b", "卫星"),
    (r"\bspacecraft\b", "航天器"),
    (r"\bearth stations\b", "地球站"),
    (r"\bearth station\b", "地球站"),
    (r"\bservice links\b", "业务链路"),
    (r"\bservice link\b", "业务链路"),
    (r"\bservices\b", "业务"),
    (r"\bservice\b", "业务"),
    (r"\blinks\b", "链路"),
    (r"\blink\b", "链路"),
    (r"\bdegrees\b", "°"),
    (r"\bdegree\b", "°"),
    (r"\bkilograms\b", "千克"),
    (r"\bkilogram\b", "千克"),
    (r"\bsmall satellites\b", "小卫星"),
    (r"\bsmall satellite\b", "小卫星"),
    (r"\bsystems\b", "系统"),
    (r"\bsystem\b", "系统"),
    (r"\bfrequency bands\b", "频段"),
    (r"\bfrequency band\b", "频段"),
    (r"\bfrequency\b", "频率"),
    (r"\bparameters\b", "参数"),
    (r"\bparameter\b", "参数"),
    (r"\bcharacteristics\b", "特性"),
    (r"\boperations\b", "运行"),
    (r"\boperation\b", "运行"),
    (r"\bmissions\b", "任务"),
    (r"\bmission\b", "任务"),
    (r"\busers\b", "用户"),
    (r"\buser\b", "用户"),
    (r"\bcurrent\b", "现有"),
    (r"\bprimary\b", "主要业务"),
    (r"\bsecondary\b", "次要业务"),
]

KNOWN_LATIN = {
    "ITU", "ITU-R", "ISO", "IEC", "IPR", "WRC", "RR", "NGSO", "GSO", "LEO",
    "TT&C", "SOS", "VHF", "UHF", "RF", "PSD", "CDF", "RAAN", "GMDSS", "AIS",
    "HADS", "GOES", "DCP", "DCPR", "EESS", "MetSat", "MetAids", "SRS", "RAS",
    "MSS", "COSPAS", "SARSAT", "TDMA", "CDMA", "SINAD", "BER", "ND", "SPACE",
    "RHCP", "LHCP", "FM", "GMSK", "FSK", "AFSK", "C", "I", "N", "BW", "TU",
    "Berlin", "SNS", "MHz", "GHz", "kHz", "Hz", "dB", "dBW", "dBm", "dBi",
    "dBc", "km", "ms", "ppm", "rpm", "pps", "K", "W", "A", "B", "C", "D",
}


def cleanup(text: str) -> str:
    out = text
    for pattern, replacement in RESIDUAL_REPLACEMENTS:
        out = re.sub(pattern, replacement, out, flags=re.I)
    out = out.replace("卫星s", "卫星").replace("系统s", "系统").replace("业务s", "业务")
    out = out.replace("small 卫星", "小卫星").replace("卫星 satellite", "卫星")
    out = re.sub(r"\b((?:19|20)\d{2})\s+s\b", r"\1年", out)
    out = re.sub(r"\b(\d+)\s+卫星\b", r"\1颗卫星", out)
    out = re.sub(r"\s+([，。；：！？、）])", r"\1", out)
    out = re.sub(r"([（])\s+", r"\1", out)
    out = re.sub(r"(?<=[\u3400-\u9fff])\s+(?=[\u3400-\u9fff])", "", out)
    out = re.sub(r"[ \t]{2,}", " ", out)
    return out.strip()


def untranslated_word_count(text: str) -> int:
    stripped = re.sub(r"https?://\S+", " ", text)
    stripped = re.sub(r"(?:ITU-R|RR|WRC)[^，。；\n]*", " ", stripped)
    words = re.findall(r"\b[A-Za-z][A-Za-z&/-]{2,}\b", stripped)
    return sum(1 for word in words if word not in KNOWN_LATIN and word.upper() not in KNOWN_LATIN)


def fallback_translate(source: str) -> str:
    try:
        translated = v5.google_translate(source, attempts=18)
    except Exception:
        translated = source
    translated = v6.restore_terms(translated, {})
    return cleanup(translated)


def translate_units_v7(units: list[dict[str, Any]], model_dir: str, cache_path: Path) -> None:
    cache_file = cache_path.with_name("translation_cache_v7.json")
    cache: dict[str, str] = {}
    if cache_file.exists():
        try:
            cache = json.loads(cache_file.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    jobs: list[tuple[int, str, str, str, dict[str, str]]] = []
    for index, unit in enumerate(units):
        source = str(unit["text"])
        key = hashlib.sha256(("v7:" + source).encode("utf-8")).hexdigest()
        exact = v5.exact_translation(source)
        if exact is not None:
            unit["zh"] = cleanup(exact)
            cache[key] = unit["zh"]
            continue
        if key in cache:
            unit["zh"] = cache[key]
            continue
        if pipeline.mostly_nonlinguistic(source):
            unit["zh"] = source
            cache[key] = source
            continue
        protected, mapping = v6.select_spans(source, index)
        jobs.append((index, key, source, protected, mapping))

    pipeline.log(f"V7 pending units: {len(jobs)}")
    cursor = 0
    translated_count = 0
    while cursor < len(jobs):
        batch: list[tuple[int, str, str, str, dict[str, str]]] = []
        characters = 0
        while cursor < len(jobs) and len(batch) < 20:
            candidate = jobs[cursor]
            extra = len(candidate[3]) + 30
            if batch and characters + extra > 3200:
                break
            batch.append(candidate)
            characters += extra
            cursor += 1

        local_indices = list(range(len(batch)))
        sources = [item[3] for item in batch]
        try:
            outputs = v5.translate_batch(local_indices, sources)
        except Exception as exc:
            pipeline.log(f"V7 batch warning: {exc!r}")
            outputs = {}

        for local_index, (unit_index, key, source, protected, mapping) in enumerate(batch):
            raw = outputs.get(local_index, "")
            if not raw or any(marker not in raw for marker in mapping):
                try:
                    raw = v5.google_translate(protected, attempts=12)
                except Exception:
                    raw = ""
            if raw and all(marker in raw for marker in mapping):
                translation = cleanup(v6.restore_terms(raw, mapping))
            else:
                translation = fallback_translate(source)

            # A few public-translation responses may preserve a full English
            # clause. Retry the unprotected sentence when that occurs.
            if untranslated_word_count(translation) >= 5:
                alternative = fallback_translate(source)
                if untranslated_word_count(alternative) < untranslated_word_count(translation):
                    translation = alternative

            if not translation:
                translation = source
            units[unit_index]["zh"] = translation
            cache[key] = translation

        translated_count += len(batch)
        if translated_count % 400 < len(batch) or translated_count == len(jobs):
            pipeline.log(f"V7 translated {translated_count}/{len(jobs)} units")
        time.sleep(0.10 + random.random() * 0.08)

    for unit in units:
        if "zh" not in unit:
            unit["zh"] = str(unit["text"])

    cache_file.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


pipeline.translate_units = translate_units_v7

if __name__ == "__main__":
    raise SystemExit(pipeline.main())
