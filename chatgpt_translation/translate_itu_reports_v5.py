#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
import os
import random
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

import translate_itu_reports as pipeline

pipeline.REPORTS["SA.2425"]["zh_title"] = "满足短期任务非GSO卫星空间操作业务频谱需求的研究"
pipeline.REPORTS["SA.2426"]["zh_title"] = "1 GHz以下短期任务非GSO卫星空间操作业务遥测、跟踪与遥控的技术特性"
pipeline.REPORTS["SA.2427"]["zh_title"] = "1 GHz以下空间操作业务现有频率划分适用性研究及可能新增和/或升级频率划分的补充共用研究"

EXACT: dict[str, str] = {
    "Foreword": "前言",
    "Introduction": "引言",
    "Summary": "总结",
    "Conclusion": "结论",
    "Conclusions": "结论",
    "Annex": "附件",
    "Appendix": "附录",
    "Parameter": "参数",
    "Parameters": "参数",
    "Value": "数值",
    "Values": "数值",
    "Unit": "单位",
    "Status": "业务地位",
    "Direction": "方向",
    "Remarks": "备注",
    "Frequency": "频率",
    "Frequency band": "频段",
    "Frequency bands": "频段",
    "Primary": "主要业务",
    "Secondary": "次要业务",
    "Yes": "是",
    "No": "否",
    "N/A": "不适用",
    "Orbit": "轨道",
    "Transmission scheme": "传输方案",
    "Simulation overview": "仿真概览",
    "Series": "系列",
    "Title": "标题",
    "TABLE": "表",
    "FIGURE": "图",
    "NOTE": "注",
    "Electronic Publication": "电子出版物",
    "Geneva, 2018": "2018年，日内瓦",
    "Series of ITU-R Reports": "ITU-R报告系列",
    "Satellite delivery": "卫星传送",
    "Recording for production, archival and play-out; film for television": "用于制作、存档和播放的录制；电视胶片",
    "Broadcasting service (sound)": "广播业务（声音）",
    "Broadcasting service (television)": "广播业务（电视）",
    "Fixed service": "固定业务",
    "Mobile, radiodetermination, amateur and related satellite services": "移动、无线电测定、业余及相关卫星业务",
    "Radiowave propagation": "无线电波传播",
    "Radio astronomy": "射电天文",
    "Remote sensing systems": "遥感系统",
    "Fixed-satellite service": "固定卫星业务",
    "Space applications and meteorology": "空间应用和气象",
    "Frequency sharing and coordination between fixed-satellite and fixed service systems": "固定卫星业务与固定业务系统之间的频率共用和协调",
    "Spectrum management": "频谱管理",
    "The role of the Radiocommunication Sector is to ensure the rational, equitable, efficient and economical use of the radio-frequency spectrum by all radiocommunication services, including satellite services, and carry out studies without limit of frequency range on the basis of which Recommendations are adopted.": "无线电通信部门的职责是确保包括卫星业务在内的所有无线电通信业务合理、公平、有效和经济地使用无线电频谱，并开展不受频率范围限制的研究，在此基础上通过建议书。",
    "The regulatory and policy functions of the Radiocommunication Sector are performed by World and Regional Radiocommunication Conferences and Radiocommunication Assemblies supported by Study Groups.": "无线电通信部门的规章和政策职能由世界和区域无线电通信大会以及无线电通信全会履行，并由各研究组提供支持。",
    "Note: This ITU-R Report was approved in English by the Study Group under the procedure detailed in Resolution ITU-R 1.": "注：本ITU-R报告英文版由研究组按照ITU-R第1号决议规定的程序批准。",
    "All rights reserved. No part of this publication may be reproduced, by any means whatsoever, without written permission of ITU.": "版权所有。未经国际电信联盟书面许可，不得以任何方式复制本出版物的任何部分。",
    "Studies to accommodate spectrum requirements in the space operation service for non-geostationary satellites with short duration missions": "满足短期任务非GSO卫星空间操作业务频谱需求的研究",
    "Technical characteristics for telemetry, tracking and command in the space operation service below 1 GHz for non-GSO satellites with short duration missions": "1 GHz以下短期任务非GSO卫星空间操作业务遥测、跟踪与遥控的技术特性",
    "Studies on the suitability of existing allocations to the space operation service below 1 GHz and additional sharing studies on possible new and/or upgraded allocations": "1 GHz以下空间操作业务现有频率划分适用性研究及可能新增和/或升级频率划分的补充共用研究",
}

# Exact fragments from the split cover titles.
EXACT.update({
    "Studies to accommodate spectrum": "满足频谱需求的研究——",
    "requirements in the space operation service": "空间操作业务",
    "for non-geostationary satellites with short": "面向短期任务非GSO卫星",
    "duration missions": "短期任务",
    "Technical characteristics for telemetry,": "遥测、",
    "tracking and command in the space": "跟踪与遥控技术特性——",
    "operation service below 1 GHz for non-GSO": "1 GHz以下空间操作业务中的短期任务非GSO",
    "satellites with short duration missions": "卫星",
    "Studies on the suitability of existing": "现有频率划分适用性研究——",
    "allocations to the space operation service": "空间操作业务",
    "below 1 GHz and additional sharing studies": "1 GHz以下及补充共用研究",
    "on possible new and/or upgraded": "可能新增和/或升级的",
    "allocations": "频率划分",
})

POST_REPLACEMENTS: list[tuple[str, str]] = [
    ("空间运行服务", "空间操作业务"),
    ("空间运营服务", "空间操作业务"),
    ("空间业务服务", "空间操作业务"),
    ("遥测、跟踪和指挥", "遥测、跟踪与遥控"),
    ("遥测、跟踪和命令", "遥测、跟踪与遥控"),
    ("遥测、跟踪和控制", "遥测、跟踪与遥控"),
    ("非 GSO", "非GSO"),
    ("非地球静止轨道", "非GSO轨道"),
    ("非地球静止卫星", "非GSO卫星"),
    ("近地轨道", "低地球轨道"),
    ("功率光谱密度", "功率谱密度"),
    ("等效各向同性辐射功率", "等效全向辐射功率"),
    ("频率分配", "频率划分"),
    ("现有分配", "现有频率划分"),
    ("主要分配", "主要业务划分"),
    ("次要分配", "次要业务划分"),
    ("外带发射", "带外发射"),
    ("伪发射", "杂散发射"),
    ("虚假发射", "杂散发射"),
    ("保护标准", "保护准则"),
    ("地球站站", "地球站"),
    ("空间站站", "空间电台"),
    ("电信指挥", "遥控"),
    ("载波干扰比率", "载波干扰比"),
    ("载波噪声比", "载噪比"),
    ("卫星地球站组合", "卫星-地球站组合"),
    ("短期飞行任务", "短期任务"),
    ("短时任务", "短期任务"),
    ("频谱要求", "频谱需求"),
]

USER_AGENTS = [
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/140.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/139.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/138.0 Safari/537.36",
]

HEADER_RE = re.compile(r"^(?:ii|iv|vi|viii|x|\d+)\s+(?:Rep\.|Report)\s+ITU-R\s+([A-Z]+\.\d+(?:-\d+)?)\s*(\d+)?$", re.I)
REPORT_RE = re.compile(r"^(?:REPORT|Report|Rep\.)\s+ITU-R\s+([A-Z]+\.\d+(?:-\d+)?)$", re.I)


def compact(text: str) -> str:
    return " ".join(text.replace("\u00ad", "").split())


def exact_translation(source: str) -> str | None:
    c = compact(source)
    if c in EXACT:
        return EXACT[c]
    m = HEADER_RE.match(c)
    if m:
        page = m.group(2)
        return f"ITU-R {m.group(1)} 报告" + (f"（原文页码 {page}）" if page else "")
    m = REPORT_RE.match(c)
    if m:
        return f"ITU-R {m.group(1)} 报告"
    if re.fullmatch(r"(?:Recommendation|Report|Resolution) ITU-R [A-Z]+\.\d+(?:-\d+)?", c):
        return c
    if re.fullmatch(r"ITU-R [A-Z]+\.\d+(?:-\d+)?", c):
        return c
    return None


def normalize(text: str) -> str:
    out = text.replace("\u00ad", "")
    for old, new in POST_REPLACEMENTS:
        out = out.replace(old, new)
    out = out.replace("（ ", "（").replace(" ）", "）")
    out = re.sub(r"\s+([，。；：！？、])", r"\1", out)
    out = re.sub(r"(?<=[\u3400-\u9fff])\s+(?=[\u3400-\u9fff])", "", out)
    out = re.sub(r"(?<=[\u3400-\u9fff])\s+([，。；：！？、）])", r"\1", out)
    out = re.sub(r"([（])\s+(?=[\u3400-\u9fff])", r"\1", out)
    out = re.sub(r"[ \t]{2,}", " ", out)
    return out.strip()


def google_translate(text: str, attempts: int = 12) -> str:
    params = urllib.parse.urlencode({
        "client": "dict-chrome-ex",
        "sl": "en",
        "tl": "zh-CN",
        "q": text,
    })
    url = "https://clients5.google.com/translate_a/t?" + params
    last: Exception | None = None
    for attempt in range(attempts):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": random.choice(USER_AGENTS),
                "Accept": "application/json,text/plain,*/*",
                "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.7",
                "Connection": "close",
            })
            with urllib.request.urlopen(req, timeout=90) as response:
                payload = json.loads(response.read().decode("utf-8"))
            if isinstance(payload, list) and payload and isinstance(payload[0], str):
                return payload[0]
            raise RuntimeError(f"Unexpected clients5 response: {payload!r}")
        except urllib.error.HTTPError as exc:
            last = exc
            if exc.code == 429:
                time.sleep(min(90, 3 * (attempt + 1)) + random.random() * 2)
            elif exc.code in (500, 502, 503, 504):
                time.sleep(min(40, 2 ** min(attempt, 5)) + random.random())
            else:
                time.sleep(2 + attempt)
        except Exception as exc:
            last = exc
            time.sleep(min(30, 1.5 * (attempt + 1)) + random.random())
    raise RuntimeError(f"clients5 translation failed after {attempts} attempts: {last!r}")


def translate_batch(indices: list[int], sources: list[str]) -> dict[int, str]:
    if not indices:
        return {}
    if len(indices) == 1:
        return {indices[0]: normalize(google_translate(sources[0]))}
    markers = [f"ZZUNIT{i:06d}ZZ" for i in range(len(indices) - 1)]
    joined_parts: list[str] = []
    for i, source in enumerate(sources):
        joined_parts.append(source)
        if i < len(markers):
            joined_parts.append(markers[i])
    joined = "\n".join(joined_parts)
    translated = google_translate(joined)
    if all(marker in translated for marker in markers):
        pieces = re.split("|".join(re.escape(marker) for marker in markers), translated)
        if len(pieces) == len(indices):
            return {idx: normalize(piece) for idx, piece in zip(indices, pieces)}
    # If a marker was altered, recursively split the batch rather than risking
    # unit-to-unit misalignment.
    middle = len(indices) // 2
    result = translate_batch(indices[:middle], sources[:middle])
    result.update(translate_batch(indices[middle:], sources[middle:]))
    return result


def split_oversize(text: str, max_chars: int = 3200) -> list[str]:
    if len(text) <= max_chars:
        return [text]
    sentences = re.split(r"(?<=[.!?;:])\s+|\n+", text)
    parts: list[str] = []
    current = ""
    for sentence in sentences:
        sentence = sentence.strip()
        if not sentence:
            continue
        if len(sentence) > max_chars:
            if current:
                parts.append(current)
                current = ""
            for start in range(0, len(sentence), max_chars):
                parts.append(sentence[start:start + max_chars])
            continue
        candidate = sentence if not current else current + " " + sentence
        if len(candidate) <= max_chars:
            current = candidate
        else:
            parts.append(current)
            current = sentence
    if current:
        parts.append(current)
    return parts


def translate_units_v5(units: list[dict[str, Any]], model_dir: str, cache_path: Path) -> None:
    cache_file = cache_path.with_name("translation_cache_v5.json")
    cache: dict[str, str] = {}
    if cache_file.exists():
        try:
            cache = json.loads(cache_file.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    pending: list[tuple[int, str, str]] = []
    long_plans: dict[int, list[str]] = {}

    for index, unit in enumerate(units):
        source = str(unit["text"])
        key = hashlib.sha256(("v5:" + source).encode("utf-8")).hexdigest()
        exact = exact_translation(source)
        if exact is not None:
            unit["zh"] = exact
            cache[key] = exact
            continue
        if key in cache:
            unit["zh"] = cache[key]
            continue
        if pipeline.mostly_nonlinguistic(source):
            unit["zh"] = source
            cache[key] = source
            continue
        pieces = split_oversize(source)
        if len(pieces) == 1:
            pending.append((index, key, pieces[0]))
        else:
            piece_keys: list[str] = []
            for piece_no, piece in enumerate(pieces):
                piece_key = hashlib.sha256((f"v5-piece:{piece_no}:" + piece).encode("utf-8")).hexdigest()
                piece_keys.append(piece_key)
                if piece_key not in cache:
                    pending.append((-1, piece_key, piece))
            long_plans[index] = piece_keys

    # Translate unique sources only.
    unique_by_key: dict[str, tuple[int, str]] = {}
    for index, key, source in pending:
        unique_by_key.setdefault(key, (index, source))
    jobs = [(index, key, source) for key, (index, source) in unique_by_key.items()]
    pipeline.log(f"V5 pending units/fragments: {len(jobs)}")

    cursor = 0
    translated_count = 0
    while cursor < len(jobs):
        batch: list[tuple[int, str, str]] = []
        chars = 0
        while cursor < len(jobs) and len(batch) < 24:
            candidate = jobs[cursor]
            extra = len(candidate[2]) + 24
            if batch and chars + extra > 3400:
                break
            batch.append(candidate)
            chars += extra
            cursor += 1
        local_indices = list(range(len(batch)))
        sources = [item[2] for item in batch]
        try:
            outputs = translate_batch(local_indices, sources)
        except Exception as exc:
            # Last-resort individual retries; preserve the English source only
            # if the external service is persistently unavailable.
            pipeline.log(f"Batch translation warning: {exc!r}")
            outputs = {}
            for i, source in enumerate(sources):
                try:
                    outputs[i] = normalize(google_translate(source, attempts=16))
                except Exception as item_exc:
                    pipeline.log(f"Individual translation failed: {item_exc!r}")
                    outputs[i] = source
        for i, (_, key, source) in enumerate(batch):
            value = outputs.get(i, source)
            if not value or len(value) > max(20000, len(source) * 12):
                value = source
            cache[key] = normalize(value)
        translated_count += len(batch)
        if translated_count % 500 < len(batch) or translated_count == len(jobs):
            pipeline.log(f"V5 translated {translated_count}/{len(jobs)} units/fragments")
        time.sleep(0.12 + random.random() * 0.08)

    for index, unit in enumerate(units):
        if "zh" in unit:
            continue
        source = str(unit["text"])
        key = hashlib.sha256(("v5:" + source).encode("utf-8")).hexdigest()
        if index in long_plans:
            translation = normalize(" ".join(cache.get(k, "") for k in long_plans[index]))
        else:
            translation = cache.get(key, source)
        # Detect pathological repetitions and keep the source as a visible,
        # reviewable fallback rather than inserting corrupted Chinese text.
        tokens = re.findall(r"[\u3400-\u9fff]{1,8}", translation)
        if tokens and len(tokens) > 80:
            from collections import Counter
            most = Counter(tokens).most_common(1)[0][1]
            if most / len(tokens) > 0.35:
                translation = source
        unit["zh"] = normalize(translation)
        cache[key] = unit["zh"]

    cache_file.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


pipeline.translate_units = translate_units_v5

if __name__ == "__main__":
    raise SystemExit(pipeline.main())
