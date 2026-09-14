#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
import random
import re
import time
from pathlib import Path
from typing import Any

import translate_itu_reports_v4 as v4
import translate_itu_reports_v5 as v5

pipeline = v5.pipeline

# Additional expressions that benefit from deterministic ITU wording.
ADDITIONAL_TERMS: list[tuple[str, str]] = [
    ("space operations service", "空间操作业务"),
    ("short duration NGSO mission", "短期任务非GSO卫星"),
    ("short duration NGSO missions", "短期任务非GSO卫星"),
    ("NGSO SD satellites", "短期任务非GSO卫星"),
    ("NGSO SD satellite", "短期任务非GSO卫星"),
    ("NGSO SD", "短期任务非GSO"),
    ("low Earth orbit", "低地球轨道"),
    ("low earth orbit", "低地球轨道"),
    ("period of validity", "有效期"),
    ("notifying administration", "通知主管部门"),
    ("Radiocommunication Bureau", "无线电通信局"),
    ("radio-frequency spectrum", "无线电频谱"),
    ("radio frequency spectrum", "无线电频谱"),
    ("radio-frequency interference", "无线电频率干扰"),
    ("radio frequency interference", "无线电频率干扰"),
    ("frequency assignment", "频率指配"),
    ("frequency assignments", "频率指配"),
    ("frequency band", "频段"),
    ("frequency bands", "频段"),
    ("carrier-to-noise ratio objective", "载噪比目标值"),
    ("operational duty cycle", "运行占空比"),
    ("mission lifetime", "任务寿命"),
    ("mission duration", "任务持续时间"),
    ("satellite network", "卫星网络"),
    ("satellite networks", "卫星网络"),
    ("space network", "空间网络"),
    ("space networks", "空间网络"),
    ("service links", "业务链路"),
    ("service link", "业务链路"),
    ("link time", "链路时间"),
    ("co-frequency", "同频"),
    ("co-channel", "同信道"),
    ("adjacent bands", "相邻频段"),
    ("adjacent band", "相邻频段"),
    ("harmful interference", "有害干扰"),
    ("aggregate interference", "聚合干扰"),
    ("cumulative distribution function", "累积分布函数"),
    ("random permutation", "随机置换"),
    ("off-axis angle", "离轴角"),
    ("slant range", "斜距"),
    ("path loss", "路径损耗"),
    ("atmospheric loss", "大气损耗"),
    ("ionospheric loss", "电离层损耗"),
    ("polarization loss", "极化损耗"),
    ("polarization losses", "极化损耗"),
    ("figure of merit", "品质因数"),
    ("spectral efficiency", "频谱效率"),
    ("bit rate", "比特率"),
    ("bitrate", "比特率"),
    ("out-of-band domain", "带外域"),
    ("spurious domain", "杂散域"),
    ("reference bandwidth", "参考带宽"),
    ("emission mask", "发射掩模"),
    ("space station transmitter", "空间电台发射机"),
    ("earth station transmitter", "地球站发射机"),
    ("space station receiver", "空间电台接收机"),
    ("earth station receiver", "地球站接收机"),
    ("data collection system", "数据收集系统"),
    ("data collection systems", "数据收集系统"),
    ("data collection platform", "数据收集平台"),
    ("data collection platforms", "数据收集平台"),
    ("meteorological aids", "气象辅助"),
    ("radiosonde", "无线电探空仪"),
    ("radiosondes", "无线电探空仪"),
    ("dropsonde", "下投式探空仪"),
    ("dropsondes", "下投式探空仪"),
    ("rocketsonde", "火箭探空仪"),
    ("rocketsondes", "火箭探空仪"),
    ("land mobile service", "陆地移动业务"),
    ("maritime mobile service", "水上移动业务"),
    ("aeronautical mobile service", "航空移动业务"),
    ("aeronautical mobile (route) service", "航空移动（航路）业务"),
    ("search and rescue", "搜索与救援"),
    ("safety of life", "生命安全"),
    ("World Radiocommunication Conference", "世界无线电通信大会"),
    ("Regional Radiocommunication Conference", "区域无线电通信大会"),
    ("Radiocommunication Assembly", "无线电通信全会"),
    ("Study Group", "研究组"),
    ("Study Groups", "研究组"),
    ("invites ITU-R", "邀请ITU-R"),
    ("invites 1", "第1项邀请"),
    ("invites 2", "第2项邀请"),
    ("invites 3", "第3项邀请"),
    ("recognizing a)", "认识到a)"),
]

# Merge dictionaries, with specific/long expressions taking precedence.
_seen: set[str] = set()
TERM_ITEMS: list[tuple[str, str]] = []
for source, target in sorted(ADDITIONAL_TERMS + v4.TERM_ITEMS, key=lambda p: len(p[0]), reverse=True):
    key = source.lower()
    if key not in _seen:
        _seen.add(key)
        TERM_ITEMS.append((source, target))
TERM_LOOKUP = {source.lower(): target for source, target in TERM_ITEMS}
TERM_RE = re.compile("|".join(re.escape(source) for source, _ in TERM_ITEMS), re.I)

# Protect report identifiers, URLs, acronyms, numerical values with units, and
# technical formula tokens. Google keeps alphanumeric ZZ...ZZ markers intact.
ID_PATTERNS = [
    r"https?://\S+",
    r"\b(?:Recommendation|Report|Resolution)\s+ITU-R\s+[A-Z]+\.\d+(?:-\d+)?\b",
    r"\bITU-R\s+[A-Z]+\.\d+(?:-\d+)?\b",
    r"\bRR\s+(?:No\.|Nos\.)\s*\d+(?:\.\d+)*(?:\s*(?:and|,|to|-)\s*\d+(?:\.\d+)*)*\b",
    r"\bWRC-\d{2}\b",
    r"\b(?:NGSO|GSO|LEO|TT&C|SOS|VHF|UHF|RF|PSD|CDF|RAAN|GMDSS|AIS|HADS|GOES|DCPR?|EESS|MetSat|MetAids|SRS|RAS|MSS|AM\(R\)S|COSPAS-SARSAT|TDMA|CDMA|SINAD|BER|ND-SPACE|RHCP|LHCP|FM|GMSK|FSK|AFSK)\b",
    r"\b(?:C/I|C/N0?|I/N|S/N|Eb/N0|G/T)\b",
    r"\be\.i\.r\.p\.\b",
    r"\b[A-Z]\.\d+(?:\.[A-Za-z0-9]+){1,6}\b",
    r"(?<!\w)[≤≥<>≈~+−-]?\s*\d+(?:[ .,:/×*+−-]\d+)*(?:\s*(?:GHz|MHz|kHz|Hz|dBW/kHz|dBW/Hz|dBW|dBm|dBi|dBc|dB/K|dB|kbit/s|bit/s/Hz|bit/s|km|m|ms|µs|s|K|W|%|degrees?|rpm|pps))",
]
ID_RES = [re.compile(pattern) for pattern in ID_PATTERNS]

# Targeted cleanup for recurring generic translations that are technically
# misleading in ITU documents.
POST_REPLACEMENTS = v5.POST_REPLACEMENTS + [
    ("空间运营业务", "空间操作业务"),
    ("空间运行业务", "空间操作业务"),
    ("空间运行服务", "空间操作业务"),
    ("新的分配", "新的频率划分"),
    ("新分配", "新频率划分"),
    ("频率分派", "频率指配"),
    ("长期分配", "长期频率指配"),
    ("纳卫星", "纳米卫星"),
    ("皮卫星", "皮米卫星"),
    ("平均异常", "平近点角"),
    ("近地点论点", "近地点幅角"),
    ("上升节点的右提升", "升交点赤经"),
    ("十年频率", "十倍频程"),
    ("频率十年", "十倍频程"),
    ("主要分配", "主要业务划分"),
    ("次要分配", "次要业务划分"),
    ("主要基础", "主要业务地位"),
    ("次要基础", "次要业务地位"),
    ("跟踪和指挥", "跟踪与遥控"),
    ("遥测、跟踪和指挥", "遥测、跟踪与遥控"),
    ("空间站", "空间电台"),
    ("推荐。 ITU-R", "ITU-R"),
    ("建议。 ITU-R", "ITU-R"),
    ("分贝", "dB"),
]


def select_spans(text: str, unit_index: int) -> tuple[str, dict[str, str]]:
    """Replace protected terms/identifiers with stable per-unit markers."""
    candidates: list[tuple[int, int, str, int]] = []
    for match in TERM_RE.finditer(text):
        candidates.append((match.start(), match.end(), TERM_LOOKUP[match.group(0).lower()], 0))
    for regex in ID_RES:
        for match in regex.finditer(text):
            candidates.append((match.start(), match.end(), match.group(0), 1))
    # Earliest wins. At same position choose longest; deterministic terms take
    # precedence over generic identifier spans of equal length.
    candidates.sort(key=lambda item: (item[0], -(item[1] - item[0]), item[3]))
    selected: list[tuple[int, int, str]] = []
    occupied_end = -1
    for start, end, replacement, _ in candidates:
        if start < occupied_end:
            continue
        selected.append((start, end, replacement))
        occupied_end = end

    mapping: dict[str, str] = {}
    chunks: list[str] = []
    position = 0
    for term_no, (start, end, replacement) in enumerate(selected):
        if start > position:
            chunks.append(text[position:start])
        marker = f"ZZT{unit_index:06d}X{term_no:04d}ZZ"
        chunks.append(marker)
        mapping[marker] = replacement
        position = end
    chunks.append(text[position:])
    return "".join(chunks), mapping


def restore_terms(text: str, mapping: dict[str, str]) -> str:
    output = text
    for marker, replacement in mapping.items():
        output = output.replace(marker, replacement)
    for old, new in POST_REPLACEMENTS:
        output = output.replace(old, new)
    output = v5.normalize(output)
    output = re.sub(r"(?<=[\u3400-\u9fff])\s+(?=[\u3400-\u9fff])", "", output)
    output = re.sub(r"(?<=[\u3400-\u9fff])\s+([，。；：！？、）])", r"\1", output)
    output = re.sub(r"([（])\s+(?=[\u3400-\u9fff])", r"\1", output)
    output = re.sub(r"[ \t]{2,}", " ", output)
    return output.strip()


def translate_units_v6(units: list[dict[str, Any]], model_dir: str, cache_path: Path) -> None:
    cache_file = cache_path.with_name("translation_cache_v6.json")
    cache: dict[str, str] = {}
    if cache_file.exists():
        try:
            cache = json.loads(cache_file.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    jobs: list[tuple[int, str, str, dict[str, str]]] = []
    for index, unit in enumerate(units):
        source = str(unit["text"])
        key = hashlib.sha256(("v6:" + source).encode("utf-8")).hexdigest()
        exact = v5.exact_translation(source)
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
        protected, mapping = select_spans(source, index)
        jobs.append((index, key, protected, mapping))

    pipeline.log(f"V6 pending units: {len(jobs)}")
    cursor = 0
    translated_count = 0
    while cursor < len(jobs):
        batch: list[tuple[int, str, str, dict[str, str]]] = []
        characters = 0
        while cursor < len(jobs) and len(batch) < 20:
            candidate = jobs[cursor]
            extra = len(candidate[2]) + 30
            if batch and characters + extra > 3200:
                break
            batch.append(candidate)
            characters += extra
            cursor += 1
        local_indices = list(range(len(batch)))
        sources = [item[2] for item in batch]
        try:
            outputs = v5.translate_batch(local_indices, sources)
        except Exception as exc:
            pipeline.log(f"V6 batch warning: {exc!r}")
            outputs = {}
            for local_index, source in enumerate(sources):
                try:
                    outputs[local_index] = v5.google_translate(source, attempts=16)
                except Exception as item_exc:
                    pipeline.log(f"V6 individual failure: {item_exc!r}")
                    outputs[local_index] = source
        for local_index, (unit_index, key, protected, mapping) in enumerate(batch):
            raw = outputs.get(local_index, protected)
            # If a term marker was unexpectedly changed, retry this unit alone.
            missing = [marker for marker in mapping if marker not in raw]
            if missing:
                try:
                    raw = v5.google_translate(protected, attempts=16)
                except Exception:
                    raw = protected
            if any(marker not in raw for marker in mapping):
                raw = protected
            translation = restore_terms(raw, mapping)
            units[unit_index]["zh"] = translation
            cache[key] = translation
        translated_count += len(batch)
        if translated_count % 400 < len(batch) or translated_count == len(jobs):
            pipeline.log(f"V6 translated {translated_count}/{len(jobs)} units")
        time.sleep(0.12 + random.random() * 0.08)

    for unit in units:
        if "zh" not in unit:
            unit["zh"] = str(unit["text"])

    cache_file.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


pipeline.translate_units = translate_units_v6

if __name__ == "__main__":
    raise SystemExit(pipeline.main())
