#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
import os
import re
from pathlib import Path
from typing import Any

import torch
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer

import translate_itu_reports as pipeline

pipeline.REPORTS["SA.2425"]["zh_title"] = "满足短期任务非GSO卫星空间操作业务频谱需求的研究"
pipeline.REPORTS["SA.2426"]["zh_title"] = "1 GHz以下短期任务非GSO卫星空间操作业务遥测、跟踪与遥控的技术特性"

# Fixed Chinese equivalents are inserted only after the surrounding English
# fragments have been translated.  This avoids both fragile placeholders and
# feeding Chinese text into an English-source Marian tokenizer.
TERMS: list[tuple[str, str]] = [
    ("non-geostationary satellites with short duration missions", "短期任务非GSO卫星"),
    ("non-GSO satellites with short duration missions", "短期任务非GSO卫星"),
    ("NGSO satellites with short duration missions", "短期任务非GSO卫星"),
    ("NGSO short duration mission satellites", "短期任务非GSO卫星"),
    ("short duration NGSO satellites", "短期任务非GSO卫星"),
    ("short duration mission satellites", "短期任务卫星"),
    ("short duration missions", "短期任务"),
    ("short duration mission", "短期任务"),
    ("non-geostationary-satellite orbit", "非GSO卫星轨道"),
    ("non-geostationary satellite orbit", "非GSO卫星轨道"),
    ("space operations service", "空间操作业务"),
    ("space operation service", "空间操作业务"),
    ("telemetry, tracking and command", "遥测、跟踪与遥控"),
    ("telemetry, tracking and control", "遥测、跟踪与遥控"),
    ("spectrum requirements", "频谱需求"),
    ("spectrum requirement", "频谱需求"),
    ("existing allocations", "现有频率划分"),
    ("new and/or upgraded allocations", "新增和/或升级的频率划分"),
    ("frequency allocations", "频率划分"),
    ("frequency allocation", "频率划分"),
    ("allocated bandwidth", "已划分带宽"),
    ("sharing and compatibility studies", "共用与兼容性研究"),
    ("sharing studies", "共用研究"),
    ("compatibility studies", "兼容性研究"),
    ("mitigation techniques", "缓解技术"),
    ("incumbent services", "现有业务"),
    ("technical and operational characteristics", "技术和运行特性"),
    ("technical characteristics", "技术特性"),
    ("operational characteristics", "运行特性"),
    ("space segment", "空间段"),
    ("ground segment", "地面段"),
    ("space-to-Earth direction", "空对地方向"),
    ("Earth-to-space direction", "地对空方向"),
    ("space-to-Earth", "空对地"),
    ("Earth-to-space", "地对空"),
    ("satellite-earth station combinations", "卫星-地球站组合"),
    ("satellite-earth station combination", "卫星-地球站组合"),
    ("baseline sustainability factor", "基准可持续系数"),
    ("aggregate interference power spectral density", "聚合干扰功率谱密度"),
    ("interference power spectral density", "干扰功率谱密度"),
    ("power spectral density", "功率谱密度"),
    ("carrier-to-interference ratio", "载波干扰比"),
    ("carrier-to-noise ratio", "载噪比"),
    ("protection criteria", "保护准则"),
    ("protection criterion", "保护准则"),
    ("necessary bandwidth", "必要带宽"),
    ("occupied bandwidth", "占用带宽"),
    ("out-of-band emissions", "带外发射"),
    ("out-of-band emission", "带外发射"),
    ("spurious emissions", "杂散发射"),
    ("spurious emission", "杂散发射"),
    ("unwanted emissions", "非必要发射"),
    ("unwanted emission", "非必要发射"),
    ("equivalent isotropically radiated power", "等效全向辐射功率"),
    ("power flux-density", "功率通量密度"),
    ("power flux density", "功率通量密度"),
    ("link budget analysis", "链路预算分析"),
    ("link budget", "链路预算"),
    ("link margin", "链路余量"),
    ("system noise temperature", "系统噪声温度"),
    ("noise temperature", "噪声温度"),
    ("antenna radiation pattern", "天线辐射方向图"),
    ("antenna pointing loss", "天线指向损耗"),
    ("antenna beamwidth", "天线波束宽度"),
    ("antenna gain", "天线增益"),
    ("minimum ground elevation angle", "最小地面仰角"),
    ("minimum elevation angle", "最小仰角"),
    ("tracking antenna", "跟踪天线"),
    ("omnidirectional", "全向"),
    ("circular polarization", "圆极化"),
    ("linear polarization", "线极化"),
    ("duty cycle", "占空比"),
    ("visibility cone", "可见锥"),
    ("simulation duration", "仿真时长"),
    ("time step", "时间步长"),
    ("orbital altitude", "轨道高度"),
    ("apogee altitude", "远地点高度"),
    ("perigee altitude", "近地点高度"),
    ("orbital inclination", "轨道倾角"),
    ("angle of inclination", "倾角"),
    ("mean anomaly", "平近点角"),
    ("argument of periapsis", "近地点幅角"),
    ("right ascension of the ascending node", "升交点赤经"),
    ("eccentricity", "偏心率"),
    ("Earth station deployment", "地球站部署"),
    ("earth station", "地球站"),
    ("space station", "空间电台"),
    ("International Telecommunication Union", "国际电信联盟"),
    ("Radiocommunication Sector", "无线电通信部门"),
    ("Radiocommunication Assembly", "无线电通信全会"),
    ("World Radiocommunication Conference", "世界无线电通信大会"),
    ("Radio Regulations", "《无线电规则》"),
    ("Earth exploration-satellite service", "卫星地球探测业务"),
    ("meteorological-satellite service", "气象卫星业务"),
    ("meteorological aids service", "气象辅助业务"),
    ("mobile-satellite service", "移动卫星业务"),
    ("amateur-satellite service", "业余卫星业务"),
    ("aeronautical mobile service", "航空移动业务"),
    ("maritime mobile service", "水上移动业务"),
    ("radio astronomy service", "射电天文业务"),
    ("space research service", "空间研究业务"),
    ("radiolocation service", "无线电定位业务"),
    ("radionavigation service", "无线电导航业务"),
    ("fixed-satellite service", "固定卫星业务"),
    ("fixed service", "固定业务"),
    ("mobile service", "移动业务"),
    ("broadcasting service", "广播业务"),
    ("primary allocation", "主要业务划分"),
    ("secondary allocation", "次要业务划分"),
    ("primary basis", "主要业务地位"),
    ("secondary basis", "次要业务地位"),
]

# Add the original glossary, keeping the first (more specific) target for a
# duplicated English phrase.
_seen: set[str] = set()
TERM_ITEMS: list[tuple[str, str]] = []
for src, dst in sorted(TERMS + pipeline.GLOSSARY, key=lambda x: len(x[0]), reverse=True):
    key = src.lower()
    if key not in _seen:
        _seen.add(key)
        TERM_ITEMS.append((src, dst))
TERM_LOOKUP = {src.lower(): dst for src, dst in TERM_ITEMS}
TERM_RE = re.compile("|".join(re.escape(src) for src, _ in TERM_ITEMS), re.I)

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
    "Remote sensing systems": "遥感系统",
    "Space applications and meteorology": "空间应用和气象",
    "Spectrum management": "频谱管理",
    "The role of the Radiocommunication Sector is to ensure the rational, equitable, efficient and economical use of the radio-frequency spectrum by all radiocommunication services, including satellite services, and carry out studies without limit of frequency range on the basis of which Recommendations are adopted.": "无线电通信部门的职责是确保包括卫星业务在内的所有无线电通信业务合理、公平、有效和经济地使用无线电频谱，并开展不受频率范围限制的研究，在此基础上通过建议书。",
    "The regulatory and policy functions of the Radiocommunication Sector are performed by World and Regional Radiocommunication Conferences and Radiocommunication Assemblies supported by Study Groups.": "无线电通信部门的规章和政策职能由世界和区域无线电通信大会以及无线电通信全会履行，并由各研究组提供支持。",
    "Note: This ITU-R Report was approved in English by the Study Group under the procedure detailed in Resolution ITU-R 1.": "注：本ITU-R报告英文版由研究组按照ITU-R第1号决议规定的程序批准。",
    "All rights reserved. No part of this publication may be reproduced, by any means whatsoever, without written permission of ITU.": "版权所有。未经国际电信联盟书面许可，不得以任何方式复制本出版物的任何部分。",
    "Studies to accommodate spectrum requirements in the space operation service for non-geostationary satellites with short duration missions": "满足短期任务非GSO卫星空间操作业务频谱需求的研究",
    "Technical characteristics for telemetry, tracking and command in the space operation service below 1 GHz for non-GSO satellites with short duration missions": "1 GHz以下短期任务非GSO卫星空间操作业务遥测、跟踪与遥控的技术特性",
    "Studies on the suitability of existing allocations to the space operation service below 1 GHz and additional sharing studies on possible new and/or upgraded allocations": "1 GHz以下空间操作业务现有频率划分适用性研究及可能新增和/或升级频率划分的补充共用研究",
}

ID_PATTERNS = [
    r"https?://\S+",
    r"\b(?:Recommendation|Report|Resolution)\s+ITU-R\s+[A-Z]+\.\d+(?:-\d+)?\b",
    r"\bITU-R\s+[A-Z]+\.\d+(?:-\d+)?\b",
    r"\bRR\s+No\.\s*\d+(?:\.\d+)*\b",
    r"\bWRC-\d{2}\b",
    r"\b(?:NGSO|GSO|LEO|TT&C|SOS|VHF|UHF|RF|PSD|CDF|RAAN|GMDSS|AIS|HADS|GOES|DCPR?|EESS|MetSat|MetAids|SRS|RAS|MSS|AM\(R\)S|COSPAS-SARSAT|TDMA|CDMA|SINAD|BER)\b",
    r"\b(?:C/I|C/N0?|I/N|S/N)\b",
    r"\be\.i\.r\.p\.\b",
    r"\b[A-Z]\.\d+(?:\.[A-Za-z0-9]+){1,5}\b",
    r"(?<!\w)[≤≥<>≈~+−-]?\s*\d+(?:[ .,:/×*+−-]\d+)*(?:\s*(?:GHz|MHz|kHz|Hz|dBW|dBm|dBi|dB/K|dB|kbit/s|bit/s/Hz|bit/s|km|m|ms|µs|s|K|W|%|degrees?))",
]
ID_RES = [re.compile(p) for p in ID_PATTERNS]

POST_REPLACEMENTS: list[tuple[str, str]] = [
    ("空间运营业务", "空间操作业务"),
    ("空间运行服务", "空间操作业务"),
    ("空间运营服务", "空间操作业务"),
    ("遥测、跟踪和命令", "遥测、跟踪与遥控"),
    ("遥测、跟踪和控制", "遥测、跟踪与遥控"),
    ("遥测、跟踪和指令", "遥测、跟踪与遥控"),
    ("非地球静止", "非GSO"),
    ("非地球同步", "非GSO"),
    ("非对地静止", "非GSO"),
    ("地球到空间", "地对空"),
    ("空间到地球", "空对地"),
    ("功率光谱密度", "功率谱密度"),
    ("等效各向同性辐射功率", "等效全向辐射功率"),
    ("外带发射", "带外发射"),
    ("虚假发射", "杂散发射"),
    ("保护标准", "保护准则"),
    ("频率分配", "频率划分"),
    ("频率任务", "频率指配"),
    ("地球站站", "地球站"),
    ("空间站站", "空间电台"),
]


def exact_translation(source: str) -> str | None:
    compact = " ".join(source.split())
    if compact in EXACT:
        return EXACT[compact]
    if re.fullmatch(r"(?:Report|Recommendation|Resolution) ITU-R [A-Z]+\.\d+(?:-\d+)?", compact):
        return compact
    if re.fullmatch(r"ITU-R [A-Z]+\.\d+(?:-\d+)?", compact):
        return compact
    return None


def protected_segments(text: str) -> list[tuple[str, str]]:
    """Return ordered (kind, text) spans; kind is 'translate' or 'fixed'."""
    matches: list[tuple[int, int, str]] = []
    for m in TERM_RE.finditer(text):
        matches.append((m.start(), m.end(), TERM_LOOKUP[m.group(0).lower()]))
    for regex in ID_RES:
        for m in regex.finditer(text):
            matches.append((m.start(), m.end(), m.group(0)))
    # Earliest match wins; at the same position use the longest span.
    matches.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    selected: list[tuple[int, int, str]] = []
    cursor = -1
    for start, end, replacement in matches:
        if start < cursor:
            continue
        selected.append((start, end, replacement))
        cursor = end
    spans: list[tuple[str, str]] = []
    pos = 0
    for start, end, replacement in selected:
        if start > pos:
            spans.append(("translate", text[pos:start]))
        spans.append(("fixed", replacement))
        pos = end
    if pos < len(text):
        spans.append(("translate", text[pos:]))
    return spans or [("translate", text)]


def clean_piece(text: str) -> str:
    out = pipeline.normalize_chinese(text)
    for old, new in POST_REPLACEMENTS:
        out = out.replace(old, new)
    out = re.sub(r"\s+([，。；：！？、])", r"\1", out)
    out = re.sub(r"([（《])\s+", r"\1", out)
    out = re.sub(r"\s+([）》])", r"\1", out)
    out = re.sub(r"[ \t]{2,}", " ", out)
    return out.strip()


def assemble(parts: list[str]) -> str:
    text = " ".join(p.strip() for p in parts if p and p.strip())
    text = clean_piece(text)
    # Remove spaces that machine translation leaves between adjacent Chinese
    # spans while retaining useful spacing around Latin identifiers and units.
    text = re.sub(r"(?<=[\u3400-\u9fff])\s+(?=[\u3400-\u9fff])", "", text)
    text = re.sub(r"(?<=[\u3400-\u9fff])\s+([，。；：！？、）])", r"\1", text)
    text = re.sub(r"([（])\s+(?=[\u3400-\u9fff])", r"\1", text)
    return text.strip()


def translate_units_v4(units: list[dict[str, Any]], model_dir: str, cache_path: Path) -> None:
    tokenizer = AutoTokenizer.from_pretrained(model_dir, local_files_only=True)
    model = AutoModelForSeq2SeqLM.from_pretrained(model_dir, local_files_only=True)
    model.eval()
    torch.set_num_threads(max(1, min(8, os.cpu_count() or 2)))

    cache_file = cache_path.with_name("translation_cache_v4.json")
    cache: dict[str, str] = {}
    if cache_file.exists():
        try:
            cache = json.loads(cache_file.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    plans: dict[int, list[tuple[str, str]]] = {}
    jobs: dict[str, str] = {}

    for index, unit in enumerate(units):
        source = str(unit["text"])
        full_key = hashlib.sha256(("v4-unit:" + source).encode("utf-8")).hexdigest()
        exact = exact_translation(source)
        if exact is not None:
            unit["zh"] = exact
            cache[full_key] = exact
            continue
        if full_key in cache:
            unit["zh"] = cache[full_key]
            continue
        if pipeline.mostly_nonlinguistic(source):
            unit["zh"] = source
            cache[full_key] = source
            continue

        plan: list[tuple[str, str]] = []
        for kind, span in protected_segments(source):
            if kind == "fixed":
                plan.append(("fixed", span))
                continue
            if not re.search(r"[A-Za-z]", span):
                plan.append(("fixed", span))
                continue
            for piece in pipeline.split_long(span, tokenizer, max_tokens=240):
                if not piece.strip():
                    continue
                piece_key = hashlib.sha256(("v4-piece:" + piece).encode("utf-8")).hexdigest()
                plan.append(("translated", piece_key))
                if piece_key not in cache:
                    jobs.setdefault(piece_key, piece)
        plans[index] = plan

    pipeline.log(f"V4 unique translation fragments: {len(jobs)}")
    job_items = list(jobs.items())
    batch_size = 24
    completed = 0
    for start in range(0, len(job_items), batch_size):
        batch = job_items[start : start + batch_size]
        sources = [source for _, source in batch]
        encoded = tokenizer(
            sources,
            return_tensors="pt",
            padding=True,
            truncation=True,
            max_length=280,
        )
        with torch.inference_mode():
            generated = model.generate(
                **encoded,
                max_new_tokens=320,
                num_beams=2,
                do_sample=False,
                early_stopping=True,
                no_repeat_ngram_size=4,
                repetition_penalty=1.05,
                use_cache=True,
            )
        outputs = tokenizer.batch_decode(generated, skip_special_tokens=True)
        for (key, source), output in zip(batch, outputs):
            cleaned = clean_piece(output)
            if not cleaned or re.search(r"_{6,}|X{8,}|ZXQPH", cleaned):
                cleaned = source.strip()
            cache[key] = cleaned
        completed += len(batch)
        if completed % 240 < batch_size or completed == len(job_items):
            pipeline.log(f"V4 translated {completed}/{len(job_items)} fragments")

    for index, unit in enumerate(units):
        if "zh" in unit:
            continue
        rendered: list[str] = []
        for kind, value in plans.get(index, []):
            rendered.append(value if kind == "fixed" else cache.get(value, ""))
        translation = assemble(rendered)
        source = str(unit["text"])
        if not translation:
            translation = source
        unit["zh"] = translation
        full_key = hashlib.sha256(("v4-unit:" + source).encode("utf-8")).hexdigest()
        cache[full_key] = translation

    cache_file.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


pipeline.translate_units = translate_units_v4

if __name__ == "__main__":
    raise SystemExit(pipeline.main())
