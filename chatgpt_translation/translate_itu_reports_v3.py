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

# Use the familiar ITU rendering of NGSO in the document titles.
pipeline.REPORTS["SA.2425"]["zh_title"] = "满足短期任务非GSO卫星空间操作业务频谱需求的研究"
pipeline.REPORTS["SA.2426"]["zh_title"] = "1 GHz以下短期任务非GSO卫星空间操作业务遥测、跟踪与遥控的技术特性"

# These translations are inserted directly into the source text before machine
# translation.  Unlike the previous placeholder approach, this does not depend
# on the model reproducing an artificial token character-for-character.
EXTRA_GLOSSARY: list[tuple[str, str]] = [
    ("non-geostationary satellites with short duration missions", "短期任务非GSO卫星"),
    ("non-GSO satellites with short duration missions", "短期任务非GSO卫星"),
    ("NGSO satellites with short duration missions", "短期任务非GSO卫星"),
    ("NGSO short duration mission satellites", "短期任务非GSO卫星"),
    ("short duration NGSO satellites", "短期任务非GSO卫星"),
    ("short duration mission satellites", "短期任务卫星"),
    ("short duration missions", "短期任务"),
    ("space operations service", "空间操作业务"),
    ("space operation service", "空间操作业务"),
    ("telemetry, tracking and command", "遥测、跟踪与遥控"),
    ("telemetry, tracking and control", "遥测、跟踪与遥控"),
    ("spectrum requirements", "频谱需求"),
    ("spectrum requirement", "频谱需求"),
    ("frequency allocations", "频率划分"),
    ("frequency allocation", "频率划分"),
    ("existing allocations", "现有频率划分"),
    ("new and/or upgraded allocations", "新增和/或升级的频率划分"),
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
    ("allocated bandwidth", "已划分带宽"),
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
    ("minimum elevation angle", "最小仰角"),
    ("minimum ground elevation angle", "最小地面仰角"),
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
    ("Earth station", "地球站"),
    ("earth station", "地球站"),
    ("Space station", "空间电台"),
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

# Exact short labels and standard front-matter sentences.  Machine translation
# models are weakest on isolated table labels, so handle them deterministically.
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
    "All rights reserved. No part of this publication may be reproduced, by any means whatsoever, without written permission of ITU.": "版权所有。未经国际电信联盟书面许可，不得以任何方式复制本出版物的任何部分。",
    "The regulatory and policy functions of the Radiocommunication Sector are performed by World and Regional Radiocommunication Conferences and Radiocommunication Assemblies supported by Study Groups.": "无线电通信部门的规章和政策职能由世界和区域无线电通信大会以及无线电通信全会履行，并由各研究组提供支持。",
}

# Keep file identifiers, acronyms and compact technical symbols readable after
# translation.  They are not replaced before translation; these substitutions
# only repair recurring model variants afterwards.
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
    ("拨款", "频率划分"),
    ("地球站站", "地球站"),
    ("空间站站", "空间电台"),
]

ALL_GLOSSARY = sorted(EXTRA_GLOSSARY + pipeline.GLOSSARY, key=lambda pair: len(pair[0]), reverse=True)


def prepare_source(text: str) -> str:
    out = text
    for source, target in ALL_GLOSSARY:
        out = re.sub(re.escape(source), target, out, flags=re.I)
    # Translate common standalone labels inside extracted table rows.
    label_patterns = [
        (r"\bTABLE\b", "表"),
        (r"\bFIGURE\b", "图"),
        (r"\bANNEX\b", "附件"),
        (r"\bAPPENDIX\b", "附录"),
        (r"\bPrimary\b", "主要业务"),
        (r"\bSecondary\b", "次要业务"),
        (r"\bDirection\b", "方向"),
        (r"\bStatus\b", "业务地位"),
        (r"\bRemarks\b", "备注"),
        (r"\bParameter\b", "参数"),
        (r"\bValue\b", "数值"),
        (r"\bUnit\b", "单位"),
    ]
    for pattern, replacement in label_patterns:
        out = re.sub(pattern, replacement, out, flags=re.I)
    return out


def clean_translation(text: str) -> str:
    out = pipeline.normalize_chinese(text)
    for source, target in POST_REPLACEMENTS:
        out = out.replace(source, target)
    out = re.sub(r"\bZXQPH\w*\b", "", out)
    out = re.sub(r"X{8,}", "", out)
    out = re.sub(r"\s+([，。；：！？、])", r"\1", out)
    out = re.sub(r"([（《])\s+", r"\1", out)
    out = re.sub(r"\s+([）》])", r"\1", out)
    out = re.sub(r"[ \t]{2,}", " ", out)
    return out.strip()


def translated_exact(source: str) -> str | None:
    compact = " ".join(source.split())
    if compact in EXACT:
        return EXACT[compact]
    # Preserve report/recommendation identifiers and compact formula-like rows.
    if re.fullmatch(r"(?:Report|Recommendation|Resolution) ITU-R [A-Z]+\.\d+(?:-\d+)?", compact):
        return compact
    if re.fullmatch(r"ITU-R [A-Z]+\.\d+(?:-\d+)?", compact):
        return compact
    return None


def translate_units_v3(units: list[dict[str, Any]], model_dir: str, cache_path: Path) -> None:
    tokenizer = AutoTokenizer.from_pretrained(model_dir, local_files_only=True)
    model = AutoModelForSeq2SeqLM.from_pretrained(model_dir, local_files_only=True)
    model.eval()
    torch.set_num_threads(max(1, min(8, os.cpu_count() or 2)))

    cache: dict[str, str] = {}
    if cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    jobs: list[tuple[int, str, str]] = []
    for index, unit in enumerate(units):
        source = str(unit["text"])
        key = hashlib.sha256(("v3:" + source).encode("utf-8")).hexdigest()
        exact = translated_exact(source)
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
        prepared = prepare_source(source)
        parts = pipeline.split_long(prepared, tokenizer, max_tokens=260)
        unit["_translation_parts"] = len(parts)
        for part_index, part in enumerate(parts):
            jobs.append((index, f"{key}:{part_index}", part))

    pipeline.log(f"V3 translation jobs: {len(jobs)} chunks")
    translated_parts: dict[int, list[str]] = {}
    batch_size = 16
    completed = 0
    for start in range(0, len(jobs), batch_size):
        batch = jobs[start : start + batch_size]
        sources = [item[2] for item in batch]
        encoded = tokenizer(
            sources,
            return_tensors="pt",
            padding=True,
            truncation=True,
            max_length=300,
        )
        with torch.inference_mode():
            generated = model.generate(
                **encoded,
                max_new_tokens=340,
                num_beams=2,
                do_sample=False,
                early_stopping=True,
                no_repeat_ngram_size=4,
                repetition_penalty=1.05,
                use_cache=True,
            )
        outputs = tokenizer.batch_decode(generated, skip_special_tokens=True)
        for item, output in zip(batch, outputs):
            unit_index = item[0]
            translated_parts.setdefault(unit_index, []).append(clean_translation(output))
        completed += len(batch)
        if completed % 200 < batch_size or completed == len(jobs):
            pipeline.log(f"V3 translated {completed}/{len(jobs)} chunks")

    for index, unit in enumerate(units):
        if "zh" in unit:
            continue
        translation = clean_translation(" ".join(translated_parts.get(index, [])))
        source = str(unit["text"])
        if not translation or re.search(r"ZXQPH|X{8,}", translation):
            translation = prepare_source(source)
        unit["zh"] = translation
        key = hashlib.sha256(("v3:" + source).encode("utf-8")).hexdigest()
        cache[key] = translation

    cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


pipeline.translate_units = translate_units_v3

if __name__ == "__main__":
    raise SystemExit(pipeline.main())
