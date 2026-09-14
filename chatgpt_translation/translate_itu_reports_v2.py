#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
import os
import re

import torch
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer

import translate_itu_reports as pipeline


def fixed_protect_glossary(text: str) -> tuple[str, dict[str, str]]:
    mapping: dict[str, str] = {}
    output = text
    counter = 0

    def add_placeholder(value: str) -> str:
        nonlocal counter
        token = f"ZXQPH{counter:04d}QXZ"
        counter += 1
        mapping[token] = value
        return token

    for source, target in sorted(pipeline.GLOSSARY, key=lambda pair: len(pair[0]), reverse=True):
        pattern = re.compile(re.escape(source), re.I)
        while True:
            match = pattern.search(output)
            if not match:
                break
            token = add_placeholder(target)
            output = output[: match.start()] + token + output[match.end() :]

    patterns = [
        r"https?://\S+",
        r"\b(?:Recommendation|Report|Resolution)\s+ITU-R\s+[A-Z]+\.\d+(?:-\d+)?\b",
        r"\bITU-R\s+[A-Z]+\.\d+(?:-\d+)?\b",
        r"\bRR\s+No\.\s*\d+(?:\.\d+)*\b",
        r"\bWRC-\d{2}\b",
        r"\b(?!ZXQPH\d{4}QXZ\b)[A-Z][A-Z0-9/()-]{1,14}\b",
    ]
    for expression in patterns:
        regex = re.compile(expression)
        search_start = 0
        while True:
            match = regex.search(output, search_start)
            if not match:
                break
            token = add_placeholder(match.group(0))
            output = output[: match.start()] + token + output[match.end() :]
            search_start = match.start() + len(token)
    return output, mapping


def fixed_translate_units(units, model_dir: str, cache_path):
    tokenizer = AutoTokenizer.from_pretrained(model_dir, local_files_only=False)
    model = AutoModelForSeq2SeqLM.from_pretrained(model_dir, local_files_only=False)
    model.eval()
    torch.set_num_threads(max(1, min(8, os.cpu_count() or 2)))

    cache: dict[str, str] = {}
    if cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
            pipeline.log(f"Loaded translation cache: {len(cache)} entries")
        except Exception:
            cache = {}

    jobs = []
    for index, unit in enumerate(units):
        source = unit["text"]
        key = hashlib.sha256(source.encode("utf-8")).hexdigest()
        if key in cache:
            unit["zh"] = cache[key]
            continue
        if pipeline.mostly_nonlinguistic(source):
            unit["zh"] = source
            cache[key] = source
            continue
        protected, mapping = fixed_protect_glossary(source)
        # Keep each source chunk well below the 512-position Marian limit.
        parts = pipeline.split_long(protected, tokenizer, max_tokens=280)
        unit["_translation_parts"] = len(parts)
        for part_index, part in enumerate(parts):
            jobs.append((index, f"{key}:{part_index}", part, mapping))

    pipeline.log(f"Translation jobs: {len(jobs)} chunks; cached units: {sum('zh' in u for u in units)}")
    translated_parts: dict[int, list[str]] = {}
    batch_size = 20
    completed = 0
    for start in range(0, len(jobs), batch_size):
        batch = jobs[start : start + batch_size]
        sources = [item[2] for item in batch]
        encoded = tokenizer(
            sources,
            return_tensors="pt",
            padding=True,
            truncation=True,
            max_length=320,
        )
        with torch.inference_mode():
            generated = model.generate(
                **encoded,
                max_new_tokens=360,
                num_beams=1,
                do_sample=False,
                use_cache=True,
            )
        outputs = tokenizer.batch_decode(generated, skip_special_tokens=True)
        for item, output in zip(batch, outputs):
            unit_index, _, _, mapping = item
            restored = pipeline.normalize_chinese(pipeline.restore_placeholders(output, mapping))
            translated_parts.setdefault(unit_index, []).append(restored)
        completed += len(batch)
        if completed % 200 < batch_size or completed == len(jobs):
            pipeline.log(f"Translated {completed}/{len(jobs)} chunks")

    for index, unit in enumerate(units):
        if "zh" in unit:
            continue
        pieces = translated_parts.get(index, [])
        translation = pipeline.normalize_chinese(" ".join(piece for piece in pieces if piece))
        if not translation:
            translation = unit["text"]
        unit["zh"] = translation
        key = hashlib.sha256(unit["text"].encode("utf-8")).hexdigest()
        cache[key] = translation

    cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    pipeline.log(f"Saved translation cache: {cache_path}")


pipeline.protect_glossary = fixed_protect_glossary
pipeline.translate_units = fixed_translate_units

if __name__ == "__main__":
    raise SystemExit(pipeline.main())
