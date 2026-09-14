#!/usr/bin/env python3
from __future__ import annotations

import re

from chatgpt_translation import translate_itu_reports as pipeline


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
        # Do not re-protect placeholders that were inserted for glossary terms.
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


pipeline.protect_glossary = fixed_protect_glossary

if __name__ == "__main__":
    raise SystemExit(pipeline.main())
