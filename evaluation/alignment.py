"""句子级对齐评估。

它回答的不再是“这条整体过没过”，而是“具体错在哪一句”：
1. 正向：回答里的每个论断句，能否追溯到它引用的证据块；
2. 反向：标准证据里的每个句子，有多少被回答覆盖到。
低重合的句子先用规则筛出来，再用模型逐句核实，兼顾成本与准确度。
"""

from __future__ import annotations

import json
import re
import time
from typing import Dict, List, Optional

SENTENCE_SPLIT = re.compile(r"(?<=[。！？；!?;])")
CITATION = re.compile(r"\[(R\d+-C\d+)\]")
DOC_CITATION = re.compile(r"\[(R\d+)\]")

# 这些写法属于对证据状态的说明，不是事实论断，不要求标注引用。
META_PATTERNS = re.compile(
    r"证据不足|无法回答|无法判断|无法确认|无法说明|无法据此|无法给出|无法进一步|"
    r"未明确|未给出|未提供|未进一步|没有提供|不存在|不足以|不予回答|仅能看出|仅能说明"
)

# 转述文献内容的事实句，必须标注来源。没有标注就是归因缺失。
FACT_PATTERNS = re.compile(
    r"已有研究|研究表明|研究指出|研究显示|研究认为|有学者|一项研究|一项基于|案例|"
    r"结果显示|结果发现|研究发现|调查显示|数据表明|文献指出|文献认为|样本|实验"
)

# 作者自己的界定、推论与范围说明，不要求标注来源。
ANALYSIS_PATTERNS = re.compile(
    r"本研究|本文|本综述|本项研究|就本研究而言|就本文而言|意味着|据此|由此可见|"
    r"需要指出的是|需要说明的是|综上所述|本章|本节|上述|这一|该结论|总体而言"
)

# 语料里的元信息行（作者、年份、DOI 等）不是可被回答覆盖的内容，统计覆盖率时跳过。
METADATA_PATTERNS = re.compile(
    r"^\s*作者：|^\s*年份：|DOI|字段覆盖率|^\s*\|.*\|$|^\s*-\s*\*\*(研究背景|研究问题|理论基础|研究方法|数据来源|研究设计|核心结论|研究局限|可借鉴内容)\*\*（[^）]*）：\s*$"
)


def split_sentences(text: str, min_length: int = 8) -> List[str]:
    parts = [part.strip() for part in SENTENCE_SPLIT.split(text or "")]
    return [part for part in parts if len(part) >= min_length]


def bigrams(text: str) -> set:
    clean = "".join(ch for ch in (text or "") if "\u4e00" <= ch <= "\u9fff" or ch.isalnum()).lower()
    return {clean[i:i + 2] for i in range(max(0, len(clean) - 1))}


def overlap(source: str, target: str) -> float:
    """source 的内容有多少比例能在 target 里找到，用于判断引用是否支撑该句。"""

    a, b = bigrams(source), bigrams(target)
    if not a or not b:
        return 0.0
    return len(a & b) / len(a)


def align_answer(answer: str, citations: List[str], chunk_texts: Dict[str, str],
                 threshold: float = 0.15, citation_pattern: re.Pattern = CITATION) -> List[Dict]:
    """逐句标注：论断句是否可追溯，说明句与格式片段单独归类。"""

    rows = []
    for sentence in split_sentences(answer):
        cited = [cid for cid in citation_pattern.findall(sentence)]
        stripped = citation_pattern.sub("", sentence).strip()
        if len(stripped) < 6:
            rows.append({"sentence": sentence, "cited": cited, "kind": "格式片段",
                         "best_overlap": 0.0, "verdict": "格式片段"})
            continue
        if META_PATTERNS.search(stripped):
            rows.append({"sentence": sentence, "cited": cited, "kind": "说明句",
                         "best_overlap": 0.0, "verdict": "说明句"})
            continue
        if not cited:
            # 没有引用时再区分：转述文献事实属于归因缺失，作者分析句不算问题。
            if FACT_PATTERNS.search(stripped):
                rows.append({"sentence": sentence, "cited": [], "kind": "文献事实句",
                             "best_overlap": 0.0, "verdict": "未标注引用"})
            elif ANALYSIS_PATTERNS.search(stripped):
                rows.append({"sentence": sentence, "cited": [], "kind": "作者分析句",
                             "best_overlap": 0.0, "verdict": "作者分析"})
            else:
                rows.append({"sentence": sentence, "cited": [], "kind": "论断句",
                             "best_overlap": 0.0, "verdict": "未标注引用"})
            continue
        best, best_id = 0.0, ""
        for cid in cited:
            score = overlap(sentence, chunk_texts.get(cid, ""))
            if score > best:
                best, best_id = score, cid
        verdict = "可追溯" if best >= threshold else "引用可疑"
        rows.append({"sentence": sentence, "cited": cited, "kind": "文献事实句", "best_chunk": best_id,
                     "best_overlap": round(best, 3), "verdict": verdict})
    return rows


def gold_coverage(gold_text: str, answer: str, threshold: float = 0.3) -> Dict:
    """反向覆盖：标准证据里的句子有多少被回答覆盖到。"""

    gold_sentences = split_sentences(gold_text, min_length=12)
    answer_sentences = split_sentences(answer)
    covered, missed = [], []
    for sentence in gold_sentences:
        if METADATA_PATTERNS.search(sentence):
            continue
        best = max((overlap(sentence, item) for item in answer_sentences), default=0.0)
        (covered if best >= threshold else missed).append(
            {"sentence": sentence, "best_overlap": round(best, 3)}
        )
    return {
        "gold_sentences": len(covered) + len(missed),
        "covered": len(covered),
        "coverage": round(len(covered) / len(gold_sentences), 3) if gold_sentences else 0.0,
        "covered_detail": covered,
        "missed_detail": missed,
    }


VERIFY_PROMPT = """你是证据核对员。下面给出一个论断句、它引用的证据，以及原论文的相关证据。
请判断该论断句是否被引用的证据支持。注意：中文句子引用英文证据时，用词不同但事实一致也算支持。
只输出 JSON：{"supported":true,"reason":"一句话理由"}"""


class SentenceVerifier:
    """对低重合句子做二次核实，避免跨语言或改写导致的误判。"""

    def __init__(self, model: str = "") -> None:
        import os
        from openai import OpenAI

        self.model = model or os.environ.get("LLM_MODEL", "deepseek-chat")
        self.client = OpenAI(
            api_key=os.environ.get("OPENAI_API_KEY", "any"),
            base_url=os.environ.get("OPENAI_BASE_URL", "http://127.0.0.1:15721/v1"),
            timeout=120, max_retries=0,
        )
        self.tokens = 0

    def verify(self, sentence: str, cited_text: str, gold_text: str):
        user = (
            f"论断句：{sentence}\n\n"
            f"它引用的证据：\n{cited_text or '（无）'}\n\n"
            f"原论文相关证据：\n{gold_text or '（无）'}"
        )
        for _ in range(3):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[{"role": "system", "content": VERIFY_PROMPT}, {"role": "user", "content": user}],
                    temperature=0.0,
                    max_tokens=250,
                )
                content = (response.choices[0].message.content or "").strip()
                self.tokens += int(getattr(response.usage, "total_tokens", 0) or 0)
                start, end = content.find("{"), content.rfind("}")
                if start >= 0 and end > start:
                    payload = json.loads(content[start:end + 1])
                    return bool(payload.get("supported")), str(payload.get("reason", ""))[:160]
            except Exception:  # noqa: BLE001
                time.sleep(2)
        return None, "核实失败"


def summarize_alignment(cases: List[Dict]) -> Dict:
    """汇总句子级指标。"""

    total_sentences = traceable = suspicious = missing_citation = meta = fragment = 0
    analysis = 0
    verified_supported = verified_total = 0
    gold_total = gold_covered = 0
    for item in cases:
        rows = item.get("sentences") or []
        total_sentences += len(rows)
        for row in rows:
            kind = row.get("kind") or "论断句"
            if kind == "格式片段":
                fragment += 1
                continue
            if kind == "说明句":
                meta += 1
                continue
            if row["verdict"] == "作者分析":
                analysis += 1
                continue
            if row["verdict"] == "未标注引用":
                missing_citation += 1
            elif row["verdict"] == "可追溯":
                traceable += 1
            else:
                suspicious += 1
                if row.get("verified") is not None:
                    verified_total += 1
                    if row["verified"]:
                        verified_supported += 1
        coverage = item.get("coverage") or {}
        gold_total += coverage.get("gold_sentences", 0)
        gold_covered += coverage.get("covered", 0)
    claims = traceable + suspicious + missing_citation
    return {
        "sentences": total_sentences,
        "claims": claims,
        "meta": meta,
        "analysis": analysis,
        "fragment": fragment,
        "missing_citation": missing_citation,
        "missing_citation_rate": round(missing_citation / claims, 3) if claims else 0.0,
        "traceable": traceable,
        "traceable_rate": round(traceable / claims, 3) if claims else 0.0,
        "suspicious": suspicious,
        "suspicious_rate": round(suspicious / claims, 3) if claims else 0.0,
        "verified": verified_total,
        "verified_supported": verified_supported,
        "gold_sentences": gold_total,
        "gold_covered": gold_covered,
        "gold_coverage": round(gold_covered / gold_total, 3) if gold_total else 0.0,
    }
