"""真实检索问答：BM25 召回 + IDF 重排 + 大模型依据证据作答。

这是被评测的对象本身，不参与评分；评分由 alignment.py 与 judge.py 负责。
"""

from __future__ import annotations

import json
import time
from typing import Dict, List

from . import runtime
from .realcorpus import Index, rerank_by_idf, search

ANSWER_SYSTEM = """你是论文证据问答助手。你只能依据给定的证据片段回答，不得编造。
要求：
1. 每个论断句后面标注来源编号，格式如 [R2-C06]；
2. 只能使用证据里出现过的编号，不得引用其他编号；
3. 证据不足以回答时，直接说明证据不足，不要给确定结论；
4. 来源编号必须紧跟在对应论断句末尾，不得单独占一行或单独成句；
5. 只输出 JSON：{"answer":"正文","cited":["R2-C06"]}"""

EXPAND_SYSTEM = """你是学术检索助手。请把用户的问题转成用于关键词检索的查询词。
要求：
1. 必须同时给出中文关键词和对应的英文关键词，两部分都不能为空，英文部分用于检索英文论文；
2. 每部分给出 3 到 6 个词或短语，覆盖研究对象、方法、数据与关键概念；
3. 只输出 JSON：{"zh":["武夷山","精制茶"],"en":["Wuyishan","refined tea"]}"""


def expand_query(client, model: str, question: str, counter: Dict[str, int]) -> List[str]:
    """跨语言查询扩展：中文问题补英文术语，避免英文文献漏召回。"""

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "system", "content": EXPAND_SYSTEM}, {"role": "user", "content": f"问题：{question}"}],
            temperature=0.0,
            max_tokens=300,
        )
        counter["tokens"] += int(getattr(response.usage, "total_tokens", 0) or 0)
        content = (response.choices[0].message.content or "").strip()
        start, end = content.find("{"), content.rfind("}")
        if start >= 0 and end > start:
            payload = json.loads(content[start:end + 1])
            terms: List[str] = []
            for key in ("zh", "en", "terms"):
                for item in payload.get(key) or []:
                    value = str(item).strip()
                    if value and value not in terms:
                        terms.append(value)
            return terms[:12]
    except Exception:  # noqa: BLE001
        pass
    return []


def answer_question(index: Index, question: str, config=None) -> Dict:
    """对单个问题执行完整问答，返回答案、召回、引用与消耗。"""

    config = config or runtime.default_config()
    client = runtime.client()
    counter = {"tokens": 0}
    started = time.time()

    terms = expand_query(client, config.model, question, counter)
    query = " ".join([question] + terms)
    recall_k = max(20, config.top_k * 6)
    candidates = [chunk.id for chunk, _ in search(index, query, top_k=recall_k)]
    ordered = rerank_by_idf(index, query, candidates, config.top_k)
    by_id = {chunk.id: chunk for chunk in index.chunks}
    hits = [by_id[cid] for cid in ordered if cid in by_id]
    retrieved = [chunk.id for chunk in hits]
    evidence = "\n\n".join(f"[{chunk.id}]（{chunk.title}）{chunk.text}" for chunk in hits)

    error = ""
    answer, cited = "", []
    try:
        response = client.chat.completions.create(
            model=config.model,
            messages=[
                {"role": "system", "content": ANSWER_SYSTEM},
                {"role": "user", "content": f"问题：{question}\n\n证据片段：\n{evidence or '（无检索结果）'}"},
            ],
            temperature=config.temperature,
            max_tokens=900,
        )
        counter["tokens"] += int(getattr(response.usage, "total_tokens", 0) or 0)
        content = (response.choices[0].message.content or "").strip()
        start, end = content.find("{"), content.rfind("}")
        if start >= 0 and end > start:
            payload = json.loads(content[start:end + 1])
            answer = str(payload.get("answer", ""))
            cited = [str(item) for item in (payload.get("cited") or [])]
        else:
            error = "empty_response"
    except Exception as exc:  # noqa: BLE001
        error = f"{type(exc).__name__}: {exc}"

    return {
        "answer": answer,
        "cited": cited,
        "retrieved": retrieved,
        "evidence_text": evidence,
        "query_terms": terms,
        "latency_ms": round((time.time() - started) * 1000, 1),
        "tokens": counter["tokens"],
        "cost": round(counter["tokens"] * 0.0000011, 6),
        "error": error,
    }
