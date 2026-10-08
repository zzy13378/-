"""真实语料层：加载真实论文证据、切块、建立检索索引。

语料来自项目里已有的真实论文提取结果，每篇论文是一个独立文档，切块后每个块
有稳定编号（例如 R2-C03）。问题从这个语料里生成，金标准就是问题所依据的那个块，
因此召回率是可客观测量的，不需要人工写期望答案。
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Tuple


@dataclass
class Chunk:
    id: str
    doc_id: str
    title: str
    text: str


@dataclass
class Index:
    chunks: List[Chunk]
    df: Dict[str, int] = field(default_factory=dict)
    avg_len: float = 0.0
    doc_tokens: List[List[str]] = field(default_factory=list)


PAPER_HEADER = re.compile(r"^### \[(R\d+)\]\s*(.+)$")

# 基准测试结论：720 字的块在 Hit@3 与 MRR 上表现最好（见 reports/retrieval_bench.json）。
DEFAULT_CHUNK_SIZE = 720


def load_papers(paths: Iterable[Path]) -> List[Dict[str, str]]:
    """从报告文件里按 ### [Rn] 抽取真实论文条目。"""

    papers: Dict[str, Dict[str, str]] = {}
    for path in paths:
        path = Path(path)
        if not path.exists():
            continue
        current_id = ""
        buffer: List[str] = []
        for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
            match = PAPER_HEADER.match(line.strip())
            # 只保留 ### [Rn] 这一层的正文；遇到上一级 ## 标题说明论文部分结束，
            # 后面的内容属于报告自己的输出（写作 Skill、比较矩阵等），不能当语料。
            if current_id and line.strip().startswith("## ") and not line.strip().startswith("### "):
                papers.setdefault(current_id, {"id": current_id, "title": "", "text": ""})
                papers[current_id]["text"] += "\n".join(buffer) + "\n"
                buffer = []
                current_id = ""
                continue
            if match:
                if current_id and buffer:
                    papers.setdefault(current_id, {"id": current_id, "title": "", "text": ""})
                    papers[current_id]["text"] += "\n".join(buffer) + "\n"
                current_id = match.group(1)
                buffer = []
                entry = papers.setdefault(current_id, {"id": current_id, "title": match.group(2).strip(), "text": ""})
                entry["title"] = match.group(2).strip()
                continue
            if current_id:
                buffer.append(line)
        if current_id and buffer:
            papers.setdefault(current_id, {"id": current_id, "title": "", "text": ""})
            papers[current_id]["text"] += "\n".join(buffer) + "\n"

    documents = []
    for paper_id, item in papers.items():
        body = _clean(item["text"])
        if len(body) < 200:
            continue
        documents.append({"doc_id": paper_id, "title": item["title"], "text": body})
    return documents


def _clean(text: str) -> str:
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def chunk_documents(documents: Iterable[Dict[str, str]], size: int = 360) -> List[Chunk]:
    """按行累积切块，超长时按句号再切，保持块内语义完整。"""

    chunks: List[Chunk] = []
    for document in documents:
        doc_id = document["doc_id"]
        title = document.get("title", "")
        buffer = ""
        index = 1
        for line in document["text"].splitlines():
            line = line.strip()
            if not line:
                continue
            if buffer and len(buffer) + len(line) > size:
                chunks.append(Chunk(f"{doc_id}-C{index:02d}", doc_id, title, buffer.strip()))
                index += 1
                buffer = line
            else:
                buffer = f"{buffer} {line}".strip()
            while len(buffer) > size * 2:
                cut = buffer.rfind("。", 0, size)
                cut = cut + 1 if cut > size * 0.5 else size
                chunks.append(Chunk(f"{doc_id}-C{index:02d}", doc_id, title, buffer[:cut].strip()))
                index += 1
                buffer = buffer[cut:]
        if buffer.strip():
            chunks.append(Chunk(f"{doc_id}-C{index:02d}", doc_id, title, buffer.strip()))
    return chunks


def tokenize(text: str) -> List[str]:
    """中文用二元字组，英文数字用词，兼顾无分词器的场景。"""

    tokens: List[str] = []
    text = text.lower()
    for word in re.findall(r"[a-z0-9]{2,}", text):
        tokens.append(word)
    chinese = re.sub(r"[^\u4e00-\u9fff]", " ", text)
    for run in chinese.split():
        if len(run) == 1:
            tokens.append(run)
        else:
            tokens.extend(run[i:i + 2] for i in range(len(run) - 1))
    return tokens


def build_index(chunks: List[Chunk]) -> Index:
    # 标题一起进索引：问题里常出现研究对象名，标题往往比正文更集中地包含它。
    doc_tokens = [tokenize(f"{chunk.title} {chunk.text}") for chunk in chunks]
    df: Dict[str, int] = {}
    for tokens in doc_tokens:
        for token in set(tokens):
            df[token] = df.get(token, 0) + 1
    avg_len = sum(len(tokens) for tokens in doc_tokens) / len(doc_tokens) if doc_tokens else 0.0
    return Index(chunks=chunks, df=df, avg_len=avg_len, doc_tokens=doc_tokens)


def search(index: Index, query: str, top_k: int = 3,
           allowed_docs: Iterable[str] | None = None) -> List[Tuple[Chunk, float]]:
    """BM25 检索。allowed_docs 用于限定只能在某些文档内检索。"""

    tokens = tokenize(query)
    total = len(index.chunks)
    if not tokens or not total:
        return []
    scores: List[Tuple[int, float]] = []
    k1, b = 1.5, 0.75
    allow = set(allowed_docs) if allowed_docs else None
    for position, chunk in enumerate(index.chunks):
        if allow is not None and chunk.doc_id not in allow:
            continue
        doc_tokens = index.doc_tokens[position]
        length = len(doc_tokens) or 1
        counts: Dict[str, int] = {}
        for token in doc_tokens:
            counts[token] = counts.get(token, 0) + 1
        score = 0.0
        for token in set(tokens):
            tf = counts.get(token, 0)
            if not tf:
                continue
            idf = math.log(1 + (total - index.df.get(token, 0) + 0.5) / (index.df.get(token, 0) + 0.5))
            score += idf * (tf * (k1 + 1)) / (tf + k1 * (1 - b + b * length / (index.avg_len or 1)))
        if score > 0:
            scores.append((position, score))
    scores.sort(key=lambda item: -item[1])
    return [(index.chunks[position], round(score, 4)) for position, score in scores[:top_k]]


def rerank_by_idf(index: Index, query: str, candidate_ids: List[str], top_k: int) -> List[str]:
    """第二阶段重排：按 IDF 加权命中覆盖率排序，抑制长块和常见词。"""

    totals = len(index.chunks) or 1
    query_tokens = set(tokenize(query))
    scored: List[Tuple[str, float]] = []
    by_id = {chunk.id: chunk for chunk in index.chunks}
    for chunk_id in candidate_ids:
        chunk = by_id.get(chunk_id)
        if chunk is None:
            continue
        tokens = set(tokenize(chunk.text))
        weight = 0.0
        for token in query_tokens & tokens:
            df = index.df.get(token, 0)
            weight += math.log(1 + (totals - df + 0.5) / (df + 0.5))
        scored.append((chunk_id, weight / ((len(tokens) ** 0.35) or 1.0)))
    scored.sort(key=lambda pair: -pair[1])
    return [chunk_id for chunk_id, _ in scored[:top_k]]


def build_from_reports(paths: Iterable[Path], size: int = DEFAULT_CHUNK_SIZE) -> Tuple[List[Chunk], Index]:
    documents = load_papers(paths)
    chunks = chunk_documents(documents, size=size)
    return chunks, build_index(chunks)
