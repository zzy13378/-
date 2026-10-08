"""内置高质量文献库。

把常用高质量中英文期刊的论文元数据与摘要预先收进本地文件，检索时先查本地库，
只有本地结果不足时才联网补充，避免每次检索都重复消耗 OpenAlex 的调用额度。

语料文件是 JSONL，每行一篇论文，字段与 Paper 兼容，可直接转成 Paper 对象。
"""

from __future__ import annotations

import json
import math
import re
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

PROJECT_DIR = Path(__file__).resolve().parent
CORPUS_PATH = PROJECT_DIR / "data" / "local_corpus.jsonl"
MANIFEST_PATH = PROJECT_DIR / "data" / "local_corpus_manifest.json"


@dataclass
class CorpusRecord:
    paper_id: str
    title: str
    authors: List[str] = field(default_factory=list)
    year: Optional[int] = None
    abstract: str = ""
    doi: str = ""
    url: str = ""
    venue: str = ""
    language: str = "en"
    cited_by_count: int = 0
    source: str = "内置文献库"
    fetched_at: str = ""
    # 全文相关字段：只有开放获取论文才会抓取全文，闭源论文保持空值并靠
    # text_quality 标注为摘要级证据。
    oa_url: str = ""
    is_oa: bool = False
    fulltext_path: str = ""
    fulltext_chars: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @staticmethod
    def from_dict(data: Dict[str, Any]) -> "CorpusRecord":
        return CorpusRecord(
            paper_id=str(data.get("paper_id") or ""),
            title=str(data.get("title") or ""),
            authors=list(data.get("authors") or []),
            year=data.get("year"),
            abstract=str(data.get("abstract") or ""),
            doi=str(data.get("doi") or ""),
            url=str(data.get("url") or ""),
            venue=str(data.get("venue") or ""),
            language=str(data.get("language") or "en"),
            cited_by_count=int(data.get("cited_by_count") or 0),
            source=str(data.get("source") or "内置文献库"),
            fetched_at=str(data.get("fetched_at") or ""),
            oa_url=str(data.get("oa_url") or ""),
            is_oa=bool(data.get("is_oa")),
            fulltext_path=str(data.get("fulltext_path") or ""),
            fulltext_chars=int(data.get("fulltext_chars") or 0),
        )

    def text(self) -> str:
        return f"{self.title} {self.abstract} {' '.join(self.authors[:3])} {self.venue}"

    def index_text(self) -> str:
        """参与相关性计算的文本：只保留标题与摘要，避免作者名和刊名带来噪声。"""

        return f"{self.title} {self.abstract}"

    def evidence_text(self) -> str:
        """提供给抽取阶段使用的证据文本：有全文用全文，否则退回摘要。"""

        if self.fulltext_path:
            path = Path(self.fulltext_path)
            if not path.is_absolute():
                path = PROJECT_DIR / self.fulltext_path
            try:
                if path.exists():
                    return path.read_text(encoding="utf-8", errors="replace")
            except OSError:
                pass
        return self.abstract

    def evidence_level(self) -> str:
        return "全文" if self.fulltext_chars else ("摘要" if self.abstract else "仅书目信息")


def dedup_key(record: CorpusRecord) -> str:
    if record.doi:
        return "doi:" + record.doi.lower().replace("https://doi.org/", "")
    return "title:" + re.sub(r"\W+", "", record.title).casefold()


def load_corpus(path: Path = CORPUS_PATH) -> List[CorpusRecord]:
    if not Path(path).exists():
        return []
    records: List[CorpusRecord] = []
    with Path(path).open("r", encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                records.append(CorpusRecord.from_dict(json.loads(line)))
            except (ValueError, TypeError):
                continue
    return records


def save_corpus(records: Iterable[CorpusRecord], path: Path = CORPUS_PATH) -> int:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    seen = set()
    count = 0
    with path.open("w", encoding="utf-8") as handle:
        for record in records:
            key = dedup_key(record)
            if not key or key in seen:
                continue
            seen.add(key)
            handle.write(json.dumps(record.to_dict(), ensure_ascii=False) + "\n")
            count += 1
    return count


def merge_corpus(new_records: Iterable[CorpusRecord], path: Path = CORPUS_PATH,
                 overwrite: bool = False) -> Tuple[int, int]:
    """把新记录并入现有语料，返回（总条数，新增条数）。

    overwrite=True 时用新记录覆盖同 key 的旧记录，用于补充开放获取链接这类新增字段。
    """

    existing = load_corpus(path)
    index = {dedup_key(record): record for record in existing}
    before = len(index)
    for record in new_records:
        key = dedup_key(record)
        if overwrite or key not in index:
            index[key] = record
    merged = list(index.values())
    total = save_corpus(merged, path)
    return total, total - before


def _tokens(text: str) -> List[str]:
    text = (text or "").lower()
    tokens = re.findall(r"[a-z][a-z0-9\-]{2,}", text)
    chinese = re.sub(r"[^\u4e00-\u9fff]", " ", text)
    for run in chinese.split():
        if len(run) >= 2:
            tokens.extend(run[i:i + 2] for i in range(len(run) - 1))
    return tokens


class LocalCorpus:
    """本地语料检索：BM25 + 质量加权，不依赖任何外部服务。"""

    def __init__(self, records: Optional[List[CorpusRecord]] = None, path: Path = CORPUS_PATH) -> None:
        self.records = records if records is not None else load_corpus(path)
        self.tokens: List[List[str]] = [_tokens(record.index_text()) for record in self.records]
        self.df: Dict[str, int] = {}
        for tokens in self.tokens:
            for token in set(tokens):
                self.df[token] = self.df.get(token, 0) + 1
        self.avg_len = (sum(len(t) for t in self.tokens) / len(self.tokens)) if self.tokens else 0.0
        # 出现频率过高的词（研究、基于、空间这类）区分度低，打分时直接忽略。
        total = max(1, len(self.records))
        self.stop_tokens = {token for token, count in self.df.items() if count / total > 0.25}

    def __len__(self) -> int:
        return len(self.records)

    @staticmethod
    def quality(record: CorpusRecord) -> float:
        current = datetime.now().year
        year = record.year or 0
        age = max(1, current - year + 1) if year else 10
        per_year = max(0, record.cited_by_count) / age
        score = min(1.0, math.log1p(per_year) / math.log1p(200)) * 0.6
        if record.venue:
            score += 0.2
        if record.doi:
            score += 0.1
        if year and (current - year) <= 5:
            score += 0.1
        return round(min(1.0, score), 3)

    def search(self, topic: str, limit: int = 8, from_year: Optional[int] = None,
               sort: str = "quality", language: str = "") -> List[CorpusRecord]:
        """先按主题锚点严格匹配，且只取带摘要的论文；严格无结果才放宽。"""

        results = self._search(topic, limit, from_year, sort, language,
                               require_anchor=True, require_abstract=True)
        if not results:
            results = self._search(topic, limit, from_year, sort, language,
                                   require_anchor=False, require_abstract=True)
        if not results:
            # 仍然没有，才允许返回无摘要记录，保证不会搜空。
            results = self._search(topic, limit, from_year, sort, language,
                                   require_anchor=False, require_abstract=False)
        return results

    def _search(self, topic: str, limit: int, from_year: Optional[int], sort: str,
                language: str, require_anchor: bool, require_abstract: bool = True) -> List[CorpusRecord]:
        raw_tokens = set(_tokens(topic))
        if not raw_tokens or not self.records:
            return []
        query_tokens = {token for token in raw_tokens if token not in self.stop_tokens} or raw_tokens
        anchors = self._anchors(topic)
        total = len(self.records)
        k1, b = 1.5, 0.75
        scored: List[Tuple[float, float, CorpusRecord]] = []
        for position, record in enumerate(self.records):
            if from_year and record.year and record.year < from_year:
                continue
            if language and record.language != language:
                continue
            if require_abstract and not record.abstract:
                continue
            tokens = self.tokens[position]
            if not tokens:
                continue
            haystack = record.index_text()
            anchor_hits = sum(1 for anchor in anchors if anchor in haystack) if anchors else 0
            counts: Dict[str, int] = {}
            for token in tokens:
                counts[token] = counts.get(token, 0) + 1
            bm25 = 0.0
            hits = 0
            for token in query_tokens:
                tf = counts.get(token, 0)
                if not tf:
                    continue
                hits += 1
                idf = math.log(1 + (total - self.df.get(token, 0) + 0.5) / (self.df.get(token, 0) + 0.5))
                bm25 += idf * (tf * (k1 + 1)) / (
                    tf + k1 * (1 - b + b * len(tokens) / (self.avg_len or 1))
                )
            coverage = hits / len(query_tokens)
            if bm25 <= 0 and not anchor_hits:
                continue
            if require_anchor:
                # 严格模式：中文主题必须有锚点命中，否则只是碰巧共享了通用词。
                if anchors and anchor_hits == 0:
                    continue
            else:
                # 放宽模式也不能没有底线：要么命中锚点，要么至少覆盖一半的判别词，
                # 否则宁可返回空结果，让上层回退到在线检索。
                if anchor_hits == 0 and coverage < 0.5:
                    continue
            anchor_ratio = (anchor_hits / len(anchors)) if anchors else 0.0
            relevance = min(1.0, bm25 / 25.0) * 0.45 + coverage * 0.2 + min(1.0, anchor_ratio * 2.0) * 0.35
            quality = self.quality(record)
            final = 0.55 * relevance + 0.45 * quality
            scored.append((final, relevance, record))
        if sort == "recent":
            scored.sort(key=lambda item: (-(item[2].year or 0), -item[0]))
        else:
            scored.sort(key=lambda item: -item[0])
        return [record for _, _, record in scored[:limit]]

    @staticmethod
    def _anchors(topic: str) -> set:
        """从中文主题里提取三元字锚点，忽略通用词，用于强制相关性。"""

        stopwords = ("的", "在", "中", "与", "和", "及", "对", "应用", "研究", "场景",
                     "方法", "分析", "问题", "综述", "不同", "效果", "局限", "空白")
        text = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", topic or "")
        for word in stopwords:
            text = text.replace(word, " ")
        anchors = set()
        for chunk in text.split():
            if len(chunk) < 3:
                continue
            for index in range(len(chunk) - 2):
                anchors.add(chunk[index:index + 3])
        return anchors

    def stats(self) -> Dict[str, Any]:
        chinese = sum(1 for record in self.records if record.language == "zh")
        venues = sorted({record.venue for record in self.records if record.venue})
        years = [record.year for record in self.records if record.year]
        return {
            "total": len(self.records),
            "chinese": chinese,
            "english": len(self.records) - chinese,
            "venues": len(venues),
            "year_min": min(years) if years else None,
            "year_max": max(years) if years else None,
        }


def load_manifest(path: Path = MANIFEST_PATH) -> Dict[str, Any]:
    path = Path(path)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return {}


def save_manifest(data: Dict[str, Any], path: Path = MANIFEST_PATH) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
