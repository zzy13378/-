#!/usr/bin/env python3
"""构建内置高质量文献库。

一次性把常用高质量中英文期刊的近年论文收进本地库，之后检索优先走本地，
避免每次重复调用 OpenAlex。采集过程有节流与重试，支持中断续跑。

用法：
    python build_local_corpus.py                 # 采集全部（中文核心 + 英文高质量）
    python build_local_corpus.py --only chinese
    python build_local_corpus.py --only english
    python build_local_corpus.py --per-venue 100 --from-year 2018
    python build_local_corpus.py --status        # 只看当前语料库状态
"""

from __future__ import annotations

import argparse
import json
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from local_corpus import (
    CORPUS_PATH,
    MANIFEST_PATH,
    CorpusRecord,
    LocalCorpus,
    load_manifest,
    merge_corpus,
    save_manifest,
)

PROJECT_DIR = Path(__file__).resolve().parent
CHINESE_JOURNAL_FILE = PROJECT_DIR / "chinese_journals.json"
USER_AGENT = "P4-Literature-Assistant/1.1 (mailto:p4-literature-assistant@example.com)"

ENGLISH_VENUES = [
    # 生态、地理、环境、资源
    "Landscape and Urban Planning",
    "Ecosystem Services",
    "Journal of Environmental Management",
    "Ecological Indicators",
    "Land Use Policy",
    "Landscape Ecology",
    "Science of the Total Environment",
    "Ecological Economics",
    "Ambio",
    "Journal of Environmental Planning and Management",
    "Environmental Science & Policy",
    "Sustainability Science",
    "People and Nature",
    "Conservation Biology",
    "Biological Conservation",
    "Global Environmental Change",
    "Nature Sustainability",
    "Tourism Management",
    "Annals of Tourism Research",
    "Journal of Sustainable Tourism",
    # 人工智能、信息系统、管理
    "Nature Machine Intelligence",
    "Journal of Machine Learning Research",
    "Artificial Intelligence",
    "IEEE Transactions on Pattern Analysis and Machine Intelligence",
    "Information Processing & Management",
    "Journal of the Association for Information Science and Technology",
    "Expert Systems with Applications",
    "Knowledge-Based Systems",
    "Decision Support Systems",
    "MIS Quarterly",
    "Information Systems Research",
    "Journal of Management Information Systems",
    "Scientometrics",
    "Research Policy",
    "Technological Forecasting and Social Change",
]


def fetch_json(url: str, timeout: int = 25, attempts: int = 3) -> Optional[Dict[str, Any]]:
    for attempt in range(attempts):
        request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            if exc.code == 429:
                wait = 10 * (attempt + 1)
                print(f"    触发限流，等待 {wait} 秒后重试")
                time.sleep(wait)
                continue
            print(f"    请求失败 HTTP {exc.code}")
            return None
        except Exception as exc:  # noqa: BLE001
            print(f"    请求异常 {type(exc).__name__}")
            time.sleep(3)
    return None


def abstract_from_index(index: Any) -> str:
    if not isinstance(index, dict):
        return ""
    words: List[Tuple[int, str]] = []
    for word, positions in index.items():
        if isinstance(positions, list):
            for position in positions:
                if isinstance(position, int):
                    words.append((position, str(word)))
    if not words:
        return ""
    words.sort()
    return " ".join(word for _, word in words)[:2000]


def authors_of(item: Dict[str, Any]) -> List[str]:
    names: List[str] = []
    for authorship in item.get("authorships") or []:
        author = (authorship or {}).get("author") or {}
        name = author.get("display_name")
        if name:
            names.append(str(name))
    return names[:20]


def to_record(item: Dict[str, Any], venue: str, language: str) -> Optional[CorpusRecord]:
    title = str(item.get("title") or "").strip()
    if not title:
        return None
    doi = str(item.get("doi") or "").replace("https://doi.org/", "")
    location = item.get("primary_location") or {}
    source = location.get("source") or {}
    oa = item.get("open_access") or {}
    best_oa = item.get("best_oa_location") or {}
    oa_url = str(best_oa.get("pdf_url") or best_oa.get("landing_page_url") or oa.get("oa_url") or "")
    return CorpusRecord(
        paper_id=str(item.get("id") or "").rsplit("/", 1)[-1],
        title=title[:500],
        authors=authors_of(item),
        year=item.get("publication_year"),
        abstract=abstract_from_index(item.get("abstract_inverted_index")),
        doi=doi,
        url=str(location.get("landing_page_url") or (f"https://doi.org/{doi}" if doi else "")),
        venue=str(source.get("display_name") or venue),
        language=language,
        cited_by_count=int(item.get("cited_by_count") or 0),
        source="内置文献库",
        fetched_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        oa_url=oa_url,
        is_oa=bool(oa.get("is_oa")),
    )


def resolve_source_id(name: str) -> Optional[str]:
    url = "https://api.openalex.org/sources?" + urllib.parse.urlencode(
        {"search": name, "per-page": 3, "mailto": "p4-literature-assistant@example.com"}
    )
    payload = fetch_json(url)
    if not payload:
        return None
    target = name.lower()
    best = None
    for item in payload.get("results") or []:
        display = str(item.get("display_name") or "").lower()
        if display == target:
            return str(item.get("id") or "").rsplit("/", 1)[-1]
        if best is None and (target in display or display in target):
            best = str(item.get("id") or "").rsplit("/", 1)[-1]
    return best


def harvest_venue(source_id: str, venue: str, language: str, per_venue: int,
                  from_year: int, sleep: float) -> List[CorpusRecord]:
    filters = [
        f"primary_location.source.id:{source_id}",
        "type:article",
        f"from_publication_date:{from_year}-01-01",
    ]
    params = {
        "filter": ",".join(filters),
        "sort": "cited_by_count:desc",
        "per-page": min(200, per_venue),
        "select": ("id,doi,title,authorships,publication_year,primary_location,"
                   "abstract_inverted_index,cited_by_count,open_access,best_oa_location"),
        "mailto": "p4-literature-assistant@example.com",
    }
    url = "https://api.openalex.org/works?" + urllib.parse.urlencode(params)
    payload = fetch_json(url)
    time.sleep(sleep)
    if not payload:
        return []
    records = []
    for item in payload.get("results") or []:
        record = to_record(item, venue, language)
        if record:
            records.append(record)
    return records


def load_chinese_journals() -> Dict[str, str]:
    if not CHINESE_JOURNAL_FILE.exists():
        return {}
    try:
        return json.loads(CHINESE_JOURNAL_FILE.read_text(encoding="utf-8"))
    except ValueError:
        return {}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--only", choices=["chinese", "english", "all"], default="all")
    parser.add_argument("--per-venue", type=int, default=100)
    parser.add_argument("--from-year", type=int, default=2018)
    parser.add_argument("--sleep", type=float, default=1.2)
    parser.add_argument("--status", action="store_true")
    parser.add_argument("--refresh", action="store_true", help="忽略清单，重新采集")
    args = parser.parse_args()

    if args.status:
        corpus = LocalCorpus()
        print(json.dumps(corpus.stats(), ensure_ascii=False, indent=1))
        print("清单:", json.dumps(load_manifest(), ensure_ascii=False)[:400])
        return

    manifest = {} if args.refresh else load_manifest()
    done = set(manifest.get("done_venues") or [])
    sources: List[Tuple[str, str, str]] = []  # (source_id, display_name, language)

    if args.only in {"chinese", "all"}:
        for english_name, source_id in load_chinese_journals().items():
            sources.append((source_id, english_name, "zh"))

    if args.only in {"english", "all"}:
        print(f"解析 {len(ENGLISH_VENUES)} 个英文期刊的 OpenAlex 源 ID")
        for name in ENGLISH_VENUES:
            source_id = resolve_source_id(name)
            if source_id:
                sources.append((source_id, name, "en"))
            else:
                print(f"  跳过（未找到）：{name}")
            time.sleep(args.sleep)

    print(f"待采集期刊：{len(sources)} 个\n")
    total_new = 0
    for index, (source_id, name, language) in enumerate(sources, 1):
        if source_id in done:
            print(f"[{index}/{len(sources)}] 跳过（已采集）{name}")
            continue
        print(f"[{index}/{len(sources)}] 采集 {name} ...", flush=True)
        records = harvest_venue(source_id, name, language, args.per_venue, args.from_year, args.sleep)
        if records:
            total, added = merge_corpus(records, overwrite=True)
            total_new += added
            print(f"    获得 {len(records)} 篇，新增 {added} 篇，语料库共 {total} 篇")
            done.add(source_id)
            save_manifest({
                "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "done_venues": sorted(done),
                "total": total,
            })
        else:
            print("    没有获取到数据")

    corpus = LocalCorpus()
    print(f"\n采集结束，新增 {total_new} 篇。当前语料库：{json.dumps(corpus.stats(), ensure_ascii=False)}")


if __name__ == "__main__":
    main()
