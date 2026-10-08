#!/usr/bin/env python3
"""核验中文核心期刊在 OpenAlex 里的真实身份与质量指标。

中文期刊在 OpenAlex 里常以英文名或拼音名收录，直接按中文名搜未必命中。
这个脚本对同一本期刊分别用中文名和英文名检索，把候选源的名称、收录量、
被引量、h 指数、两年平均被引等指标打出来，用于判断收录的是不是真期刊、
是不是高质量期刊。

用法：
    python probe_journals.py --check-current   # 核验当前已收录的期刊
    python probe_journals.py --probe           # 用中英文名检索候选期刊
"""

from __future__ import annotations

import argparse
import json
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional

PROJECT_DIR = Path(__file__).resolve().parent
CHINESE_JOURNAL_FILE = PROJECT_DIR / "chinese_journals.json"
REPORT_FILE = PROJECT_DIR / "data" / "journal_audit.json"
USER_AGENT = "P4-Literature-Assistant/1.1 (mailto:p4-literature-assistant@example.com)"

# 候选中文高质量期刊：中文名、英文名、所属核心体系
CANDIDATES = [
    # 生态、地理、环境、资源
    ("生态学报", "Acta Ecologica Sinica", "CSCD/北大核心"),
    ("应用生态学报", "Chinese Journal of Applied Ecology", "CSCD/北大核心"),
    ("生态学杂志", "Chinese Journal of Ecology", "CSCD/北大核心"),
    ("地理学报", "Acta Geographica Sinica", "CSCD/北大核心"),
    ("地理研究", "Geographical Research", "CSCD/北大核心"),
    ("自然资源学报", "Journal of Natural Resources", "CSCD/北大核心"),
    ("资源科学", "Resources Science", "CSCD/北大核心"),
    ("环境科学", "Environmental Science", "CSCD/北大核心"),
    ("环境科学学报", "Acta Scientiae Circumstantiae", "CSCD/北大核心"),
    ("中国环境科学", "China Environmental Science", "CSCD/北大核心"),
    ("生态经济", "Ecological Economy", "北大核心"),
    ("中国人口·资源与环境", "China Population Resources and Environment", "CSSCI/北大核心"),
    ("旅游学刊", "Tourism Tribune", "CSSCI/北大核心"),
    ("人文地理", "Human Geography", "CSSCI/北大核心"),
    ("城市规划", "City Planning Review", "CSCD/北大核心"),
    ("遥感学报", "Journal of Remote Sensing", "CSCD/北大核心"),
    ("地球信息科学学报", "Journal of Geo-information Science", "CSCD/北大核心"),
    # 情报、图书馆、管理、经济
    ("情报学报", "Journal of the China Society for Scientific and Technical Information", "CSSCI/北大核心"),
    ("图书情报工作", "Library and Information Service", "CSSCI/北大核心"),
    ("中国图书馆学报", "Journal of Library Science in China", "CSSCI/北大核心"),
    ("情报理论与实践", "Information Studies: Theory & Application", "CSSCI/北大核心"),
    ("数据分析与知识发现", "Data Analysis and Knowledge Discovery", "CSSCI/北大核心"),
    ("管理世界", "Management World", "CSSCI/北大核心"),
    ("中国软科学", "China Soft Science", "CSSCI/北大核心"),
    ("管理科学学报", "Journal of Management Sciences in China", "CSSCI/北大核心"),
    ("南开管理评论", "Nankai Business Review", "CSSCI/北大核心"),
    ("经济研究", "Economic Research Journal", "CSSCI/北大核心"),
    ("中国工业经济", "China Industrial Economics", "CSSCI/北大核心"),
    ("科研管理", "Science Research Management", "CSSCI/北大核心"),
    ("科学学研究", "Studies in Science of Science", "CSSCI/北大核心"),
    ("系统工程理论与实践", "Systems Engineering - Theory & Practice", "CSCD/北大核心"),
    ("中国行政管理", "Chinese Public Administration", "CSSCI/北大核心"),
    ("公共管理学报", "Journal of Public Management", "CSSCI/北大核心"),
    # 计算机、信息、教育
    ("计算机学报", "Chinese Journal of Computers", "CSCD/北大核心"),
    ("软件学报", "Journal of Software", "CSCD/北大核心"),
    ("计算机研究与发展", "Journal of Computer Research and Development", "CSCD/北大核心"),
    ("自动化学报", "Acta Automatica Sinica", "CSCD/北大核心"),
    ("中国科学：信息科学", "Scientia Sinica Informationis", "CSCD/北大核心"),
    ("心理学报", "Acta Psychologica Sinica", "CSSCI/北大核心"),
    ("中国远程教育", "Distance Education in China", "CSSCI/北大核心"),
    ("开放教育研究", "Open Education Research", "CSSCI/北大核心"),
    ("电化教育研究", "e-Education Research", "CSSCI/北大核心"),
]


def fetch_json(url: str, sleep: float = 1.2, attempts: int = 3) -> Optional[Dict[str, Any]]:
    for attempt in range(attempts):
        request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        try:
            with urllib.request.urlopen(request, timeout=25) as response:
                payload = json.loads(response.read().decode("utf-8"))
                time.sleep(sleep)
                return payload
        except urllib.error.HTTPError as exc:
            if exc.code == 429:
                time.sleep(10 * (attempt + 1))
                continue
            return None
        except Exception:  # noqa: BLE001
            time.sleep(3)
    return None


def metrics_of(item: Dict[str, Any]) -> Dict[str, Any]:
    stats = item.get("summary_stats") or {}
    return {
        "id": str(item.get("id") or "").rsplit("/", 1)[-1],
        "display_name": item.get("display_name"),
        "alternate_titles": (item.get("alternate_titles") or [])[:3],
        "works_count": item.get("works_count"),
        "cited_by_count": item.get("cited_by_count"),
        "h_index": stats.get("h_index"),
        "two_year_mean_citedness": round(stats.get("2yr_mean_citedness") or 0, 2),
        "country": item.get("country_code"),
        "type": item.get("type"),
    }


def source_by_id(source_id: str) -> Optional[Dict[str, Any]]:
    payload = fetch_json(f"https://api.openalex.org/sources/{source_id}")
    return metrics_of(payload) if payload else None


def search_sources(query: str, per_page: int = 3) -> List[Dict[str, Any]]:
    url = "https://api.openalex.org/sources?" + urllib.parse.urlencode(
        {"search": query, "per-page": per_page, "mailto": "p4-literature-assistant@example.com"}
    )
    payload = fetch_json(url)
    if not payload:
        return []
    return [metrics_of(item) for item in payload.get("results") or []]


def pick_best(results: List[Dict[str, Any]], chinese: str, english: str) -> Optional[Dict[str, Any]]:
    """从候选里挑最匹配的一本：名称对得上且指标最高。"""

    target_zh, target_en = chinese.lower(), english.lower()
    best, best_score = None, -1.0
    for item in results:
        names = [str(item.get("display_name") or "").lower()]
        names += [str(value).lower() for value in (item.get("alternate_titles") or [])]
        name_hit = any(target_zh in name or name in target_zh or target_en in name or name in target_en
                       for name in names if name)
        if not name_hit:
            continue
        score = (item.get("h_index") or 0) + (item.get("works_count") or 0) / 1000.0
        if score > best_score:
            best, best_score = item, score
    return best


def check_current(sleep: float) -> List[Dict[str, Any]]:
    if not CHINESE_JOURNAL_FILE.exists():
        print("没有找到 chinese_journals.json")
        return []
    cached = json.loads(CHINESE_JOURNAL_FILE.read_text(encoding="utf-8"))
    rows = []
    for name, source_id in cached.items():
        info = source_by_id(source_id)
        if info:
            info["configured_name"] = name
            rows.append(info)
            print(f"{name[:45]:45s} -> {info['display_name'][:38]:38s} "
                  f"h={info['h_index']} 收录={info['works_count']} 被引={info['cited_by_count']} "
                  f"国家={info['country']}")
        else:
            print(f"{name}: 查询失败")
        time.sleep(sleep)
    return rows


def probe(sleep: float) -> List[Dict[str, Any]]:
    rows = []
    for chinese, english, tier in CANDIDATES:
        zh_results = search_sources(chinese)
        en_results = search_sources(english)
        best = pick_best(zh_results + en_results, chinese, english)
        if best:
            best.update({"chinese_name": chinese, "english_name": english, "tier": tier,
                         "zh_hit": bool(pick_best(zh_results, chinese, english)),
                         "en_hit": bool(pick_best(en_results, chinese, english))})
            rows.append(best)
            print(f"{chinese[:16]:16s} | {best['display_name'][:34]:34s} | "
                  f"h={best['h_index']:>4} 收录={best['works_count']:>7} "
                  f"被引={best['cited_by_count']:>8} | 中文命中={best['zh_hit']} 英文命中={best['en_hit']}")
        else:
            print(f"{chinese}: 未找到匹配期刊（中文/英文检索均无命中）")
    return rows


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check-current", action="store_true")
    parser.add_argument("--probe", action="store_true")
    parser.add_argument("--sleep", type=float, default=1.2)
    args = parser.parse_args()

    REPORT_FILE.parent.mkdir(parents=True, exist_ok=True)
    if args.check_current:
        rows = check_current(args.sleep)
        REPORT_FILE.write_text(json.dumps({"current": rows}, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"\n已写入 {REPORT_FILE}")
    elif args.probe:
        rows = probe(args.sleep)
        REPORT_FILE.write_text(json.dumps({"probe": rows}, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"\n找到 {len(rows)} 本可定位的中文期刊，已写入 {REPORT_FILE}")
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
