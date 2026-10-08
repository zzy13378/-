#!/usr/bin/env python3
"""按期刊核验结果清理内置文献库。

OpenAlex 对部分中文期刊只建了“空壳条目”：名称对得上，但收录量与被引量极低，
摘要基本为空。这些条目会污染检索结果，用它们的 source id 去采集也只能拿到
少量残缺记录。本脚本依据 journal_audit.json 的核验指标，把这类期刊从语料库和
期刊清单里清理出去，原始文件会先备份。

判定规则：h 指数 >= 10，或被引总量 >= 1000，视为有效期刊。

用法：
    python clean_local_corpus.py --dry-run   # 只打印判定结果
    python clean_local_corpus.py             # 执行清理
"""

from __future__ import annotations

import argparse
import json
import shutil
from datetime import datetime
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent
AUDIT_FILE = PROJECT_DIR / "data" / "journal_audit.json"
CORPUS_FILE = PROJECT_DIR / "data" / "local_corpus.jsonl"
JOURNAL_FILE = PROJECT_DIR / "chinese_journals.json"
MIN_H_INDEX = 10
MIN_CITED = 1000


def load_audit() -> list:
    if not AUDIT_FILE.exists():
        raise SystemExit("缺少 data/journal_audit.json，请先运行 probe_journals.py --check-current")
    return json.loads(AUDIT_FILE.read_text(encoding="utf-8")).get("current", [])


def classify(rows: list):
    keep, drop = [], []
    for row in rows:
        h = row.get("h_index") or 0
        cited = row.get("cited_by_count") or 0
        (keep if (h >= MIN_H_INDEX or cited >= MIN_CITED) else drop).append(row)
    return keep, drop


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    rows = load_audit()
    keep, drop = classify(rows)

    print(f"核验期刊 {len(rows)} 本：有效 {len(keep)} 本，空壳 {len(drop)} 本\n")
    print("【有效期刊】")
    for row in sorted(keep, key=lambda r: -(r.get("h_index") or 0)):
        print(f"  {str(row.get('configured_name'))[:34]:34s} -> {str(row.get('display_name'))[:30]:30s} "
              f"h={row.get('h_index')} 收录={row.get('works_count')} 被引={row.get('cited_by_count')}")
    print("\n【空壳期刊（清理）】")
    for row in sorted(drop, key=lambda r: (r.get("h_index") or 0)):
        print(f"  {str(row.get('configured_name'))[:34]:34s} -> {str(row.get('display_name'))[:30]:30s} "
              f"h={row.get('h_index')} 收录={row.get('works_count')} 被引={row.get('cited_by_count')}")

    drop_venues = {str(row.get("display_name") or "") for row in drop}
    drop_ids = {str(row.get("id") or "") for row in drop}

    records = [json.loads(line) for line in CORPUS_FILE.read_text(encoding="utf-8").splitlines() if line.strip()]
    removed = [r for r in records if r.get("language") == "zh" and (r.get("venue") in drop_venues or r.get("venue") in drop_venues)]
    kept = [r for r in records if r not in removed]

    print(f"\n语料库：共 {len(records)} 篇，将移除中文空壳期刊记录 {len(removed)} 篇，保留 {len(kept)} 篇")

    if args.dry_run:
        print("\n这是 dry-run，没有写入任何文件。")
        return

    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    shutil.copy2(CORPUS_FILE, CORPUS_FILE.with_name(f"local_corpus_backup_{stamp}.jsonl"))
    shutil.copy2(JOURNAL_FILE, JOURNAL_FILE.with_name(f"chinese_journals_backup_{stamp}.json"))
    with CORPUS_FILE.open("w", encoding="utf-8") as handle:
        for record in kept:
            handle.write(json.dumps(record, ensure_ascii=False) + "\n")

    cached = json.loads(JOURNAL_FILE.read_text(encoding="utf-8"))
    keep_names = {str(row.get("configured_name")) for row in keep}
    new_cached = {name: source_id for name, source_id in cached.items() if name in keep_names}
    JOURNAL_FILE.write_text(json.dumps(new_cached, ensure_ascii=False, indent=2), encoding="utf-8")

    audit_path = PROJECT_DIR / "data" / "journal_audit_decision.json"
    audit_path.write_text(json.dumps({
        "decided_at": stamp,
        "rule": f"h_index >= {MIN_H_INDEX} or cited_by_count >= {MIN_CITED}",
        "keep": keep,
        "drop": drop,
        "corpus_before": len(records),
        "corpus_after": len(kept),
    }, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n已清理。期刊清单剩余 {len(new_cached)} 本，语料库剩余 {len(kept)} 篇。")
    print(f"备份：local_corpus_backup_{stamp}.jsonl / chinese_journals_backup_{stamp}.json")
    print(f"判定记录：{audit_path}")


if __name__ == "__main__":
    main()
