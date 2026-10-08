#!/usr/bin/env python3
"""审查本项目的运行产出（命令行入口）。

界面上的“产出审查”页签与这里调用同一套逻辑（evaluation/p4_adapter.py）。

用法：
    python -m evaluation.audit_project --list
    python -m evaluation.audit_project                    # 审查最新一次运行
    python -m evaluation.audit_project --run p4_run_...   # 审查指定运行
    python -m evaluation.audit_project --verify           # 对可疑句做模型核实
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from dotenv import load_dotenv

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))
load_dotenv(PROJECT_DIR / ".env")

from evaluation import p4_adapter  # noqa: E402


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run", default="", help="运行记录 id，默认最新一次")
    parser.add_argument("--list", action="store_true")
    parser.add_argument("--verify", action="store_true", help="对可疑句做模型核实")
    parser.add_argument("--out", default="")
    args = parser.parse_args()

    project = p4_adapter.find_project()
    if project is None:
        raise SystemExit("找不到项目根目录，请设置 P4_PROJECT_DIR")

    runs = p4_adapter.list_runs(project)
    if not runs:
        raise SystemExit("还没有运行记录，请先执行：python run_real_pipeline.py --topic \"你的主题\"")

    if args.list:
        print(f"项目：{project}")
        for run in runs:
            print(f"  {run['id']} | {run['generated_at']} | {run['papers']} 篇 | "
                  f"字段覆盖率 {run['field_coverage']:.0%} | {run['topic'][:34]}")
        return

    target = next((run for run in runs if run["id"] == args.run), runs[0])
    print(f"运行：{target['id']}（{target['generated_at']}）\n")

    audit = p4_adapter.audit_run(p4_adapter.load_run(Path(target["path"])), verify=args.verify)
    alignment = audit["alignment"]

    print("=" * 62)
    print("本项目产出审查")
    print("=" * 62)
    print(f"主题：{audit['topic']}")
    print(f"论文数：{audit['papers']} | 证据级别：{audit['evidence_levels']}")
    print(f"引用编号：{audit['citation_total']} 处，涉及 {audit['cited_papers']} 篇")
    print(f"无效引用：{audit['invalid_citations'] or '无'} | 未被引用：{audit['unused_papers'] or '无'}")
    print(f"篇均引用：{audit['citation_counts']}")
    print()
    print(f"论断句 {alignment['claims']} 句 | 可追溯 {alignment['traceable_rate']:.0%} | "
          f"未标注引用 {alignment['missing_citation']}（{alignment['missing_citation_rate']:.0%}） | "
          f"引用可疑 {alignment['suspicious']}")
    print(f"作者分析句 {alignment.get('analysis', 0)} | 说明句 {alignment['meta']} | 格式片段 {alignment['fragment']}")
    if alignment["verified"]:
        print(f"可疑句核实为真支持：{alignment['verified_supported']}/{alignment['verified']}")
    print(f"证据覆盖率：{audit['coverage']['coverage']:.1%}")

    problems = p4_adapter.problem_sentences(audit)
    if problems:
        print(f"\n需要修改的句子（共 {len(problems)} 条，展示前 10 条）：")
        for row in problems[:10]:
            tag = row["verdict"]
            if row.get("verified") is not None:
                tag += f"（核实：{'支持' if row['verified'] else '不支持'}）"
            print(f"  [{tag}] {row['sentence'][:100]}")

    out_path = Path(args.out) if args.out else PROJECT_DIR / "reports" / "evaluation" / "p4_output_audit.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(audit, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n审查结果：{out_path}")


if __name__ == "__main__":
    main()
