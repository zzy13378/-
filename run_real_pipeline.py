#!/usr/bin/env python3
"""真实跑一遍科研文献整合流程，并把结果落盘供评测平台审查。

用法：
    python run_real_pipeline.py --topic "生态系统文化服务与国家公园空间评估" --max-papers 8

输出：
    reports/p4_real_run.json   结构化结果（论文、九字段、初稿、评审、引用校验）
    reports/p4_real_report.md  完整 Markdown 报告
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

PROJECT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_DIR))
load_dotenv(PROJECT_DIR / ".env")

import p4_literature_assistant as p4  # noqa: E402

REPORT_DIR = PROJECT_DIR / "reports"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--topic", default="生态系统文化服务与国家公园空间评估")
    parser.add_argument("--max-papers", type=int, default=8)
    parser.add_argument("--years", type=int, default=8)
    parser.add_argument("--local-only", action="store_true", help="只用内置文献库，不联网")
    args = parser.parse_args()

    if args.local_only:
        os.environ["P4_LOCAL_CORPUS_ONLY"] = "true"

    print(f"主题：{args.topic}")
    print(f"模型：{os.environ.get('LLM_MODEL', '(未设置)')} | 接口：{os.environ.get('OPENAI_BASE_URL', '(未设置)')}")
    print(f"内置文献库：{json.dumps(p4.local_corpus_stats(), ensure_ascii=False)}")
    print(f"联网检索：{'否' if args.local_only else '本地优先，不足时联网'}\n")

    orchestrator = p4.MultiAgentOrchestrator()
    started = time.time()
    result = orchestrator.run(
        topic=args.topic,
        template="文献综述",
        use_online=True,
        max_papers=args.max_papers,
        years=args.years,
        sort="quality",
        field_decisions={},
        card_decisions={},
    )
    elapsed = time.time() - started

    summary = result.get("execution_summary", {})
    research = result.get("research", {})
    print(f"完成，用时 {elapsed:.0f} 秒")
    print(f"  论文数：{summary.get('paper_count')} | 来源：{summary.get('source_mode')}")
    print(f"  内容分析覆盖率：{summary.get('content_analysis_coverage'):.0%} | 字段覆盖率：{summary.get('structured_field_coverage'):.0%}")
    print(f"  来源覆盖率：{summary.get('citation_coverage'):.0%} | 评审分：{summary.get('quality_score')}/10 | 写作质量分：{summary.get('writing_quality_score')}/10")

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "elapsed_seconds": round(elapsed, 1),
        "topic": result.get("topic"),
        "article": result.get("article", ""),
        "papers": [
            {
                "ref_id": paper.ref_id,
                "title": paper.title,
                "authors": paper.authors,
                "year": paper.year,
                "venue": paper.venue,
                "source": paper.source,
                "doi": paper.doi,
                "abstract": paper.abstract,
            }
            for paper in research.get("papers", [])
        ],
        "structured_papers": research.get("structured_papers", []),
        "comparisons": research.get("comparisons", []),
        "outline": result.get("outline", {}),
        "review": result.get("review", {}),
        "citation_check": result.get("citation_check", {}),
        "writing_quality": result.get("writing_quality", {}),
        "execution_summary": summary,
    }
    json_path = REPORT_DIR / "p4_real_run.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=1, default=str), encoding="utf-8")

    # 另外存一份带时间戳的历史记录，供评测平台按运行批次审查与对比。
    runs_dir = REPORT_DIR / "runs"
    runs_dir.mkdir(parents=True, exist_ok=True)
    stamped = runs_dir / f"p4_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    stamped.write_text(json.dumps(payload, ensure_ascii=False, indent=1, default=str), encoding="utf-8")

    md_path = REPORT_DIR / "p4_real_report.md"
    md_path.write_text(p4.generate_report(result), encoding="utf-8")

    print(f"\n结构化结果：{json_path}")
    print(f"历史记录：{stamped}")
    print(f"完整报告：{md_path}")


if __name__ == "__main__":
    main()
