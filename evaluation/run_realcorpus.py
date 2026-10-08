#!/usr/bin/env python3
"""真实语料评测：真实论文语料 + 真实检索 + 真实模型 + 语义判定。

问题由语料自动生成（金标准就是问题所依据的语料块），检索命中率与引用有效性
都是可客观测量的，答案质量由模型对照标准证据判定。不使用任何人工编写的期望关键词。

用法：
    python -m evaluation.run_realcorpus --limit 15
    python -m evaluation.run_realcorpus --reuse --limit 15   # 复用上次的问题
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))
load_dotenv(PROJECT_DIR / ".env")

from evaluation import p4_adapter, runtime  # noqa: E402
from evaluation.alignment import align_answer, gold_coverage, summarize_alignment  # noqa: E402
from evaluation.judge import Judge  # noqa: E402
from evaluation.question_gen import QuestionGenerator  # noqa: E402
from evaluation.rag_agent import answer_question  # noqa: E402
from evaluation.realcorpus import build_from_reports  # noqa: E402

REPORT_DIR = PROJECT_DIR / "reports" / "evaluation"
CORPUS_FILE = PROJECT_DIR / "_report_new.md"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=15)
    parser.add_argument("--top-k", type=int, default=3)
    parser.add_argument("--reuse", action="store_true", help="复用上次生成的问题")
    parser.add_argument("--corpus", default=str(CORPUS_FILE))
    args = parser.parse_args()

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    question_file = REPORT_DIR / "realcorpus_questions.json"
    result_file = REPORT_DIR / "realcorpus_run.json"

    ready, info = runtime.model_ready()
    if not ready:
        raise SystemExit(f"模型未配置：{info}")
    print(f"模型：{info}\n")

    chunks, index = build_from_reports([Path(args.corpus)])
    print(f"语料：{len({c.doc_id for c in chunks})} 篇论文，{len(chunks)} 个块\n")

    if args.reuse and question_file.exists():
        questions = json.loads(question_file.read_text(encoding="utf-8"))[: args.limit]
        print(f"复用已有问题 {len(questions)} 条\n")
    else:
        generator = QuestionGenerator()
        questions = generator.generate(chunks, count=args.limit)
        question_file.write_text(json.dumps(questions, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"生成问题 {len(questions)} 条（消耗 token {generator.tokens}）\n")

    config = runtime.default_config(top_k=args.top_k)
    judge = Judge()
    chunk_texts = {chunk.id: chunk.text for chunk in index.chunks}
    rows = []

    for number, item in enumerate(questions, 1):
        result = answer_question(index, item["question"], config)
        retrieved, cited = result["retrieved"], result["cited"]
        gold_chunk = item["gold_chunk"]
        gold_doc = item.get("gold_doc") or gold_chunk.split("-")[0]
        hit = gold_chunk in retrieved
        doc_hit = any(cid.split("-")[0] == gold_doc for cid in retrieved)
        rank = retrieved.index(gold_chunk) + 1 if hit else 0
        invalid = [cid for cid in cited if cid not in retrieved]

        verdict, reason = (None, "未判定")
        if not result["error"]:
            verdict, reason = judge.answer_correct(
                item["question"], item.get("gold_text", ""), result["evidence_text"], result["answer"]
            )

        sentences = align_answer(result["answer"], cited, chunk_texts)
        coverage = gold_coverage(item.get("gold_text", ""), result["answer"])

        rows.append({
            "id": item["id"], "question": item["question"], "gold_chunk": gold_chunk,
            "retrieved": retrieved, "cited": cited, "hit": hit, "doc_hit": doc_hit, "rank": rank,
            "invalid_cites": invalid, "judge_pass": verdict, "judge_reason": reason,
            "answer": result["answer"], "error": result["error"],
            "latency_ms": result["latency_ms"], "tokens": result["tokens"],
            "sentences": sentences, "coverage": coverage,
        })
        flag = "命中" if hit else "未命中"
        state = "通过" if verdict else ("失败" if verdict is False else "未判定")
        print(f"  [{number}/{len(questions)}] {item['id']} 检索{flag} 答案{state} | {item['question'][:34]}")

    total = len(rows)
    hit_count = sum(1 for r in rows if r["hit"])
    doc_hit_count = sum(1 for r in rows if r["doc_hit"])
    mrr = sum((1 / r["rank"]) for r in rows if r["rank"]) / total if total else 0
    judged = [r for r in rows if r["judge_pass"] is not None]
    passed = sum(1 for r in judged if r["judge_pass"])
    invalid_total = sum(len(r["invalid_cites"]) for r in rows)
    cited_total = sum(len(r["cited"]) for r in rows)
    alignment = summarize_alignment(rows)

    print("\n" + "=" * 60)
    print("真实语料评测结果")
    print("=" * 60)
    print(f"问题数：{total}")
    print(f"检索命中 Hit@{args.top_k}（块级）：{hit_count / total:.1%}｜论文级：{doc_hit_count / total:.1%}")
    print(f"检索 MRR：{mrr:.3f}")
    print(f"答案正确率：{passed / len(judged):.1%}（{passed}/{len(judged)}）" if judged else "答案正确率：未判定")
    print(f"无效引用：{invalid_total} 条 / 共引用 {cited_total} 条")
    print(f"论断句 {alignment['claims']}｜可追溯 {alignment['traceable_rate']:.1%}｜"
          f"未标注引用 {alignment['missing_citation']}（{alignment['missing_citation_rate']:.1%}）｜"
          f"引用可疑 {alignment['suspicious']}")
    print(f"证据覆盖率：{alignment['gold_coverage']:.1%}")
    print(f"总 token：{sum(r['tokens'] for r in rows) + judge.tokens}")

    result_file.write_text(json.dumps({
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "corpus": args.corpus, "top_k": args.top_k, "total": total,
        "hit_rate": round(hit_count / total, 3) if total else 0,
        "doc_hit_rate": round(doc_hit_count / total, 3) if total else 0,
        "mrr": round(mrr, 3),
        "judge_pass_rate": round(passed / len(judged), 3) if judged else 0,
        "invalid_citations": invalid_total, "total_citations": cited_total,
        "sentence_alignment": alignment, "rows": rows,
    }, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n结果：{result_file}")


if __name__ == "__main__":
    main()
