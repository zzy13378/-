"""连接第一个项目：发现它的运行记录，并把产出转成可审查的结构。

第一个项目（科研文献整合系统）每跑一次会在 reports/runs/ 下留一份带时间戳的
JSON。这个适配层负责找到项目目录、列出历史运行、把结果标准化，并做句子级审查。
界面和命令行都调用同一套逻辑，避免两处实现不一致。
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Dict, List, Optional

from .alignment import (
    DOC_CITATION,
    SentenceVerifier,
    align_answer,
    gold_coverage,
    summarize_alignment,
)

def find_project(start: Optional[Path] = None) -> Optional[Path]:
    """定位项目根目录：优先环境变量，其次本模块的上层目录。

    评测模块已经并入项目内部（evaluation/），所以项目根目录就是它的父目录。
    """

    env = os.environ.get("P4_PROJECT_DIR")
    if env:
        candidate = Path(env).expanduser()
        if (candidate / "p4_literature_assistant.py").exists():
            return candidate.resolve()
    base = Path(start).resolve() if start else Path(__file__).resolve().parent.parent
    for path in (base, base.parent):
        if (path / "p4_literature_assistant.py").exists():
            return path
    return None


def list_runs(project_dir: Path) -> List[Dict]:
    """列出可审查的运行记录，新的排在前面。"""

    runs_dir = Path(project_dir) / "reports" / "runs"
    latest = Path(project_dir) / "reports" / "p4_real_run.json"
    entries: List[Dict] = []
    paths = sorted(runs_dir.glob("p4_run_*.json"), reverse=True) if runs_dir.exists() else []
    if not paths and latest.exists():
        paths = [latest]
    for path in paths:
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            continue
        summary = payload.get("execution_summary") or {}
        entries.append({
            "id": path.stem,
            "path": str(path),
            "generated_at": payload.get("generated_at") or "",
            "topic": payload.get("topic") or "",
            "papers": len(payload.get("papers") or []),
            "chars": len(payload.get("article") or ""),
            "field_coverage": summary.get("structured_field_coverage", 0),
            "citation_coverage": summary.get("citation_coverage", 0),
            "quality_score": summary.get("quality_score", 0),
            "writing_quality_score": summary.get("writing_quality_score", 0),
        })
    return entries


def load_run(path: Path) -> Dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def strip_markdown(article: str) -> str:
    body = (article or "").split("## 参考文献", 1)[0]
    lines = []
    for line in body.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or line.startswith(">") or line.startswith("**"):
            continue
        lines.append(line)
    return "\n".join(lines)


def build_evidence(run: Dict) -> Dict[str, str]:
    """结构化字段 + 摘要构成证据；与写作阶段可见的证据保持一致。"""

    evidence: Dict[str, str] = {}
    for item in run.get("structured_papers") or []:
        paper_id = str(item.get("paper_id") or "")
        fields = item.get("fields") or {}
        usable = [str(value) for key, value in fields.items() if value and "原文未明确" not in str(value)]
        evidence[paper_id] = "\n".join(usable)
    for paper in run.get("papers") or []:
        paper_id = str(paper.get("ref_id") or "")
        parts = [evidence.get(paper_id, "")]
        if paper.get("abstract"):
            parts.append(str(paper["abstract"]))
        parts.append(str(paper.get("title") or ""))
        evidence[paper_id] = "\n".join(part for part in parts if part)
    return evidence


def audit_run(run: Dict, verify: bool = False, verify_budget: int = 30) -> Dict:
    """对一次运行做句子级审查，返回指标与逐句明细。"""

    article = run.get("article") or ""
    papers = run.get("papers") or []
    valid_ids = {str(paper.get("ref_id")) for paper in papers}
    evidence = build_evidence(run)
    body = strip_markdown(article)

    citations = DOC_CITATION.findall(body)
    invalid = sorted({cid for cid in citations if cid not in valid_ids})
    cited_ids = {cid for cid in citations if cid in valid_ids}
    unused = sorted(valid_ids - cited_ids)

    rows = align_answer(body, list(cited_ids), evidence, threshold=0.15, citation_pattern=DOC_CITATION)

    # 证据级别统计：全文 / 摘要 / 仅书目信息
    levels: Dict[str, int] = {}
    for item in run.get("structured_papers") or []:
        level = ((item.get("metadata") or {}).get("text_quality")) or "未知"
        levels[level] = levels.get(level, 0) + 1

    tokens = 0
    if verify:
        verifier = SentenceVerifier()
        budget = verify_budget
        for row in rows:
            if row["verdict"] != "引用可疑" or budget <= 0:
                continue
            cited_text = "\n".join(evidence.get(cid, "") for cid in row["cited"])
            supported, why = verifier.verify(row["sentence"], cited_text, cited_text)
            row["verified"] = supported
            row["verify_reason"] = why
            budget -= 1
        tokens = verifier.tokens

    coverage = gold_coverage("\n".join(evidence.values()), body, threshold=0.25)
    alignment = summarize_alignment([{"sentences": rows, "coverage": coverage}])
    counts = {cid: citations.count(cid) for cid in sorted(valid_ids)}

    return {
        "topic": run.get("topic"),
        "generated_at": run.get("generated_at"),
        "papers": len(papers),
        "article_chars": len(article),
        "citation_total": len(citations),
        "cited_papers": len(cited_ids),
        "invalid_citations": invalid,
        "unused_papers": unused,
        "citation_counts": counts,
        "evidence_levels": levels,
        "alignment": alignment,
        "coverage": coverage,
        "sentences": rows,
        "verify_tokens": tokens,
        "execution_summary": run.get("execution_summary") or {},
        "review": run.get("review") or {},
    }


def problem_sentences(audit: Dict) -> List[Dict]:
    """挑出需要人工关注的问题句。"""

    problems = []
    for row in audit.get("sentences") or []:
        if row["verdict"] in {"未标注引用", "引用可疑"}:
            problems.append(row)
    return problems
