"""评测与审查页面（真实模式）。

只做三件事，全部基于真实数据：
1. 看内置文献库状态；
2. 用真实语料自动出题，跑真实检索与真实模型，给出命中率、答案正确率与句子级指标；
3. 审查本项目真实运行产出的初稿，定位引用与归因问题。

由主界面 app.py 调用 render() 渲染，不单独跑一个应用。
"""

from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

PROJECT_DIR = Path(__file__).resolve().parent.parent
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from evaluation import p4_adapter, runtime  # noqa: E402
from evaluation.alignment import align_answer, gold_coverage, summarize_alignment  # noqa: E402
from evaluation.judge import Judge  # noqa: E402
from evaluation.question_gen import QuestionGenerator  # noqa: E402
from evaluation.rag_agent import answer_question  # noqa: E402
from evaluation.realcorpus import build_from_reports  # noqa: E402

EVAL_DIR = PROJECT_DIR / "reports" / "evaluation"
CORPUS_SOURCE = PROJECT_DIR / "_report_new.md"


@st.cache_resource(show_spinner=False)
def load_local_corpus():
    """加载内置文献库（构建检索索引较慢，缓存复用）。"""

    import local_corpus

    corpus = local_corpus.LocalCorpus()
    return corpus, corpus.stats()


@st.cache_resource(show_spinner=False)
def load_real_corpus(path: str, size: int = 720):
    """加载用于评测的真实论文语料（按报告文件切块建索引）。"""

    return build_from_reports([Path(path)], size=size)


def render_corpus_tab() -> None:
    st.subheader("内置文献库")
    try:
        corpus, stats = load_local_corpus()
    except Exception as exc:  # noqa: BLE001
        st.error(f"加载失败：{exc}")
        return
    columns = st.columns(5)
    columns[0].metric("论文总数", f"{stats.get('total', 0)} 篇")
    columns[1].metric("中文", f"{stats.get('chinese', 0)} 篇")
    columns[2].metric("英文", f"{stats.get('english', 0)} 篇")
    columns[3].metric("期刊数", stats.get("venues", 0))
    full = sum(1 for record in corpus.records if record.fulltext_chars)
    columns[4].metric("已有全文", f"{full} 篇")
    st.caption(
        f"年份范围 {stats.get('year_min')} 至 {stats.get('year_max')}。"
        "语料库直接服务检索环节：命中足够时自动跳过在线检索，减少 OpenAlex 调用。"
    )
    st.markdown("**维护命令**")
    st.code(
        "python build_local_corpus.py --per-venue 100 --from-year 2018   # 采集\n"
        "python fetch_fulltext.py --status                                # 看全文覆盖率\n"
        "python fetch_fulltext.py --source unpaywall --limit 50           # 抓开放获取全文\n"
        "python ingest_pdfs.py                                            # 图书馆下载的 PDF 入库",
        language="bash",
    )


def render_retrieval_tab() -> None:
    st.subheader("真实语料检索评测")
    ready, info = runtime.model_ready()
    if not ready:
        st.error(f"模型未配置：{info}。请在 .env 里配置 OPENAI_BASE_URL / OPENAI_API_KEY / LLM_MODEL。")
        return
    st.caption(f"模型：{info}｜问题由语料自动生成，金标准就是问题所依据的语料块，判定不使用人工期望关键词。")

    controls = st.columns([1, 1, 1, 2])
    with controls[0]:
        limit = st.number_input("题目数量", min_value=1, max_value=15, value=5, step=1)
    with controls[1]:
        top_k = st.number_input("Top-K", min_value=1, max_value=8, value=3, step=1)
    with controls[2]:
        reuse = st.checkbox("复用已有问题", value=False)
    with controls[3]:
        corpus_path = st.text_input("评测语料", value=str(CORPUS_SOURCE))

    if st.button("开始评测", type="primary", key="eval_run_real"):
        question_file = EVAL_DIR / "realcorpus_questions.json"
        EVAL_DIR.mkdir(parents=True, exist_ok=True)
        progress = st.progress(0.0, text="准备中")
        try:
            chunks, index = load_real_corpus(corpus_path)
        except Exception as exc:  # noqa: BLE001
            st.error(f"语料加载失败：{exc}")
            return

        if reuse and question_file.exists():
            questions = json.loads(question_file.read_text(encoding="utf-8"))[: int(limit)]
        else:
            generator = QuestionGenerator()
            questions = generator.generate(chunks, count=int(limit))
            question_file.write_text(json.dumps(questions, ensure_ascii=False, indent=1), encoding="utf-8")

        config = runtime.default_config(top_k=int(top_k))
        judge = Judge()
        chunk_texts = {chunk.id: chunk.text for chunk in index.chunks}
        rows = []
        for number, item in enumerate(questions, 1):
            progress.progress(number / len(questions), text=f"第 {number}/{len(questions)} 题：{item['question'][:30]}")
            result = answer_question(index, item["question"], config)
            retrieved, cited = result["retrieved"], result["cited"]
            gold = item["gold_chunk"]
            verdict = None
            reason = "未判定"
            if not result["error"]:
                verdict, reason = judge.answer_correct(
                    item["question"], item.get("gold_text", ""), result["evidence_text"], result["answer"]
                )
            rows.append({
                "id": item["id"], "question": item["question"], "gold_chunk": gold,
                "retrieved": retrieved, "cited": cited,
                "hit": gold in retrieved,
                "doc_hit": any(cid.split("-")[0] == gold.split("-")[0] for cid in retrieved),
                "invalid_cites": [cid for cid in cited if cid not in retrieved],
                "judge_pass": verdict, "judge_reason": reason, "answer": result["answer"],
                "error": result["error"], "latency_ms": result["latency_ms"], "tokens": result["tokens"],
                "sentences": align_answer(result["answer"], cited, chunk_texts),
                "coverage": gold_coverage(item.get("gold_text", ""), result["answer"]),
            })
        progress.empty()

        total = len(rows)
        judged = [r for r in rows if r["judge_pass"] is not None]
        alignment = summarize_alignment(rows)
        payload = {
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total": total,
            "hit_rate": round(sum(1 for r in rows if r["hit"]) / total, 3) if total else 0,
            "doc_hit_rate": round(sum(1 for r in rows if r["doc_hit"]) / total, 3) if total else 0,
            "judge_pass_rate": round(sum(1 for r in judged if r["judge_pass"]) / len(judged), 3) if judged else 0,
            "invalid_citations": sum(len(r["invalid_cites"]) for r in rows),
            "sentence_alignment": alignment,
            "tokens": sum(r["tokens"] for r in rows) + judge.tokens,
            "rows": rows,
        }
        st.session_state["real_eval"] = payload
        (EVAL_DIR / "realcorpus_run.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8"
        )

    payload = st.session_state.get("real_eval")
    if not payload:
        st.info("点“开始评测”会调用真实模型，按题量产生少量费用。先跑 3 到 5 题即可验证。")
        return

    alignment = payload["sentence_alignment"]
    columns = st.columns(6)
    columns[0].metric("块级命中率", f"{payload['hit_rate']:.0%}")
    columns[1].metric("论文级命中率", f"{payload['doc_hit_rate']:.0%}")
    columns[2].metric("答案正确率", f"{payload['judge_pass_rate']:.0%}")
    columns[3].metric("无效引用", payload["invalid_citations"])
    columns[4].metric("未标注引用", alignment["missing_citation"])
    columns[5].metric("总 token", payload["tokens"])

    st.markdown("**逐题结果**")
    st.dataframe(pd.DataFrame([{
        "编号": row["id"], "问题": row["question"][:60],
        "块级命中": "是" if row["hit"] else "否",
        "论文级命中": "是" if row["doc_hit"] else "否",
        "答案": "通过" if row["judge_pass"] else ("失败" if row["judge_pass"] is False else "未判定"),
        "耗时(ms)": row["latency_ms"], "token": row["tokens"],
    } for row in payload["rows"]]), width="stretch", hide_index=True)
    st.session_state.setdefault("real_eval", payload)


def render_audit_tab() -> None:
    st.subheader("产出审查")
    st.caption("读取本项目真实运行记录，检查引用有效性、漏引文献与无来源论断。")
    project = p4_adapter.find_project()
    runs = p4_adapter.list_runs(project) if project else []
    if not runs:
        st.info("还没有运行记录。先跑一次：python run_real_pipeline.py --topic \"你的主题\"")
        return

    labels = {
        run["id"]: f"{run['id']}｜{run['generated_at']}｜{run['papers']} 篇｜{run['topic'][:26]}"
        for run in runs
    }
    chosen = st.selectbox("选择运行记录", list(labels), format_func=lambda key: labels[key], key="audit_run_choice")
    target = next(run for run in runs if run["id"] == chosen)

    columns = st.columns(5)
    columns[0].metric("论文数", f"{target['papers']} 篇")
    columns[1].metric("字段覆盖率", f"{target['field_coverage']:.0%}")
    columns[2].metric("来源覆盖率", f"{target['citation_coverage']:.0%}")
    columns[3].metric("评审分", f"{target['quality_score']}/10")
    columns[4].metric("写作质量分", f"{target['writing_quality_score']}/10")

    verify = st.checkbox("对可疑句调用模型核实", value=False, key="audit_verify_flag")
    if st.button("执行审查", type="primary", key="audit_do"):
        with st.spinner("审查中…"):
            audit = p4_adapter.audit_run(p4_adapter.load_run(Path(target["path"])), verify=verify)
        st.session_state["audit_result"] = audit
        st.session_state["audit_target"] = target["id"]

    audit = st.session_state.get("audit_result")
    if not audit or st.session_state.get("audit_target") != target["id"]:
        return

    alignment = audit["alignment"]
    columns = st.columns(6)
    columns[0].metric("引用编号", f"{audit['citation_total']} 处")
    columns[1].metric("涉及文献", f"{audit['cited_papers']}/{audit['papers']}")
    columns[2].metric("无效编号", len(audit["invalid_citations"]))
    columns[3].metric("未被引用", len(audit["unused_papers"]))
    columns[4].metric("论断句", alignment["claims"])
    columns[5].metric("可追溯率", f"{alignment['traceable_rate']:.0%}")
    st.caption(
        f"未标注引用 {alignment['missing_citation']} 句（{alignment['missing_citation_rate']:.0%}）｜"
        f"引用可疑 {alignment['suspicious']} 句｜作者分析句 {alignment.get('analysis', 0)} 句（不计入问题）｜"
        f"证据级别：{audit['evidence_levels']}"
    )
    if alignment["verified"]:
        st.caption(f"可疑句经模型核实为真支持：{alignment['verified_supported']}/{alignment['verified']}")

    problems = p4_adapter.problem_sentences(audit)
    if problems:
        st.markdown(f"**需要修改的句子（{len(problems)} 条）**")
        st.dataframe(pd.DataFrame([{
            "类型": row["verdict"],
            "核实": "" if row.get("verified") is None else ("支持" if row["verified"] else "不支持"),
            "重合度": row["best_overlap"],
            "引用": "、".join(row["cited"]),
            "句子": row["sentence"],
        } for row in problems]), width="stretch", hide_index=True)
    else:
        st.success("没有发现未标注引用或引用对不上的句子。")
    if audit["unused_papers"]:
        st.warning(f"有证据但未被正文引用：{'、'.join(audit['unused_papers'])}")


def render() -> None:
    """主界面调用的入口。"""

    st.header("🧪 评测与审查（真实模式）")
    st.caption(
        "这一页只处理真实数据：真实论文语料、真实检索、真实大模型，以及本项目真实跑出来的产出。"
        "没有模拟用例，也没有离线演示。"
    )
    tabs = st.tabs(["内置文献库", "真实语料检索评测", "产出审查"])
    with tabs[0]:
        render_corpus_tab()
    with tabs[1]:
        render_retrieval_tab()
    with tabs[2]:
        render_audit_tab()
