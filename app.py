"""P4: 科研文献整合与论文初稿辅助。"""
import json
import streamlit as st
import pandas as pd
import subprocess
import sys
import os
import re
from datetime import datetime
from pathlib import Path
import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

PROJECT_DIR = Path(__file__).resolve().parent

# ============================================================
# 强制设置环境变量（确保 ccSwitch 配置生效）
# ============================================================
from dotenv import load_dotenv

load_dotenv()

OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
LLM_MODEL = os.getenv("LLM_MODEL", "gpt-4o-mini")

# ============================================================
# 导入 P4 科研文献整合系统（直接导入）
# ============================================================
sys.path.insert(0, str(PROJECT_DIR))

try:
    from p4_literature_assistant import (
        MultiAgentOrchestrator as P4Orchestrator,
        count_chinese_words,
        generate_report as generate_p4_report,
        llm as P4_LLM,
        deduplicate_papers,
        load_papers_from_bytes,
        STRUCTURED_FIELD_LABELS,
        WRITING_SKILL_CATALOG,
        WRITING_STYLE_PACKS,
        LLMError,
    )
    HAS_WRITER = True
except ImportError as e:
    HAS_WRITER = False
    WRITER_IMPORT_ERROR = str(e)

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(
    page_title="科研文献整合助手",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# 工具函数：清除 ANSI 转义码
# ============================================================
def clean_ansi(text: str) -> str:
    ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
    return ansi_escape.sub('', text)

FIELD_STATUS_OPTIONS = {"确认": "confirmed", "修改": "edited", "排除": "excluded"}
CARD_STATUS_OPTIONS = {"采用": "adopted", "排除": "excluded"}
OUTPUT_TYPE_OPTIONS = {"大纲": "outline", "结构化要点": "keypoints", "论文初稿": "draft"}
OUTLINE_FOCUS_FIELDS = ("method", "finding", "limitation")


def render_llm_error(error):
    """模型调用失败时给出明确结果，而不是静默降级成模板写作。"""
    st.error(
        f"模型调用失败：已重试 {error.get('attempts', 0)} 次仍未成功，本次没有生成内容。"
        "系统不会退回模板写作，避免产出看似完整但不可用的初稿。"
    )
    cols = st.columns(3)
    cols[0].metric("模型", error.get("model", "-") or "-")
    cols[1].metric("重试次数", error.get("attempts", 0))
    cols[2].metric("接口状态", "不可用")
    st.markdown(f"**失败原因：** `{error.get('last_error', '未知')}`")
    st.caption(f"接口地址：{error.get('base_url', '-')}")
    st.markdown(
        "排查建议：确认本地模型代理已启动（默认 `http://127.0.0.1:15721/v1`）；"
        "确认 `LLM_MODEL` 为 `deepseek-v4-flash`；推理模型需要较大的 token 预算，"
        "如仍超时可在 `.env` 中调大 `P4_LLM_TIMEOUT`，或调整 `P4_LLM_ATTEMPTS`。"
    )


def render_confirmation_stage(prepared):
    """人工确认节点：输出类型、结构化字段与论点证据。"""
    research = prepared.get("research", {})
    writing_skills = prepared.get("writing_skills", {})
    summary = prepared.get("execution_summary", {})
    outline = prepared.get("outline", {})
    structured_papers = research.get("structured_papers", [])
    cards = writing_skills.get("claim_evidence", {}).get("cards", [])
    default_sections = outline.get("sections", [])

    st.markdown("---")
    st.subheader("🧑‍⚖️ 人工确认节点")
    st.caption(
        "流程在生成内容前暂停。输出类型决定生成粒度：大纲只产出章节结构与证据顺序，不调用写作模型；"
        "结构化要点额外给出每章可引用证据；论文初稿才生成正文并执行引用覆盖与衔接评审。"
    )

    top_cols = st.columns([1.0, 1.0, 2.4])
    with top_cols[0]:
        output_label = st.radio(
            "输出类型", list(OUTPUT_TYPE_OPTIONS.keys()), index=2, key="p4_output_type"
        )
    with top_cols[1]:
        structure_source = st.radio(
            "结构来源", ["默认模板", "自定义章节"], index=0, key="p4_structure_source"
        )
    custom_sections = []
    with top_cols[2]:
        if structure_source == "自定义章节":
            raw_sections = st.text_area(
                "章节列表（每行一个章节）",
                value="\n".join(default_sections),
                key="p4_custom_sections",
                height=150,
            )
            custom_sections = [line.strip() for line in raw_sections.splitlines() if line.strip()]
            st.caption(f"当前 {len(custom_sections)} 个章节，生成时会按这些章节重新匹配文献素材。")
        else:
            st.caption("默认章节结构：" + "、".join(default_sections))

    style_cols = st.columns([1.2, 1.2, 2.0])
    with style_cols[0]:
        style_label = st.selectbox("写作风格", list(WRITING_STYLE_PACKS.keys()), index=0, key="p4_style")
    with style_cols[1]:
        enable_rewrite = st.checkbox("启用评审后重写", value=True, key="p4_rewrite")
    with style_cols[2]:
        st.caption("写作风格控制段落长度、句式与禁用表达；重写会针对评审标记的章节重新生成一次。")

    generated = prepared.get("generated_topic") or {}
    if generated.get("topic"):
        st.info(f"系统根据文献自动生成的研究主题：**{generated['topic']}**（来源：{generated.get('source', '模型生成')}）")
        if generated.get("research_question"):
            st.caption(f"研究问题：{generated['research_question']}")
    elif prepared.get("topic"):
        st.caption(f"当前研究主题：{prepared.get('topic')}")

    focus_fields = OUTLINE_FOCUS_FIELDS if output_label == "大纲" else tuple(STRUCTURED_FIELD_LABELS.keys())
    if output_label == "大纲":
        st.caption(
            "当前选择“大纲”：章节结构不依赖九字段逐条确认，因此只需确认研究方法、核心结论和研究局限，其余字段默认通过。"
        )

    alignment = research.get("topic_alignment") or {}
    gap_ratio = research.get("evidence_gap_ratio", 0.0)
    if alignment.get("ratio", 1.0) < 0.4:
        st.error(
            f"文献与主题相关性不足：{alignment.get('aligned', 0)}/{alignment.get('total', 0)} 篇文献包含主题关键词"
            f"（{alignment.get('ratio', 0):.0%}）。当前材料无法支撑这个主题的综述，继续生成只会得到大量"
            "“证据不足”的表述。建议换用与主题一致的文献，或把研究主题改成与已有文献匹配的方向。"
        )
    if gap_ratio > 0.5:
        st.warning(
            f"文献可用信息不足：结构化字段中有 {gap_ratio:.0%} 标记为待核验或未明确，"
            "通常意味着只拿到摘要、没有正文。建议上传 PDF 全文后重新提取。"
        )

    search_failures = research.get("search_failures") or []
    if search_failures:
        st.warning("部分数据源未返回结果：" + "；".join(search_failures))
    search_queries = research.get("search_queries") or []
    if search_queries:
        with st.expander("本次使用的检索式", expanded=False):
            for query in search_queries:
                st.write("- " + str(query))
            st.caption("中文检索式用于中文核心期刊定向检索，英文检索式用于 OpenAlex / Crossref / arXiv。")

    info_cols = st.columns(4)
    info_cols[0].metric("待确认论文", f"{summary.get('paper_count', 0)} 篇")
    info_cols[1].metric("待确认字段", f"{len(structured_papers) * len(focus_fields)} 个")
    info_cols[2].metric("待确认论点卡片", f"{len(cards)} 张")
    info_cols[3].metric("字段覆盖率", f"{summary.get('structured_field_coverage', 0):.0%}")

    field_decisions = {}
    card_decisions = {}

    tab_field, tab_card = st.tabs(["① 结构化字段确认", "② 论点证据确认"])

    with tab_field:
        st.caption("默认状态为“确认”。内容不准确时改为“修改”并填写正确内容；原文无法支持该字段时改为“排除”。")
        if not structured_papers:
            st.info("暂无结构化字段可确认。")
        for item in structured_papers:
            paper_id = item.get("paper_id", "")
            fields = item.get("fields", {})
            sources = item.get("field_sources", {})
            with st.expander(
                f"[{paper_id}] {item.get('title', '未命名论文')}｜字段 {item.get('field_count', 0)}/9",
                expanded=False,
            ):
                for field_key in focus_fields:
                    field_label = STRUCTURED_FIELD_LABELS[field_key]
                    st.markdown(f"**{field_label}** ｜ 来源：{sources.get(field_key, '未提取')}")
                    st.write(fields.get(field_key, ""))
                    choice = st.radio(
                        f"{field_label} 处理方式",
                        list(FIELD_STATUS_OPTIONS.keys()),
                        index=0,
                        horizontal=True,
                        key=f"p4_field_choice_{paper_id}_{field_key}",
                        label_visibility="collapsed",
                    )
                    value = ""
                    if choice == "修改":
                        value = st.text_area(
                            f"{field_label} 修改后的内容",
                            value=fields.get(field_key, ""),
                            key=f"p4_field_value_{paper_id}_{field_key}",
                            height=90,
                        )
                    field_decisions.setdefault(paper_id, {})[field_key] = {
                        "status": FIELD_STATUS_OPTIONS[choice],
                        "value": value,
                    }

    with tab_card:
        st.caption("默认状态为“采用”。证据不足以支撑该论点时改为“排除”，排除的卡片不会进入章节证据顺序。")
        if not cards:
            st.info("暂无论点证据卡片可确认。")
        current_paper = None
        for card in cards:
            paper_id = card.get("paper_id", "")
            if paper_id != current_paper:
                current_paper = paper_id
                st.markdown(f"**[{paper_id}] {card.get('title', '未命名论文')}**")
            row = st.columns([5, 1])
            with row[0]:
                st.markdown(f"{card.get('field_label', '')}：{card.get('claim', '')}")
                st.caption(
                    f"来源状态：{card.get('source_status', '')} ｜ 核验状态：{card.get('verification_status', '')} ｜ "
                    f"{card.get('writing_guidance', '')}"
                )
            with row[1]:
                choice = st.radio(
                    "处理方式",
                    list(CARD_STATUS_OPTIONS.keys()),
                    index=0,
                    key=f"p4_card_choice_{paper_id}_{card.get('field', '')}",
                    label_visibility="collapsed",
                )
            card_decisions[f"{paper_id}|{card.get('field', '')}"] = CARD_STATUS_OPTIONS[choice]

    edited_fields = sum(
        1 for paper in field_decisions.values() for decision in paper.values() if decision["status"] == "edited"
    )
    excluded_fields = sum(
        1 for paper in field_decisions.values() for decision in paper.values() if decision["status"] == "excluded"
    )
    excluded_cards = sum(1 for value in card_decisions.values() if value == "excluded")
    st.caption(
        f"当前选择：输出 {output_label} ｜ 结构来源 {structure_source} ｜ 修改字段 {edited_fields} 个 ｜ "
        f"排除字段 {excluded_fields} 个 ｜ 排除论点证据 {excluded_cards} 张。"
    )

    action_cols = st.columns([1, 1, 3])
    with action_cols[0]:
        if st.button(f"✅ 生成{output_label}", type="primary", key="p4_compose"):
            with st.spinner(f"正在按确认结果生成{output_label}…"):
                try:
                    composed = P4Orchestrator().compose(
                        prepared,
                        field_decisions=field_decisions,
                        card_decisions=card_decisions,
                        output_type=OUTPUT_TYPE_OPTIONS[output_label],
                        custom_sections=custom_sections or None,
                        style=style_label,
                        enable_rewrite=enable_rewrite,
                    )
                    st.session_state["p4_result"] = composed
                    st.session_state.pop("p4_error", None)
                    st.rerun()
                except LLMError as exc:
                    st.session_state["p4_error"] = exc.to_dict()
                    st.rerun()
                except Exception as exc:
                    st.error(f"生成失败：{exc}")
                    import traceback
                    st.code(traceback.format_exc())
    with action_cols[1]:
        if st.button("↺ 重新开始", key="p4_reset"):
            st.session_state.pop("p4_prepared", None)
            st.session_state.pop("p4_result", None)
            st.rerun()


# ============================================================
# P4 单项目入口
# ============================================================
st.sidebar.title("📚 科研文献整合助手")
st.sidebar.markdown("---")
st.sidebar.caption("面向科研新手的多文献整合与结构化写作辅助")
st.sidebar.caption("支持 PDF、Word、TXT、Markdown、JSON、CSV")
try:
    _manifest_path = PROJECT_DIR / "data" / "local_corpus_manifest.json"
    if _manifest_path.exists():
        _manifest = json.loads(_manifest_path.read_text(encoding="utf-8"))
        _corpus_file = PROJECT_DIR / "data" / "local_corpus.jsonl"
        _lines = 0
        if _corpus_file.exists():
            with _corpus_file.open("r", encoding="utf-8") as _handle:
                _lines = sum(1 for _line in _handle if _line.strip())
        st.sidebar.caption(f"📚 内置文献库：{_lines} 篇（已采集 {len(_manifest.get('done_venues') or [])} 个期刊）")
except Exception:
    pass
st.sidebar.markdown("---")
st.sidebar.caption(f"📅 {datetime.now().strftime('%Y-%m-%d %H:%M')}")
st.sidebar.caption("P4 Literature Assistant · v1.0")

# 页面选择：文献整合主流程，或真实模式的评测与审查
page = st.sidebar.radio("功能", ["文献整合", "评测与审查"], key="p4_page")
if page == "评测与审查":
    try:
        from evaluation.app import render as render_evaluation
    except ImportError as exc:
        st.error(f"评测模块加载失败：{exc}")
        st.stop()
    render_evaluation()
    st.stop()

project = "p4"

# ============================================================
# 页面内容
# ============================================================

# ---------- p4 科研文献整合 ----------
if project == "p4":
    st.header("✍️ 科研文献整合与论文初稿辅助")
    st.caption("面向科研新手：把已找到的多篇论文整理成可比较、可追溯的综述初稿")

    if not HAS_WRITER:
        st.warning("⚠️ 未找到 P4 文献整合模块，请检查项目文件是否完整。")
        st.error(f"导入错误详情: {WRITER_IMPORT_ERROR}")
        st.write(f"项目目录: {PROJECT_DIR}")
        st.write("目录下的 Python 文件:")
        for f in os.listdir(PROJECT_DIR):
            if f.endswith(".py"):
                st.write(f"  - {f}")
    else:
        st.info("典型场景：你已经收集了多篇关于某个课题的论文，但不知道如何分类、比较研究结论，也难以把材料组织成论文的研究背景或文献综述。")

        with st.form("p4_literature_form"):
            col1, col2 = st.columns(2)
            with col1:
                topic = st.text_input(
                    "研究主题 / 检索关键词（可留空）",
                    "",
                    help="留空时系统会根据上传的文献自动生成研究主题与题目。若不上传文献，请在这里填写检索关键词。",
                )
                research_question = st.text_area(
                    "研究问题（可留空）",
                    "",
                    height=110,
                    help="留空时系统会根据文献自动生成。",
                )
            with col2:
                template = st.selectbox("初稿结构", ["文献综述", "研究背景"])
                max_papers = st.number_input("最多整合论文数", min_value=1, max_value=20, value=8, step=1)
                use_online = st.checkbox("接入真实论文数据库检索", value=True)
                local_only = st.checkbox(
                    "仅用内置文献库（不联网）",
                    value=False,
                    help="内置语料库已收录高质量中英文期刊论文，命中足够时系统会自动跳过联网检索，"
                    "减少 OpenAlex 调用；勾选后完全离线检索。",
                )
                sources = st.multiselect(
                    "检索数据源",
                    ["中文核心期刊", "OpenAlex", "Crossref", "arXiv", "DOAJ"],
                    default=["中文核心期刊", "OpenAlex", "Crossref"],
                    help="中文核心期刊：按国内核心刊名定向检索（生态学报、中国科学、情报学报、计算机研究与发展等）；"
                    "OpenAlex 与 Crossref 覆盖国际期刊并带引用量；arXiv 覆盖 AI 前沿预印本；DOAJ 为开放获取期刊。",
                )
                year_range = st.selectbox("发表时间范围", ["近 5 年", "近 3 年", "近 10 年", "不限"], index=0)
                sort_label = st.selectbox("排序方式", ["按质量（引用密度＋期刊＋时效）", "按发表时间"], index=0)
                uploaded_files = st.file_uploader(
                    "导入已收集论文（PDF / Word / TXT / Markdown / JSON / CSV，可多选）",
                    type=["pdf", "docx", "txt", "md", "json", "csv"],
                    accept_multiple_files=True,
                    help="可直接拖入已下载的论文文件。系统优先提取标题、作者、年份、摘要和 DOI；扫描版 PDF 需要先完成文字识别。",
                )
            submitted = st.form_submit_button("🚀 开始文献整合")

        if submitted and not uploaded_files and not str(topic).strip():
            st.error("请至少上传一篇文献；如果没有现成文献，请在上方填写研究主题或检索关键词，系统会据此在线检索。")
            submitted = False

        if submitted:
            st.session_state.pop("p4_result", None)
            st.session_state.pop("p4_prepared", None)
            st.session_state.pop("p4_error", None)
            with st.spinner("正在执行：策划 → 文献调研 → 结构化提取 → 论点证据绑定…"):
                try:
                    imported_papers = None
                    parse_errors = []
                    parsed_files = []
                    if uploaded_files:
                        imported_papers = []
                        for uploaded_file in uploaded_files:
                            try:
                                papers = load_papers_from_bytes(uploaded_file.getvalue(), uploaded_file.name)
                                imported_papers.extend(papers)
                                parsed_files.append(uploaded_file.name)
                            except ValueError as exc:
                                parse_errors.append(f"{uploaded_file.name}：{exc}")
                        before_deduplication = len(imported_papers)
                        imported_papers = deduplicate_papers(imported_papers)
                        if parsed_files:
                            st.caption(
                                f"已读取 {len(parsed_files)} 个文件、识别 {len(imported_papers)} 篇论文"
                                + (f"，合并了 {before_deduplication - len(imported_papers)} 篇重复文献" if before_deduplication > len(imported_papers) else "")
                            )
                        for message in parse_errors:
                            st.warning(f"未导入：{message}")
                        if not imported_papers:
                            imported_papers = None
                            st.warning("上传文件中没有读取到有效论文，将继续使用在线检索或内置示例。")

                    os.environ["P4_LOCAL_CORPUS_ONLY"] = "true" if local_only else "false"
                    prepared = P4Orchestrator().prepare(
                        topic=topic,
                        research_question=research_question,
                        template=template,
                        papers=imported_papers,
                        use_online=use_online,
                        max_papers=int(max_papers),
                        sources=sources,
                        years={"近 5 年": 5, "近 3 年": 3, "近 10 年": 10, "不限": None}[year_range],
                        sort="quality" if sort_label.startswith("按质量") else "recent",
                    )
                    st.session_state["p4_prepared"] = prepared
                    st.session_state.pop("p4_error", None)
                except LLMError as exc:
                    st.session_state.pop("p4_prepared", None)
                    st.session_state["p4_error"] = exc.to_dict()
                except (UnicodeDecodeError, ValueError, json.JSONDecodeError) as e:
                    st.error(f"论文文件解析失败：{e}")
                except Exception as e:
                    st.error(f"运行失败：{e}")
                    import traceback
                    st.code(traceback.format_exc())

        llm_error = st.session_state.get("p4_error")
        if llm_error:
            render_llm_error(llm_error)

        prepared = st.session_state.get("p4_prepared")
        if prepared:
            render_confirmation_stage(prepared)

        result = st.session_state.get("p4_result")
        if result:
            summary = result.get("execution_summary", {})
            review = result.get("review", {})
            research = result.get("research", {})
            confirmation = result.get("confirmation", {}) or {}
            output_label_result = summary.get("output_label", "论文初稿")
            review_applicable = summary.get("review_applicable", True)
            st.success(f"✅ {output_label_result}已生成（已应用人工确认结果）")
            if confirmation:
                st.caption(
                    f"输出 {confirmation.get('output_label', output_label_result)}"
                    f"（{confirmation.get('structure_source', '默认模板')}）；"
                    f"字段确认 {confirmation.get('confirmed_fields', 0)} 个、修改 {confirmation.get('edited_fields', 0)} 个、"
                    f"排除 {confirmation.get('excluded_fields', 0)} 个；论点证据采用 {confirmation.get('adopted_cards', 0)} 张、"
                    f"排除 {confirmation.get('excluded_cards', 0)} 张。"
                )
                if confirmation.get("excluded_evidence") or confirmation.get("notes"):
                    with st.expander("人工确认记录", expanded=False):
                        if confirmation.get("excluded_evidence"):
                            st.write("已排除字段：" + "、".join(confirmation["excluded_evidence"]))
                        for note in confirmation.get("notes", []):
                            st.write(note)
            metric_cols = st.columns(6) if review_applicable else st.columns(5)
            metric_cols[0].metric("整合论文", f"{summary.get('paper_count', 0)} 篇")
            metric_cols[1].metric("主题归纳", f"{len(research.get('themes', []))} 个")
            metric_cols[2].metric("研究空白", f"{len(research.get('gaps', []))} 条")
            if review_applicable:
                metric_cols[3].metric("来源覆盖率", f"{summary.get('citation_coverage', 0):.0%}")
                metric_cols[4].metric("评审得分", f"{summary.get('quality_score', 0)}/10")
                metric_cols[5].metric("写作质量分", f"{summary.get('writing_quality_score', 0)}/10")
            else:
                metric_cols[3].metric("输出类型", output_label_result)
                metric_cols[4].metric(
                    "已绑定证据章节",
                    f"{summary.get('evidence_bound_sections', 0)}/{summary.get('section_count', 0)}",
                )
            model_mode = "真实 API" if P4_LLM.available else "规则降级模式"
            st.caption(
                f"论文来源：{summary.get('source_mode', '未知')}；模型模式：{model_mode}；"
                f"内容分析覆盖率：{summary.get('content_analysis_coverage', research.get('content_analysis_coverage', 0)):.0%}"
            )
            alignment = research.get("topic_alignment") or {}
            gap_ratio = research.get("evidence_gap_ratio", 0.0)
            review_flags = review.get("quality_flags", []) if isinstance(review, dict) else []
            for flag in review_flags:
                if "相关性不足" in flag or "可用信息不足" in flag:
                    st.error(flag)
            if alignment.get("ratio", 1.0) < 0.4:
                st.caption(
                    f"主题相关性：{alignment.get('aligned', 0)}/{alignment.get('total', 0)} 篇命中主题关键词"
                    f"（{alignment.get('ratio', 0):.0%}）"
                )
            if not P4_LLM.available:
                reason = getattr(P4_LLM, "last_error", "") or "未配置模型服务"
                st.warning(
                    f"当前未调用写作模型，正文由规则模板生成，语言质量会明显低于模型写作。原因：{reason}。"
                    f"当前模型地址：{getattr(P4_LLM, 'base_url', '未设置')}，模型：{getattr(P4_LLM, 'model', '未设置')}。"
                    "请在 .env 中配置可用的 OPENAI_API_KEY 与 OPENAI_BASE_URL，或启动本地模型代理后重新生成。"
                )

            quality = result.get("writing_quality", {}) or {}
            citation = result.get("citation_check", {}) or {}
            if quality or citation:
                with st.expander("✍️ 写作质量与引用校验", expanded=False):
                    if quality:
                        q_cols = st.columns(4)
                        q_cols[0].metric("主题句覆盖率", f"{quality.get('topic_coverage', 0):.0%}")
                        q_cols[1].metric("分析句占比", f"{quality.get('analysis_ratio', 0):.0%}")
                        q_cols[2].metric("重复句比例", f"{quality.get('duplicate_ratio', 0):.0%}")
                        q_cols[3].metric("段落数", f"{quality.get('paragraph_count', 0)}")
                        st.caption(f"评分依据：{quality.get('score_basis', '')}")
                        if quality.get("moves_used"):
                            st.write("已使用的论证动作：" + "、".join(quality["moves_used"]))
                        st.write(f"连接词多样性：{quality.get('connective_diversity', 0):.0%}")
                        if quality.get("banned_hits"):
                            st.warning("命中禁用表达：" + "、".join(quality["banned_hits"]))
                        if quality.get("ai_flavor_hits"):
                            st.warning("命中 AI 味表达：" + "、".join(quality["ai_flavor_hits"]))
                    if citation:
                        st.write(
                            f"引用校验：有效引用 {citation.get('cited_count', 0)}/{citation.get('valid_count', 0)}；"
                            f"无效编号 {len(citation.get('invalid_ids', []))} 个；待核验句 {len(citation.get('unsupported_sentences', []))} 句。"
                        )
                        for item in citation.get("unsupported_sentences", [])[:5]:
                            st.warning(f"证据重合度 {item.get('best_overlap')}：{item.get('sentence')}")
                    for entry in result.get("rewrite_log", []) or []:
                        st.info(
                            f"已重写章节：{'、'.join(entry.get('sections', []))}；"
                            f"评审分 {entry.get('score_before')} → {entry.get('score_after')}。"
                        )

            with st.expander("📚 已纳入的论文来源", expanded=True):
                papers = result.get("papers", [])
                if papers:
                    paper_rows = [
                        {
                            "编号": paper.ref_id,
                            "论文": paper.title,
                            "作者": ", ".join(paper.authors[:6]) or "未知",
                            "年份": paper.year or "未知",
                            "来源": paper.source,
                            "期刊/会议": getattr(paper, "venue", "") or "未标注",
                            "被引": getattr(paper, "citation_count", 0) or 0,
                            "解析置信度": paper.metadata_confidence,
                        }
                        for paper in papers
                    ]
                    st.dataframe(pd.DataFrame(paper_rows), width="stretch", hide_index=True)
                    warnings = [
                        f"{paper.ref_id}《{paper.title}》：{warning}"
                        for paper in papers
                        for warning in getattr(paper, "parse_warnings", [])
                    ]
                    warnings = list(dict.fromkeys(warnings))
                    if warnings:
                        with st.expander("解析质量提示", expanded=False):
                            st.warning(f"共发现 {len(warnings)} 条解析提示，以下展示前 12 条。")
                            for warning in warnings[:12]:
                                st.warning(warning)
                            if len(warnings) > 12:
                                st.caption(f"其余 {len(warnings) - 12} 条提示与上述内容相同或属于同类书目信息提醒。")

            tab_structured, tab_skills, tab_theme, tab_matrix, tab_gap, tab_article, tab_review = st.tabs(
                ["文献结构化提取", "写作 Skill 工作台", "主题整合", "文献比较矩阵", "研究空白", output_label_result, "评审与迭代"]
            )
            with tab_structured:
                structured_papers = research.get("structured_papers", [])
                if structured_papers:
                    st.caption(
                        f"已按统一字段整理 {len(structured_papers)} 篇论文；字段覆盖率 "
                        f"{research.get('structured_field_coverage', 0):.0%}。未在原文找到的内容会标记为待核验。"
                    )
                    for index, item in enumerate(structured_papers):
                        metadata = item.get("metadata", {})
                        with st.expander(
                            f"[{item.get('paper_id', f'R{index + 1}')}] {item.get('title', '未命名论文')}",
                            expanded=index == 0,
                        ):
                            st.caption(
                                f"作者：{', '.join(metadata.get('authors', [])) or '未知'}｜"
                                f"年份：{metadata.get('year') or '未知'}｜"
                                f"DOI：{metadata.get('doi') or '未识别'}｜"
                                f"文本状态：{metadata.get('text_quality', '未知')}｜"
                                f"字段：{item.get('field_count', 0)}/9"
                            )
                            fields = item.get("fields", {})
                            sources = item.get("field_sources", {})
                            for field_key, field_label in STRUCTURED_FIELD_LABELS.items():
                                st.markdown(f"**{field_label}** · {sources.get(field_key, '未提取')}")
                                st.write(fields.get(field_key, "原文未明确说明。"))
                            if item.get("warnings"):
                                st.warning("；".join(item["warnings"]))
                else:
                    st.info("暂无结构化提取结果，请先上传论文并开始文献整合。")
            with tab_skills:
                writing_skills = result.get("writing_skills", {})
                st.markdown(f"### 本次流程使用的 {len(writing_skills.get('catalog', WRITING_SKILL_CATALOG))} 个写作 Skill")
                st.caption("Skill 不是单独的聊天按钮，而是嵌入文献整理与论文初稿流程的具体能力模块。")
                for skill in writing_skills.get("catalog", WRITING_SKILL_CATALOG):
                    with st.expander(f"{skill.get('name', '未命名 Skill')} · {skill.get('purpose', '')}", expanded=True):
                        st.markdown(f"**输入：** {skill.get('input', '')}")
                        st.markdown(f"**输出：** {skill.get('output', '')}")

                binding = writing_skills.get("claim_evidence", {})
                st.markdown("### 1. 论点-证据绑定结果")
                st.caption(
                    f"共生成 {len(binding.get('cards', []))} 张论点卡片；"
                    f"已支持 {binding.get('supported_count', 0)} 张，"
                    f"覆盖率 {binding.get('coverage', 0):.0%}。"
                )
                binding_rows = [
                    {
                        "论文": card.get("paper_id", ""),
                        "字段": card.get("field_label", ""),
                        "论点": card.get("claim", ""),
                        "证据状态": card.get("verification_status", ""),
                        "来源": card.get("source_status", ""),
                        "使用建议": card.get("writing_guidance", ""),
                    }
                    for card in binding.get("cards", [])
                ]
                if binding_rows:
                    st.dataframe(pd.DataFrame(binding_rows), width="stretch", hide_index=True)
                else:
                    st.info("暂无可绑定的论点与证据。")

                synthesis = writing_skills.get("cross_paper_synthesis", {})
                st.markdown("### 2. 跨文献综合写作结果")
                for item in synthesis.get("syntheses", []):
                    st.markdown(f"**{item.get('id', '')}｜{item.get('theme', '相关研究')}**")
                    st.write(item.get("paragraph", ""))
                    st.caption(f"涉及来源：{', '.join(item.get('paper_ids', []))}；{item.get('writing_advice', '')}")
                if not synthesis.get("syntheses"):
                    st.info("暂无同主题文献综合段落。")

                structure = writing_skills.get("section_argument_structure", {})
                st.markdown("### 3. 章节论证结构")
                for plan in structure.get("plans", []):
                    with st.expander(plan.get("section", "未命名章节"), expanded=False):
                        st.markdown(f"**章节目标：** {plan.get('goal', '')}")
                        st.markdown(f"**中心论点：** {plan.get('core_claim', '')}")
                        st.markdown("**写作步骤：**")
                        for step in plan.get("writing_steps", []):
                            st.write(f"- {step}")
                        st.markdown(f"**过渡句建议：** {plan.get('transition', '')}")
                        evidence_order = plan.get("evidence_order", [])
                        if evidence_order:
                            st.caption("建议证据顺序：" + " → ".join(item.get("paper_id", "") for item in evidence_order))

                planning = writing_skills.get("paragraph_planning", {})
                st.markdown("### 4. 论证动作编排结果")
                planning_sections = planning.get("sections", {})
                if planning_sections:
                    for section_name, payload in planning_sections.items():
                        with st.expander(section_name, expanded=False):
                            for plan in payload.get("paragraphs", []):
                                st.markdown(f"**{plan.get('id', '')}｜{plan.get('move', '')}**")
                                st.write(plan.get("topic_sentence", ""))
                                st.caption(
                                    "证据：" + ("、".join(plan.get("evidence_ids", [])) or "未绑定")
                                    + (f"｜衔接：{plan.get('transition')}" if plan.get("transition") else "")
                                )
                else:
                    st.info("暂无段落规划结果。")

                logic = writing_skills.get("logic_coherence_review", {})
                st.markdown("### 5. 逻辑与衔接评审")
                st.metric("逻辑评审分", f"{logic.get('score', 0)}/10")
                st.write(f"**评审结论：** {logic.get('verdict', '待评审')}")
                for strength in logic.get("strengths", []):
                    st.success(strength)
                logic_rows = [
                    {
                        "等级": issue.get("level", ""),
                        "章节": issue.get("section", ""),
                        "问题": issue.get("issue", ""),
                        "修改建议": issue.get("suggestion", ""),
                    }
                    for issue in logic.get("issues", [])
                ]
                if logic_rows:
                    st.dataframe(pd.DataFrame(logic_rows), width="stretch", hide_index=True)
                else:
                    st.success("暂未发现明显的逻辑、引用或章节衔接问题。")
            with tab_theme:
                themes = research.get("themes", [])
                if themes:
                    for theme in themes:
                        st.markdown(f"#### {theme.get('name', '未命名主题')}")
                        st.write(theme.get("summary", ""))
                        st.caption(f"依据来源：{', '.join(theme.get('paper_ids', [])) or '未绑定来源'}")
                else:
                    st.info("暂未提炼出主题，请补充论文摘要或调整研究主题。")
            with tab_matrix:
                comparisons = research.get("comparisons", [])
                matrix_rows = [
                    {
                        "来源": item.get("paper_id", ""),
                        "研究问题": item.get("research_question", ""),
                        "方法": item.get("method", ""),
                        "主要结论": item.get("finding", ""),
                        "局限": item.get("limitation", ""),
                    }
                    for item in comparisons
                ]
                if matrix_rows:
                    st.dataframe(pd.DataFrame(matrix_rows), width="stretch", hide_index=True)
                else:
                    st.info("暂未生成比较矩阵。")
            with tab_gap:
                gaps = research.get("gaps", [])
                if gaps:
                    for index, gap in enumerate(gaps, 1):
                        st.markdown(f"**{index}.** {gap}")
                else:
                    st.info("暂未提炼出研究空白，需要结合全文和实验细节进一步核验。")
            with tab_article:
                st.markdown(result.get("article", ""))
            with tab_review:
                if not review_applicable:
                    st.info(review.get("reason", "当前输出类型未执行全文评审。"))
                    for suggestion in review.get("suggestions", []):
                        st.markdown(f"**建议：** {suggestion}")
                else:
                    score_rows = [{"评审维度": key, "得分": value} for key, value in review.get("scores", {}).items()]
                    if score_rows:
                        st.dataframe(pd.DataFrame(score_rows), width="stretch", hide_index=True)
                    st.write(f"**评审结论：** {review.get('verdict', '需要修改')}")
                    if review.get("quality_flags"):
                        st.warning("；".join(review["quality_flags"]))
                    for index, suggestion in enumerate(review.get("suggestions", []), 1):
                        st.markdown(f"**迭代建议 {index}：** {suggestion}")
                    if review.get("missing_sources"):
                        st.warning(f"正文尚未引用这些已有证据的文献：{', '.join(review['missing_sources'])}")
                    if review.get("non_writable_sources"):
                        st.info(
                            f"以下文献仅有书目信息、没有摘要或正文，未参与写作，也不计入来源覆盖率："
                            f"{', '.join(review['non_writable_sources'])}"
                        )

            st.download_button(
                "📥 下载完整 Markdown 报告",
                data=generate_p4_report(result),
                file_name="p4-literature-report.md",
                mime="text/markdown",
            )

# ---------- p6 视频管道 ----------
elif project == "🎬 p6 视频管道":
    st.header("🎬 视频知识管道")
    st.caption("将视频自动转化为可检索的结构化知识")

    # 视频处理流程（固定展示）
    st.markdown("### 📋 视频处理流程")
    steps = [
        "📌 Step 1: 提取视频元数据 (时长/分辨率)",
        "📌 Step 2: 提取音频 (ffmpeg)",
        "📌 Step 3: 语音转文字 (Whisper)",
        "📌 Step 4: 生成摘要与关键点 (LLM)",
        "📌 Step 5: 保存知识条目 (Markdown + JSON)"
    ]
    for step in steps:
        st.write(step)

    st.markdown("---")

    # 配置区域
    col1, col2 = st.columns(2)
    with col1:
        mock_mode = st.checkbox("🔧 模拟模式（无需 ffmpeg/whisper）", value=True)
    with col2:
        video_path = st.text_input(
            "📹 视频文件路径",
            value=str(PROJECT_DIR / "videos" / "AI_Agent_趋势分析.mp4"),
            help="输入本地视频文件的完整路径"
        )

    # 处理按钮
    if st.button("🎬 处理视频", key="p6_run"):
        if not mock_mode and not os.path.exists(video_path):
            st.error(f"❌ 视频文件不存在: {video_path}")
        else:
            with st.spinner("正在处理视频...（可能需要 1-2 分钟）"):
                try:
                    env = os.environ.copy()
                    env["PYTHONIOENCODING"] = "utf-8"
                    env["PYTHONUTF8"] = "1"

                    cmd = [sys.executable, str(PROJECT_DIR / "p6-video-pipeline.py")]
                    if mock_mode:
                        cmd.append("--mock")
                    else:
                        cmd.append("--real")
                    cmd.append(video_path)

                    result = subprocess.run(
                        cmd,
                        capture_output=True,
                        text=True,
                        timeout=600,
                        encoding='utf-8',
                        errors='replace',
                        env=env,
                        cwd=str(PROJECT_DIR)
                    )
                    raw_output = result.stdout or result.stderr or ""

                    if not raw_output.strip():
                        st.warning("⚠️ 脚本没有输出，请检查 `p6-video-pipeline.py` 是否存在。")
                    else:
                        clean_output = clean_ansi(raw_output)

                        # 解析关键信息
                        summary_match = re.search(r"📝 摘要: (.+?)(?:\n|$)", clean_output)
                        tags_match = re.search(r"🏷️  标签: (.+?)(?:\n|$)", clean_output)

                        if summary_match:
                            st.success(f"📝 摘要: {summary_match.group(1)}")
                        if tags_match:
                            st.info(f"🏷️ 标签: {tags_match.group(1)}")

                        # 检查是否有生成的知识文件
                        knowledge_dir = PROJECT_DIR / "knowledge_output"
                        if knowledge_dir.exists():
                            md_files = list(knowledge_dir.glob("*.md"))
                            if md_files:
                                latest_file = max(md_files, key=lambda f: f.stat().st_mtime)
                                st.success(f"📄 知识报告已生成: `{latest_file.name}`")
                                with open(latest_file, "r", encoding="utf-8") as f:
                                    report_content = f.read()
                                st.download_button(
                                    label="📥 下载知识报告 (Markdown)",
                                    data=report_content,
                                    file_name=latest_file.name,
                                    mime="text/markdown"
                                )

                        with st.expander("📋 查看完整处理日志"):
                            st.text(clean_output)

                except subprocess.TimeoutExpired:
                    st.error("⏰ 脚本执行超时（>120秒）")
                except Exception as e:
                    st.error(f"❌ 运行失败: {e}")
                    import traceback
                    st.code(traceback.format_exc())

    # 使用说明（展开式）—— 改用普通字符串避免三引号问题
    with st.expander("💡 使用说明"):
        st.markdown(
            "**两种模式说明：**\n\n"
            "| 模式 | 说明 | 依赖 |\n"
            "| :--- | :--- | :--- |\n"
            "| **模拟模式** | 使用内置示例数据快速演示流程 | 无 |\n"
            "| **真实模式** | 处理真实视频文件（需要 ffmpeg + whisper） | ffmpeg, whisper |\n\n"
            "**真实模式安装依赖：**\n\n"
            "```powershell\n"
            "# 安装 ffmpeg (Windows)\n"
            "winget install ffmpeg\n\n"
            "# 安装 whisper\n"
            "pip install openai-whisper\n"
            "```"
        )
