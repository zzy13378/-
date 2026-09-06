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

# ============================================================
# P4 单项目入口
# ============================================================
st.sidebar.title("📚 科研文献整合助手")
st.sidebar.markdown("---")
st.sidebar.caption("面向科研新手的多文献整合与结构化写作辅助")
st.sidebar.caption("支持 PDF、Word、TXT、Markdown、JSON、CSV")
st.sidebar.markdown("---")
st.sidebar.caption(f"📅 {datetime.now().strftime('%Y-%m-%d %H:%M')}")
st.sidebar.caption("P4 Literature Assistant · v1.0")
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
                topic = st.text_input("研究主题", "大语言模型智能体在科研场景中的应用")
                research_question = st.text_area(
                    "研究问题",
                    "现有研究主要采用哪些方法解决科研场景中的知识获取与任务协作问题？不同方法的效果、局限和研究空白是什么？",
                    height=110,
                )
            with col2:
                template = st.selectbox("初稿结构", ["文献综述", "研究背景"])
                max_papers = st.number_input("最多整合论文数", min_value=1, max_value=20, value=8, step=1)
                use_online = st.checkbox("接入 OpenAlex 真实论文数据库检索", value=True)
                uploaded_files = st.file_uploader(
                    "导入已收集论文（PDF / Word / TXT / Markdown / JSON / CSV，可多选）",
                    type=["pdf", "docx", "txt", "md", "json", "csv"],
                    accept_multiple_files=True,
                    help="可直接拖入已下载的论文文件。系统优先提取标题、作者、年份、摘要和 DOI；扫描版 PDF 需要先完成文字识别。",
                )
            submitted = st.form_submit_button("🚀 开始文献整合")

        if submitted:
            with st.spinner("正在执行：策划 → 文献调研 → 结构化写作 → 来源评审…"):
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

                    orchestrator = P4Orchestrator()
                    result = orchestrator.run(
                        topic=topic,
                        research_question=research_question,
                        template=template,
                        papers=imported_papers,
                        use_online=use_online,
                        max_papers=int(max_papers),
                    )
                    st.session_state["p4_result"] = result
                except (UnicodeDecodeError, ValueError, json.JSONDecodeError) as e:
                    st.error(f"论文文件解析失败：{e}")
                except Exception as e:
                    st.error(f"运行失败：{e}")
                    import traceback
                    st.code(traceback.format_exc())

        result = st.session_state.get("p4_result")
        if result:
            summary = result.get("execution_summary", {})
            review = result.get("review", {})
            research = result.get("research", {})
            st.success("✅ 文献整合完成")
            metric_cols = st.columns(5)
            metric_cols[0].metric("整合论文", f"{summary.get('paper_count', 0)} 篇")
            metric_cols[1].metric("主题归纳", f"{len(research.get('themes', []))} 个")
            metric_cols[2].metric("研究空白", f"{len(research.get('gaps', []))} 条")
            metric_cols[3].metric("来源覆盖率", f"{summary.get('citation_coverage', 0):.0%}")
            metric_cols[4].metric("评审得分", f"{summary.get('quality_score', 0)}/10")
            st.caption(
                f"论文来源：{summary.get('source_mode', '未知')}；模型模式：{'真实 API' if P4_LLM.available else '规则降级模式'}；"
                f"内容分析覆盖率：{summary.get('content_analysis_coverage', research.get('content_analysis_coverage', 0)):.0%}"
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

            tab_theme, tab_matrix, tab_gap, tab_article, tab_review = st.tabs(
                ["主题整合", "文献比较矩阵", "研究空白", "论文初稿", "评审与迭代"]
            )
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
                score_rows = [{"评审维度": key, "得分": value} for key, value in review.get("scores", {}).items()]
                if score_rows:
                    st.dataframe(pd.DataFrame(score_rows), width="stretch", hide_index=True)
                st.write(f"**评审结论：** {review.get('verdict', '需要修改')}")
                if review.get("quality_flags"):
                    st.warning("；".join(review["quality_flags"]))
                for index, suggestion in enumerate(review.get("suggestions", []), 1):
                    st.markdown(f"**迭代建议 {index}：** {suggestion}")
                if review.get("missing_sources"):
                    st.warning(f"正文尚未覆盖来源：{', '.join(review['missing_sources'])}")

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
