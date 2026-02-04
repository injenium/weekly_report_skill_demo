import os
import json
from datetime import datetime
import streamlit as st
import pandas as pd
st.write("RUNNING FROM:", os.path.abspath(__file__))
st.write("LAST MOD:", os.path.getmtime(__file__))
from tools import (
    read_table,
    normalize_columns,
    compute_weekly_kpis,
    dataframe_to_markdown_table,
    load_skill_pack,
    call_ollama_chat,
    build_report_prompt,
    make_docx_from_markdown_text,
)

st.set_page_config(page_title="Skills Demo | 项目周报生成器", layout="wide")

st.title("🧩 Skills Demo：项目周报生成器（本地模型 + Ollama）")

# ---- Sidebar: config ----
st.sidebar.header("配置")
ollama_host = st.sidebar.text_input("Ollama Host", value=os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434"))
ollama_model = st.sidebar.text_input("Ollama Model", value=os.getenv("OLLAMA_MODEL", "qwen3:14b"))
temperature = st.sidebar.slider("temperature", 0.0, 1.5, 0.3, 0.05)

st.sidebar.divider()
st.sidebar.header("Skill（技能包）")
skill_name = st.sidebar.selectbox("选择 Skill", ["None", "weekly_report"])

st.sidebar.caption("提示：Skills 的本质是把 SOP + 模板 + 口径固化为可插拔模块。")
st.sidebar.divider()
show_trace = st.sidebar.checkbox("显示执行轨迹", value=True)

# ---- Upload ----
colL, colR = st.columns([1, 1])
with colL:
    st.subheader("1) 上传任务清单（Excel / CSV）")
    file = st.file_uploader("上传文件", type=["xlsx", "xls", "csv"])
    st.caption("建议列：project / module / owner / status / priority / due_date / progress / blocker / risk")

# ---- Read & preview ----
trace = []
df = None
if file is not None:
    trace.append(f"Loaded file: {file.name}")
    df = read_table(file)
    trace.append(f"Parsed rows: {len(df)}")

    df = normalize_columns(df)
    trace.append("Normalized columns")

    with colL:
        st.write("数据预览（前 50 行）")
        st.dataframe(df.head(50), width=True)

# ---- Generate report ----
with colR:
    st.subheader("2) 生成周报")
    default_prompt = "给我生成本周项目周报：总体进度、里程碑、Top 风险、按负责人统计、下周行动清单（按公司模板输出）。"
    user_request = st.text_area("自然语言", value=default_prompt, height=120)

    gen = st.button("🚀 生成周报", type="primary", disabled=(df is None))

    if gen and df is not None:
        # Load skill pack (optional)
        skill_pack = None
        if skill_name != "None":
            skill_pack = load_skill_pack(skill_name)
            trace.append(f"Loaded skill pack: {skill_name}")

        # Compute KPIs locally (tool layer)
        kpis = compute_weekly_kpis(df)
        trace.append("Computed weekly KPIs")

        # Prepare compact context for LLM
        table_md = dataframe_to_markdown_table(df, max_rows=25)
        prompt = build_report_prompt(user_request, kpis, table_md, skill_pack)
        trace.append("Built LLM prompt")

        # Call local Ollama
        try:
            response_text = call_ollama_chat(
                host=ollama_host,
                model=ollama_model,
                system=prompt["system"],
                user=prompt["user"],
                temperature=temperature,
            )
            trace.append("Ollama chat completed")
        except Exception as e:
            st.error(f"调用 Ollama 失败：{e}")
            st.stop()

        st.markdown("### ✅ 周报输出（Markdown）")
        st.markdown(response_text)

        # Downloads
        st.download_button(
            "⬇️ 下载 Markdown",
            data=response_text.encode("utf-8"),
            file_name=f"weekly_report_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
            mime="text/markdown",
        )

        # Optional DOCX
        docx_bytes = make_docx_from_markdown_text(response_text)
        st.download_button(
            "⬇️ 下载 DOCX（纯文本版）",
            data=docx_bytes,
            file_name=f"weekly_report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# ---- Trace ----
if show_trace:
    st.divider()
    st.subheader("执行轨迹")
    st.code("\n".join([f"- {x}" for x in trace]) if trace else "- （等待操作）")
