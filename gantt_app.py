import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(layout="wide")
st.title("📊 项目进度管理系统（完整版）")

import io

st.subheader("📥 下载Excel模板")

template_df = pd.DataFrame({
    "Project": ["项目A"],
    "Task": ["任务1"],
    "Start": ["2024-04-01"],
    "Finish": ["2024-04-05"],
    "Progress": [50],
    "Owner": ["张三"]
})

output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    template_df.to_excel(writer, index=False, sheet_name="template")

excel_data = output.getvalue()

st.download_button(
    label="📥 下载项目模板Excel",
    data=excel_data,
    file_name="gantt_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
# ===== 上传文件 =====
uploaded_file = st.file_uploader("上传Excel文件", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("📄 数据预览")
    st.dataframe(df)

    # ===== 检查必要列 =====
    required_cols = ["Project", "Task", "Start", "Finish", "Progress", "Owner"]
    if not all(col in df.columns for col in required_cols):
        st.error(f"Excel必须包含列：{required_cols}")
        st.stop()

    # ===== 日期转换（避免报错）=====
    df["Start"] = pd.to_datetime(df["Start"], errors="coerce")
    df["Finish"] = pd.to_datetime(df["Finish"], errors="coerce")

    # ===== 删除无效数据 =====
    df = df.dropna(subset=["Start", "Finish"])

    # ===== 多项目筛选 =====
    projects = df["Project"].dropna().unique()
    selected_projects = st.multiselect(
        "选择项目（可多选）",
        projects,
        default=projects
    )

    df_filtered = df[df["Project"].isin(selected_projects)].copy()

    # ===== 配色方式 =====
    color_option = st.selectbox(
        "选择配色方式",
        ["按项目", "按负责人", "按进度"]
    )

    if color_option == "按项目":
        color_col = "Project"
    elif color_option == "按负责人":
        color_col = "Owner"
    else:
        color_col = "Progress"

    # ===== 甘特图 =====
    fig = px.timeline(
        df_filtered,
        x_start="Start",
        x_end="Finish",
        y="Task",
        color=color_col,
        hover_data=["Project", "Owner", "Progress"]
    )

    fig.update_yaxes(autorange="reversed")

    # 显示进度
    fig.update_traces(
        text=df_filtered["Progress"].astype(str) + "%",
        textposition="inside"
    )

    st.plotly_chart(fig, use_container_width=True)

    # ===== 导出（已修复稳定版）=====
    st.subheader("📤 导出图表")

    try:
        img_bytes = fig.to_image(format="png")
        st.download_button(
            "下载 PNG 图片",
            data=img_bytes,
            file_name="gantt_chart.png",
            mime="image/png"
        )
    except:
        st.warning("⚠️ PNG导出失败，请安装 kaleido：pip install kaleido")

    try:
        pdf_bytes = fig.to_image(format="pdf")
        st.download_button(
            "下载 PDF 文件",
            data=pdf_bytes,
            file_name="gantt_chart.pdf",
            mime="application/pdf"
        )
    except:
        st.warning("⚠️ PDF导出失败，请安装 kaleido：pip install kaleido")

    # ===== 项目分析 =====
    st.subheader("📋 项目分析报告")

    today = pd.to_datetime(datetime.today().date())

    # 延误任务
    delayed_tasks = df_filtered[
        (df_filtered["Finish"] < today) &
        (df_filtered["Progress"] < 100)
    ]

    # 风险任务
    df_filtered["days_left"] = (df_filtered["Finish"] - today).dt.days

    risk_tasks = df_filtered[
        (df_filtered["days_left"] <= 2) &
        (df_filtered["Progress"] < 80) &
        (df_filtered["Progress"] < 100)
    ]

    # ===== 输出 =====
    st.markdown("### ⏰ 延误任务")

    if delayed_tasks.empty:
        st.success("没有延误任务 👍")
    else:
        st.error(f"发现 {len(delayed_tasks)} 个延误任务")
        st.dataframe(delayed_tasks[["Project", "Task", "Owner", "Finish", "Progress"]])

    st.markdown("### ⚠️ 延期风险任务")

    if risk_tasks.empty:
        st.success("暂无高风险任务 👍")
    else:
        st.warning(f"发现 {len(risk_tasks)} 个风险任务")
        st.dataframe(risk_tasks[["Project", "Task", "Owner", "days_left", "Progress"]])

    # ===== 自动总结 =====
    st.subheader("🧠 自动总结")

    total_tasks = len(df_filtered)
    completed = len(df_filtered[df_filtered["Progress"] == 100])
    avg_progress = df_filtered["Progress"].mean()

    summary = f"""
当前共 {total_tasks} 个任务，已完成 {completed} 个，
整体平均进度为 {avg_progress:.1f}%。

延误任务 {len(delayed_tasks)} 个，
存在风险任务 {len(risk_tasks)} 个。
"""

    st.info(summary)

else:
    st.info("请先上传Excel文件 👆")

    st.subheader("🧠 自动生成项目周报（AI风格）")


    def generate_report(df, delayed_tasks, risk_tasks):
        total = len(df)
        completed = len(df[df["Progress"] == 100])
        avg_progress = df["Progress"].mean()

        # 项目维度统计
        project_summary = df.groupby("Project")["Progress"].mean().round(1).to_dict()

        report = f"""
    【项目周报】

    一、整体概况
    当前共 {total} 个任务，已完成 {completed} 个任务。
    整体平均进度为 {avg_progress:.1f}%。

    二、项目进展情况
    """

        for k, v in project_summary.items():
            report += f"- {k} 平均完成度：{v}%\n"

        report += "\n三、风险情况\n"

        if len(delayed_tasks) > 0:
            report += f"- 存在 {len(delayed_tasks)} 个延误任务，需立即跟进。\n"
        else:
            report += "- 无明显延误任务。\n"

        if len(risk_tasks) > 0:
            report += f"- 存在 {len(risk_tasks)} 个潜在风险任务，建议提前干预。\n"
        else:
            report += "- 暂无明显风险任务。\n"

        report += "\n四、管理建议\n"

        if len(risk_tasks) > 0 or len(delayed_tasks) > 0:
            report += "- 建议优先处理高风险与已延误任务，避免影响整体交付进度。\n"
        else:
            report += "- 当前项目运行稳定，可按计划推进。\n"

        return report


    if st.button("📄 生成项目周报"):
        report = generate_report(df_filtered, delayed_tasks, risk_tasks)
        st.text_area("📊 周报内容", report, height=300)