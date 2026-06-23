from __future__ import annotations

import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import streamlit as st


PROJECT_ROOT = Path(__file__).resolve().parent
SCRIPTS_ROOT = PROJECT_ROOT / "scripts"
if str(SCRIPTS_ROOT) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_ROOT))

from mesp_automation_engine import analyze_folder  # noqa: E402
from spd03015_exporter import build_spd03015_bytes  # noqa: E402
from supporting_exporter import build_supporting_bytes  # noqa: E402
from workpaper_exporter import build_workpaper_bytes  # noqa: E402


st.set_page_config(
    page_title="AI-MESP Workpaper Copilot",
    page_icon="📊",
    layout="wide",
)


def save_uploaded_files(uploaded_files, target_dir: Path) -> int:
    saved = 0
    for uploaded_file in uploaded_files:
        name = uploaded_file.name
        if name.lower().endswith(".zip"):
            zip_path = target_dir / name
            zip_path.write_bytes(uploaded_file.getbuffer())
            with zipfile.ZipFile(zip_path) as archive:
                for member in archive.infolist():
                    if member.is_dir():
                        continue
                    member_path = Path(member.filename)
                    if member_path.is_absolute() or ".." in member_path.parts:
                        continue
                    destination = target_dir / member_path
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    with archive.open(member) as source, destination.open("wb") as output:
                        shutil.copyfileobj(source, output)
                    saved += 1
            zip_path.unlink(missing_ok=True)
            continue

        destination = target_dir / Path(name).name
        destination.write_bytes(uploaded_file.getbuffer())
        saved += 1
    return saved


def render_issues(issues: list[dict]) -> None:
    if not issues:
        st.success("未发现异常或待补充事项。")
        return
    st.dataframe(
        [
            {
                "类型": item.get("type"),
                "样本": item.get("sample"),
                "问题": item.get("problem"),
                "建议": item.get("suggestion"),
                "状态": item.get("status"),
            }
            for item in issues
        ],
        use_container_width=True,
        hide_index=True,
    )


def render_workbooks(items: list[dict]) -> None:
    if not items:
        st.info("暂无 SAP 表格映射结果。")
        return
    st.dataframe(
        [
            {
                "样本": item.get("sample"),
                "报表": item.get("report"),
                "订单号": item.get("order") or "-",
                "文件": item.get("file"),
                "Sheet": item.get("sheet"),
                "行数": item.get("row_count"),
                "映射字段数": len(item.get("mapping") or {}),
                "缺失字段数": len(item.get("missing_fields") or []),
            }
            for item in items
        ],
        use_container_width=True,
        hide_index=True,
    )


def render_trace(items: list[dict]) -> None:
    if not items:
        st.info("暂无 Evidence Traceability 结果。")
        return
    st.dataframe(
        [
            {
                "样本": item.get("sample"),
                "报表": item.get("report"),
                "底稿字段": item.get("workpaper_field"),
                "来源文件": item.get("source_file"),
                "来源 Sheet": item.get("source_sheet"),
                "来源字段": item.get("source_header"),
                "来源列": item.get("source_column"),
            }
            for item in items[:300]
        ],
        use_container_width=True,
        hide_index=True,
    )


st.title("AI-MESP Workpaper Copilot")
st.caption("面向 MESP 底稿的智能取数、自动重算与异常复核助手")

with st.sidebar:
    st.header("分析参数")
    period = st.text_input("审计期间", value="2025.01.01-2025.12.31")
    program = st.selectbox(
        "测试程序",
        ["SPD03015", "SPD03012", "SPD03014", "SPD03011", "SPD03013"],
        index=0,
    )
    st.info("建议上传 zip 文件以保留样本文件夹结构，例如 样本1、样本2。")

uploaded_files = st.file_uploader(
    "上传支持文件或 zip 包",
    type=["xlsx", "xlsm", "csv", "png", "jpg", "jpeg", "pdf", "zip"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"已选择 {len(uploaded_files)} 个文件。")

analyze_clicked = st.button("Analyze", type="primary", disabled=not uploaded_files)

if analyze_clicked:
    with tempfile.TemporaryDirectory(prefix="mesp_streamlit_") as temp:
        temp_dir = Path(temp)
        saved_count = save_uploaded_files(uploaded_files, temp_dir)
        if saved_count == 0:
            st.error("没有可分析的文件。")
            st.stop()

        with st.spinner("正在识别 SAP 支持文件并生成复核结果..."):
            result = analyze_folder(temp_dir, period=period, program=program)
            supporting_bytes = build_supporting_bytes(result, temp_dir)
            spp_dir = temp_dir / "_generated_spp"
            spp_dir.mkdir(parents=True, exist_ok=True)
            (spp_dir / "AI-MESP_SPP_Supporting.xlsx").write_bytes(supporting_bytes)
            spd03015_bytes = build_spd03015_bytes(spp_dir)

    summary = result.get("summary", {})
    cols = st.columns(5)
    cols[0].metric("样本数", summary.get("sample_count", 0))
    cols[1].metric("识别文件", summary.get("recognized_file_count", 0))
    cols[2].metric("缺失项", summary.get("missing_file_count", 0))
    cols[3].metric("表格数", summary.get("workbook_count", 0))
    cols[4].metric("追溯项", len(result.get("evidence_trace") or []))

    tab_issues, tab_workbooks, tab_trace, tab_json = st.tabs(
        ["异常与追问", "SAP 表格映射", "Evidence Traceability", "JSON"]
    )

    with tab_issues:
        render_issues(result.get("issues") or [])

    with tab_workbooks:
        render_workbooks(result.get("workbook_results") or result.get("co03_results") or [])

    with tab_trace:
        render_trace(result.get("evidence_trace") or [])

    with tab_json:
        workbook_bytes = build_workpaper_bytes(result)
        st.download_button(
            "下载 SPP Supporting Excel",
            data=supporting_bytes,
            file_name="AI-MESP_SPP_Supporting.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary",
        )
        st.download_button(
            "下载 SPD03012+SPD03014+SPD03015_IRM(SAP)",
            data=spd03015_bytes,
            file_name="AI-MESP_SPD03012_SPD03014_SPD03015_IRM(SAP).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary",
        )
        st.download_button(
            "下载 Excel 底稿",
            data=workbook_bytes,
            file_name="AI-MESP_Workpaper_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.download_button(
            "下载 mesp_automation_result.json",
            data=json.dumps(result, ensure_ascii=False, indent=2, default=str).encode("utf-8"),
            file_name="mesp_automation_result.json",
            mime="application/json",
        )
        st.json(result)
else:
    st.info("请上传 CO03、KSBT、3611、CKM3 支持文件。若包含多个样本，推荐上传 zip 包。")
