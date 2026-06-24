from __future__ import annotations

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
        ["SPD03012", "SPD03014", "SPD03015"],
        index=0,
    )
    st.info("建议上传 zip 包以保留样本文件夹结构；单样本也可以直接上传该样本的全部支持文件。")

uploaded_files = st.file_uploader(
    "上传支持文件或 zip 包",
    type=["xlsx", "xlsm", "csv", "png", "jpg", "jpeg", "pdf", "zip"],
    accept_multiple_files=True,
)

st.info(
    "命名要求：zip 包内建议按 `样本1/`、`样本2/` 建文件夹；每个样本至少包含 CO03、KSBT、3611、CKM3 表格文件，可同时包含对应截图。"
    "\n\n"
    "文件名需包含样本号、订单编号和报表类型，例如："
    "\n"
    "`样本3/3.订单编号11001846-CO03-表格.xlsx`、"
    "`样本3/3.订单编号11001846-CO03-截图.png`、"
    "`样本3/3.订单编号11001846-KSBT-表格.xlsx`、"
    "`样本3/3.订单编号11001846-3611-表格.xlsx`、"
    "`样本3/3.订单编号11001846-物料ID-13012857-CKM3-表格.xlsx`。"
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
            selected_spd_bytes = build_spd03015_bytes(spp_dir, program=program, period=period)

    st.session_state["analysis_bundle"] = {
        "result": result,
        "program": program,
        "supporting_bytes": supporting_bytes,
        "selected_spd_bytes": selected_spd_bytes,
    }

bundle = st.session_state.get("analysis_bundle")

if bundle:
    result = bundle["result"]
    selected_program = bundle["program"]
    supporting_bytes = bundle["supporting_bytes"]
    selected_spd_bytes = bundle["selected_spd_bytes"]

    summary = result.get("summary", {})
    cols = st.columns(5)
    cols[0].metric("样本数", summary.get("sample_count", 0))
    cols[1].metric("识别文件", summary.get("recognized_file_count", 0))
    cols[2].metric("缺失项", summary.get("missing_file_count", 0))
    cols[3].metric("表格数", summary.get("workbook_count", 0))
    cols[4].metric("追溯项", len(result.get("evidence_trace") or []))

    tab_issues, tab_workbooks, tab_trace, tab_download = st.tabs(
        ["异常与追问", "SAP 表格映射", "Evidence Traceability", "底稿结果下载"]
    )

    with tab_issues:
        render_issues(result.get("issues") or [])

    with tab_workbooks:
        render_workbooks(result.get("workbook_results") or result.get("co03_results") or [])

    with tab_trace:
        render_trace(result.get("evidence_trace") or [])

    with tab_download:
        st.download_button(
            "下载 SPP Supporting Excel",
            data=supporting_bytes,
            file_name="AI-MESP_SPP_Supporting.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary",
        )
        st.download_button(
            f"下载 {selected_program}_IRM(SAP)",
            data=selected_spd_bytes,
            file_name=f"AI-MESP_{selected_program}_IRM(SAP).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
else:
    st.info("请上传符合命名要求的 CO03、KSBT、3611、CKM3 支持文件或 zip 包。")
