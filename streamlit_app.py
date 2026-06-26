from __future__ import annotations

import shutil
import sys
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
import streamlit as st


PROJECT_ROOT = Path(__file__).resolve().parent
SCRIPTS_ROOT = PROJECT_ROOT / "scripts"
if str(SCRIPTS_ROOT) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_ROOT))

from mesp_automation_engine import REQUIRED_FIELDS, analyze_folder, analyze_workbook  # noqa: E402
from deepseek_client import (  # noqa: E402
    clean_ocr_text_with_deepseek,
    load_deepseek_config,
    normalize_filename_with_deepseek,
    test_deepseek_connection,
)
import ocr_client  # noqa: E402
from ckm3_table_builder import build_ckm3_workbook_bytes, extract_ckm3_rows  # noqa: E402
from spd03015_exporter import build_spd03015_bytes  # noqa: E402
from supporting_exporter import build_supporting_bytes  # noqa: E402


is_image_file = ocr_client.is_image_file
load_online_ocr_config = ocr_client.load_online_ocr_config
online_ocr_available = ocr_client.online_ocr_available
recognize_uploaded_image = ocr_client.recognize_uploaded_image
recognize_co03_product_id = getattr(ocr_client, "recognize_co03_product_id", None)
recognize_ckm3_material_id = getattr(ocr_client, "recognize_ckm3_material_id", None)


st.set_page_config(
    page_title="AI-MESP Workpaper Copilot",
    page_icon="📊",
    layout="wide",
)

query_theme = st.query_params.get("theme", st.session_state.get("ui_theme", "light"))
if isinstance(query_theme, list):
    query_theme = query_theme[0] if query_theme else "light"
if query_theme not in {"light", "dark"}:
    query_theme = "light"

st.session_state["ui_theme"] = query_theme
theme_is_dark = query_theme == "dark"
theme_class = "theme-dark" if theme_is_dark else "theme-light"
next_theme = "light" if theme_is_dark else "dark"
theme_label = "主题切换"

st.markdown(
    """
    <style>
    [data-testid="stHeader"] { display: none; }
    [data-testid="stSidebar"] { display: none; }
    .block-container {
        max-width: none;
        padding: 0 0 3rem 0;
    }
    .stApp {
        background: #f4f7fb;
        color: #172033;
    }
    .mesp-hero {
        padding: 2.3rem 4.8rem 2.1rem;
        background: linear-gradient(120deg, #00338d 0%, #005eb8 65%, #00a3e0 100%);
        color: #fff;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 2rem;
        border-bottom: 1px solid rgba(255,255,255,0.14);
    }
    .mesp-hero h1 {
        margin: 0 0 .45rem;
        font-size: 2.1rem;
        line-height: 1.1;
        font-weight: 800;
        letter-spacing: 0;
    }
    .mesp-hero p {
        margin: 0;
        font-size: 1.02rem;
        opacity: .9;
    }
    .mesp-badge {
        border: 1px solid rgba(255,255,255,.38);
        background: rgba(255,255,255,.12);
        border-radius: .55rem;
        padding: .7rem 1rem;
        font-weight: 700;
        white-space: nowrap;
    }
    .mesp-shell {
        width: min(92rem, calc(100vw - 4.5rem));
        margin: 2rem auto 0;
    }
    .step-card {
        background: #fff;
        border: 1px solid #d7e1ee;
        border-radius: .65rem;
        padding: 1.35rem;
        box-shadow: 0 1rem 2rem rgba(20,32,51,.07);
    }
    .step-item {
        display: grid;
        grid-template-columns: 2.2rem 1fr;
        gap: .9rem;
        align-items: start;
        padding: 1rem .75rem;
        border-radius: .5rem;
        margin-bottom: .45rem;
    }
    .step-item.active { background: #edf4ff; color: #00338d; }
    .step-no {
        width: 2rem;
        height: 2rem;
        border-radius: 50%;
        border: 1px solid #6f819d;
        display: grid;
        place-items: center;
        font-weight: 800;
        color: #00338d;
    }
    .step-item strong { display: block; font-size: 1rem; margin-bottom: .25rem; }
    .step-item span { color: #5c6f8c; font-size: .9rem; }
    .main-title h2 {
        margin: .2rem 0 .45rem;
        font-size: 1.65rem;
        font-weight: 850;
        letter-spacing: 0;
    }
    .main-title p {
        color: #587091;
        margin: 0 0 1.3rem;
        line-height: 1.6;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-color: #d7e1ee;
        border-radius: .65rem;
        box-shadow: 0 1rem 2rem rgba(20,32,51,.06);
        background: #fff;
    }
    div[data-testid="stMetric"] {
        background: #fff;
        border: 1px solid #d7e1ee;
        border-radius: .55rem;
        padding: 1rem 1.1rem;
        box-shadow: 0 .6rem 1.4rem rgba(20,32,51,.045);
    }
    div[data-testid="stMetric"] label { color: #587091; }
    div[data-testid="stMetricValue"] { color: #00338d; }
    .program-help {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: .85rem;
        margin: .2rem 0 1rem;
    }
    .program-card {
        border: 1px solid #d7e1ee;
        border-radius: .55rem;
        padding: .9rem 1rem;
        min-height: 6.4rem;
        background: #fff;
    }
    .program-card strong { display: block; margin-bottom: .35rem; }
    .program-code { color: #00338d; font-weight: 850; }
    .pill {
        display: inline-block;
        padding: .22rem .55rem;
        border-radius: .4rem;
        border: 1px solid #cfd9e7;
        background: #eef3fa;
        color: #172033;
        margin: .45rem .25rem 0 0;
        font-size: .82rem;
    }
    .upload-rules {
        border: 1px solid #cfe1f7;
        background: #eef6ff;
        color: #17446f;
        border-radius: .55rem;
        padding: 1rem 1.15rem;
        line-height: 1.7;
        margin: .6rem 0 1rem;
    }
    .upload-rules code {
        background: rgba(255,255,255,.78);
        border: 1px solid #cfd9e7;
        border-radius: .35rem;
        padding: .08rem .32rem;
        color: #00338d;
    }
    .stButton > button, .stDownloadButton > button {
        border-radius: .5rem;
        font-weight: 800;
        min-height: 2.7rem;
    }
    .stButton > button[kind="primary"],
    .stDownloadButton > button[kind="primary"] {
        background: #e7f1ff !important;
        border-color: #9fc5f8 !important;
        color: #00338d !important;
    }
    .stButton > button[kind="primary"]:hover,
    .stDownloadButton > button[kind="primary"]:hover {
        background: #d8eaff !important;
        border-color: #6ea8ee !important;
        color: #002f6c !important;
    }
    .stButton > button[kind="primary"]:focus,
    .stDownloadButton > button[kind="primary"]:focus {
        box-shadow: 0 0 0 .18rem rgba(0, 94, 184, .18) !important;
    }
    [data-testid="stRadio"] [role="radiogroup"] label span:first-child {
        border-color: #7fb3f0 !important;
    }
    [data-testid="stRadio"] [role="radiogroup"] label span:first-child:has(input:checked) {
        background-color: #005eb8 !important;
        border-color: #005eb8 !important;
    }
    .fork-button {
        position: fixed;
        top: .45rem;
        right: .55rem;
        z-index: 9999;
        display: inline-flex;
        align-items: center;
        gap: .42rem;
        padding: .45rem .72rem;
        border: 1px solid rgba(255,255,255,.25);
        border-radius: .2rem;
        background: rgba(8, 12, 20, .72);
        color: #fff !important;
        font-size: .78rem;
        font-weight: 800;
        text-decoration: none !important;
        line-height: 1;
        box-shadow: 0 .25rem .75rem rgba(0,0,0,.22);
        backdrop-filter: blur(6px);
    }
    .fork-button:hover {
        background: rgba(0, 51, 141, .92);
        border-color: rgba(255,255,255,.45);
    }
    .fork-button svg {
        width: .95rem;
        height: .95rem;
        fill: currentColor;
    }
    .theme-button {
        position: fixed;
        top: .45rem;
        right: 5.05rem;
        z-index: 10000;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: .45rem .72rem;
        border: 1px solid rgba(255,255,255,.25);
        border-radius: .2rem;
        background: rgba(8, 12, 20, .72);
        color: #fff !important;
        font-size: .78rem;
        font-weight: 800;
        text-decoration: none !important;
        line-height: 1;
        box-shadow: 0 .25rem .75rem rgba(0,0,0,.22);
        backdrop-filter: blur(6px);
    }
    .theme-button:hover {
        background: rgba(0, 51, 141, .92);
        border-color: rgba(255,255,255,.45);
    }
    .theme-dark {
        background: #05070b;
        color: #f8fafc;
        min-height: 100vh;
    }
    .theme-dark .mesp-hero {
        background: linear-gradient(120deg, #030712 0%, #06142f 54%, #00338d 100%);
        border-bottom-color: rgba(255,255,255,.08);
    }
    .theme-dark .step-card,
    .theme-dark div[data-testid="stVerticalBlockBorderWrapper"],
    .theme-dark .program-card,
    .theme-dark div[data-testid="stMetric"] {
        background: #0b0f19 !important;
        border-color: #20293a !important;
        box-shadow: 0 1rem 2rem rgba(0,0,0,.38) !important;
    }
    .theme-dark .step-item.active {
        background: #0e2444;
        color: #93c5fd;
    }
    .theme-dark .step-no {
        border-color: #4b638a;
        color: #93c5fd;
    }
    .theme-dark .step-item span,
    .theme-dark .main-title p,
    .theme-dark div[data-testid="stMetric"] label {
        color: #9aa8bd !important;
    }
    .theme-dark .program-code,
    .theme-dark div[data-testid="stMetricValue"] {
        color: #60a5fa !important;
    }
    .theme-dark .pill {
        background: #111827;
        border-color: #2a3548;
        color: #dbeafe;
    }
    .theme-dark .upload-rules {
        background: #0b1728;
        border-color: #1e3a5f;
        color: #bfdbfe;
    }
    .theme-dark .upload-rules code {
        background: #0f172a;
        border-color: #2a3548;
        color: #93c5fd;
    }
    .theme-dark input,
    .theme-dark textarea,
    .theme-dark select,
    .theme-dark [data-baseweb="select"] > div,
    .theme-dark [data-testid="stFileUploaderDropzone"] {
        background: #090d15 !important;
        color: #f8fafc !important;
        border-color: #263246 !important;
    }
    .theme-dark [data-testid="stDataFrame"],
    .theme-dark [data-testid="stTable"] {
        background: #0b0f19 !important;
    }
    @media (max-width: 900px) {
        .mesp-hero { padding: 1.6rem 1.4rem; align-items: flex-start; flex-direction: column; }
        .mesp-shell { width: calc(100vw - 1.5rem); margin-top: 1rem; }
        .program-help { grid-template-columns: 1fr; }
        .fork-button { top: .35rem; right: .35rem; }
        .theme-button { top: .35rem; right: 4.85rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if theme_is_dark:
    st.markdown(
        """
        <style>
        .stApp,
        [data-testid="stAppViewContainer"] {
            background: #05070b !important;
            color: #f8fafc !important;
        }
        [data-testid="stHeader"] {
            background: transparent !important;
        }
        .block-container {
            background: #05070b !important;
        }
        .mesp-hero {
            background: linear-gradient(120deg, #020617 0%, #09204d 56%, #005eb8 100%) !important;
            border-bottom: 1px solid rgba(255,255,255,.08) !important;
        }
        .mesp-shell {
            color: #f8fafc !important;
        }
        .step-card,
        div[data-testid="stVerticalBlockBorderWrapper"],
        .program-card,
        div[data-testid="stMetric"] {
            background: #0f172a !important;
            border-color: #243047 !important;
            box-shadow: 0 1rem 2rem rgba(0,0,0,.34) !important;
        }
        .step-item.active {
            background: #102647 !important;
            color: #dbeafe !important;
        }
        .step-no {
            border-color: #5b7aaa !important;
            color: #93c5fd !important;
            background: rgba(96,165,250,.08) !important;
        }
        .step-item strong,
        .main-title h2,
        .program-card strong,
        h1, h2, h3, h4, h5, h6,
        label,
        [data-testid="stMarkdownContainer"] {
            color: #f8fafc !important;
        }
        .step-item span,
        .main-title p,
        div[data-testid="stMetric"] label,
        [data-testid="stWidgetLabel"] p,
        [data-testid="stFileUploader"] small {
            color: #b6c4d8 !important;
        }
        [data-testid="stFileUploader"] small,
        [data-testid="stFileUploaderDropzone"] small,
        [data-testid="stFileUploaderDropzone"] span {
            color: #dbeafe !important;
            opacity: 1 !important;
        }
        .program-code,
        div[data-testid="stMetricValue"] {
            color: #60a5fa !important;
        }
        .pill {
            background: #172033 !important;
            border-color: #334155 !important;
            color: #dbeafe !important;
        }
        .upload-rules {
            background: #0b1f37 !important;
            border-color: #1d4f82 !important;
            color: #d8ecff !important;
        }
        .upload-rules strong {
            color: #ffffff !important;
        }
        .upload-rules code {
            background: #07111f !important;
            border-color: #2d4567 !important;
            color: #93c5fd !important;
        }
        input,
        textarea,
        select,
        [data-baseweb="select"] > div,
        [data-testid="stFileUploaderDropzone"] {
            background: #111827 !important;
            color: #f8fafc !important;
            border-color: #334155 !important;
        }
        [data-baseweb="select"] span,
        [data-baseweb="select"] div,
        [data-testid="stTextInput"] input {
            color: #f8fafc !important;
        }
        [data-testid="stFileUploaderDropzone"] button,
        .stButton > button,
        .stDownloadButton > button {
            background: #172033 !important;
            border-color: #334155 !important;
            color: #f8fafc !important;
        }
        .stButton > button[kind="primary"],
        .stDownloadButton > button[kind="primary"] {
            background: #005eb8 !important;
            border-color: #2b8edb !important;
            color: #ffffff !important;
        }
        [data-testid="stRadio"] label,
        [data-testid="stRadio"] p {
            color: #dbeafe !important;
        }
        [data-testid="stDataFrame"],
        [data-testid="stTable"] {
            background: #0f172a !important;
            color: #f8fafc !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
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


def extract_zip_bytes(zip_bytes: bytes, target_dir: Path) -> int:
    saved = 0
    with zipfile.ZipFile(BytesIO(zip_bytes)) as archive:
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
    return saved


def build_analysis_bundle_from_folder(source_dir: Path, period: str, program: str) -> dict:
    result = analyze_folder(source_dir, period=period, program=program)
    supporting_bytes = build_supporting_bytes(result, source_dir)
    spp_dir = source_dir / "_generated_spp"
    spp_dir.mkdir(parents=True, exist_ok=True)
    (spp_dir / "AI-MESP_SPP_Supporting.xlsx").write_bytes(supporting_bytes)
    summary = result.get("summary", {})
    selected_spd_bytes = (
        build_spd03015_bytes(source_dir, program=program, period=period)
        if summary.get("sample_count", 0) > 0 and summary.get("recognized_file_count", 0) > 0
        else b""
    )
    return {
        "result": result,
        "program": program,
        "supporting_bytes": supporting_bytes,
        "selected_spd_bytes": selected_spd_bytes,
    }


def run_analysis_from_cleaned_zip(zip_bytes: bytes | None, period: str, program: str) -> None:
    if not zip_bytes:
        st.session_state["filename_cleanup_error"] = "请先运行文件名清洗并生成标准命名文件包。"
        return
    try:
        with tempfile.TemporaryDirectory(prefix="mesp_cleaned_zip_") as temp:
            temp_dir = Path(temp)
            saved_count = extract_zip_bytes(zip_bytes, temp_dir)
            if saved_count == 0:
                st.session_state["filename_cleanup_error"] = "标准命名文件包为空，无法计算。"
                return
            st.session_state["analysis_bundle"] = build_analysis_bundle_from_folder(temp_dir, period, program)
            st.session_state["current_step"] = 4
    except Exception as exc:
        st.session_state["filename_cleanup_error"] = str(exc)


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


def mark_scenario_started() -> None:
    st.session_state["scenario_started"] = True
    st.session_state["params_confirmed"] = False
    st.session_state["ckm3_ocr_confirmed"] = False
    st.session_state["filename_cleanup_confirmed"] = False
    st.session_state["has_uploaded_files"] = False
    st.session_state.pop("analysis_bundle", None)


def go_to_step(step: int) -> None:
    if step <= get_unlocked_step():
        st.session_state["current_step"] = step


def go_next_from_scenario() -> None:
    st.session_state["scenario_started"] = True
    st.session_state["current_step"] = 2


PROGRAM_OPTIONS = {
    "SPD03012 - 标准成本法固定加工成本": "SPD03012",
    "SPD03014 - 标准成本法可变加工成本": "SPD03014",
    "SPD03015 - 存货成本差异分摊": "SPD03015",
}


def sync_test_program_from_label() -> None:
    st.session_state["test_program"] = PROGRAM_OPTIONS.get(
        st.session_state.get("test_program_label"),
        "SPD03012",
    )
    mark_scenario_started()


def go_next_from_params() -> None:
    st.session_state["scenario_started"] = True
    st.session_state["params_confirmed"] = True
    st.session_state["current_step"] = 3


def get_unlocked_step() -> int:
    if st.session_state.get("analysis_bundle"):
        return 4
    if st.session_state.get("params_confirmed"):
        return 3
    if st.session_state.get("scenario_started"):
        return 2
    return 1


def get_active_step() -> int:
    if "current_step" not in st.session_state:
        st.session_state["current_step"] = 1
    unlocked_step = get_unlocked_step()
    if st.session_state["current_step"] > unlocked_step:
        st.session_state["current_step"] = unlocked_step
    return st.session_state["current_step"]


def step_class(step: int, active_step: int) -> str:
    return "step-item active" if step == active_step else "step-item"


def run_deepseek_connection_test() -> None:
    if not deepseek_config:
        st.session_state["deepseek_status"] = {
            "ok": False,
            "message": "DeepSeek API Key 未配置。",
        }
        return
    try:
        response = test_deepseek_connection(deepseek_config)
        st.session_state["deepseek_status"] = {
            "ok": True,
            "message": f"连接成功：{response.strip()}",
        }
    except Exception as exc:
        st.session_state["deepseek_status"] = {
            "ok": False,
            "message": str(exc),
        }


def display_missing_fields(report: str, missing_fields: list[str]) -> list[str]:
    aliases = REQUIRED_FIELDS.get(report, {})
    return [(aliases.get(field) or [field])[0] for field in missing_fields]


def run_filename_cleanup(uploaded_files) -> None:
    if not deepseek_config:
        st.session_state["filename_cleanup_results"] = []
        st.session_state["filename_cleanup_error"] = "DeepSeek API Key 未配置。"
        return

    results = []
    try:
        for uploaded_file in uploaded_files or []:
            results.append(normalize_filename_with_deepseek(uploaded_file.name, deepseek_config))
        st.session_state["filename_cleanup_results"] = results
        st.session_state.pop("filename_cleanup_error", None)
    except Exception as exc:
        st.session_state["filename_cleanup_error"] = str(exc)


def _cleanup_evidence_kind(filename: str) -> str:
    extension = Path(filename).suffix.lower().lstrip(".")
    if extension in {"xlsx", "xlsm", "xls", "csv"}:
        return "表格"
    if extension in {"png", "jpg", "jpeg", "pdf"}:
        return "截图"
    return "未知"


def _standard_cleanup_filename(
    *,
    sample_no: str,
    order_id: str,
    material_id: str,
    product_id: str = "",
    report: str,
    evidence_kind: str,
    extension: str,
) -> str:
    sample_text = str(sample_no or "1").strip()
    order_text = str(order_id or "待补充").strip()
    extension_text = extension.lstrip(".").lower()
    if report == "CKM3":
        material_text = str(material_id or "待补充").strip()
        return (
            f"样本{sample_text}/{sample_text}.订单编号{order_text}-"
            f"物料ID-{material_text}-CKM3-{evidence_kind}.{extension_text}"
        )
    if report == "CO03" and product_id:
        product_text = str(product_id).strip()
        return (
            f"样本{sample_text}/{sample_text}.订单编号{order_text}-"
            f"物料编码-{product_text}-CO03-{evidence_kind}.{extension_text}"
        )
    return f"样本{sample_text}/{sample_text}.订单编号{order_text}-{report}-{evidence_kind}.{extension_text}"


def _deduplicate_zip_name(name: str, used_names: set[str]) -> str:
    if name not in used_names:
        used_names.add(name)
        return name
    path = Path(name)
    folder = path.parent.as_posix()
    stem = path.stem
    suffix = path.suffix
    counter = 2
    while True:
        candidate_name = f"{stem}-{counter}{suffix}"
        candidate = f"{folder}/{candidate_name}" if folder != "." else candidate_name
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        counter += 1


def build_standard_named_zip_bytes(cleanup_results: list[dict]) -> bytes:
    output = BytesIO()
    used_names: set[str] = set()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for item in cleanup_results:
            file_bytes = item.get("file_bytes")
            standard_filename = item.get("standard_filename")
            if not file_bytes or not standard_filename:
                continue
            zip_name = _deduplicate_zip_name(str(standard_filename), used_names)
            archive.writestr(zip_name, file_bytes)
    return output.getvalue()


def run_filename_cleanup_by_bucket(
    sample_no: str,
    order_id: str,
    bucket_files: dict[str, list],
) -> None:
    st.session_state["filename_cleanup_results"] = []
    st.session_state.pop("standard_named_zip_bytes", None)
    st.session_state.pop("filename_cleanup_error", None)

    results = []
    try:
        with tempfile.TemporaryDirectory(prefix="mesp_filename_check_") as temp:
            temp_dir = Path(temp)
            detected_co03_product_id = ""
            detected_ckm3_material_id = ""
            if online_ocr_available(online_ocr_config):
                for uploaded_file in bucket_files.get("CO03") or []:
                    if is_image_file(uploaded_file.name):
                        if recognize_co03_product_id:
                            detected_co03_product_id = recognize_co03_product_id(
                                uploaded_file.name,
                                uploaded_file.getvalue(),
                                online_ocr_config,
                            )
                        if detected_co03_product_id:
                            break
                for uploaded_file in bucket_files.get("CKM3") or []:
                    if is_image_file(uploaded_file.name):
                        if recognize_ckm3_material_id:
                            detected_ckm3_material_id = recognize_ckm3_material_id(
                                uploaded_file.name,
                                uploaded_file.getvalue(),
                                online_ocr_config,
                            )
                        if detected_ckm3_material_id:
                            break

            for report, files in bucket_files.items():
                for uploaded_file in files or []:
                    file_bytes = uploaded_file.getvalue()
                    if deepseek_config:
                        cleaned = normalize_filename_with_deepseek(uploaded_file.name, deepseek_config)
                    else:
                        cleaned = {"source_file": uploaded_file.name}

                    extension = Path(uploaded_file.name).suffix.lower().lstrip(".")
                    evidence_kind = _cleanup_evidence_kind(uploaded_file.name)
                    cleaned["上传位置"] = report
                    cleaned["sample_no"] = str(sample_no or cleaned.get("sample_no") or "1").strip()
                    cleaned["report_type"] = report
                    cleaned["order_id"] = order_id.strip() if order_id.strip() else cleaned.get("order_id") or "待补充"
                    if report == "CKM3":
                        detected_material_id = str(cleaned.get("material_id") or detected_ckm3_material_id or "").strip()
                        if (
                            not detected_material_id
                            and recognize_ckm3_material_id
                            and is_image_file(uploaded_file.name)
                            and online_ocr_available(online_ocr_config)
                        ):
                            detected_material_id = recognize_ckm3_material_id(
                                uploaded_file.name,
                                file_bytes,
                                online_ocr_config,
                            )
                        cleaned["material_id"] = detected_material_id or "待补充"
                    if report == "CO03":
                        detected_product_id = str(cleaned.get("product_id") or detected_co03_product_id or "").strip()
                        if (
                            not detected_product_id
                            and recognize_co03_product_id
                            and is_image_file(uploaded_file.name)
                            and online_ocr_available(online_ocr_config)
                        ):
                            detected_product_id = recognize_co03_product_id(
                                uploaded_file.name,
                                file_bytes,
                                online_ocr_config,
                            )
                        cleaned["product_id"] = detected_product_id
                    cleaned["evidence_kind"] = evidence_kind
                    cleaned["extension"] = extension
                    cleaned["standard_filename"] = _standard_cleanup_filename(
                        sample_no=cleaned["sample_no"],
                        order_id=cleaned["order_id"],
                        material_id=cleaned.get("material_id") or "",
                        product_id=cleaned.get("product_id") or "",
                        report=report,
                        evidence_kind=evidence_kind,
                        extension=extension,
                    )
                    cleaned["file_bytes"] = file_bytes

                    field_status = "非表格文件，未检查字段"
                    missing_fields = []
                    if uploaded_file.name.lower().endswith((".xlsx", ".xlsm", ".csv")):
                        safe_name = Path(uploaded_file.name).name
                        temp_path = temp_dir / safe_name
                        temp_path.write_bytes(file_bytes)
                        report_for_check = report if report in {"CO03", "KSBT", "3611", "CKM3"} else cleaned["report_type"]
                        workbook_result = analyze_workbook(temp_path, report_for_check)
                        missing_fields = workbook_result.get("missing_fields") or []
                        field_status = "完整" if not missing_fields else "缺失字段"

                    cleaned["field_status"] = field_status
                    cleaned["missing_fields"] = missing_fields
                    cleaned["missing_field_labels"] = display_missing_fields(cleaned["report_type"], missing_fields)
                    cleaned["source"] = "DeepSeek 文件名清洗 + Excel 字段检查"
                    results.append(cleaned)
        st.session_state["filename_cleanup_results"] = results
        st.session_state["standard_named_zip_bytes"] = build_standard_named_zip_bytes(results)
    except Exception as exc:
        st.session_state["filename_cleanup_error"] = str(exc)


def run_ckm3_ocr_to_excel(uploaded_files) -> None:
    st.session_state["ckm3_ocr_excels"] = []
    st.session_state.pop("ckm3_ocr_error", None)
    if not online_ocr_available(online_ocr_config):
        st.session_state["ckm3_ocr_error"] = "Qwen-OCR 服务未配置，无法识别 CKM3 截图。"
        return

    image_files = [item for item in uploaded_files or [] if is_image_file(item.name)]
    if not image_files:
        st.session_state["ckm3_ocr_error"] = "请上传 CKM3 截图文件（PNG/JPG/JPEG/PDF）。"
        return

    results = []
    try:
        for uploaded_file in image_files:
            uploaded_file.seek(0)
            ocr_result = recognize_uploaded_image(uploaded_file, online_ocr_config)
            ocr_text = ocr_result.get("text") or ""
            rows = extract_ckm3_rows(ocr_text)
            if not rows:
                results.append(
                    {
                        "source_file": uploaded_file.name,
                        "status": "未识别到 CKM3 表格行",
                        "row_count": 0,
                        "ocr_text": ocr_text,
                    }
                )
                continue
            output_name = f"{Path(uploaded_file.name).stem}-CKM3-表格.xlsx"
            results.append(
                {
                    "source_file": uploaded_file.name,
                    "status": "已生成",
                    "file_name": output_name,
                    "row_count": len(rows),
                    "bytes": build_ckm3_workbook_bytes(ocr_text),
                    "ocr_text": ocr_text,
                }
            )
        st.session_state["ckm3_ocr_excels"] = results
    except Exception as exc:
        st.session_state["ckm3_ocr_error"] = str(exc)


def run_ocr_cleanup(uploaded_files) -> None:
    st.session_state["ocr_cleanup_results"] = []
    st.session_state.pop("ocr_cleanup_error", None)
    if not online_ocr_available(online_ocr_config):
        st.session_state["ocr_cleanup_error"] = "Qwen-OCR 服务未配置，无法运行图片 OCR。"
        return
    if not deepseek_config:
        st.session_state["ocr_cleanup_error"] = "DeepSeek API Key 未配置，无法清洗 OCR 文本。"
        return

    image_files = [item for item in uploaded_files or [] if is_image_file(item.name)]
    if not image_files:
        st.session_state["ocr_cleanup_error"] = "未找到可 OCR 的图片文件。"
        return

    results = []
    try:
        for uploaded_file in image_files:
            ocr_result = recognize_uploaded_image(uploaded_file, online_ocr_config)
            cleaned = clean_ocr_text_with_deepseek(
                uploaded_file.name,
                ocr_result.get("text") or "",
                deepseek_config,
            )
            results.append(
                {
                    "source_file": uploaded_file.name,
                    "ocr_text": ocr_result.get("text") or "",
                    "ocr_line_count": ocr_result.get("line_count", 0),
                    "cleaned": cleaned,
                }
            )
        st.session_state["ocr_cleanup_results"] = results
    except Exception as exc:
        st.session_state["ocr_cleanup_error"] = str(exc)


def run_intelligent_cleanup(uploaded_files) -> None:
    st.session_state.pop("intelligent_cleanup_error", None)
    st.session_state.pop("intelligent_cleanup_excel", None)
    st.session_state["generated_support_excels"] = []
    st.session_state["filename_cleanup_results"] = []
    st.session_state["ocr_cleanup_results"] = []

    if not deepseek_config:
        st.session_state["intelligent_cleanup_error"] = "DeepSeek API Key 未配置，无法执行智能清洗。"
        return

    filename_results = []
    ocr_results = []
    try:
        for uploaded_file in uploaded_files or []:
            filename_results.append(normalize_filename_with_deepseek(uploaded_file.name, deepseek_config))

        if online_ocr_available(online_ocr_config):
            image_files = [item for item in uploaded_files or [] if is_image_file(item.name)]
            for uploaded_file in image_files:
                ocr_result = recognize_uploaded_image(uploaded_file, online_ocr_config)
                cleaned = clean_ocr_text_with_deepseek(
                    uploaded_file.name,
                    ocr_result.get("text") or "",
                    deepseek_config,
                )
                ocr_results.append(
                    {
                        "source_file": uploaded_file.name,
                        "ocr_text": ocr_result.get("text") or "",
                        "ocr_line_count": ocr_result.get("line_count", 0),
                        "cleaned": cleaned,
                    }
                )
                if (cleaned.get("report_type") or "").upper() == "CKM3":
                    ckm3_rows = extract_ckm3_rows(ocr_result.get("text") or "")
                    if ckm3_rows:
                        standard_filename = cleaned.get("standard_filename") or uploaded_file.name
                        file_name = standard_filename.replace("截图", "表格")
                        if not file_name.lower().endswith(".xlsx"):
                            file_name = f"{Path(file_name).stem}.xlsx"
                        st.session_state["generated_support_excels"].append(
                            {
                                "source_file": uploaded_file.name,
                                "file_name": file_name,
                                "report_type": "CKM3",
                                "row_count": len(ckm3_rows),
                                "bytes": build_ckm3_workbook_bytes(ocr_result.get("text") or ""),
                            }
                        )

        st.session_state["filename_cleanup_results"] = filename_results
        st.session_state["ocr_cleanup_results"] = ocr_results
        st.session_state["intelligent_cleanup_excel"] = build_cleanup_excel_bytes(filename_results, ocr_results)
    except Exception as exc:
        st.session_state["intelligent_cleanup_error"] = str(exc)


def build_cleanup_excel_bytes(filename_results: list[dict], ocr_results: list[dict]) -> bytes:
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "识别清洗结果"
    headers = [
        "来源",
        "原始文件",
        "样本号",
        "订单编号",
        "物料编码",
        "物料ID",
        "报表类型",
        "证据类型",
        "建议标准文件名",
        "置信度",
        "OCR行数",
        "备注",
    ]
    _write_header(summary_sheet, headers)
    row_index = 2
    for item in filename_results:
        _write_cleanup_row(summary_sheet, row_index, "DeepSeek 文件名清洗", item)
        row_index += 1
    for item in ocr_results:
        cleaned = item.get("cleaned") or {}
        _write_cleanup_row(
            summary_sheet,
            row_index,
            "Qwen-OCR + DeepSeek",
            cleaned,
            source_file=item.get("source_file"),
            ocr_line_count=item.get("ocr_line_count"),
        )
        row_index += 1
    summary_sheet.freeze_panes = "A2"
    for column in summary_sheet.columns:
        summary_sheet.column_dimensions[column[0].column_letter].width = min(
            max(len(str(cell.value or "")) for cell in column) + 2,
            42,
        )

    ocr_sheet = workbook.create_sheet("OCR原文")
    _write_header(ocr_sheet, ["原始文件", "OCR原文"])
    for row_index, item in enumerate(ocr_results, start=2):
        ocr_sheet.cell(row=row_index, column=1, value=item.get("source_file"))
        ocr_sheet.cell(row=row_index, column=2, value=item.get("ocr_text"))
    ocr_sheet.column_dimensions["A"].width = 42
    ocr_sheet.column_dimensions["B"].width = 100
    ocr_sheet.freeze_panes = "A2"

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def build_filename_cleanup_excel_bytes(cleanup_results: list[dict]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "标准文件名清洗结果"
    headers = [
        "上传位置",
        "原始文件",
        "样本号",
        "订单编号",
        "物料编码",
        "物料ID",
        "识别类型",
        "证据类型",
        "建议标准文件名",
        "字段检查",
        "缺失字段",
    ]
    _write_header(sheet, headers)
    for row_index, item in enumerate(cleanup_results, start=2):
        values = [
            item.get("上传位置"),
            item.get("source_file"),
            item.get("sample_no"),
            item.get("order_id"),
            item.get("product_id"),
            item.get("material_id"),
            item.get("report_type"),
            item.get("evidence_kind"),
            item.get("standard_filename"),
            item.get("field_status"),
            "、".join(item.get("missing_field_labels") or []),
        ]
        for column_index, value in enumerate(values, start=1):
            sheet.cell(row=row_index, column=column_index, value=value)
    sheet.freeze_panes = "A2"
    for column in sheet.columns:
        sheet.column_dimensions[column[0].column_letter].width = min(
            max(len(str(cell.value or "")) for cell in column) + 2,
            52,
        )
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def _excel_preview_rows(file_bytes: bytes, max_rows: int = 30) -> list[dict]:
    workbook = load_workbook(BytesIO(file_bytes), data_only=True, read_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    rows = []
    for row_index, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        values = list(row)
        if not any(value not in (None, "") for value in values):
            continue
        rows.append({f"列{index + 1}": value for index, value in enumerate(values)})
        if len(rows) >= max_rows:
            break
    workbook.close()
    return rows


def render_cleanup_file_preview(item: dict) -> None:
    file_name = item.get("standard_filename") or item.get("source_file") or "文件预览"
    file_bytes = item.get("file_bytes") or b""
    suffix = Path(str(file_name)).suffix.lower()
    st.caption(item.get("source_file") or "")
    if suffix in {".png", ".jpg", ".jpeg"}:
        st.image(file_bytes, caption=file_name, use_container_width=True)
    elif suffix in {".xlsx", ".xlsm"}:
        rows = _excel_preview_rows(file_bytes)
        if rows:
            st.dataframe(rows, use_container_width=True, hide_index=True)
        else:
            st.info("该 Excel 文件未读取到可预览数据。")
    elif suffix == ".csv":
        text = file_bytes.decode("utf-8-sig", errors="replace")
        st.text("\n".join(text.splitlines()[:40]))
    elif suffix == ".pdf":
        st.info("PDF 文件暂不支持在线预览，请下载标准命名文件包后查看。")
    else:
        st.info("该文件类型暂不支持在线预览。")


if hasattr(st, "dialog"):
    @st.dialog("文件预览")
    def show_cleanup_preview_dialog(item: dict) -> None:
        render_cleanup_file_preview(item)
        if st.button("关闭预览"):
            st.session_state.pop("cleanup_preview_index", None)
            st.rerun()


def _write_header(sheet, headers: list[str]) -> None:
    fill = PatternFill("solid", fgColor="00338D")
    font = Font(color="FFFFFF", bold=True)
    for column_index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=column_index, value=header)
        cell.fill = fill
        cell.font = font


def _write_cleanup_row(
    sheet,
    row_index: int,
    source: str,
    item: dict,
    *,
    source_file: str | None = None,
    ocr_line_count: int | None = None,
) -> None:
    values = [
        source,
        source_file or item.get("source_file"),
        item.get("sample_no"),
        item.get("order_id"),
        item.get("product_id"),
        item.get("material_id"),
        item.get("report_type"),
        item.get("evidence_kind"),
        item.get("standard_filename"),
        item.get("confidence"),
        ocr_line_count,
        item.get("notes"),
    ]
    for column_index, value in enumerate(values, start=1):
        sheet.cell(row=row_index, column=column_index, value=value)


deepseek_config = load_deepseek_config(st.secrets)
online_ocr_config = load_online_ocr_config(st.secrets)


st.markdown(
    f"""
    <a class="theme-button" href="?theme={next_theme}" target="_self" rel="noopener">{theme_label}</a>
    <a class="fork-button" href="https://github.com/tintindd/ai-mesp-workpaper-copilot" target="_blank" rel="noopener">
      <span>Fork</span>
      <svg viewBox="0 0 16 16" aria-hidden="true">
        <path d="M8 0C3.58 0 0 3.64 0 8.13c0 3.59 2.29 6.63 5.47 7.7.4.08.55-.18.55-.39 0-.19-.01-.83-.01-1.51-2.01.38-2.53-.5-2.69-.96-.09-.24-.48-.96-.82-1.15-.28-.15-.68-.53-.01-.54.63-.01 1.08.59 1.23.83.72 1.23 1.87.88 2.33.67.07-.53.28-.88.51-1.08-1.78-.2-3.64-.91-3.64-4.03 0-.89.31-1.62.82-2.19-.08-.2-.36-1.04.08-2.16 0 0 .67-.22 2.2.84A7.5 7.5 0 0 1 8 3.89c.68 0 1.36.09 2 .27 1.52-1.06 2.19-.84 2.19-.84.44 1.12.16 1.96.08 2.16.51.57.82 1.3.82 2.19 0 3.13-1.87 3.82-3.65 4.03.29.25.54.75.54 1.52 0 1.1-.01 1.99-.01 2.26 0 .22.15.47.55.39A8.08 8.08 0 0 0 16 8.13C16 3.64 12.42 0 8 0Z"/>
      </svg>
    </a>
    <div class="mesp-hero">
      <div>
        <h1>AI-MESP Workpaper Copilot</h1>
        <p>面向 MESP 底稿的智能取数、自动重算与异常复核助手</p>
      </div>
      <div class="mesp-badge">Local Demo · SAP Evidence Mapping</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="mesp-shell">', unsafe_allow_html=True)
step_col, work_col = st.columns([1.05, 3.75], gap="large")
active_step = get_active_step()
unlocked_step = get_unlocked_step()

with step_col:
    with st.container(border=True):
        nav_items = [
            (1, "1  选择场景", "成本方法、程序和审计期间"),
            (2, "2  输入参数", "样本文件与命名规则"),
            (3, "3  证据处理与上传", "OCR、清洗和计算结果"),
            (4, "4  复核结果", "异常、映射和追溯链"),
        ]
        for step, label, caption in nav_items:
            st.button(
                label,
                key=f"nav_step_{step}",
                type="primary" if step == active_step else "secondary",
                disabled=step > unlocked_step,
                use_container_width=True,
                on_click=go_to_step,
                args=(step,),
            )
            st.caption(caption if step <= unlocked_step else f"{caption}（未解锁）")

with work_col:
    if active_step == 1:
        with st.container(border=True):
            st.markdown(
                """
                <div class="main-title">
                  <h2>选择测试场景</h2>
                  <p>先选择成本方法、审计期间和 MESP 测试程序。系统会据此判断需要检查的支持文件和字段。</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

            method_col, period_col = st.columns([1.1, 1.1], gap="large")
            with method_col:
                st.selectbox(
                    "成本方法",
                    ["标准成本法"],
                    index=0,
                    key="cost_method",
                    on_change=mark_scenario_started,
                )
            with period_col:
                st.text_input(
                    "审计期间",
                    value="2025.01.01-2025.12.31",
                    key="audit_period",
                    on_change=mark_scenario_started,
                )

            st.markdown(
                """
                <div class="program-help">
                  <div class="program-card">
                    <strong>标准成本法 - 固定加工成本</strong>
                    <div class="program-code">SPD03012</div>
                    <span class="pill">CO03</span><span class="pill">KSBT</span><span class="pill">3611</span>
                  </div>
                  <div class="program-card">
                    <strong>标准成本法 - 可变加工成本</strong>
                    <div class="program-code">SPD03014</div>
                    <span class="pill">CO03</span><span class="pill">KSBT</span><span class="pill">3611</span>
                  </div>
                  <div class="program-card">
                    <strong>存货成本差异分摊</strong>
                    <div class="program-code">SPD03015</div>
                    <span class="pill">CKM3</span><span class="pill">3611</span>
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            current_program = st.session_state.get("test_program", "SPD03012")
            default_label = next(
                (label for label, code in PROGRAM_OPTIONS.items() if code == current_program),
                "SPD03012 - 标准成本法固定加工成本",
            )
            st.selectbox(
                "测试程序",
                list(PROGRAM_OPTIONS),
                index=list(PROGRAM_OPTIONS).index(default_label),
                key="test_program_label",
                on_change=sync_test_program_from_label,
            )
            st.session_state["test_program"] = PROGRAM_OPTIONS.get(
                st.session_state.get("test_program_label"),
                current_program,
            )

            st.button("下一步", type="primary", on_click=go_next_from_scenario)

    elif active_step == 2:
        with st.container(border=True):
            st.markdown(
                """
                <div class="main-title">
                  <h2>输入参数与命名规则</h2>
                  <p>确认审计期间、测试程序和样本文件命名要求。确认后才会解锁上传证据。</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

            summary_cols = st.columns(3)
            summary_cols[0].metric("成本方法", st.session_state.get("cost_method", "标准成本法"))
            summary_cols[1].metric("测试程序", st.session_state.get("test_program", "SPD03012"))
            summary_cols[2].metric("审计期间", st.session_state.get("audit_period", "2025.01.01-2025.12.31"))

            st.markdown(
                """
                <div class="upload-rules">
                  <strong>命名要求</strong><br>
                  zip 包内建议按 <code>样本1/</code>、<code>样本2/</code> 建文件夹；每个样本至少包含
                  <code>CO03</code>、<code>KSBT</code>、<code>3611</code>、<code>CKM3</code> 表格文件，可同时包含对应截图。<br>
                  文件名需包含样本号、订单编号和报表类型，例如：
                  <code>样本3/3.订单编号11001846-CO03-表格.xlsx</code>、
                  <code>样本3/3.订单编号11001846-CO03-截图.png</code>、
                  <code>样本3/3.订单编号11001846-KSBT-表格.xlsx</code>、
                  <code>样本3/3.订单编号11001846-3611-表格.xlsx</code>、
                  <code>样本3/3.订单编号11001846-物料ID-13012857-CKM3-表格.xlsx</code>。
                </div>
                """,
                unsafe_allow_html=True,
            )

            back_col, next_col = st.columns([1, 1])
            with back_col:
                st.button("返回上一步", on_click=go_to_step, args=(1,))
            with next_col:
                st.button("确认参数，下一步", type="primary", on_click=go_next_from_params)

    elif active_step == 3:
        with st.container(border=True):
            st.markdown(
                """
                <div class="main-title">
                  <h2>证据处理与上传</h2>
                  <p>在同一个工作区完成可选的 OCR、文件名标准化清洗，以及计算结果生成。若已准备好标准文件，可直接使用计算结果。</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.markdown(
                """
                <div class="upload-rules">
                  <strong>推荐处理顺序</strong><br>
                  1. 若缺少 CKM3 Excel，先在 <code>OCR</code> 中把截图转成表格。<br>
                  2. 若文件命名不规范，在 <code>文件名清洗</code> 中导出标准命名 ZIP。<br>
                  3. 将最终标准文件或 zip 包上传到 <code>计算结果</code>，生成 SPP 和底稿结果。
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.info(f"当前选择的测试程序：{st.session_state.get('test_program', 'SPD03012')}")

            tab_ocr, tab_cleanup, tab_upload = st.tabs(["OCR", "文件名清洗", "计算结果"])

            with tab_ocr:
                ocr_col, config_col = st.columns([1.25, 1])
                with ocr_col:
                    st.info("仅用于 CKM3 截图，识别后生成可替代 CKM3-表格.xlsx 的支持性 Excel。CO03、KSBT、3611 仍优先使用 SAP 导出的 Excel/CSV。")
                with config_col:
                    if online_ocr_available(online_ocr_config):
                        st.success("Qwen-OCR 服务已配置，智谱视觉可作为备用通道")
                    else:
                        st.warning("Qwen-OCR 服务未配置，OCR 暂不可用。")

                ckm3_ocr_files = st.file_uploader(
                    "上传 CKM3 截图（PNG/JPG/JPEG/PDF）",
                    type=["png", "jpg", "jpeg", "pdf"],
                    accept_multiple_files=True,
                    key="ckm3_ocr_files",
                )
                st.button(
                    "仅识别 CKM3 截图并生成 Excel",
                    disabled=not ckm3_ocr_files or not online_ocr_available(online_ocr_config),
                    type="primary",
                    on_click=run_ckm3_ocr_to_excel,
                    args=(ckm3_ocr_files,),
                )

                ckm3_ocr_error = st.session_state.get("ckm3_ocr_error")
                if ckm3_ocr_error:
                    st.error(ckm3_ocr_error)

                ckm3_ocr_excels = st.session_state.get("ckm3_ocr_excels") or []
                if ckm3_ocr_excels:
                    st.dataframe(
                        [
                            {
                                "来源截图": item.get("source_file"),
                                "状态": item.get("status"),
                                "生成文件": item.get("file_name") or "-",
                                "明细行数": item.get("row_count"),
                            }
                            for item in ckm3_ocr_excels
                        ],
                        use_container_width=True,
                        hide_index=True,
                    )
                    for index, item in enumerate(ckm3_ocr_excels, start=1):
                        if item.get("bytes"):
                            st.download_button(
                                f"下载 OCR 生成 Excel #{index}",
                                data=item["bytes"],
                                file_name=Path(item.get("file_name") or f"ckm3_ocr_{index}.xlsx").name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="secondary",
                            )
                    with st.expander("查看 OCR 原文"):
                        for item in ckm3_ocr_excels:
                            st.markdown(f"**{item.get('source_file')}**")
                            st.text(item.get("ocr_text") or "")

            with tab_cleanup:
                config_col, action_col = st.columns([1.2, 1])
                with config_col:
                    if deepseek_config:
                        st.success(f"DeepSeek 已配置：{deepseek_config.model}")
                    else:
                        st.warning("DeepSeek API Key 未配置，将仅按输入信息和上传位置生成标准文件名。")
                        st.caption("在 Streamlit Cloud 的 Secrets 中添加 DEEPSEEK_API_KEY 后，可辅助识别原始文件名。")
                with action_col:
                    st.button(
                        "测试 DeepSeek 连接",
                        disabled=not deepseek_config,
                        on_click=run_deepseek_connection_test,
                    )
                    deepseek_status = st.session_state.get("deepseek_status")
                    if deepseek_status:
                        if deepseek_status.get("ok"):
                            st.caption(deepseek_status.get("message"))
                        else:
                            st.error(deepseek_status.get("message"))

                sample_col, order_col = st.columns(2)
                with sample_col:
                    cleanup_sample_no = st.text_input(
                        "样本编号",
                        value="1",
                        key="cleanup_sample_no",
                        placeholder="例如 1",
                    )
                with order_col:
                    cleanup_order_id = st.text_input(
                        "订单编号",
                        key="cleanup_order_id",
                        placeholder="例如 11000437",
                    )
                st.markdown(
                    """
                    <div class="upload-rules">
                      <strong>文件名清洗说明</strong><br>
                      请按文件所属类型上传到对应位置。输出文件包将按标准格式重命名，例如
                      <code>样本1/1.订单编号11000437-CO03-表格.xlsx</code>、
                      <code>样本1/1.订单编号11000437-3611-截图.png</code>、
                      <code>样本1/1.订单编号11000437-物料ID-13014012-CKM3-截图.png</code>。
                      CKM3 截图会自动识别物料ID；若 CO03、KSBT、3611、CKM3 都上传截图和 Excel，将导出 4×2 个标准命名文件。
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                bucket_cols = st.columns(4)
                bucket_files = {}
                for column, report in zip(bucket_cols, ["CO03", "KSBT", "3611", "CKM3"]):
                    with column:
                        bucket_files[report] = st.file_uploader(
                            f"{report} 文件",
                            type=["xlsx", "xlsm", "csv", "png", "jpg", "jpeg", "pdf"],
                            accept_multiple_files=True,
                            key=f"cleanup_{report.lower()}_files",
                        )

                total_cleanup_files = sum(len(files or []) for files in bucket_files.values())
                st.button(
                    "运行文件名清洗",
                    disabled=total_cleanup_files == 0,
                    type="primary",
                    on_click=run_filename_cleanup_by_bucket,
                    args=(cleanup_sample_no, cleanup_order_id, bucket_files),
                )

                cleanup_error = st.session_state.get("filename_cleanup_error")
                if cleanup_error:
                    st.error(cleanup_error)

                cleanup_results = st.session_state.get("filename_cleanup_results") or []
                if cleanup_results:
                    cleanup_table_rows = [
                        {
                            "上传位置": item.get("上传位置"),
                            "原始文件": item.get("source_file"),
                            "样本号": item.get("sample_no") or "-",
                            "订单编号": item.get("order_id") or "-",
                            "物料编码": item.get("product_id") or "-",
                            "物料ID": item.get("material_id") or "-",
                            "识别类型": item.get("report_type") or "-",
                            "证据类型": item.get("evidence_kind") or "-",
                            "建议标准文件名": item.get("standard_filename") or "-",
                            "字段检查": item.get("field_status") or "-",
                            "缺失字段": "、".join(item.get("missing_field_labels") or []) or "-",
                        }
                        for item in cleanup_results
                    ]
                    st.dataframe(cleanup_table_rows, use_container_width=True, hide_index=True)

                    st.markdown("#### 标准文件名预览")
                    st.caption("点击标准文件名可预览文件内容。图片会直接展示，Excel 会展示首个 Sheet 的前 30 行。")
                    for index, item in enumerate(cleanup_results):
                        row_cols = st.columns([1.1, 3.2, 1.2, 1])
                        row_cols[0].write(item.get("上传位置") or "-")
                        row_cols[1].button(
                            item.get("standard_filename") or item.get("source_file") or f"文件 {index + 1}",
                            key=f"preview_cleanup_{index}",
                            use_container_width=True,
                            on_click=lambda i=index: st.session_state.update({"cleanup_preview_index": i}),
                        )
                        row_cols[2].write(item.get("evidence_kind") or "-")
                        row_cols[3].write(item.get("field_status") or "-")

                    preview_index = st.session_state.get("cleanup_preview_index")
                    if preview_index is not None and preview_index < len(cleanup_results):
                        preview_item = cleanup_results[preview_index]
                        if hasattr(st, "dialog"):
                            show_cleanup_preview_dialog(preview_item)
                        else:
                            with st.expander("文件预览", expanded=True):
                                if st.button("关闭预览"):
                                    st.session_state.pop("cleanup_preview_index", None)
                                    st.rerun()
                                render_cleanup_file_preview(preview_item)

                    standard_zip_bytes = st.session_state.get("standard_named_zip_bytes")
                    if standard_zip_bytes:
                        st.download_button(
                            "导出标准命名文件包 ZIP",
                            data=standard_zip_bytes,
                            file_name=f"样本{cleanup_sample_no or '1'}_标准命名文件包.zip",
                            mime="application/zip",
                            type="primary",
                        )
                    calculate_choice = st.radio(
                        "是否进行计算",
                        ["否", "是"],
                        horizontal=True,
                        key="calculate_after_cleanup",
                    )
                    if calculate_choice == "是":
                        st.button(
                            "使用清洗后的标准文件直接计算",
                            disabled=not standard_zip_bytes,
                            type="primary",
                            on_click=run_analysis_from_cleaned_zip,
                            args=(
                                standard_zip_bytes,
                                st.session_state.get("audit_period", "2025.01.01-2025.12.31"),
                                st.session_state.get("test_program", "SPD03012"),
                            ),
                        )
                    st.download_button(
                        "下载标准文件名清洗清单 Excel",
                        data=build_filename_cleanup_excel_bytes(cleanup_results),
                        file_name="AI-MESP_标准文件名清洗结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="secondary",
                    )

            with tab_upload:
                uploaded_files = st.file_uploader(
                    "上传最终支持文件或 zip 包并计算",
                    type=["xlsx", "xlsm", "csv", "png", "jpg", "jpeg", "pdf", "zip"],
                    accept_multiple_files=True,
                    key="supporting_files",
                )

                if uploaded_files:
                    st.session_state["has_uploaded_files"] = True
                    st.write(f"已选择 {len(uploaded_files)} 个文件。")

                analyze_clicked = st.button("Analyze", type="primary", disabled=not uploaded_files)

        if analyze_clicked:
            period = st.session_state.get("audit_period", "2025.01.01-2025.12.31")
            program = st.session_state.get("test_program", "SPD03012")
            with tempfile.TemporaryDirectory(prefix="mesp_streamlit_") as temp:
                temp_dir = Path(temp)
                saved_count = save_uploaded_files(uploaded_files, temp_dir)
                if saved_count == 0:
                    st.error("没有可分析的文件。")
                    st.stop()

                with st.spinner("正在识别 SAP 支持文件并生成复核结果..."):
                    st.session_state["analysis_bundle"] = build_analysis_bundle_from_folder(temp_dir, period, program)
            st.session_state["current_step"] = 4
            st.rerun()

    elif active_step == 4:
        bundle = st.session_state.get("analysis_bundle")
        if not bundle:
            st.info("请先上传证据并点击 Analyze，完成后这里会展示复核结果。")
        else:
            result = bundle["result"]
            selected_program = bundle["program"]
            supporting_bytes = bundle["supporting_bytes"]
            selected_spd_bytes = bundle["selected_spd_bytes"]

            summary = result.get("summary", {})
            cols = st.columns(5)
            cols[0].metric("样本数", summary.get("sample_count", 0))
            cols[1].metric("识别文件", summary.get("recognized_file_count", 0))
            cols[2].metric("异常项", summary.get("issue_count", len(result.get("issues") or [])))
            cols[3].metric("表格数", summary.get("workbook_count", 0))
            cols[4].metric("已追溯项", summary.get("trace_count", len(result.get("evidence_trace") or [])))

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
                if summary.get("sample_count", 0) == 0 or summary.get("recognized_file_count", 0) == 0:
                    st.warning("未识别到有效样本，暂不生成底稿结果。请先根据异常提示修正上传文件命名或内容。")
                else:
                    st.info(f"本次生成的测试程序：{selected_program}")
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

st.markdown("</div>", unsafe_allow_html=True)
