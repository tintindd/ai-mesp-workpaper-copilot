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

with step_col:
    st.markdown(
        """
        <div class="step-card">
          <div class="step-item active">
            <div class="step-no">1</div>
            <div><strong>选择场景</strong><span>成本方法、程序和审计期间</span></div>
          </div>
          <div class="step-item">
            <div class="step-no">2</div>
            <div><strong>输入参数</strong><span>样本文件与命名规则</span></div>
          </div>
          <div class="step-item">
            <div class="step-no">3</div>
            <div><strong>上传证据</strong><span>CO03、KSBT、3611、CKM3</span></div>
          </div>
          <div class="step-item">
            <div class="step-no">4</div>
            <div><strong>复核结果</strong><span>异常、映射和追溯链</span></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with work_col:
    with st.container(border=True):
        st.markdown(
            """
            <div class="main-title">
              <h2>选择测试场景</h2>
              <p>先选择成本方法，再选择对应的 MESP 测试程序。系统会据此判断需要检查的支持文件和字段。</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        method_col, period_col = st.columns([1.1, 1.1], gap="large")
        with method_col:
            st.selectbox("成本方法", ["标准成本法"], index=0)
        with period_col:
            period = st.text_input("审计期间", value="2025.01.01-2025.12.31")

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

        program = st.radio(
            "测试程序",
            ["SPD03012", "SPD03014", "SPD03015"],
            horizontal=True,
            label_visibility="collapsed",
        )

        st.markdown("#### 上传证据")
        uploaded_files = st.file_uploader(
            "上传支持文件或 zip 包",
            type=["xlsx", "xlsm", "csv", "png", "jpg", "jpeg", "pdf", "zip"],
            accept_multiple_files=True,
        )

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
                summary = result.get("summary", {})
                selected_spd_bytes = (
                    build_spd03015_bytes(spp_dir, program=program, period=period)
                    if summary.get("sample_count", 0) > 0 and summary.get("recognized_file_count", 0) > 0
                    else b""
                )

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

st.markdown("</div>", unsafe_allow_html=True)
