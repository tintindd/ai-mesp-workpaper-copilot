const fileInput = document.getElementById("files");
const analyzeBtn = document.getElementById("analyzeBtn");
const clearBtn = document.getElementById("clearBtn");
const statusText = document.getElementById("status");

const reports = ["CO03", "KSBT", "3611", "CKM3"];

function setStatus(message) {
  statusText.textContent = message;
}

function classifyPath(path) {
  const upper = path.toUpperCase();
  const report = reports.find((item) => upper.includes(item)) || "未知报表";
  const kind = /表格|导出|EXPORT|EXCEL|XLS|XLSX|CSV/.test(upper)
    ? "表格"
    : /截图|SCREEN|IMAGE|PNG|JPG|JPEG|PDF/.test(upper)
      ? "截图"
      : "未知类型";
  const sampleMatch = path.match(/样本\s*(\d+)/i) || path.match(/sample\s*(\d+)/i) || path.split(/[\\/]/).pop().match(/^(\d+)[.\-_ ]/);
  const sample = sampleMatch ? `样本 ${Number(sampleMatch[1])}` : "未识别样本";
  return { sample, report, kind };
}

function renderFilePreview() {
  const files = Array.from(fileInput.files || []);
  const preview = document.getElementById("filePreview");
  preview.innerHTML = "";

  if (!files.length) {
    preview.innerHTML = '<p class="hint">尚未选择文件。</p>';
    return;
  }

  const groups = new Map();
  files.forEach((file) => {
    const path = file.webkitRelativePath || file.name;
    const meta = classifyPath(path);
    const key = meta.sample;
    const group = groups.get(key) || { total: 0, evidence: new Set() };
    group.total += 1;
    if (meta.report !== "未知报表" && meta.kind !== "未知类型") {
      group.evidence.add(`${meta.report}-${meta.kind}`);
    }
    groups.set(key, group);
  });

  [...groups.entries()].sort().forEach(([sample, group]) => {
    const row = document.createElement("div");
    row.className = "file-row";
    row.innerHTML = `<span>${sample}: ${group.total} 个文件</span><span>${group.evidence.size}/8 项证据</span>`;
    preview.appendChild(row);
  });
}

function formatAmount(value) {
  if (typeof value !== "number") return "-";
  return value.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function renderTotals(totals) {
  const entries = Object.entries(totals || {});
  if (!entries.length) return "-";
  return entries.slice(0, 4).map(([key, value]) => `${key}: ${formatAmount(value)}`).join("<br>");
}

function statusLabel(status) {
  if (status === "missing") return '<span class="status missing">缺失</span>';
  if (status === "warn") return '<span class="status warn">待确认</span>';
  return '<span class="status ok">已匹配</span>';
}

function renderIssues(items) {
  const body = document.getElementById("issueBody");
  if (!items.length) {
    body.innerHTML = '<tr><td colspan="5">暂无异常。</td></tr>';
    return;
  }
  body.innerHTML = items.map((item) => `
    <tr>
      <td>${item.type}</td>
      <td>${item.sample}</td>
      <td>${item.problem}</td>
      <td>${item.suggestion}</td>
      <td>${statusLabel(item.status)}</td>
    </tr>
  `).join("");
}

function renderWorkbooks(items) {
  const body = document.getElementById("workbookBody");
  if (!items.length) {
    body.innerHTML = '<tr><td colspan="7">暂无表格映射结果。</td></tr>';
    return;
  }
  body.innerHTML = items.map((item) => `
    <tr>
      <td>${item.sample}</td>
      <td>${item.report}</td>
      <td>${item.order || "-"}</td>
      <td>${item.file}</td>
      <td>${item.row_count ?? "-"}</td>
      <td>${Object.keys(item.mapping || {}).length} 个字段</td>
      <td>${renderTotals(item.totals)}</td>
    </tr>
  `).join("");
}

function renderTrace(items) {
  const body = document.getElementById("traceBody");
  if (!items.length) {
    body.innerHTML = '<tr><td colspan="6">暂无证据追溯结果。</td></tr>';
    return;
  }
  body.innerHTML = items.slice(0, 180).map((item) => `
    <tr>
      <td>${item.sample}</td>
      <td>${item.report}</td>
      <td>${item.workpaper_field}</td>
      <td>${item.source_file}</td>
      <td>${item.source_sheet}</td>
      <td>${item.source_header}</td>
    </tr>
  `).join("");
}

function renderResult(data) {
  const summary = data.summary || {};
  document.getElementById("sampleMetric").textContent = summary.sample_count ?? 0;
  document.getElementById("fileMetric").textContent = summary.recognized_file_count ?? 0;
  document.getElementById("missingMetric").textContent = summary.missing_file_count ?? 0;
  document.getElementById("traceMetric").textContent = Array.isArray(data.evidence_trace) ? data.evidence_trace.length : 0;
  renderIssues(data.issues || []);
  renderWorkbooks(data.workbook_results || data.co03_results || []);
  renderTrace(data.evidence_trace || []);
}

async function analyze() {
  const files = Array.from(fileInput.files || []);
  if (!files.length) {
    alert("请先选择支持文件。");
    return;
  }

  const period = document.getElementById("period").value.trim();
  const program = document.getElementById("program").value;
  const formData = new FormData();

  files.forEach((file) => {
    const relativePath = file.webkitRelativePath || file.name;
    formData.append("files", file, file.name);
    formData.append("relative_paths", relativePath);
  });

  analyzeBtn.disabled = true;
  setStatus("正在上传并分析...");

  try {
    const query = new URLSearchParams({ period, program });
    const response = await fetch(`/api/analyze?${query.toString()}`, {
      method: "POST",
      body: formData,
    });
    const data = await response.json();
    if (!response.ok || data.error) {
      throw new Error(data.error || `HTTP ${response.status}`);
    }
    renderResult(data);
    setStatus(`分析完成：识别 ${data.summary?.recognized_file_count ?? 0} 个文件。`);
  } catch (error) {
    console.error(error);
    setStatus(`分析失败：${error.message}`);
  } finally {
    analyzeBtn.disabled = false;
  }
}

function clearAll() {
  fileInput.value = "";
  renderFilePreview();
  renderResult({ summary: {}, issues: [], workbook_results: [], evidence_trace: [] });
  setStatus("尚未分析。");
}

fileInput.addEventListener("change", renderFilePreview);
analyzeBtn.addEventListener("click", analyze);
clearBtn.addEventListener("click", clearAll);
renderFilePreview();
