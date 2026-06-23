from __future__ import annotations

import re
from dataclasses import dataclass
from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries

from mesp_automation_engine import find_order, find_report, find_sample, numeric


BLUE = "1F4E78"
DARK_BLUE = "17365D"
LIGHT_BLUE = "D9EAF7"
LIGHT_YELLOW = "FFF2CC"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"
BORDER = "B7C9D6"
TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "templates" / "spd03015_irm_sap_template.xlsx"
CKM3_LABEL_START_ROWS = {1: 10, 2: 56, 3: 117, 4: 167, 5: 213}
CKM3_TITLE_ROWS = {1: 1, 2: 49, 3: 99, 4: 148, 5: 197}


@dataclass
class Spd03015Sample:
    sample: int
    order: str
    material_id: str
    period: str
    ckm3_path: Path | None
    beginning_qty: float
    beginning_variance: float
    receipt_qty: float
    receipt_variance: float
    outbound_qty: float
    outbound_variance: float


def _norm(value: Any) -> str:
    return str(value or "").strip()


def _material_id_from_name(name: str) -> str:
    match = re.search(r"物料ID[-_\s]*([A-Za-z0-9]+)", name, flags=re.IGNORECASE)
    return match.group(1) if match else ""


def _safe_div(numerator: float, denominator: float) -> float:
    return numerator / denominator if denominator else 0.0


def _extract_ckm3_amounts(path: Path) -> dict[str, float]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return {}

    headers = [_norm(value) for value in rows[0]]
    index = {header: i for i, header in enumerate(headers) if header}

    def get(row: tuple, header: str) -> Any:
        position = index.get(header)
        if position is None or position >= len(row):
            return None
        return row[position]

    data_rows = rows[1:]

    def find_exact(category: str) -> tuple | None:
        return next((row for row in data_rows if _norm(get(row, "类别")) == category), None)

    def find_level_zero(category: str) -> tuple | None:
        return next(
            (
                row
                for row in data_rows
                if _norm(get(row, "类别")) == category and numeric(get(row, "层级")) == 0
            ),
            None,
        ) or find_exact(category)

    beginning = find_level_zero("期初库存")
    receipt = find_exact("收货")
    consumption = find_exact("消耗")
    ending = find_level_zero("期末库存")

    beginning_qty = numeric(get(beginning, "交易数量")) if beginning else 0.0
    beginning_variance = (
        numeric(get(beginning, "实际值")) - numeric(get(beginning, "初始评估")) if beginning else 0.0
    )

    if abs(beginning_variance) < 0.0000001:
        for row in data_rows:
            category = _norm(get(row, "类别"))
            if category in {"收货", "库存累计"}:
                break
            if "重新评估" in category or "上一期间结算" in category:
                beginning_variance += numeric(get(row, "价格差异"))

    receipt_qty = numeric(get(receipt, "交易数量")) if receipt else 0.0
    receipt_variance = numeric(get(receipt, "价格差异")) if receipt else 0.0

    outbound_row = consumption or ending
    outbound_qty = numeric(get(outbound_row, "交易数量")) if outbound_row else 0.0
    outbound_variance = numeric(get(outbound_row, "价格差异")) if outbound_row else 0.0

    return {
        "beginning_qty": beginning_qty,
        "beginning_variance": beginning_variance,
        "receipt_qty": receipt_qty,
        "receipt_variance": receipt_variance,
        "outbound_qty": outbound_qty,
        "outbound_variance": outbound_variance,
    }


def _discover_samples(source_folder: Path) -> list[Spd03015Sample]:
    by_sample: dict[int, dict[str, Any]] = {}

    for path in sorted(source_folder.rglob("*")):
        if not path.is_file() or path.name.startswith("~$"):
            continue

        sample = find_sample(path.relative_to(source_folder))
        if not sample:
            continue

        info = by_sample.setdefault(sample, {"sample": sample, "order": "", "material_id": "", "ckm3_path": None})
        order = find_order(path)
        material_id = _material_id_from_name(path.name)
        if order and not info["order"]:
            info["order"] = order
        if material_id and not info["material_id"]:
            info["material_id"] = material_id
        if find_report(path.name) == "CKM3" and path.suffix.lower() in {".xlsx", ".xlsm", ".csv"}:
            info["ckm3_path"] = path

    samples: list[Spd03015Sample] = []
    for sample_no in sorted(by_sample):
        info = by_sample[sample_no]
        amounts = _extract_ckm3_amounts(info["ckm3_path"]) if info.get("ckm3_path") else {}
        samples.append(
            Spd03015Sample(
                sample=sample_no,
                order=info.get("order", ""),
                material_id=info.get("material_id", ""),
                period="",
                ckm3_path=info.get("ckm3_path"),
                beginning_qty=amounts.get("beginning_qty", 0.0),
                beginning_variance=amounts.get("beginning_variance", 0.0),
                receipt_qty=amounts.get("receipt_qty", 0.0),
                receipt_variance=amounts.get("receipt_variance", 0.0),
                outbound_qty=amounts.get("outbound_qty", 0.0),
                outbound_variance=amounts.get("outbound_variance", 0.0),
            )
        )

    return samples


def _style_header(sheet, row: int, max_col: int) -> None:
    thin = Side(style="thin", color=BORDER)
    for col in range(1, max_col + 1):
        cell = sheet.cell(row, col)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(name="Arial", bold=True, color=WHITE, size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)


def _style_body(sheet, min_row: int, max_row: int, max_col: int) -> None:
    thin = Side(style="thin", color=BORDER)
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, max_col=max_col):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin, left=thin, right=thin)


def _set_widths(sheet) -> None:
    widths = {
        "A": 6,
        "B": 16,
        "C": 12,
        "D": 13,
        "E": 14,
        "F": 13,
        "G": 16,
        "H": 18,
        "I": 18,
        "J": 38,
        "K": 14,
        "L": 17,
        "M": 13,
        "N": 19,
        "O": 13,
        "P": 24,
        "Q": 24,
        "R": 24,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width


def _merge(sheet, ranges: list[str]) -> None:
    for cell_range in ranges:
        sheet.merge_cells(cell_range)


def _build_spd_sheet(workbook: Workbook, samples: list[Spd03015Sample], ck_rows: dict[int, dict[str, int]]) -> None:
    sheet = workbook.active
    sheet.title = "SPD03015_IRM(SAP)"
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A19"
    _set_widths(sheet)

    _merge(sheet, ["A13:Q13", "A14:Q14", "B16:G16", "H16:I16", "K16:M16", "O16:Q16", "A28:Q28", "A30:R30"])

    sheet["A2"] = "kpmg    Inventory cost variance allocation (standard cost method) - recalculate\n存货成本差异分摊（标准成本法）- 重新计算"
    sheet["A3"] = "(This substantive procedure template is used to test the inventory cost variance allocation) (本实质性程序模板用于测试存货成本差异分摊)"
    sheet["A5"] = "Procedure ID 程序编号"
    sheet["C5"] = "SPD03015"
    sheet["A6"] = "Currency\n币种"
    sheet["C6"] = "Rmb\n人民币"
    sheet["A7"] = "Unit (e.g. 000s, millions)\n单位（例如千、百万）"
    sheet["C7"] = " "
    sheet["A9"] = "Period start date\n期间开始日"
    sheet["A10"] = "Period end date\n期间截止日"
    sheet["A13"] = "The testing steps are the steps listed for the procedure in the workflow screen. Add additional rows and columns as necessary if more testing steps are needed.\n测试步骤是指工作流程工作屏中列出的本程序相关步骤。如需增加测试步骤，请根据需要增加表格的行数和列数。"
    sheet["A14"] = "If the sample size for the substantive procedure is greater than the number of rows in the table below, copy and paste rows rather than inserting blank rows, to ensure that the formulas and pick-list options flow down to all rows.\n如果实质性程序的样本量超过下表行数，请进行复制粘贴（而非插入空白行）以确保公式和列表选项覆盖所有行。"

    sheet["B16"] = "Per production report/sub-ledger\n基于生产报告/明细账"
    sheet["H16"] = "Supporting document\n相关文件"
    sheet["K16"] = "4 = Reperform the calculation"
    sheet["O16"] = "Testing steps\n测试步骤"

    headers = [
        "序号",
        "物料 ID",
        "期间",
        "期初库存",
        "期初差异金额",
        "本期入库",
        "本月新增差异金额",
        "本月出库分摊的差异金额",
        "Working paper reference\n工作底稿索引",
        "[Other RDE, if applicable (please specify)\n其他相关数据要素，如适用（请注明）]",
        "The stage of completion\n完成阶段",
        "重新计算差异分摊率",
        "本月出库",
        "\n重新计算本月出库分摊的差异金额",
        "Variance\n差异",
        "评价期末存货差异计算方法的合理性。",
        "在根据完成阶段（即在制品、制成品）分配人工/间接制造费用的情况下，评估分配的合理性。",
        "3. Vouch the data used in the calculation to relevant documentation.\n将计算中使用的数据核对至相应的支持性文件。",
    ]
    for col, value in enumerate(headers, 1):
        sheet.cell(17, col, value)

    support_row = [
        "",
        "",
        "",
        "A-差异分摊表/CKM3",
        "B-差异分摊表/CKM3",
        "C-差异分摊表/CKM3",
        "D-差异分摊表/CKM3",
        "E-差异分摊表/CKM3",
        "",
        "",
        "",
        "F=（B+D)/(A+C)",
        "G-差异分摊表/CKM3",
        "H=F*G",
        "I=H-E",
        "",
        "",
        "",
    ]
    for col, value in enumerate(support_row, 1):
        sheet.cell(18, col, value)

    _style_header(sheet, 17, 18)
    for col in range(1, 19):
        sheet.cell(18, col).fill = PatternFill("solid", fgColor=LIGHT_YELLOW)
        sheet.cell(18, col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    other_rde = (
        "【<差异分摊报表>表\n"
        "期初库存数量Inventory quantity at the beginning of the period\n"
        "本期入库数量Inventory receipt quantity in the current period\n"
        "期初库存差异Inventory amount variance at the beginning of the period\n"
        "本期入库差异Inventory receipt amount variance in the current period\n"
        "本期出库数量Inventory outbound quantity in the current period\n"
        "本期出库差异Inventory outbound amount variance in the current period\n\n"
        "【BOM】\n审批人\n审批时间"
    )

    for row_offset, sample in enumerate(samples):
        row = 19 + row_offset
        ck = ck_rows[sample.sample]
        sheet.cell(row, 1, sample.sample)
        sheet.cell(row, 2, sample.material_id or "")
        sheet.cell(row, 3, sample.period)
        sheet.cell(row, 4, f"='CKM3'!Z{ck['beginning_qty']}")
        sheet.cell(row, 5, f"='CKM3'!Z{ck['beginning_variance']}")
        sheet.cell(row, 6, f"='CKM3'!Z{ck['receipt_qty']}")
        sheet.cell(row, 7, f"='CKM3'!Z{ck['receipt_variance']}")
        sheet.cell(row, 8, f"='CKM3'!Z{ck['outbound_variance']}")
        sheet.cell(row, 9, "N/A")
        sheet.cell(row, 10, other_rde)
        sheet.cell(row, 12, f"=(E{row}+G{row})/(D{row}+F{row})")
        sheet.cell(row, 13, f"='CKM3'!Z{ck['outbound_qty']}")
        sheet.cell(row, 14, f"=M{row}*L{row}")
        sheet.cell(row, 15, f"=N{row}-H{row}")
        sheet.cell(row, 16, "Reasonable合理")
        sheet.cell(row, 17, "Reasonable合理")
        sheet.cell(row, 18, "Agrees相符")

    last_data_row = 18 + len(samples)
    if samples:
        _style_body(sheet, 18, last_data_row, 18)

    for row in range(19, last_data_row + 1):
        sheet.row_dimensions[row].height = 80
        for col in [4, 5, 6, 7, 8, 12, 13, 14, 15]:
            sheet.cell(row, col).number_format = "#,##0.00"

    sheet["A25"] = "Note: 项目组需根据被审计单位制定的存货差异分配方法修改重新计算的公式。"
    sheet["A27"] = "Note:\n注释："
    sheet["A28"] = "The design of this procedure template assumes that the relevant considerations for the procedure will be appropriately documented on the workflow screen for the procedure, unless they are directly addressed within this procedure template.\n本程序模板的设计基于以下假设：本程序的相关考虑事项将恰当记录于与本程序有关的工作流程工作屏中，除非它们已由本程序模板直接应对。 "
    sheet["A30"] = "The engagement team determines whether additional tick marks, other notations and/or additional information are relevant to document on this procedure template based on any custom steps that were added to the substantive procedure and/or the results obtained from performing the steps above.\n项目组应根据自主添加的任何实质性程序步骤和/或通过执行这些步骤获取的结果，确定是否需在本程序模板中添加额外的标记、其他注释及/或其他信息。 "

    for cell in ["A2", "A13", "A14", "A25", "A27", "A28", "A30"]:
        sheet[cell].alignment = Alignment(wrap_text=True, vertical="top")
    sheet["A2"].font = Font(name="Arial", bold=True, size=14, color=DARK_BLUE)
    for row in [13, 14, 28, 30]:
        sheet.cell(row, 1).fill = PatternFill("solid", fgColor=LIGHT_GRAY)


def _build_ckm3_sheet(workbook: Workbook, samples: list[Spd03015Sample]) -> dict[int, dict[str, int]]:
    sheet = workbook.create_sheet("CKM3")
    sheet.sheet_view.showGridLines = False
    sheet.column_dimensions["A"].width = 48
    sheet.column_dimensions["Y"].width = 24
    sheet.column_dimensions["Z"].width = 16

    ck_rows: dict[int, dict[str, int]] = {}
    for index, sample in enumerate(samples, 1):
        start = {1: 1, 2: 49, 3: 99, 4: 148, 5: 197}.get(index, 1 + (index - 1) * 49)
        sheet.cell(start, 1, f"{sample.sample}.订单编号：{sample.order or ''}")
        sheet.cell(start, 1).font = Font(name="Arial", bold=True, color=DARK_BLUE)

        labels = [
            ("beginning_qty", "期数库存交易数量", sample.beginning_qty),
            ("beginning_variance", "期初价格差异", sample.beginning_variance),
            ("receipt_qty", "收货库存交易数量", sample.receipt_qty),
            ("receipt_variance", "收货价格差异", sample.receipt_variance),
            ("outbound_qty", "消耗库存交易数量", sample.outbound_qty),
            ("outbound_variance", "消耗价格差异", sample.outbound_variance),
        ]

        label_start = start + 9
        ck_rows[sample.sample] = {}
        for offset, (key, label, value) in enumerate(labels):
            row = label_start + offset
            sheet.cell(row, 25, label)
            sheet.cell(row, 26, value)
            sheet.cell(row, 26).number_format = "#,##0.00"
            ck_rows[sample.sample][key] = row

    return ck_rows


def _ckm3_label_start(index: int) -> int:
    return CKM3_LABEL_START_ROWS.get(index, 10 + (index - 1) * 49)


def _ckm3_title_row(index: int) -> int:
    return CKM3_TITLE_ROWS.get(index, 1 + (index - 1) * 49)


def _populate_template_ckm3(sheet, samples: list[Spd03015Sample]) -> dict[int, dict[str, int]]:
    ck_rows: dict[int, dict[str, int]] = {}
    labels = [
        ("beginning_qty", "期数库存交易数量", "beginning_qty"),
        ("beginning_variance", "期初价格差异", "beginning_variance"),
        ("receipt_qty", "收货库存交易数量", "receipt_qty"),
        ("receipt_variance", "收货价格差异", "receipt_variance"),
        ("outbound_qty", "消耗库存交易数量", "outbound_qty"),
        ("outbound_variance", "消耗价格差异", "outbound_variance"),
    ]

    for index, sample in enumerate(samples, 1):
        title_row = _ckm3_title_row(index)
        label_start = _ckm3_label_start(index)
        sheet.cell(title_row, 1, f"{sample.sample}.订单编号：{sample.order or ''}")
        ck_rows[sample.sample] = {}

        for offset, (key, label, attr) in enumerate(labels):
            row = label_start + offset
            sheet.cell(row, 25, label)
            sheet.cell(row, 26, getattr(sample, attr))
            sheet.cell(row, 26).number_format = "#,##0.00"
            ck_rows[sample.sample][key] = row

    return ck_rows


def _copy_row_style(sheet, source_row: int, target_row: int, max_col: int = 37, copy_static_values: bool = False) -> None:
    sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height
    for col in range(1, max_col + 1):
        source = sheet.cell(source_row, col)
        target = sheet.cell(target_row, col)
        if source.has_style:
            target.font = copy(source.font)
            target.fill = copy(source.fill)
            target.border = copy(source.border)
            target.alignment = copy(source.alignment)
            target.number_format = source.number_format
            target.protection = copy(source.protection)
        if copy_static_values and col in {9, 10, 16, 17, 18}:
            target.value = source.value


def _ensure_template_sample_rows(sheet, sample_count: int) -> None:
    if sample_count <= 5:
        return

    extra_rows = sample_count - 5
    insert_at = 24
    template_row = 23
    static_values = {col: sheet.cell(template_row, col).value for col in [9, 10, 16, 17, 18]}
    merged_ranges_to_shift = [
        str(cell_range) for cell_range in sheet.merged_cells.ranges if cell_range.min_row >= insert_at
    ]

    for cell_range in merged_ranges_to_shift:
        sheet.unmerge_cells(cell_range)

    sheet.insert_rows(insert_at, extra_rows)

    for cell_range in merged_ranges_to_shift:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        shifted_range = (
            f"{get_column_letter(min_col)}{min_row + extra_rows}:"
            f"{get_column_letter(max_col)}{max_row + extra_rows}"
        )
        sheet.merge_cells(shifted_range)

    for row in range(insert_at, insert_at + extra_rows):
        _copy_row_style(sheet, template_row, row, copy_static_values=False)
        for col, value in static_values.items():
            sheet.cell(row, col, value)


def _populate_template_spd(sheet, samples: list[Spd03015Sample]) -> None:
    blue_fill = PatternFill("solid", fgColor="FF00338D")
    for row in [2, 3]:
        for col in range(1, 38):
            sheet.cell(row, col).fill = blue_fill

    _ensure_template_sample_rows(sheet, len(samples))

    last_sample_row = 18 + max(len(samples), 5)
    for row in range(19, last_sample_row + 1):
        for col in [2, 3, 4, 5, 6, 7, 8, 12, 13, 14, 15]:
            sheet.cell(row, col).value = None

    for row_offset, sample in enumerate(samples):
        row = 19 + row_offset
        sheet.cell(row, 1, sample.sample)
        sheet.cell(row, 2, sample.material_id or "")
        sheet.cell(row, 3, sample.period or "")
        sheet.cell(row, 4, sample.beginning_qty)
        sheet.cell(row, 5, sample.beginning_variance)
        sheet.cell(row, 6, sample.receipt_qty)
        sheet.cell(row, 7, sample.receipt_variance)
        sheet.cell(row, 8, sample.outbound_variance)
        sheet.cell(row, 12, f"=(E{row}+G{row})/(D{row}+F{row})")
        sheet.cell(row, 13, sample.outbound_qty)
        sheet.cell(row, 14, f"=M{row}*L{row}")
        sheet.cell(row, 15, f"=N{row}-H{row}")


def _build_from_template(samples: list[Spd03015Sample]) -> Workbook:
    workbook = load_workbook(TEMPLATE_PATH, data_only=False)
    _populate_template_spd(workbook["SPD03015_IRM(SAP)"], samples)
    if "CKM3" in workbook.sheetnames:
        del workbook["CKM3"]
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    return workbook


def build_spd03015_workbook(source_folder: Path) -> Workbook:
    samples = _discover_samples(source_folder)
    if TEMPLATE_PATH.exists():
        return _build_from_template(samples)

    workbook = Workbook()
    ck_rows = _build_ckm3_sheet(workbook, samples)
    _build_spd_sheet(workbook, samples, ck_rows)
    if "CKM3" in workbook.sheetnames:
        del workbook["CKM3"]
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    return workbook


def build_spd03015_bytes(source_folder: Path) -> bytes:
    workbook = build_spd03015_workbook(source_folder)
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()
