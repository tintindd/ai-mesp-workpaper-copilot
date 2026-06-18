from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook


FILE_PATTERN = re.compile(
    r"(?P<sample>\d+)\.订单编号(?P<order>\d+)-(?P<report>[A-Za-z0-9]+)-(?P<kind>截图|表格)\.(?P<ext>\w+)$"
)

EXPECTED_EVIDENCE = [
    "CO03-截图",
    "CO03-表格",
    "KSBT-截图",
    "KSBT-表格",
    "3611-截图",
    "3611-表格",
    "CKM3-截图",
    "CKM3-表格",
]

CO03_REQUIRED_FIELDS = {
    "cost_center": ["成本中心"],
    "cost_element": ["成本要素"],
    "cost_element_text": ["成本要素 (文本)"],
    "material": ["物料"],
    "total_plan_cost": ["总计划成本"],
    "total_actual_cost": ["总实际成本"],
    "plan_actual_difference": ["计划/实际差异"],
    "actual_total_quantity": ["实际总计数量"],
    "actual_fixed_cost": ["实际固定成本"],
    "actual_variable_cost": ["实际变动成本"],
    "activity_type": ["活动类型"],
}


def parse_file(path: Path) -> dict | None:
    match = FILE_PATTERN.match(path.name)
    if not match:
        return None

    return {
        "file": path.name,
        "path": str(path),
        "sample": int(match.group("sample")),
        "order": match.group("order"),
        "report": match.group("report").upper(),
        "kind": match.group("kind"),
        "ext": match.group("ext").lower(),
    }


def numeric(value) -> float:
    return value if isinstance(value, (int, float)) else 0.0


def detect_columns(headers: list) -> tuple[dict, list]:
    header_index = {value: index for index, value in enumerate(headers) if value not in (None, "")}
    mapping = {}
    missing = []

    for canonical, aliases in CO03_REQUIRED_FIELDS.items():
        matched = next((alias for alias in aliases if alias in header_index), None)
        if matched:
            mapping[canonical] = {
                "source_header": matched,
                "column_index": header_index[matched] + 1,
            }
        else:
            missing.append(canonical)

    return mapping, missing


def analyze_co03_workbook(path: Path) -> dict:
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    headers = list(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True)))
    header_index = {value: index for index, value in enumerate(headers) if value not in (None, "")}
    mapping, missing_fields = detect_columns(headers)

    totals = defaultdict(float)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        for header in ["总计划成本", "总实际成本", "计划/实际差异", "实际总计数量", "实际固定成本", "实际变动成本"]:
            index = header_index.get(header)
            if index is not None and index < len(row):
                totals[header] += numeric(row[index])

    return {
        "sheet": sheet.title,
        "row_count": max((sheet.max_row or 1) - 1, 0),
        "column_count": sheet.max_column,
        "mapping": mapping,
        "missing_fields": missing_fields,
        "totals": {key: round(value, 2) for key, value in totals.items()},
    }


def build_issue(issue_type: str, sample, problem: str, suggestion: str, status: str = "warn") -> dict:
    return {
        "type": issue_type,
        "sample": sample,
        "problem": problem,
        "suggestion": suggestion,
        "status": status,
    }


def analyze_folder(folder: Path, period: str = "", program: str = "") -> dict:
    parsed_files = []
    ignored_files = []

    for path in sorted(folder.iterdir()):
        if not path.is_file():
            continue
        parsed = parse_file(path)
        if parsed:
            parsed_files.append(parsed)
        else:
            ignored_files.append(path.name)

    by_sample = defaultdict(list)
    for item in parsed_files:
        by_sample[item["sample"]].append(item)

    issues = []
    co03_results = []
    evidence_trace = []

    for sample in sorted(by_sample):
        sample_files = by_sample[sample]
        order = sample_files[0]["order"]
        present = {f"{item['report']}-{item['kind']}" for item in sample_files}

        for evidence_key in EXPECTED_EVIDENCE:
            if evidence_key not in present:
                report, kind = evidence_key.split("-")
                issues.append(
                    build_issue(
                        "支持文件缺失",
                        sample,
                        f"样本 {sample} 缺少 {report} {kind} 支持。",
                        f"请补充命名类似“{sample}.订单编号{order}-{report}-{kind}”的文件。",
                        "missing",
                    )
                )

        for item in sample_files:
            if item["report"] == "CO03" and item["kind"] == "表格" and item["ext"] == "xlsx":
                result = analyze_co03_workbook(Path(item["path"]))
                result.update({"sample": sample, "order": order, "file": item["file"]})
                co03_results.append(result)

                for canonical, meta in result["mapping"].items():
                    evidence_trace.append(
                        {
                            "sample": sample,
                            "order": order,
                            "workpaper_field": canonical,
                            "source_file": item["file"],
                            "source_sheet": result["sheet"],
                            "source_header": meta["source_header"],
                            "source_column": meta["column_index"],
                            "status": "Mapped",
                        }
                    )

                if result["missing_fields"]:
                    issues.append(
                        build_issue(
                            "CO03 字段缺失",
                            sample,
                            f"CO03 表格缺少关键字段：{', '.join(result['missing_fields'])}。",
                            "请在 SAP 导出 CO03 时选择全部字段布局，或补充字段别名映射。",
                        )
                    )

    if period:
        issues.append(
            build_issue(
                "期间一致性检查",
                "全部",
                f"需要检查 KSBT、3611、CKM3 的会计期间是否与审计期间 {period} 一致。",
                "当前训练包中这些报表为截图，建议后续补充表格导出或接入 OCR 后自动核对。",
            )
        )

    if program == "SPD03015":
        issues.append(
            build_issue(
                "关键计算字段待确认",
                "全部",
                "SPD03015 需要期初差异金额、期初库存、本期入库、本月出库和账面分摊金额。",
                "当前 CKM3 只有截图；建议补充 CKM3 表格导出后自动计算差异分摊率。",
            )
        )

    questions = [
        {
            "title": f"追问 {index + 1}：{issue['type']}",
            "body": issue["problem"],
            "ask": issue["suggestion"],
        }
        for index, issue in enumerate(issues)
    ]

    return {
        "input_folder": str(folder),
        "period": period,
        "program": program,
        "summary": {
            "sample_count": len(by_sample),
            "recognized_file_count": len(parsed_files),
            "ignored_file_count": len(ignored_files),
            "missing_file_count": sum(1 for issue in issues if issue["type"] == "支持文件缺失"),
            "warning_count": sum(1 for issue in issues if issue["type"] != "支持文件缺失"),
            "co03_workbook_count": len(co03_results),
        },
        "ignored_files": ignored_files,
        "issues": issues,
        "questions": questions,
        "co03_results": co03_results,
        "evidence_trace": evidence_trace,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze MESP support files and generate automation output.")
    parser.add_argument(
        "--input",
        default=str(Path(__file__).resolve().parents[1] / "training_files" / "训练文件"),
        help="Folder containing MESP support files.",
    )
    parser.add_argument("--period", default="", help="Audit period, e.g. 2025.01.01-2025.12.31.")
    parser.add_argument("--program", default="", help="Selected MESP program, e.g. SPD03015.")
    parser.add_argument(
        "--out",
        default=str(Path(__file__).resolve().parents[1] / "mesp_automation_result.json"),
        help="Output JSON path.",
    )
    args = parser.parse_args()

    result = analyze_folder(Path(args.input), period=args.period, program=args.program)
    output_path = Path(args.out)
    output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(f"Wrote {output_path}")
    print(json.dumps(result["summary"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
