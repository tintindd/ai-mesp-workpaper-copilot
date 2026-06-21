from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path

from PIL import Image

from mesp_automation_engine import (
    EXPECTED_EVIDENCE,
    REPORTS,
    REQUIRED_FIELDS,
    analyze_workbook,
    parse_file,
    summarize_parse_failure,
)


WORKSPACE = Path(__file__).resolve().parents[1]
TRAINING_ROOT = WORKSPACE / "training_files" / "训练文件"
OUTPUT_PATH = WORKSPACE / "mesp_training_profile.json"


def inspect_image(path: Path) -> dict:
    image = Image.open(path)
    return {"width": image.width, "height": image.height, "mode": image.mode}


def build_profile() -> dict:
    files = []
    ignored_files = []
    evidence_by_sample = defaultdict(set)
    workbook_profiles = defaultdict(list)

    for path in sorted(TRAINING_ROOT.rglob("*")):
        if not path.is_file():
            continue

        parsed = parse_file(path, TRAINING_ROOT)
        item = {
            "file": path.name,
            "relative_path": str(path.relative_to(WORKSPACE)),
            "size": path.stat().st_size,
            "parsed": bool(parsed),
        }
        if not parsed:
            ignored_files.append(summarize_parse_failure(path, TRAINING_ROOT))
            files.append(item)
            continue

        item.update(parsed)
        evidence_by_sample[item["sample"]].add(f"{item['report']}-{item['kind']}")

        if item["kind"] == "表格" and item["ext"] in {"xlsx", "xlsm", "csv"}:
            item["workbook"] = analyze_workbook(path, item["report"])
            workbook_profiles[item["report"]].append(
                {
                    "file": item["file"],
                    "relative_path": item["relative_path"],
                    "sample": item["sample"],
                    "order": item["order"],
                    **item["workbook"],
                }
            )
        elif item["kind"] == "截图":
            item["image"] = inspect_image(path)

        files.append(item)

    sample_profiles = []
    for sample in sorted(evidence_by_sample):
        present = sorted(evidence_by_sample[sample])
        missing = [key for key in EXPECTED_EVIDENCE if key not in evidence_by_sample[sample]]
        sample_files = [file for file in files if file.get("sample") == sample]
        order = next((file.get("order") for file in sample_files if file.get("order")), "")
        sample_profiles.append(
            {
                "sample": sample,
                "order": order,
                "present_evidence": present,
                "missing_evidence": missing,
            }
        )

    report_profiles = {}
    for report in REPORTS:
        workbooks = workbook_profiles.get(report, [])
        common_headers = []
        if workbooks:
            header_sets = [set(book["headers"]) for book in workbooks]
            common_headers = sorted(set.intersection(*header_sets))
        report_profiles[report] = {
            "workbook_count": len(workbooks),
            "common_headers": common_headers,
            "field_aliases": REQUIRED_FIELDS.get(report, {}),
            "workbooks": workbooks,
        }

    return {
        "source_folder": str(TRAINING_ROOT),
        "training_summary": {
            "sample_count": len(sample_profiles),
            "file_count": len(files),
            "recognized_file_count": sum(1 for item in files if item.get("parsed")),
            "ignored_file_count": len(ignored_files),
            "workbook_count": sum(len(items) for items in workbook_profiles.values()),
            "expected_evidence": EXPECTED_EVIDENCE,
            "known_limitation": "",
        },
        "samples": sample_profiles,
        "report_profiles": report_profiles,
        "ignored_files": ignored_files,
        "files": files,
    }


def main() -> None:
    if not TRAINING_ROOT.exists():
        raise FileNotFoundError(f"Training folder not found: {TRAINING_ROOT}")

    profile = build_profile()
    OUTPUT_PATH.write_text(json.dumps(profile, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(f"Wrote {OUTPUT_PATH}")
    print(
        json.dumps(
            {
                "sample_count": profile["training_summary"]["sample_count"],
                "file_count": profile["training_summary"]["file_count"],
                "recognized_file_count": profile["training_summary"]["recognized_file_count"],
                "ignored_file_count": profile["training_summary"]["ignored_file_count"],
                "workbook_count": profile["training_summary"]["workbook_count"],
                "report_workbooks": {
                    report: info["workbook_count"] for report, info in profile["report_profiles"].items()
                },
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
