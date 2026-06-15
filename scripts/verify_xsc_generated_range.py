from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def run_checked(args: list[str], cwd: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(args, cwd=str(cwd), text=True, capture_output=True, encoding="utf-8", errors="replace")


def source_object_count(path: Path) -> int:
    with zipfile.ZipFile(path) as zf:
        document = zf.read("word/document.xml").decode("utf-8", "replace")
    return len(re.findall(r"<w:object\b", document))


def source_report_count(path: Path) -> int | None:
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return None
    equations = data.get("equations") if isinstance(data, dict) else None
    return len(equations) if isinstance(equations, list) else None


def source_report_exclusions(path: Path) -> list[dict]:
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return []
    exclusions = data.get("excludedSourceObjects") if isinstance(data, dict) else None
    return exclusions if isinstance(exclusions, list) else []


def exclusions_have_evidence(exclusions: list[dict]) -> bool:
    for item in exclusions:
        if not item.get("reason"):
            return False
        expected_ole = item.get("expectedOleTarget")
        expected_wmf = item.get("expectedWmfTarget")
        if expected_ole and item.get("oleTarget") != expected_ole:
            return False
        if expected_wmf and item.get("wmfTarget") != expected_wmf:
            return False
        if not expected_ole and not expected_wmf:
            if item.get("isMathType") is not False:
                return False
            if not item.get("mathTypeError"):
                return False
    return True


def summary_count(summary: dict, key: str) -> int | None:
    value = summary.get(key, {})
    return value.get("count", value.get("n")) if isinstance(value, dict) else None


def max_error_within(value: float | None, limit: float = 1.0) -> bool:
    return value is not None and value <= limit


MATH_SPAN_RE = re.compile(r"(\$\$.*?\$\$|\$.*?\$)", re.S)


def request_math_count(path: Path) -> int:
    request = json.loads(path.read_text(encoding="utf-8"))
    total = 0
    for section in request.get("sections", []):
        for question in section.get("questions", []):
            for field in ("content", "knowledgePoint", "difficulty", "analyze", "solution", "correct"):
                total += len(MATH_SPAN_RE.findall(question.get(field) or ""))
            for tag in question.get("tags") or []:
                total += len(MATH_SPAN_RE.findall(str(tag)))
    return total


def request_build_summary(request_dir: Path) -> dict[int, dict]:
    path = request_dir / "summary.json"
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}
    summary = {}
    if isinstance(data, list):
        for item in data:
            source = str(item.get("source") or "")
            match = re.search(r"(\d+)\.docx$", source)
            if match:
                summary[int(match.group(1))] = item
    return summary


def find_generated_docx(run_dir: Path, index: int, start: int | None = None, end: int | None = None) -> Path:
    names = [
        f"xsc测试集完整重建_{index:02d}.docx",
        f"xsc测试集完整重建_{index}.docx",
    ]
    if start is not None and end is not None:
        range_dir = run_dir / "docx" / f"{start}-{end}"
        for name in names:
            direct = range_dir / name
            if direct.exists():
                return direct
    for name in names:
        direct = run_dir / name
        if direct.exists():
            return direct
    matches = []
    for name in names:
        matches.extend(run_dir.rglob(name))
    if not matches:
        raise FileNotFoundError(f"missing generated docx for index {index} under {run_dir}")
    unique_matches = sorted(set(path.resolve() for path in matches), key=str)
    if len(unique_matches) != 1:
        formatted = "\n".join(str(path) for path in unique_matches[:20])
        raise RuntimeError(f"ambiguous generated docx for index {index} under {run_dir}:\n{formatted}")
    return unique_matches[0]


def find_source_docx(manifest_item: dict, search_roots: list[Path]) -> tuple[Path | None, str, list[str]]:
    direct = Path(manifest_item["sourcePath"])
    if direct.exists():
        return direct, "manifest", []

    source_name = str(manifest_item.get("sourceName") or direct.name)
    candidates: list[Path] = []
    for root in search_roots:
        if root.exists():
            candidates.extend(root.rglob(source_name))
    unique_candidates = sorted(set(path.resolve() for path in candidates), key=str)
    if len(unique_candidates) == 1:
        return unique_candidates[0], "relocated-by-name", [str(unique_candidates[0])]
    if not unique_candidates:
        return None, "missing", []
    return None, "ambiguous", [str(path) for path in unique_candidates[:20]]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo", type=Path, default=Path.cwd())
    parser.add_argument("--python", type=Path, required=True)
    parser.add_argument("--run-dir", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--request-dir", type=Path, required=True)
    parser.add_argument("--start", type=int, required=True)
    parser.add_argument("--end", type=int, required=True)
    parser.add_argument("--summary-name", default="")
    parser.add_argument("--source-search-root", action="append", type=Path, default=[])
    parser.add_argument("--allow-relocated-source", action="store_true")
    args = parser.parse_args()

    repo = args.repo.resolve()
    run_dir = args.run_dir.resolve()
    acceptance_dir = run_dir / "acceptance"
    compare_dir = acceptance_dir / "size-compare"
    acceptance_dir.mkdir(parents=True, exist_ok=True)
    compare_dir.mkdir(parents=True, exist_ok=True)

    manifest = json.loads(args.manifest.read_text(encoding="utf-8-sig"))
    manifest_items = {int(item["index"]): item for item in manifest}
    default_source_roots = sorted({Path(item["sourcePath"]).parent for item in manifest_items.values()}, key=str)
    source_search_roots = default_source_roots + [path.resolve() for path in args.source_search_root]
    build_summary = request_build_summary(args.request_dir)
    rows = []
    for index in range(args.start, args.end + 1):
        manifest_item = manifest_items[index]
        source_docx, source_status, source_candidates = find_source_docx(manifest_item, source_search_roots)
        generated_docx = find_generated_docx(run_dir, index, args.start, args.end)
        request_json = args.request_dir / f"full-{index:02d}.request.json"
        request_expected = request_math_count(request_json)
        source_report = acceptance_dir / f"source-report-{index:02d}.json"
        if source_docx is None:
            rows.append(
                {
                    "index": index,
                    "sourceName": manifest_item.get("sourceName") or Path(manifest_item["sourcePath"]).name,
                    "sourcePath": None,
                    "sourcePathStatus": source_status,
                    "sourcePathCandidates": source_candidates,
                    "sourceTrusted": False,
                    "sourceObjectExpected": None,
                    "excludedSourceObjectCount": 0,
                    "excludedSourceObjectSamples": [],
                    "excludedSourceObjectsHaveEvidence": False,
                    "requestMathExpected": request_expected,
                    "sourceReportExit": 1,
                    "scanExit": None,
                    "compareExit": None,
                    "wmfExit": None,
                    "toolsSucceeded": False,
                    "countsMatch": False,
                    "generatedMatchesRequest": False,
                    "sourceObjectGap": None,
                    "allSourceFormulasCompared": False,
                    "sourceWmfWithin1Pct": False,
                    "sourceShapeWithin1Pct": False,
                    "noOrdinalLatexMismatches": False,
                    "allRequestFormulasHaveMetrics": False,
                    "noLeaks": False,
                    "realMathTypeOle": False,
                    "targetWmfWithin1Pct": False,
                    "targetShapeWithin1Pct": False,
                    "vectorNoDib": False,
                    "noWmfLatexLeaks": False,
                    "noUnreviewedWmfSuspiciousText": False,
                    "noAppendedMissingEquations": False,
                    "excludedObjectsReviewed": False,
                }
            )
            continue
        source_report_build = run_checked([
            str(args.python), "scripts/build_xsc_source_report.py",
            "--index", str(index),
            "--tex", str(manifest_item["latexPath"]),
            "--source-docx", str(source_docx),
            "--out", str(source_report),
        ], repo)
        source_expected = source_report_count(source_report)
        source_exclusions = source_report_exclusions(source_report)

        scan_json = acceptance_dir / f"latex-leak-scan-{index:02d}.json"
        scan = run_checked([
            str(args.python), "scripts/scan_docx_latex_leaks.py", str(generated_docx),
            "--expect-ole-count", str(request_expected), "--expect-wmf-count", str(request_expected),
            "--out", str(scan_json),
        ], repo)

        compare = run_checked([
            str(args.python), "scripts/compare_docx_pair_metrics.py", str(source_docx),
            str(generated_docx), str(compare_dir),
            "--source-report", str(source_report),
            "--generated-request", str(request_json),
        ], repo)

        wmf_json = acceptance_dir / f"wmf-records-{index:02d}.json"
        wmf = run_checked([
            str(args.python), "scripts/wmf_record_report.py", str(generated_docx),
            "--out", str(wmf_json),
        ], repo)

        scan_data = json.loads(scan_json.read_text(encoding="utf-8")) if scan_json.exists() else {"error": scan.stderr[-500:]}
        compare_json = compare_dir / f"{source_docx.stem}_summary.json"
        compare_data = json.loads(compare_json.read_text(encoding="utf-8")) if compare_json.exists() else {"error": compare.stderr[-500:]}
        wmf_data = json.loads(wmf_json.read_text(encoding="utf-8")) if wmf_json.exists() else {"error": wmf.stderr[-500:]}

        row = {
            "index": index,
            "sourceName": source_docx.name,
            "sourcePath": str(source_docx),
            "sourcePathStatus": source_status,
            "sourcePathCandidates": source_candidates,
            "sourceTrusted": source_status == "manifest" or args.allow_relocated_source,
            "sourceObjectExpected": source_expected,
            "excludedSourceObjectCount": len(source_exclusions),
            "excludedSourceObjectSamples": source_exclusions[:10],
            "excludedSourceObjectsHaveEvidence": exclusions_have_evidence(source_exclusions),
            "requestMathExpected": request_expected,
            "sourceReportExit": source_report_build.returncode,
            "scanExit": scan.returncode,
            "compareExit": compare.returncode,
            "wmfExit": wmf.returncode,
            "leakCount": scan_data.get("leakCount"),
            "hasEquationDsmt4": scan_data.get("hasEquationDsmt4"),
            "hasOleObject": scan_data.get("hasOleObject"),
            "equationDsmt4Count": scan_data.get("equationDsmt4Count"),
            "oleObjectXmlCount": scan_data.get("oleObjectXmlCount"),
            "validMathTypeOleCount": scan_data.get("validMathTypeOleCount"),
            "invalidMathTypeOleCount": scan_data.get("invalidMathTypeOleCount"),
            "sourceOle": compare_data.get("source_meta", {}).get("objects"),
            "generatedOle": compare_data.get("generated_meta", {}).get("objects"),
            "generatedWmf": compare_data.get("generated_meta", {}).get("media_wmf"),
            "missing": compare_data.get("missing_generated_objects"),
            "extra": compare_data.get("extra_generated_objects"),
            "targetObjects": compare_data.get("target_metric_objects"),
            "sourceComparedObjects": summary_count(compare_data, "ordinal_wmf_width_ratio"),
            "sourceWmfWidthWithin1Pct": compare_data.get("ordinal_wmf_width_ratio", {}).get("within_1pct"),
            "sourceWmfHeightWithin1Pct": compare_data.get("ordinal_wmf_height_ratio", {}).get("within_1pct"),
            "sourceShapeWidthWithin1Pct": compare_data.get("ordinal_shape_width_ratio", {}).get("within_1pct"),
            "sourceShapeHeightWithin1Pct": compare_data.get("ordinal_shape_height_ratio", {}).get("within_1pct"),
            "sourceWmfWidthMaxAbsErrorPct": compare_data.get("ordinal_wmf_width_ratio", {}).get("max_abs_error_pct"),
            "sourceWmfHeightMaxAbsErrorPct": compare_data.get("ordinal_wmf_height_ratio", {}).get("max_abs_error_pct"),
            "sourceShapeWidthMaxAbsErrorPct": compare_data.get("ordinal_shape_width_ratio", {}).get("max_abs_error_pct"),
            "sourceShapeHeightMaxAbsErrorPct": compare_data.get("ordinal_shape_height_ratio", {}).get("max_abs_error_pct"),
            "ordinalLatexMismatches": compare_data.get("ordinal_latex_mismatches"),
            "targetWmfWidthWithin1Pct": compare_data.get("target_wmf_width_ratio", {}).get("within_1pct"),
            "targetWmfHeightWithin1Pct": compare_data.get("target_wmf_height_ratio", {}).get("within_1pct"),
            "targetWmfWidthMaxAbsErrorPct": compare_data.get("target_wmf_width_ratio", {}).get("max_abs_error_pct"),
            "targetWmfHeightMaxAbsErrorPct": compare_data.get("target_wmf_height_ratio", {}).get("max_abs_error_pct"),
            "targetShapeWidthMaxAbsErrorPct": compare_data.get("target_shape_width_ratio", {}).get("max_abs_error_pct"),
            "targetShapeHeightMaxAbsErrorPct": compare_data.get("target_shape_height_ratio", {}).get("max_abs_error_pct"),
            "targetShapeWidthWithin1Pct": compare_data.get("target_shape_width_ratio", {}).get("within_1pct"),
            "targetShapeHeightWithin1Pct": compare_data.get("target_shape_height_ratio", {}).get("within_1pct"),
            "vectorText": wmf_data.get("vectorText"),
            "vectorContent": wmf_data.get("vectorContent"),
            "stretchDib": wmf_data.get("stretchDib"),
            "bitmapRecords": wmf_data.get("bitmapRecords"),
            "bitmapWmf": wmf_data.get("bitmapWmf"),
            "wmfLatexLeakCount": wmf_data.get("wmfLatexLeakCount"),
            "wmfLatexLeakSamples": wmf_data.get("wmfLatexLeakSamples", [])[:5],
            "wmfSuspiciousScriptTextCount": wmf_data.get("wmfSuspiciousScriptTextCount"),
            "wmfSuspiciousScriptTextSamples": wmf_data.get("wmfSuspiciousScriptTextSamples", [])[:5],
            "wmfSuspiciousReviewRequiredCount": wmf_data.get("wmfSuspiciousReviewRequiredCount"),
            "wmfSuspiciousReviewRequiredSamples": wmf_data.get("wmfSuspiciousReviewRequiredSamples", [])[:5],
            "appendedMissingEquations": build_summary.get(index, {}).get("appended_missing_equations"),
            "cursorMultiConsumes": build_summary.get(index, {}).get("cursor_multi_consumes"),
        }
        row["sourceObjectGap"] = (
            row["sourceObjectExpected"] - row["requestMathExpected"]
            if row["sourceObjectExpected"] is not None else None
        )
        row["countsMatch"] = (
            row["sourceObjectExpected"] is not None
            and
            row["sourceObjectExpected"] == row["requestMathExpected"]
            == row["sourceOle"] == row["generatedOle"] == row["generatedWmf"]
        )
        row["generatedMatchesRequest"] = (
            row["requestMathExpected"] == row["generatedOle"] == row["generatedWmf"]
        )
        row["noLeaks"] = row["leakCount"] == 0 and row["scanExit"] == 0
        row["realMathTypeOle"] = (
            row["hasEquationDsmt4"] is True
            and row["hasOleObject"] is True
            and row["equationDsmt4Count"] == row["requestMathExpected"]
            and row["oleObjectXmlCount"] == row["requestMathExpected"]
            and row["validMathTypeOleCount"] == row["requestMathExpected"]
            and row["invalidMathTypeOleCount"] == 0
        )
        row["toolsSucceeded"] = row["sourceReportExit"] == row["scanExit"] == row["compareExit"] == row["wmfExit"] == 0
        row["allSourceFormulasCompared"] = (
            row["sourceObjectExpected"] is not None
            and row["sourceComparedObjects"] == row["sourceObjectExpected"]
        )
        row["sourceWmfWithin1Pct"] = row["allSourceFormulasCompared"] and (
            row["sourceObjectExpected"] == row["sourceWmfWidthWithin1Pct"] == row["sourceWmfHeightWithin1Pct"]
            and max_error_within(row["sourceWmfWidthMaxAbsErrorPct"])
            and max_error_within(row["sourceWmfHeightMaxAbsErrorPct"])
        )
        row["sourceShapeWithin1Pct"] = row["allSourceFormulasCompared"] and (
            row["sourceObjectExpected"] == row["sourceShapeWidthWithin1Pct"] == row["sourceShapeHeightWithin1Pct"]
        )
        row["noOrdinalLatexMismatches"] = row["ordinalLatexMismatches"] == 0
        row["allRequestFormulasHaveMetrics"] = row["targetObjects"] == row["requestMathExpected"]
        row["targetWmfWithin1Pct"] = row["allRequestFormulasHaveMetrics"] and (
            row["requestMathExpected"] == row["targetWmfWidthWithin1Pct"] == row["targetWmfHeightWithin1Pct"]
            and max_error_within(row["targetWmfWidthMaxAbsErrorPct"])
            and max_error_within(row["targetWmfHeightMaxAbsErrorPct"])
        )
        row["targetShapeWithin1Pct"] = row["allRequestFormulasHaveMetrics"] and (
            row["requestMathExpected"] == row["targetShapeWidthWithin1Pct"] == row["targetShapeHeightWithin1Pct"]
            and max_error_within(row["targetShapeWidthMaxAbsErrorPct"])
            and max_error_within(row["targetShapeHeightMaxAbsErrorPct"])
        )
        row["vectorNoDib"] = (
            row["generatedWmf"] == row["vectorContent"]
            and row["stretchDib"] == 0
            and row["bitmapRecords"] == 0
            and row["bitmapWmf"] == 0
        )
        row["noWmfLatexLeaks"] = row["wmfLatexLeakCount"] == 0
        row["noUnreviewedWmfSuspiciousText"] = row["wmfSuspiciousReviewRequiredCount"] == 0
        row["noAppendedMissingEquations"] = row["appendedMissingEquations"] == 0
        row["excludedObjectsReviewed"] = row["excludedSourceObjectsHaveEvidence"] is True
        rows.append(row)

    result = {
        "run": str(run_dir),
        "start": args.start,
        "end": args.end,
        "allCountsMatch": all(row["countsMatch"] for row in rows),
        "allGeneratedMatchesRequest": all(row["generatedMatchesRequest"] for row in rows),
        "allSourceObjectsAccounted": all(
            row["sourceObjectGap"] == 0 and row["countsMatch"] and row["excludedObjectsReviewed"]
            for row in rows
        ),
        "allSourcePathsTrusted": all(row["sourceTrusted"] for row in rows),
        "unknownSourceObjectGapCount": sum(1 for row in rows if row["sourceObjectGap"] is None),
        "totalSourceObjectGap": None
        if any(row["sourceObjectGap"] is None for row in rows)
        else sum(row["sourceObjectGap"] for row in rows),
        "allToolsSucceeded": all(row["toolsSucceeded"] for row in rows),
        "allSourceFormulasCompared": all(row["allSourceFormulasCompared"] for row in rows),
        "allSourceWmfWithin1Pct": all(row["sourceWmfWithin1Pct"] for row in rows),
        "advisoryAllSourceShapeWithin1Pct": all(row["sourceShapeWithin1Pct"] for row in rows),
        "advisorySourceShapeNote": (
            "Diagnostic only: source shape boxes differ between original Word objects and regenerated "
            "target display boxes, so this value does not affect bad/exit status."
        ),
        "sourceShapeGateRequired": False,
        "allNoOrdinalLatexMismatches": all(row["noOrdinalLatexMismatches"] for row in rows),
        "allRequestFormulasHaveMetrics": all(row["allRequestFormulasHaveMetrics"] for row in rows),
        "allNoLeaks": all(row["noLeaks"] for row in rows),
        "allRealMathTypeOle": all(row["realMathTypeOle"] for row in rows),
        "allTargetWmfWithin1Pct": all(row["targetWmfWithin1Pct"] for row in rows),
        "allTargetShapeWithin1Pct": all(row["targetShapeWithin1Pct"] for row in rows),
        "allVectorNoDib": all(row["vectorNoDib"] for row in rows),
        "allNoWmfLatexLeaks": all(row["noWmfLatexLeaks"] for row in rows),
        "allNoUnreviewedWmfSuspiciousText": all(row["noUnreviewedWmfSuspiciousText"] for row in rows),
        "allNoAppendedMissingEquations": all(row["noAppendedMissingEquations"] for row in rows),
        "allExcludedObjectsReviewed": all(row["excludedObjectsReviewed"] for row in rows),
        "files": rows,
    }
    summary_name = args.summary_name or f"batch-{args.start:03d}-{args.end:03d}-acceptance-summary.json"
    (acceptance_dir / summary_name).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({key: value for key, value in result.items() if key != "files"}, ensure_ascii=False, indent=2))
    bad = [
        row for row in rows
        if not (row["toolsSucceeded"] and row["countsMatch"] and row["generatedMatchesRequest"] and row["sourceObjectGap"] == 0
                and row["allSourceFormulasCompared"] and row["sourceWmfWithin1Pct"]
                and row["noOrdinalLatexMismatches"] and row["allRequestFormulasHaveMetrics"] and row["noLeaks"]
                and row["realMathTypeOle"] and row["targetWmfWithin1Pct"] and row["targetShapeWithin1Pct"]
                and row["vectorNoDib"] and row["noWmfLatexLeaks"] and row["noUnreviewedWmfSuspiciousText"]
                and row["noAppendedMissingEquations"] and row["excludedObjectsReviewed"] and row["sourceTrusted"])
    ]
    print(json.dumps({"bad": bad}, ensure_ascii=False, indent=2))
    return 0 if not bad else 1


if __name__ == "__main__":
    raise SystemExit(main())
