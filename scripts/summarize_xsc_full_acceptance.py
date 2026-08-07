from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path


DEFAULT_RANGES = [
    (1, 10),
    (11, 12),
    (13, 20),
    (21, 30),
    (31, 40),
    (41, 50),
    (51, 60),
    (61, 70),
    (71, 80),
    (81, 90),
    (91, 100),
    (101, 110),
    (111, 120),
    (121, 130),
    (131, 140),
    (141, 155),
]

LEGACY_STAMP_RANGES = {
    "20260612-013004": (1, 10),
}


def parse_manifest(path: Path) -> list[tuple[str, int, int]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    out = []
    for item in data:
        out.append((str(item["stamp"]), int(item["start"]), int(item["end"])))
    return out


def discover_latest_batches(summary_dir: Path, ranges: list[tuple[int, int]]) -> list[tuple[str, int, int]]:
    candidates: dict[tuple[int, int], list[tuple[float, str, Path]]] = {r: [] for r in ranges}
    for path in summary_dir.glob("*.json"):
        if path.name == "xsc-full-acceptance.json":
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            continue
        declared = data.get("range")
        if declared:
            key = (int(declared.get("start") or 0), int(declared.get("end") or 0))
        else:
            key = LEGACY_STAMP_RANGES.get(str(data.get("stamp") or path.stem))
            if key is None:
                continue
        if key in candidates:
            candidates[key].append((path.stat().st_mtime, str(data.get("stamp") or path.stem), path))
    out = []
    for start, end in ranges:
        items = candidates[(start, end)]
        if not items:
            out.append(("", start, end))
            continue
        _, stamp, _ = sorted(items, key=lambda item: (item[0], item[1]))[-1]
        out.append((stamp, start, end))
    return out


def add_counts(target: dict, source: dict, keys: list[str]) -> None:
    for key in keys:
        target[key] = target.get(key, 0) + int(source.get(key) or 0)


def expected_docs(start: int, end: int) -> list[int]:
    return list(range(start, end + 1))


def main() -> None:
    parser = argparse.ArgumentParser(description="Aggregate xsc acceptance summaries and enforce full-corpus gates.")
    parser.add_argument("--summary-dir", type=Path, default=Path(r"J:\latextomathtype\analysis\acceptance-summary"))
    parser.add_argument("--out", type=Path, default=Path(r"J:\latextomathtype\analysis\acceptance-summary\xsc-full-acceptance.json"))
    parser.add_argument("--manifest", type=Path, help="Optional JSON list of {stamp,start,end}.")
    parser.add_argument("--start", type=int, default=1)
    parser.add_argument("--end", type=int, default=155)
    parser.add_argument("--max-size-error", type=float, default=0.01)
    parser.add_argument("--allow-non-wmf", action="store_true")
    args = parser.parse_args()

    batches = parse_manifest(args.manifest) if args.manifest else discover_latest_batches(args.summary_dir, DEFAULT_RANGES)
    covered: list[int] = []
    size_totals: dict[str, int] = {}
    mtef_totals: dict[str, int] = {}
    exact_mtef_totals: dict[str, int] = {}
    failure_classes: Counter[str] = Counter()
    suspect_classes: Counter[str] = Counter()
    batch_reports = []
    missing_files = []

    for stamp, start, end in batches:
        path = args.summary_dir / f"{stamp}.json"
        if not path.exists():
            missing_files.append(str(path))
            continue
        data = json.loads(path.read_text(encoding="utf-8"))
        declared = data.get("range") or {"start": start, "end": end}
        actual_start = int(declared.get("start") or start)
        actual_end = int(declared.get("end") or end)
        covered.extend(expected_docs(actual_start, actual_end))

        size = data["size"]
        mtef = data["mtef"]
        exact_mtef = data.get("exactMtef") or {}
        add_counts(size_totals, size, [
            "pairedObjects",
            "targetMetricObjects",
            "targetWidthWithin1pct",
            "targetHeightWithin1pct",
            "nonWmfGenerated",
            "wmfWidthExact",
            "wmfHeightExact",
            "shapeWidthExact",
            "shapeHeightExact",
            "dxaOrigExact",
            "dyaOrigExact",
            "baselineCompared",
            "baselineExact",
        ])
        add_counts(mtef_totals, mtef, [
            "pairs",
            "cleanPairs",
            "alignmentSuspect",
            "hardSuspectPairs",
            "acceptedHeaderPrefixPairs",
            "lowTailPairs",
            "lowCoreTailPairs",
        ])
        add_counts(exact_mtef_totals, exact_mtef, [
            "expectedObjects",
            "generatedObjects",
            "validGeneratedOle",
            "normalizedStructureEqual",
            "rawMtefEqual",
        ])
        failure_classes.update(mtef.get("failureClassCounts") or {})
        suspect_classes.update(mtef.get("suspectClassCounts") or {})
        batch_reports.append({
            "stamp": stamp,
            "range": {"start": actual_start, "end": actual_end},
            "size": size,
            "mtef": {
                key: mtef.get(key)
                for key in [
                    "pairs",
                    "cleanPairs",
                    "alignmentSuspect",
                    "hardSuspectPairs",
                    "lowTailPairs",
                    "lowCoreTailPairs",
                    "failureClassCounts",
                    "suspectClassCounts",
                ]
            },
            "exactMtef": {
                key: exact_mtef.get(key)
                for key in [
                    "expectedObjects",
                    "generatedObjects",
                    "validGeneratedOle",
                    "normalizedStructureEqual",
                    "rawMtefEqual",
                    "passed",
                ]
            },
        })

    expected = set(expected_docs(args.start, args.end))
    covered_set = set(covered)
    duplicate_docs = sorted(doc for doc in covered_set if covered.count(doc) > 1)
    missing_docs = sorted(expected - covered_set)
    extra_docs = sorted(covered_set - expected)
    paired = int(size_totals.get("pairedObjects") or 0)

    gates = {
        "coveredExpectedDocs": not missing_docs and not extra_docs and not duplicate_docs,
        "noMissingSummaryFiles": not missing_files,
        "allGeneratedPreviewsAreWmf": args.allow_non_wmf or int(size_totals.get("nonWmfGenerated") or 0) == 0,
        "allWmfPhysicalBoundsExact": paired == int(size_totals.get("wmfWidthExact") or 0)
            and paired == int(size_totals.get("wmfHeightExact") or 0),
        "allObjectShapeSizesExact": paired == int(size_totals.get("shapeWidthExact") or 0)
            and paired == int(size_totals.get("shapeHeightExact") or 0),
        "allOriginalSizesExact": paired == int(size_totals.get("dxaOrigExact") or 0)
            and paired == int(size_totals.get("dyaOrigExact") or 0),
        "allBaselinesExact": paired == int(size_totals.get("baselineCompared") or 0)
            and paired == int(size_totals.get("baselineExact") or 0),
        "mtefPairsMatchSizePairs": paired == int(mtef_totals.get("pairs") or 0),
        "allGeneratedOleValid": paired == int(exact_mtef_totals.get("validGeneratedOle") or 0),
        "allNormalizedMtefStructuresEqual": paired == int(exact_mtef_totals.get("normalizedStructureEqual") or 0),
    }
    passed = all(gates.values())
    allowed_prefix = int(suspect_classes.get("source_header_or_style_prefix") or 0)
    allowed_char_stream = int(suspect_classes.get("char_stream_style_gap") or 0)
    effective_hard = int(mtef_totals.get("hardSuspectPairs") or 0)
    effective_low_tail = int(mtef_totals.get("lowTailPairs") or 0)

    out = {
        "dataset": "xsc",
        "range": {"start": args.start, "end": args.end},
        "thresholds": {"maxTargetWmfPhysicalSizeError": args.max_size_error},
        "passed": passed,
        "gates": gates,
        "coverage": {
            "missingDocs": missing_docs,
            "extraDocs": extra_docs,
            "duplicateDocs": duplicate_docs,
            "missingSummaryFiles": missing_files,
        },
        "size": size_totals,
        "mtef": {
            **mtef_totals,
            "failureClassCounts": dict(sorted(failure_classes.items())),
            "suspectClassCounts": dict(sorted(suspect_classes.items())),
        },
        "exactMtef": exact_mtef_totals,
        "mtefEffective": {
            "allowedHeaderOrStylePrefixPairs": allowed_prefix,
            "allowedCharStreamStylePairs": allowed_char_stream,
            "allowedNonStructuralPairs": allowed_prefix + allowed_char_stream,
            "remainingHardSuspectPairs": effective_hard,
            "remainingLowTailPairs": effective_low_tail,
            "remainingStructuralGapPairs": effective_hard + effective_low_tail,
        },
        "batches": batch_reports,
    }
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.out)
    if not passed:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
