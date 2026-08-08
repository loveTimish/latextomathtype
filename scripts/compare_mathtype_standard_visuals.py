from __future__ import annotations

import argparse
import io
import json
import re
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from collections import deque
from pathlib import Path

import numpy as np
from PIL import Image


def numeric_key(name: str) -> int:
    match = re.search(r"(\d+)(?=\.[^.]+$)", name)
    return int(match.group(1)) if match else 2**31 - 1


def extract_object_wmfs(docx: Path) -> list[tuple[str, bytes] | None]:
    with zipfile.ZipFile(docx) as archive:
        if "word/document.xml" in archive.namelist() and "word/_rels/document.xml.rels" in archive.namelist():
            relationships = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
            targets = {
                item.attrib.get("Id", ""): item.attrib.get("Target", "")
                for item in relationships
            }
            document = ET.fromstring(archive.read("word/document.xml"))
            ordered: list[tuple[str, bytes] | None] = []
            for obj in document.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}object"):
                image = next(obj.iter("{urn:schemas-microsoft-com:vml}imagedata"), None)
                relation_id = image.attrib.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", ""
                ) if image is not None else ""
                target = targets.get(relation_id, "").replace("\\", "/").lstrip("/")
                name = target if target.startswith("word/") else "word/" + target
                if name.lower().endswith(".wmf") and name in archive.namelist():
                    ordered.append((name, archive.read(name)))
                else:
                    ordered.append(None)
            if ordered:
                return ordered
        names = sorted(
            (name for name in archive.namelist()
             if name.startswith("word/media/") and name.lower().endswith(".wmf")),
            key=numeric_key,
        )
        return [(name, archive.read(name)) for name in names]


def extract_wmfs(docx: Path) -> list[tuple[str, bytes]]:
    return [item for item in extract_object_wmfs(docx) if item is not None]


def render_wmf(data: bytes, magick: str, density: int) -> np.ndarray:
    with tempfile.TemporaryDirectory(prefix="mathtype-wmf-") as directory:
        source = Path(directory) / "formula.wmf"
        target = Path(directory) / "formula.png"
        source.write_bytes(data)
        command = [
            magick, "-density", str(density), str(source),
            "-background", "white", "-alpha", "remove", "-alpha", "off", str(target),
        ]
        completed = subprocess.run(command, capture_output=True, text=True, check=False)
        if completed.returncode != 0 or not target.exists():
            raise RuntimeError(f"ImageMagick WMF render failed: {completed.stderr.strip()}")
        return np.asarray(Image.open(target).convert("L"), dtype=np.float64)


def common_canvas(left: np.ndarray, right: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    height = max(left.shape[0], right.shape[0])
    width = max(left.shape[1], right.shape[1])
    a = np.full((height, width), 255.0)
    b = np.full((height, width), 255.0)
    a[: left.shape[0], : left.shape[1]] = left
    b[: right.shape[0], : right.shape[1]] = right
    return a, b


def shift_on_canvas(image: np.ndarray, dy: int, dx: int) -> np.ndarray:
    shifted = np.full_like(image, 255.0)
    src_y0 = max(0, -dy)
    src_y1 = min(image.shape[0], image.shape[0] - dy)
    src_x0 = max(0, -dx)
    src_x1 = min(image.shape[1], image.shape[1] - dx)
    if src_y1 > src_y0 and src_x1 > src_x0:
        shifted[src_y0 + dy:src_y1 + dy, src_x0 + dx:src_x1 + dx] = \
            image[src_y0:src_y1, src_x0:src_x1]
    return shifted


def global_ssim(a: np.ndarray, b: np.ndarray) -> float:
    mu_a = float(a.mean())
    mu_b = float(b.mean())
    var_a = float(a.var())
    var_b = float(b.var())
    covariance = float(((a - mu_a) * (b - mu_b)).mean())
    c1 = (0.01 * 255.0) ** 2
    c2 = (0.03 * 255.0) ** 2
    numerator = (2 * mu_a * mu_b + c1) * (2 * covariance + c2)
    denominator = (mu_a * mu_a + mu_b * mu_b + c1) * (var_a + var_b + c2)
    return numerator / denominator if denominator else 1.0


def foreground_iou(a: np.ndarray, b: np.ndarray, threshold: int) -> float:
    mask_a = a < threshold
    mask_b = b < threshold
    union = np.logical_or(mask_a, mask_b).sum()
    return float(np.logical_and(mask_a, mask_b).sum() / union) if union else 1.0


def components(mask: np.ndarray) -> list[np.ndarray]:
    height, width = mask.shape
    seen = np.zeros_like(mask, dtype=bool)
    found: list[np.ndarray] = []
    for y in range(height):
        for x in range(width):
            if not mask[y, x] or seen[y, x]:
                continue
            points: list[tuple[int, int]] = []
            queue = deque([(y, x)])
            seen[y, x] = True
            while queue:
                cy, cx = queue.popleft()
                points.append((cy, cx))
                for dy, dx in ((-1, 0), (1, 0), (0, -1), (0, 1)):
                    ny, nx = cy + dy, cx + dx
                    if 0 <= ny < height and 0 <= nx < width and mask[ny, nx] and not seen[ny, nx]:
                        seen[ny, nx] = True
                        queue.append((ny, nx))
            component = np.zeros_like(mask, dtype=bool)
            ys, xs = zip(*points)
            component[np.asarray(ys), np.asarray(xs)] = True
            found.append(component)
    return found


def missing_major_components(standard: np.ndarray, generated: np.ndarray, threshold: int) -> int:
    source_mask = standard < threshold
    target_mask = generated < threshold
    foreground = int(source_mask.sum())
    minimum_area = max(4, int(round(foreground * 0.01)))
    missing = 0
    for component in components(source_mask):
        area = int(component.sum())
        if area < minimum_area:
            continue
        overlap = int(np.logical_and(component, target_mask).sum())
        if overlap / area < 0.5:
            missing += 1
    return missing


def registered_metrics(standard: np.ndarray, generated: np.ndarray, threshold: int,
                       radius: int) -> tuple[float, float, int, int, int]:
    best = None
    for dy in range(-radius, radius + 1):
        for dx in range(-radius, radius + 1):
            candidate = shift_on_canvas(generated, dy, dx)
            iou = foreground_iou(standard, candidate, threshold)
            ssim = global_ssim(standard, candidate)
            rank = (iou, ssim, -abs(dy) - abs(dx))
            if best is None or rank > best[0]:
                best = (rank, ssim, iou, candidate, dy, dx)
    assert best is not None
    _, ssim, iou, candidate, dy, dx = best
    return ssim, iou, missing_major_components(standard, candidate, threshold), dy, dx


def main() -> None:
    parser = argparse.ArgumentParser(description="Strict MathType standard WMF visual comparison.")
    parser.add_argument("standard", type=Path)
    parser.add_argument("generated", type=Path)
    parser.add_argument("out", type=Path)
    parser.add_argument("--magick", default="magick")
    parser.add_argument("--density", type=int, default=144)
    parser.add_argument("--white-threshold", type=int, default=250)
    parser.add_argument("--min-ssim", type=float, default=0.90)
    parser.add_argument("--min-iou", type=float, default=0.85)
    parser.add_argument("--registration-radius", type=int, choices=range(0, 3), default=1)
    parser.add_argument("--no-fail", action="store_true")
    args = parser.parse_args()

    standard = extract_object_wmfs(args.standard)
    generated = extract_object_wmfs(args.generated)
    count = max(len(standard), len(generated))
    rows = []
    passed = 0
    for index in range(count):
        if (index >= len(standard) or index >= len(generated)
                or standard[index] is None or generated[index] is None):
            rows.append({"index": index, "passed": False, "error": "missing paired WMF"})
            continue
        standard_item = standard[index]
        generated_item = generated[index]
        assert standard_item is not None and generated_item is not None
        try:
            source_image = render_wmf(standard_item[1], args.magick, args.density)
            generated_image = render_wmf(generated_item[1], args.magick, args.density)
            source_image, generated_image = common_canvas(source_image, generated_image)
            ssim, iou, missing, shift_y, shift_x = registered_metrics(
                source_image, generated_image, args.white_threshold, args.registration_radius)
            formula_passed = ssim >= args.min_ssim and iou >= args.min_iou and missing == 0
            passed += int(formula_passed)
            rows.append({
                "index": index,
                "standardMedia": standard_item[0],
                "generatedMedia": generated_item[0],
                "ssim": round(ssim, 6),
                "foregroundIou": round(iou, 6),
                "missingMajorComponents": missing,
                "registrationShift": {"x": shift_x, "y": shift_y},
                "passed": formula_passed,
            })
        except Exception as exception:  # report every formula instead of aborting the batch
            rows.append({"index": index, "passed": False, "error": str(exception)})

    report = {
        "schemaVersion": 1,
        "thresholds": {"ssim": args.min_ssim, "foregroundIou": args.min_iou,
                       "missingMajorComponents": 0,
                       "registrationRadiusPx": args.registration_radius},
        "standardWmfCount": sum(item is not None for item in standard),
        "generatedWmfCount": sum(item is not None for item in generated),
        "pairedCount": sum(
            standard[index] is not None and generated[index] is not None
            for index in range(min(len(standard), len(generated)))),
        "passedCount": passed,
        "failedCount": count - passed,
        "formulas": rows,
    }
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({key: report[key] for key in (
        "standardWmfCount", "generatedWmfCount", "pairedCount", "passedCount", "failedCount")},
        ensure_ascii=False))
    if report["failedCount"] and not args.no_fail:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
