#!/usr/bin/env python3
"""V4.0.0 toolchain smoke test with render-fallback contract checks.

Usage:
    python smoke_test.py <output_dir> [--allow-unrendered]
        [--environment PATH]
        [--artifact-node PATH --artifact-node-modules PATH]

This is a toolchain test, not a visual-quality benchmark. It verifies that
python-pptx and Pillow can create a file and that the configured renderer can
produce page images for inspection.

It also verifies that 0x80070520 persists a PowerPoint COM environment
limitation and that a subsequent render skips the entire COM backend family.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from leveled_pptx import normalize_chart_axis_ids
from level_profiles import SKILL_VERSION
from render_preview import (
    POWERPOINT_COM_ERROR_CODE,
    POWERPOINT_COM_HINT,
    RenderBackendError,
    render_with_fallback,
)

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def text(slide, value, x, y, w, h, size, color, bold=False):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.LEFT
    run = paragraph.add_run()
    run.text = value
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return shape


def box(slide, x, y, w, h, fill, radius=True):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def make_test_image(path: Path) -> None:
    width, height = 1600, 900
    image = Image.new("RGB", (width, height), "#101722")
    draw = ImageDraw.Draw(image)
    for x in range(width):
        blend = x / max(width - 1, 1)
        color = (
            int(14 + 18 * blend),
            int(30 + 48 * blend),
            int(52 + 78 * blend),
        )
        draw.line((x, 0, x, height), fill=color)
    draw.ellipse((930, -260, 1770, 580), fill="#17B9C8")
    draw.ellipse((1140, 330, 1850, 1040), fill="#FF6E4A")
    image.save(path, quality=95)


def verify_fallback_contract() -> dict:
    """Exercise native failure, partial rejection, and complete promotion."""

    def page(directory: Path, index: int) -> None:
        Image.new("RGB", (80, 45), (30, 60, 90)).save(
            directory / f"slide_{index:02d}.png"
        )

    def fail(directory: Path) -> int:
        raise RenderBackendError(
            "native",
            error_code=POWERPOINT_COM_ERROR_CODE,
            hint=POWERPOINT_COM_HINT,
            detail="simulated 0x80070520 stack",
            stack="simulated 0x80070520 stack",
        )

    def must_not_retry(directory: Path) -> int:
        raise AssertionError("PowerPoint COM retry was incorrectly attempted")

    def partial(directory: Path) -> int:
        page(directory, 1)
        return 1

    def complete(directory: Path) -> int:
        page(directory, 1)
        page(directory, 2)
        return 2

    with tempfile.TemporaryDirectory(prefix="pro_ppt_smoke_contract_") as tmp:
        target = Path(tmp) / "preview"
        target.mkdir()
        for stale_index in range(1, 5):
            page(target, stale_index)
        backend, kind, count, failures = render_with_fallback(
            target,
            [
                ("native", "native_powerpoint", fail),
                ("native-retry", "native_powerpoint", must_not_retry),
                ("partial", "compatible_libreoffice", partial),
                ("alternative", "alternative_artifact_tool", complete),
            ],
            2,
        )
        result = {
            "valid": bool(
                backend == "alternative"
                and kind == "alternative_artifact_tool"
                and count == 2
                and len(failures) == 3
                and any(
                    item.get("backend") == "native-retry"
                    and item.get("attempted") is False
                    and item.get("skipped_retry") is True
                    for item in failures
                )
                and len(list(target.glob("slide_*.png"))) == 2
            ),
            "backend": backend,
            "backend_kind": kind,
            "promoted_pages": count,
            "failure_count": len(failures),
            "failures": failures,
        }
        if not result["valid"]:
            raise RuntimeError(f"Render fallback contract failed: {result}")
        return result


def build_deck(pptx_path: Path, image_path: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = f"pro-ppt-studio {SKILL_VERSION} smoke test"
    prs.core_properties.subject = "Generation, structure, render, and fallback verification"

    dark = RGBColor(0x10, 0x17, 0x22)
    light = RGBColor(0xF8, 0xF5, 0xEF)
    ink = RGBColor(0x17, 0x1A, 0x21)
    accent = RGBColor(0xFF, 0x5A, 0x36)

    cover = prs.slides.add_slide(prs.slide_layouts[6])
    cover.shapes.add_picture(str(image_path), 0, 0, prs.slide_width, prs.slide_height)
    overlay = box(cover, 0, 4.9, 13.333, 2.6, dark, radius=False)
    text(cover, "V4 DESIGN GENE SYSTEM", 0.8, 5.2, 10.0, 0.8, 42, light, True)
    text(cover, "Toolchain verification deck", 0.85, 6.15, 7.5, 0.5, 18, light)

    data = prs.slides.add_slide(prs.slide_layouts[6])
    data.background.fill.solid()
    data.background.fill.fore_color.rgb = light
    text(data, "Data should lead to a conclusion", 0.75, 0.55, 10.8, 0.7, 30, ink, True)
    chart_data = CategoryChartData()
    chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
    chart_data.add_series("Growth", (18, 27, 39, 54))
    chart = data.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.8), Inches(1.7), Inches(8.1), Inches(4.9), chart_data
    ).chart
    normalize_chart_axis_ids(chart)
    chart.has_legend = False
    chart.value_axis.has_major_gridlines = False
    chart.plots[0].has_data_labels = True
    series = chart.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = accent
    text(data, "54%", 9.55, 2.0, 2.8, 1.0, 58, ink, True)
    text(data, "FY growth index", 9.65, 3.05, 2.8, 0.4, 15, ink)
    text(data, "Illustrative data for render verification only.", 9.65, 5.75, 2.8, 0.8, 11, ink)

    flow = prs.slides.add_slide(prs.slide_layouts[6])
    flow.background.fill.solid()
    flow.background.fill.fore_color.rgb = dark
    text(flow, "Plan → Build → Render → Inspect", 0.75, 0.55, 11.5, 0.8, 30, light, True)
    labels = ("PLAN", "BUILD", "RENDER", "INSPECT")
    for i, label in enumerate(labels):
        x = 0.8 + i * 3.0
        box(flow, x, 2.6, 2.2, 1.6, accent if i == 2 else RGBColor(0x20, 0x35, 0x4A))
        text(flow, label, x + 0.15, 3.0, 1.9, 0.55, 23, light, True)
        if i < len(labels) - 1:
            text(flow, "→", x + 2.25, 3.0, 0.7, 0.55, 26, light, True)

    prs.save(pptx_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Verify local PPT toolchain")
    parser.add_argument("out_dir", help="Directory for smoke-test artifacts")
    parser.add_argument(
        "--allow-unrendered",
        action="store_true",
        help="Return success when generation and static checks pass but rendering is unavailable",
    )
    parser.add_argument("--artifact-node", help="Optional Node executable for Artifact Tool fallback")
    parser.add_argument(
        "--artifact-node-modules",
        help="Optional node_modules directory containing @oai/artifact-tool",
    )
    parser.add_argument(
        "--artifact-helper",
        default=str(Path(__file__).with_name("artifact_tool_render.mjs")),
        help="Path to artifact_tool_render.mjs",
    )
    parser.add_argument(
        "--environment",
        help="environment.json path used to persist PowerPoint COM state",
    )
    args = parser.parse_args()
    fallback_contract = verify_fallback_contract()

    out_dir = Path(args.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    image_path = out_dir / "smoke_image.png"
    pptx_path = out_dir / "smoke_test.pptx"
    preview_dir = out_dir / "preview"
    environment_path = (
        Path(args.environment).resolve()
        if args.environment
        else out_dir / "environment.json"
    )

    make_test_image(image_path)
    build_deck(pptx_path, image_path)

    script_dir = Path(__file__).resolve().parent
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    check = subprocess.run(
        [sys.executable, str(script_dir / "check_pptx.py"), str(pptx_path),
         "--json", str(out_dir / "check_report.json")],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )
    print(check.stdout)
    if check.stderr:
        print(check.stderr, file=sys.stderr)

    render_cmd = [
        sys.executable,
        str(script_dir / "render_preview.py"),
        str(pptx_path),
        str(preview_dir),
    ]
    if args.artifact_node:
        render_cmd.extend(["--artifact-node", args.artifact_node])
    if args.artifact_node_modules:
        render_cmd.extend(["--artifact-node-modules", args.artifact_node_modules])
    if args.artifact_helper:
        render_cmd.extend(["--artifact-helper", args.artifact_helper])
    render_cmd.extend(["--environment", str(environment_path)])
    render = subprocess.run(
        render_cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )
    print(render.stdout)
    if render.stderr:
        print(render.stderr, file=sys.stderr)

    pngs = sorted(preview_dir.glob("slide_*.png"))
    manifest_path = preview_dir / "render_manifest.json"
    render_manifest = {}
    if manifest_path.is_file():
        render_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest_pngs = render_manifest.get("pngs", [])
    expected_pages = render_manifest.get("expected_count")
    rendered_pages = render_manifest.get("rendered_count", 0)
    expected_status_by_kind = {
        "native_powerpoint": "rendered_native_powerpoint",
        "compatible_libreoffice": "rendered_compatible_libreoffice",
        "alternative_artifact_tool": "rendered_alternative_artifact_tool",
    }
    backend_kind = render_manifest.get("backend_kind")
    render_backend_kind = render_manifest.get("render_backend_kind")
    manifest_contract_ok = bool(
        render_manifest.get("skill_version") == SKILL_VERSION
        and
        Path(render_manifest.get("environment_file", "")).resolve() == environment_path
        and
        expected_pages == 3
        and rendered_pages == 3
        and len(manifest_pngs) == 3
        and len(pngs) == 3
        and render_manifest.get("verification_status")
        in {
            "rendered_native_powerpoint",
            "rendered_compatible_libreoffice",
            "rendered_alternative_artifact_tool",
        }
        and render_backend_kind == backend_kind
        and render_backend_kind in expected_status_by_kind
        and render_manifest.get("verification_status")
        == expected_status_by_kind.get(render_backend_kind)
        and (
            render_backend_kind != "alternative_artifact_tool"
            or render_manifest.get("native_powerpoint_verified") is False
        )
        and render_manifest.get("visual_quality_status") == "not_assessed"
    )
    environment_state = {}
    if environment_path.is_file():
        environment_state = json.loads(environment_path.read_text(encoding="utf-8"))
    powerpoint_state = environment_state.get("powerpoint_com") or {}
    failure_records = render_manifest.get("backend_failures", [])
    com_failure_complete = True
    if powerpoint_state:
        com_failure_complete = any(
            item.get("error_code") == POWERPOINT_COM_ERROR_CODE
            and bool(item.get("hint"))
            and bool(item.get("detail"))
            and bool(item.get("stack"))
            for item in failure_records
        )
    environment_contract_ok = (
        not powerpoint_state
        or (
            powerpoint_state.get("status") == "temporary_fault_recorded"
            and powerpoint_state.get("display_status") == "临时故障"
            and powerpoint_state.get("fault_class") == "transient_runtime_startup_failure"
            and powerpoint_state.get("environment_constraint")
            == "permanent_for_current_environment"
            and powerpoint_state.get("environment_constraint_display")
            == "当前环境的永久性限制"
            and powerpoint_state.get("retry_policy") == "skip_powerpoint_com"
            and powerpoint_state.get("auto_retry_allowed") is False
            and powerpoint_state.get("error_code") == POWERPOINT_COM_ERROR_CODE
        )
    )
    cached_skip = {"applicable": bool(powerpoint_state)}
    probe_preservation = {"applicable": bool(powerpoint_state)}
    if powerpoint_state:
        second_dir = out_dir / "preview_cached_skip"
        second_cmd = [
            sys.executable,
            str(script_dir / "render_preview.py"),
            str(pptx_path),
            str(second_dir),
            "--environment",
            str(environment_path),
        ]
        if args.artifact_node:
            second_cmd.extend(["--artifact-node", args.artifact_node])
        if args.artifact_node_modules:
            second_cmd.extend(["--artifact-node-modules", args.artifact_node_modules])
        if args.artifact_helper:
            second_cmd.extend(["--artifact-helper", args.artifact_helper])
        second_render = subprocess.run(
            second_cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
        second_manifest_path = second_dir / "render_manifest.json"
        second_manifest = (
            json.loads(second_manifest_path.read_text(encoding="utf-8"))
            if second_manifest_path.is_file()
            else {}
        )
        skipped_native = [
            item
            for item in second_manifest.get("skipped_backends", [])
            if item.get("backend_kind") == "native_powerpoint"
            and "skip_powerpoint_com" in item.get("reason", "")
        ]
        cached_skip = {
            "applicable": True,
            "exit_code": second_render.returncode,
            "backend": second_manifest.get("backend"),
            "backend_kind": second_manifest.get("backend_kind"),
            "render_backend_kind": second_manifest.get("render_backend_kind"),
            "verification_status": second_manifest.get("verification_status"),
            "powerpoint_com_attempted": second_manifest.get(
                "powerpoint_com_attempted"
            ),
            "native_skipped_entries": len(skipped_native),
            "valid": bool(
                second_render.returncode == 0
                and second_manifest.get("powerpoint_com_attempted") is False
                and second_manifest.get("render_backend_kind") != "native_powerpoint"
                and len(skipped_native) >= 1
            ),
        }
        probe_cmd = [
            sys.executable,
            str(script_dir / "probe_environment.py"),
            "--json",
            str(environment_path),
        ]
        probe_result = subprocess.run(
            probe_cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
        reprobed = (
            json.loads(environment_path.read_text(encoding="utf-8"))
            if environment_path.is_file()
            else {}
        )
        reprobed_state = reprobed.get("powerpoint_com") or {}
        probe_preservation = {
            "applicable": True,
            "exit_code": probe_result.returncode,
            "state_preserved": (
                reprobed_state.get("retry_policy") == "skip_powerpoint_com"
                and reprobed_state.get("error_code") == POWERPOINT_COM_ERROR_CODE
            ),
            "powerpoint_runtime_state": reprobed.get("powerpoint", {}).get(
                "runtime_state"
            ),
            "powerpoint_candidates_removed": not any(
                "PowerPoint COM" in item
                for item in reprobed.get("capabilities", {}).get(
                    "rendering_candidates", []
                )
            ),
        }
        probe_preservation["valid"] = bool(
            probe_preservation["exit_code"] == 0
            and probe_preservation["state_preserved"]
            and probe_preservation["powerpoint_runtime_state"]
            == "permanently_limited_for_current_environment"
            and probe_preservation["powerpoint_candidates_removed"]
        )
    result = {
        "skill_version": SKILL_VERSION,
        "pptx_created": pptx_path.is_file(),
        "check_exit_code": check.returncode,
        "render_exit_code": render.returncode,
        "render_backend": render_manifest.get("backend"),
        "render_backend_kind": render_manifest.get("backend_kind"),
        "render_backend_kind_explicit": render_backend_kind,
        "render_verification_status": render_manifest.get("verification_status", "unverified"),
        "native_powerpoint_verified": render_manifest.get("native_powerpoint_verified"),
        "render_backend_failures": render_manifest.get("backend_failures", []),
        "expected_pages": expected_pages,
        "rendered_pages": len(pngs),
        "manifest_rendered_pages": rendered_pages,
        "manifest_contract_ok": manifest_contract_ok,
        "powerpoint_com_state": powerpoint_state,
        "environment_contract_ok": environment_contract_ok,
        "com_failure_record_complete": com_failure_complete,
        "cached_skip_validation": cached_skip,
        "probe_preservation": probe_preservation,
        "fallback_contract_self_test": fallback_contract,
        "visual_quality_status": render_manifest.get("visual_quality_status", "not_assessed"),
        "contact_sheet": (preview_dir / "contact_sheet.png").is_file(),
    }
    (out_dir / "smoke_result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))

    rendering_ok = (
        result["render_exit_code"] == 0
        and result["rendered_pages"] == 3
        and result["manifest_contract_ok"]
        and result["environment_contract_ok"]
        and result["com_failure_record_complete"]
        and (
            not result["cached_skip_validation"].get("applicable")
            or result["cached_skip_validation"].get("valid") is True
        )
        and (
            not result["probe_preservation"].get("applicable")
            or result["probe_preservation"].get("valid") is True
        )
    )
    if not result["pptx_created"] or result["check_exit_code"] != 0:
        raise SystemExit(1)
    if not rendering_ok and not args.allow_unrendered:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
