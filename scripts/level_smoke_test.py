#!/usr/bin/env python3
"""Build and audit V4.0.0 design-gene, template and motion coverage."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

from level_profiles import (
    LEVEL_PROFILES,
    SKILL_VERSION,
    get_profile,
    validate_profile_matrix,
)
from leveled_pptx import LeveledDeckBuilder, build_strategy_sample
from check_pptx import compare_decks
from template_parser import parse_template
from style_migrator import migrate_template
import dynamic_pptx


L4_BRANCHES = (
    "technology_launch",
    "luxury_brand",
    "cultural_creative",
    "product_roadshow",
)


def make_test_image(path: Path) -> Path:
    image = Image.new("RGB", (1600, 900), "#24435A")
    for x in range(1600):
        ratio = x / 1599
        image.putpixel((x, 220), (int(30 + 80 * ratio), int(80 + 50 * ratio), 110))
    image.save(path)
    return path


def inspect_deck(path: Path) -> dict:
    presentation = Presentation(str(path))
    fonts: set[str] = set()
    font_sizes: set[float] = set()
    picture_counts: list[int] = []
    picture_crop_max = 0.0
    chart_count = 0
    table_count = 0
    charts: list[dict] = []
    rounded_shape_count = 0
    shape_counts: list[int] = []
    slide_text_lengths: list[int] = []
    fill_colors: set[str] = set()
    transitions: list[str | None] = []
    animated_slides = 0
    animation_steps: list[int] = []
    for slide in presentation.slides:
        transition = dynamic_pptx.transition_summary(slide)
        animation = dynamic_pptx.animation_summary(slide)
        transitions.append(transition.get("transition"))
        animated_slides += int(bool(animation.get("animated")))
        animation_steps.append(int(animation.get("steps", 0)))
        pictures = 0
        shape_counts.append(len(slide.shapes))
        slide_text = []
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pictures += 1
                picture_crop_max = max(
                    picture_crop_max,
                    float(shape.crop_left),
                    float(shape.crop_right),
                    float(shape.crop_top),
                    float(shape.crop_bottom),
                )
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table_count += 1
            if shape.shape_type == MSO_SHAPE_TYPE.CHART:
                chart_count += 1
                chart = shape.chart
                charts.append(
                    {
                        "type": str(chart.chart_type),
                        "legend": bool(chart.has_legend),
                        "gridlines": bool(chart.value_axis.has_major_gridlines),
                        "data_labels": bool(chart.plots[0].has_data_labels),
                        "minimum_scale": float(chart.value_axis.minimum_scale),
                    }
                )
                if getattr(chart, "font", None) and chart.font.name:
                    fonts.add(chart.font.name)
                    if chart.font.size:
                        font_sizes.add(round(float(chart.font.size.pt), 1))
            if getattr(shape, "has_text_frame", False):
                slide_text.append(shape.text.strip())
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.name:
                            fonts.add(run.font.name)
                        if run.font.size:
                            font_sizes.add(round(float(run.font.size.pt), 1))
            try:
                if "ROUNDED_RECTANGLE" in str(shape.auto_shape_type):
                    rounded_shape_count += 1
            except Exception:
                pass
            try:
                if shape.fill.type is not None:
                    fill_colors.add(str(shape.fill.fore_color.rgb))
            except Exception:
                pass
        picture_counts.append(pictures)
        slide_text_lengths.append(len("\n".join(item for item in slide_text if item)))
    return {
        "core_title": presentation.core_properties.title or "",
        "core_keywords": presentation.core_properties.keywords or "",
        "slide_count": len(presentation.slides),
        "font_families": sorted(fonts),
        "font_family_count": len(fonts),
        "font_sizes": sorted(font_sizes),
        "font_size_count": len(font_sizes),
        "picture_counts": picture_counts,
        "picture_crop_max": round(picture_crop_max, 4),
        "chart_count": chart_count,
        "table_count": table_count,
        "charts": charts,
        "rounded_shape_count": rounded_shape_count,
        "shape_counts": shape_counts,
        "slide_text_lengths": slide_text_lengths,
        "fill_colors": sorted(fill_colors),
        "transitions": transitions,
        "animated_slides": animated_slides,
        "animation_steps": animation_steps,
    }


def layout_sequence_check(audit: list[dict], maximum: int) -> list[str]:
    errors: list[str] = []
    run = 0
    previous = None
    for slide in audit:
        layout = slide["layout"] if isinstance(slide, dict) else str(slide)
        run = run + 1 if layout == previous else 1
        previous = layout
        if run > maximum:
            errors.append(
                f"layout {layout!r} repeated {run} times consecutively; max={maximum}"
            )
    return errors


def boundary_rejection_checks() -> dict:
    checks = {}
    try:
        builder = LeveledDeckBuilder("L1")
        builder.add_process_steps("Forbidden", [("A", "B")])
    except ValueError as exc:
        checks["L1_rejects_process_steps"] = str(exc)
    else:
        raise AssertionError("L1 unexpectedly accepted a process component")

    try:
        builder = LeveledDeckBuilder("L2")
        builder.add_data_dashboard(
            "Forbidden",
            [("Metric", "1", "Unit")],
            ["A", "B"],
            [("Series", [1, 2])],
        )
    except ValueError as exc:
        checks["L2_rejects_data_dashboard"] = str(exc)
    else:
        raise AssertionError("L2 unexpectedly accepted a data dashboard")
    return checks


def audit_level(
    level: str,
    out_dir: Path,
    image_path: Path,
) -> dict:
    profile = get_profile(level)
    branch = L4_BRANCHES[0] if level == "L4" else None
    filename = (
        f"{level}_{profile.strategy.key}_{branch}_sample.pptx"
        if branch
        else f"{level}_{profile.strategy.key}_sample.pptx"
    )
    artifact_path = out_dir / filename
    result = build_strategy_sample(
        level,
        artifact_path,
        image_path=image_path,
        style_branch=branch,
    )
    audit_meta = result["audit"]
    audit_slides = audit_meta["slides"]
    artifact = inspect_deck(artifact_path)
    errors: list[str] = []

    minimum_slides = {"L1": 6, "L2": 6, "L3": 10, "L4": 8}
    if artifact["slide_count"] < minimum_slides[level]:
        errors.append(
            f"expected at least {minimum_slides[level]} slides, got {artifact['slide_count']}"
        )
    if SKILL_VERSION not in artifact["core_title"]:
        errors.append("PPTX core properties do not record the V4.0.0 skill version")
    if artifact["font_family_count"] > profile.font.max_font_families:
        errors.append(
            f"used {artifact['font_family_count']} font families; "
            f"maximum is {profile.font.max_font_families}"
        )
    if artifact["font_size_count"] > profile.font.hierarchy_levels + 1:
        errors.append(
            f"used {artifact['font_size_count']} font sizes; "
            f"maximum is {profile.font.hierarchy_levels + 1} including metadata"
        )
    if artifact["chart_count"] < 1:
        errors.append("expected at least one native chart")
    for chart in artifact["charts"]:
        if chart["legend"] != profile.charts.show_legend:
            errors.append("chart legend does not match strategy profile")
        if chart["gridlines"] != profile.charts.show_gridlines:
            errors.append("chart gridlines do not match strategy profile")
        if chart["data_labels"] != profile.charts.show_data_labels:
            errors.append("chart data labels do not match strategy profile")
        if profile.charts.axis_zero_based and chart["minimum_scale"] != 0:
            errors.append("chart axis is not zero-based")
    if sum(artifact["picture_counts"]) == 0:
        errors.append("test image was not inserted")

    layouts = [item["layout"] for item in audit_slides]
    forbidden_layouts = [layout for layout in layouts if layout in profile.layout.forbidden_layouts]
    if forbidden_layouts:
        errors.append(f"forbidden layouts used: {forbidden_layouts}")
    errors.extend(
        layout_sequence_check(audit_slides, profile.layout.max_adjacent_same_layout)
    )

    if level == "L1":
        if set(layouts) != {"title_body", "simple_split"}:
            errors.append("L1 used layouts outside its two-layout whitelist")
        if artifact["rounded_shape_count"] > 0:
            errors.append("L1 generated rounded decorative components")
        if profile.information.max_body_chars is not None:
            errors.append("L1 unexpectedly set a body-character truncation budget")
    elif level == "L2":
        if not any(layout == "structured_cards" for layout in layouts):
            errors.append("L2 did not exercise basic card layout")
        if artifact["rounded_shape_count"] > 0:
            errors.append("L2 generated rounded components")
    elif level == "L3":
        required = {
            "metric_cards",
            "data_dashboard",
            "process_steps",
            "timeline",
            "architecture_compare",
            "case_visual",
            "section_transition",
            "icon_text_modules",
        }
        if not required.issubset(set(layouts)):
            errors.append(f"L3 did not cover required layouts: {sorted(required - set(layouts))}")
        if len(set(layouts)) < 7:
            errors.append(f"L3 layout diversity too low: {len(set(layouts))} unique")
        if artifact["rounded_shape_count"] == 0:
            errors.append("L3 did not exercise rounded components")
        if profile.font.numeric_font not in artifact["font_families"]:
            errors.append("L3 monospace numeric font was not used")
    elif level == "L4":
        if audit_meta["style_branch"] != branch:
            errors.append("L4 style branch audit does not match requested branch")
        if len(set(layouts)) < 4:
            errors.append(f"L4 narrative layout diversity too low: {len(set(layouts))}")

    budget = profile.information.max_body_chars
    if budget is not None:
        content_lengths = artifact["slide_text_lengths"][1:-1]
        if content_lengths and max(content_lengths) > budget + 80:
            errors.append(
                f"content text exceeds density allowance: "
                f"{max(content_lengths)} > {budget}+80 metadata allowance"
            )

    return {
        "artifact_path": str(artifact_path),
        "profile_fingerprint": profile.fingerprint(),
        "acceptance_range": list(profile.acceptance_range),
        "strategy": profile.as_dict()["strategy"],
        "font_rule": profile.as_dict()["font"],
        "grid_rule": profile.as_dict()["grid"],
        "component_rule": profile.as_dict()["components"],
        "layout_rule": profile.as_dict()["layout"],
        "information_rule": profile.as_dict()["information"],
            "density_coefficient": profile.information.density_coefficient,
        "artifact_audit": artifact,
        "slide_audit": audit_meta,
        "errors": errors,
    }


def audit_l4_branches(out_dir: Path, image_path: Path) -> dict:
    branch_dir = out_dir / "l4_styles"
    branch_dir.mkdir(parents=True, exist_ok=True)
    report = {}
    hashes = set()
    for branch in L4_BRANCHES:
        output = branch_dir / f"L4_{branch}_style_sample.pptx"
        result = build_strategy_sample(
            "L4",
            output,
            image_path=image_path,
            style_branch=branch,
        )
        digest = hashlib.sha256(output.read_bytes()).hexdigest()
        hashes.add(digest)
        expected = get_profile("L4", branch)
        actual = inspect_deck(output)
        errors = []
        if result["audit"]["style_branch"] != branch:
            errors.append("style branch metadata mismatch")
        if actual["font_family_count"] > expected.font.max_font_families:
            errors.append("style branch exceeded font-family limit")
        if actual["chart_count"] < 1:
            errors.append("style branch did not exercise chart theme")
        if expected.palette["accent"].lstrip("#").upper() not in {
            value.upper() for value in actual["fill_colors"]
        }:
            errors.append("style branch accent color was not detected in generated shapes")
        layouts = [item["layout"] for item in result["audit"]["slides"]]
        errors.extend(
            layout_sequence_check(layouts, expected.layout.max_adjacent_same_layout)
        )
        report[branch] = {
            "artifact_path": str(output),
            "sha256": digest,
            "palette_accent": expected.palette["accent"],
            "font_stack": {
                "title": expected.font.title_font,
                "body": expected.font.body_font,
                "accent": expected.font.accent_font,
            },
            "layouts": layouts,
            "artifact_audit": actual,
            "errors": errors,
        }
    if len(hashes) != 4:
        for item in report.values():
            item["errors"].append("L4 style branch output is not unique")
    report["valid"] = all(not item["errors"] for item in report.values() if isinstance(item, dict))
    return report


def audit_dynamic_levels(out_dir: Path, image_path: Path) -> dict:
    motion_dir = out_dir / "dynamic"
    motion_dir.mkdir(parents=True, exist_ok=True)
    report: dict[str, dict] = {}
    for level in ("L2", "L3", "L4"):
        path = motion_dir / f"{level}_motion_sample.pptx"
        build_strategy_sample(
            level,
            path,
            image_path=image_path,
            style_branch="technology_launch" if level == "L4" else None,
            mode="dynamic",
        )
        artifact = inspect_deck(path)
        rule = get_profile(level).animation
        errors = []
        if artifact["animated_slides"] == 0:
            errors.append("dynamic sample has no animated slides")
        if not any(artifact["transitions"]):
            errors.append("dynamic sample has no transitions")
        invalid = [
            transition
            for transition in artifact["transitions"]
            if transition and transition not in rule.transition_types
        ]
        if invalid:
            errors.append(f"transitions outside permission: {sorted(set(invalid))}")
        maximum = rule.max_element_animations
        if maximum is not None and max(artifact["animation_steps"], default=0) > maximum:
            errors.append(
                f"animation density exceeds {maximum}: "
                f"{max(artifact['animation_steps'], default=0)}"
            )
        report[level] = {
            "artifact_path": str(path),
            "transitions": artifact["transitions"],
            "animated_slides": artifact["animated_slides"],
            "max_animation_steps": max(artifact["animation_steps"], default=0),
            "errors": errors,
        }
    try:
        LeveledDeckBuilder("L1", mode="dynamic")
    except ValueError:
        l1_static_only = True
    else:
        l1_static_only = False
    report["l1_static_only"] = l1_static_only
    report["valid"] = (
        l1_static_only
        and all(
            not value["errors"]
            for key, value in report.items()
            if key.startswith("L") and isinstance(value, dict)
        )
    )
    return report


def audit_template_migration(out_dir: Path, image_path: Path) -> dict:
    source = out_dir / "L4_theme_narrative_technology_launch_sample.pptx"
    if not source.exists():
        build_strategy_sample(
            "L4",
            source,
            image_path=image_path,
            style_branch="technology_launch",
        )
    parsed = parse_template(source)
    migrated_l2 = migrate_template(parsed, "L2")
    migrated_l1 = migrate_template(parsed, "L1")
    output = out_dir / "template_migrated_L2.pptx"
    builder = LeveledDeckBuilder("L2", template_style=parsed)
    builder.add_cover("模板迁移验证", parsed.recommendations.get("level_hint", "L2"))
    builder.add_content(
        "等级边界过滤",
        ["保留模板基础风格。", "移除 L3/L4 专属效果。"],
        content_type="narrative",
    )
    builder.save(output)
    artifact = inspect_deck(output)
    errors = []
    if len(set(migrated_l2["palette"].values())) > get_profile("L2").palette_system.max_colors + 2:
        errors.append("L2 template palette exceeded its colour boundary")
    if len(set(migrated_l2["fonts"].values())) > 2:
        errors.append("L2 template fonts exceeded the two-family boundary")
    if not migrated_l1["discarded"]:
        errors.append("L1 migration did not discard source effects")
    if artifact["rounded_shape_count"] > 0:
        errors.append("L2 template migration introduced rounded cards")
    return {
        "source": str(source),
        "output": str(output),
        "l2_palette": migrated_l2["palette"],
        "l2_fonts": migrated_l2["fonts"],
        "l2_discarded": migrated_l2["discarded"],
        "l1_discarded": migrated_l1["discarded"],
        "artifact_audit": artifact,
        "errors": errors,
        "valid": not errors,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Audit V4.0.0 design-gene, template and motion generation"
    )
    parser.add_argument("out_dir", help="Directory for strategy samples")
    parser.add_argument("--image", dest="image_path", help="Optional local image")
    parser.add_argument("--json", dest="json_path", help="Optional JSON audit output")
    args = parser.parse_args()

    matrix = validate_profile_matrix()
    if not matrix["valid"]:
        raise SystemExit("Invalid profile matrix: " + "; ".join(matrix["errors"]))

    out_dir = Path(args.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    image_path = (
        Path(args.image_path).resolve()
        if args.image_path
        else make_test_image(out_dir / "strategy_test_image.png")
    )
    if not image_path.is_file():
        raise SystemExit(f"Image does not exist: {image_path}")

    report = {
        "version": SKILL_VERSION,
        "profile_matrix": matrix,
        "levels": {},
        "boundary_rejection": boundary_rejection_checks(),
    }
    for level in LEVEL_PROFILES:
        report["levels"][level] = audit_level(level, out_dir, image_path)
    report["l4_style_branches"] = audit_l4_branches(out_dir, image_path)
    report["dynamic_levels"] = audit_dynamic_levels(out_dir, image_path)
    report["template_migration"] = audit_template_migration(out_dir, image_path)
    level_paths = [
        report["levels"][level]["artifact_path"]
        for level in ("L1", "L2", "L3", "L4")
    ]
    report["visual_similarity"] = compare_decks(level_paths)
    report["valid"] = (
        all(not item["errors"] for item in report["levels"].values())
        and report["l4_style_branches"]["valid"]
        and report["dynamic_levels"]["valid"]
        and report["template_migration"]["valid"]
        and report["visual_similarity"]["valid"]
    )

    text = json.dumps(report, ensure_ascii=False, indent=2)
    print(text)
    if args.json_path:
        Path(args.json_path).write_text(text + "\n", encoding="utf-8")
    raise SystemExit(0 if report["valid"] else 1)


if __name__ == "__main__":
    main()
