#!/usr/bin/env python3
"""Map reference-template signals onto a level-safe design configuration.

A template is evidence about visual language, not permission to cross a level
boundary. This module keeps the useful palette, typography, grid, component and
rhythm signals while explicitly discarding effects that the target level does
not own.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path
from typing import Any, Mapping

from level_profiles import get_profile, normalize_level


PALETTE_ROLES = {
    "L1": ("background", "text", "line"),
    "L2": (
        "background",
        "surface",
        "text",
        "secondary",
        "accent",
        "line",
        "contrast",
    ),
    "L3": (
        "background",
        "surface",
        "text",
        "secondary",
        "accent",
        "line",
        "contrast",
        "highlight",
    ),
    "L4": (
        "background",
        "surface",
        "text",
        "secondary",
        "accent",
        "line",
        "contrast",
        "highlight",
    ),
}

COMPONENT_TOKENS = {
    "basic_card": "basic_card",
    "card": "basic_card",
    "table": "table",
    "chart": "standard_chart",
    "standard_chart": "standard_chart",
    "image_frame": "framed_image",
    "framed_image": "framed_image",
    "icon": "minimal_icon",
    "minimal_icon": "minimal_icon",
    "metric": "metric_card",
    "metric_card": "metric_card",
    "dashboard": "data_dashboard",
    "data_dashboard": "data_dashboard",
    "process": "process_step",
    "process_step": "process_step",
    "timeline": "timeline",
    "collage": "editorial_collage",
    "editorial_collage": "editorial_collage",
    "custom_infographic": "custom_infographic",
}

EFFECT_TOKENS = {
    "shadow",
    "glow",
    "texture",
    "mask",
    "gradient",
    "outline",
    "collage",
    "trigger",
    "motion_path",
}


def _clean_hex(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    match = re.fullmatch(r"#?([0-9a-fA-F]{6})", value.strip())
    return f"#{match.group(1).upper()}" if match else None


def _rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    return (int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def _luminance(value: str) -> float:
    red, green, blue = _rgb(value)
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue


def _unique(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = value.upper()
        if key not in seen:
            seen.add(key)
            result.append(value)
    return result


def _migrate_palette(
    source: Mapping[str, Any], level: str
) -> tuple[dict[str, str], list[str], list[str]]:
    profile = get_profile(level)
    defaults = profile.palette
    accepted: dict[str, str] = {}
    discarded: list[str] = []
    warnings: list[str] = []

    for role in PALETTE_ROLES[level]:
        value = _clean_hex(source.get(role))
        if value:
            accepted[role] = value

    if level == "L1":
        background = accepted.get("background", "#FFFFFF")
        text = accepted.get("text", "#111111")
        if _luminance(background) < _luminance(text):
            background, text = "#FFFFFF", "#000000"
        accepted = {
            "background": background,
            "surface": background,
            "text": text,
            "secondary": text,
            "accent": text,
            "line": background if background.upper() != text.upper() else text,
            "contrast": text,
        }
        if source.get("accent") and _clean_hex(source.get("accent")) not in {
            background,
            text,
        }:
            discarded.append("accent colour")
        return accepted, discarded, warnings

    max_colors = profile.palette_system.max_colors
    allowed_colors: list[str] = []
    for role in PALETTE_ROLES[level]:
        value = accepted.get(role)
        if not value:
            continue
        if len(_unique(allowed_colors + [value])) <= max_colors:
            allowed_colors.append(value)
        else:
            accepted.pop(role, None)
            discarded.append(f"palette.{role}")
    if len(accepted) < 2:
        warnings.append("Template palette had too little usable contrast; level palette was retained.")

    palette = dict(defaults)
    palette.update(accepted)
    return palette, discarded, warnings


def _migrate_fonts(
    source: Mapping[str, Any], level: str, style_branch: str | None
) -> tuple[dict[str, str], list[str]]:
    profile = get_profile(level, style_branch)
    defaults = {
        "title": profile.font.title_font,
        "body": profile.font.body_font,
        "accent": profile.font.accent_font,
        "numeric": profile.font.numeric_font,
    }
    aliases = {
        "title": ("title", "majorFont", "heading", "title_font"),
        "body": ("body", "minorFont", "body_font", "paragraph"),
        "accent": ("accent", "accent_font", "display"),
        "numeric": ("numeric", "numeric_font", "mono"),
    }
    requested: dict[str, str] = {}
    for role, keys in aliases.items():
        for key in keys:
            value = source.get(key)
            if isinstance(value, str) and value.strip():
                requested[role] = value.strip()
                break

    if level == "L1":
        family = requested.get("body") or requested.get("title") or defaults["body"]
        return {key: family for key in defaults}, [
            item for item in requested.values() if item != family
        ]

    selected = dict(defaults)
    families: list[str] = []
    discarded: list[str] = []
    for role in ("title", "body", "accent", "numeric"):
        candidate = requested.get(role, selected[role])
        if candidate not in families and len(families) >= profile.font.max_font_families:
            candidate = selected[role]
            discarded.append(f"font.{role}")
        selected[role] = candidate
        if candidate not in families:
            families.append(candidate)
    return selected, discarded


def _migrate_grid(
    source: Mapping[str, Any], level: str
) -> tuple[dict[str, Any], list[str], list[str]]:
    profile = get_profile(level)
    default = {
        "columns": profile.grid.columns,
        "margin_left_in": profile.grid.margin_left_in,
        "margin_right_in": profile.grid.margin_right_in,
        "margin_top_in": profile.grid.margin_top_in,
        "margin_bottom_in": profile.grid.margin_bottom_in,
        "gutter_in": profile.grid.gutter_in,
        "component_radius_in": profile.grid.component_radius_in,
        "shadow_policy": profile.grid.shadow_policy,
    }
    accepted: dict[str, Any] = {}
    discarded: list[str] = []
    warnings: list[str] = []

    columns = source.get("columns") or source.get("grid_columns")
    if isinstance(columns, (int, float)):
        if level in {"L1", "L2", "L3"} and int(columns) != profile.grid.columns:
            discarded.append("grid.columns")
        elif 2 <= int(columns) <= 24:
            accepted["columns"] = int(columns)

    for field_name in (
        "margin_left_in",
        "margin_right_in",
        "margin_top_in",
        "margin_bottom_in",
        "gutter_in",
    ):
        value = source.get(field_name)
        if isinstance(value, (int, float)) and 0 <= float(value) <= 2:
            accepted[field_name] = round(float(value), 3)

    radius = source.get("component_radius_in")
    if isinstance(radius, (int, float)):
        requested_radius = max(0.0, min(float(radius), 0.35))
        if level == "L1" and requested_radius > 0:
            discarded.append("grid.component_radius_in")
        elif level == "L2" and requested_radius > 0.06:
            discarded.append("grid.component_radius_in")
            accepted["component_radius_in"] = profile.grid.component_radius_in
        elif level == "L3" and requested_radius > 0.15:
            accepted["component_radius_in"] = 0.12
            warnings.append("Rounded corners were capped at the L3 limit.")
        else:
            accepted["component_radius_in"] = requested_radius

    effects = source.get("effects") or source.get("style_tokens") or []
    if isinstance(effects, str):
        effects = [effects]
    for effect in effects:
        token = str(effect).strip().lower()
        if token in EFFECT_TOKENS and token not in {
            "shadow"
            if profile.grid.shadow_policy != "none"
            else "",
            "texture"
            if profile.font.allow_texture
            else "",
            "gradient"
            if profile.charts.allow_gradient
            else "",
            "outline"
            if profile.font.allow_outline
            else "",
            "mask"
            if profile.images.allow_mask
            else "",
            "collage"
            if profile.images.allow_collage
            else "",
            "trigger"
            if profile.animation.allow_trigger
            else "",
            "motion_path"
            if profile.animation.allow_motion_path
            else "",
        }:
            discarded.append(f"effect.{token}")

    default.update(accepted)
    return default, discarded, warnings


def _migrate_components(
    source: Any, level: str
) -> tuple[list[str], list[str]]:
    profile = get_profile(level)
    if isinstance(source, str):
        items = [source]
    elif isinstance(source, Mapping):
        items = list(source.keys()) + [str(value) for value in source.values()]
    elif isinstance(source, list):
        items = source
    else:
        items = []
    accepted: list[str] = []
    discarded: list[str] = []
    for raw in items:
        token = str(raw).strip().lower().replace("-", "_").replace(" ", "_")
        mapped = COMPONENT_TOKENS.get(token, token)
        if (
            mapped in profile.components.allowed_components
            and mapped not in profile.components.forbidden_components
            and mapped not in accepted
        ):
            accepted.append(mapped)
        else:
            discarded.append(f"component.{token}")
    return accepted, discarded


def _migrate_motion(
    source: Mapping[str, Any], level: str
) -> tuple[dict[str, Any], list[str], list[str]]:
    profile = get_profile(level)
    motion = get_profile(level).animation
    defaults = {
        "default_transition": motion.default_transition,
        "transition_types": list(motion.transition_types),
        "entrance_types": list(motion.entrance_types),
        "max_element_animations": motion.max_element_animations,
        "stagger_ms": motion.stagger_ms,
    }
    requested = source.get("motion") if isinstance(source.get("motion"), Mapping) else source
    discarded: list[str] = []
    warnings: list[str] = []
    if not profile.animation.transition_types:
        if requested:
            discarded.append("motion")
        return defaults, discarded, warnings

    requested_transition = str(
        requested.get("default_transition")
        or requested.get("transition_hint")
        or ""
    ).strip().lower().replace(" ", "_")
    if requested_transition and requested_transition in motion.transition_types:
        defaults["default_transition"] = requested_transition
    elif requested_transition:
        discarded.append(f"motion.transition.{requested_transition}")
        warnings.append("Template transition was outside the target level and was remapped.")
    duration = requested.get("auto_advance_ms")
    if isinstance(duration, int) and 1000 <= duration <= 60000:
        defaults["auto_advance_ms"] = duration
    return defaults, discarded, warnings


def _source_dict(template: Any) -> dict[str, Any]:
    if isinstance(template, Mapping):
        return dict(template)
    if hasattr(template, "as_dict"):
        return dict(template.as_dict())
    raise TypeError("template must be a mapping or expose as_dict()")


def migrate_template(
    template: Any,
    level: str,
    *,
    style_branch: str | None = None,
) -> dict[str, Any]:
    """Return a level-safe style config consumable by ``LeveledDeckBuilder``."""

    normalized = normalize_level(level)
    source = _source_dict(template)
    source_palette = source.get("palette") or {}
    source_fonts = source.get("fonts") or source.get("font") or {}
    source_grid = source.get("grid") or source.get("layout") or {}
    source_rhythm = source.get("rhythm") or {}
    if not isinstance(source_palette, Mapping):
        source_palette = {}
    if not isinstance(source_fonts, Mapping):
        source_fonts = {}
    if not isinstance(source_grid, Mapping):
        source_grid = {}
    if not isinstance(source_rhythm, Mapping):
        source_rhythm = {}

    palette, palette_discarded, palette_warnings = _migrate_palette(
        source_palette, normalized
    )
    fonts, font_discarded = _migrate_fonts(source_fonts, normalized, style_branch)
    grid, grid_discarded, grid_warnings = _migrate_grid(source_grid, normalized)
    components, component_discarded = _migrate_components(
        source.get("components") or [], normalized
    )
    motion, motion_discarded, motion_warnings = _migrate_motion(
        {**source_rhythm, **source}, normalized
    )

    retained = {
        "palette_roles": sorted(
            role
            for role in palette
            if source_palette.get(role)
            and str(palette[role]).upper()
            == str(_clean_hex(source_palette.get(role)) or "").upper()
        ),
        "fonts": fonts,
        "grid": grid,
        "components": components,
        "motion": motion,
        "rhythm": {
            "archetype": source_rhythm.get("archetype")
            or source.get("layout_archetype"),
            "density_curve": list(
                get_profile(normalized).design_gene.information_density_curve
            ),
        },
    }
    discarded = _unique(
        palette_discarded
        + font_discarded
        + grid_discarded
        + component_discarded
        + motion_discarded
    )
    payload = {
        "skill_version": "4.0.0",
        "target_level": normalized,
        "style_branch": style_branch,
        "source_kind": source.get("kind"),
        "palette": palette,
        "fonts": fonts,
        "grid": grid,
        "components": components,
        "rhythm": retained["rhythm"],
        "motion": motion,
        "retained": retained,
        "discarded": discarded,
        "warnings": palette_warnings + grid_warnings + motion_warnings,
        "confidence": round(max(0.15, 1.0 - len(discarded) / 24), 2),
    }
    fingerprint_payload = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    payload["fingerprint"] = hashlib.sha256(
        fingerprint_payload.encode("utf-8")
    ).hexdigest()[:16]
    return payload


def migrate_file(
    template_path: str | Path,
    level: str,
    *,
    style_branch: str | None = None,
) -> dict[str, Any]:
    from template_parser import parse_template

    return migrate_template(
        parse_template(template_path), level, style_branch=style_branch
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Migrate a template style under a pro-ppt-studio level boundary"
    )
    parser.add_argument("template", help="PPTX/PDF/image/video/JSON template path")
    parser.add_argument("level", choices=["L1", "L2", "L3", "L4"])
    parser.add_argument("--branch", help="L4 style branch")
    parser.add_argument("--json", dest="json_path", help="Write migrated style JSON")
    args = parser.parse_args()
    payload = migrate_file(args.template, args.level, style_branch=args.branch)
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    print(text)
    if args.json_path:
        output = Path(args.json_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(text + "\n", encoding="utf-8")
        print(f"JSON report: {output.resolve()}")


if __name__ == "__main__":
    main()
