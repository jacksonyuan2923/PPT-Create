#!/usr/bin/env python3
"""Typography engine for pro-ppt-studio V4.

Responsibilities:

* discover which font families are actually installed (registry + font folder)
  so a declared stack is never a promise the renderer cannot keep;
* apply a full run style (CJK + Latin typeface, size, weight, colour,
  tracking, caps) in one call;
* expose the paragraph/body properties that make typography feel designed
  instead of default: line spacing, paragraph spacing, autofit, anchoring.
"""

from __future__ import annotations

import functools
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Pt

from design_tokens import FontSpec, hex_to_rgb

# RIBBI qualifiers that belong to the *same* GDI family as their base face.
_RIBBI = (
    " Bold Italic",
    " Bold",
    " Italic",
    " Regular",
    " BoldItalic",
    "Italic",
)

# Fonts that are guaranteed to exist on a Windows + Office CJK machine.
SAFE_FAMILIES = (
    "Microsoft YaHei",
    "Microsoft YaHei Light",
    "SimSun",
    "SimHei",
    "KaiTi",
    "Arial",
    "Times New Roman",
    "Segoe UI",
)


def _strip_ribbi(face: str) -> str:
    for qualifier in _RIBBI:
        if face.endswith(qualifier):
            return face[: -len(qualifier)].strip()
    return face


@functools.lru_cache(maxsize=1)
def installed_families() -> frozenset[str]:
    """Return the set of font families visible to the operating system."""

    families: set[str] = set(SAFE_FAMILIES)
    if sys.platform != "win32":
        return frozenset(families)

    try:
        import winreg

        path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
            index = 0
            while True:
                try:
                    name, _value, _type = winreg.EnumValue(key, index)
                except OSError:
                    break
                index += 1
                face = re.sub(r"\s*\(TrueType\)|\s*\(OpenType\)|\s*\(OpenType\)|$", "", name)
                for part in face.split("&"):
                    part = part.strip()
                    if not part:
                        continue
                    families.add(part)
                    families.add(_strip_ribbi(part))
    except OSError:
        pass

    fonts_dir = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    if fonts_dir.is_dir():
        try:
            from PIL import ImageFont
        except ImportError:  # pragma: no cover - Pillow is a hard dependency
            ImageFont = None
        for font_file in fonts_dir.iterdir():
            if font_file.suffix.lower() not in {".ttf", ".ttc", ".otf"}:
                continue
            stem = font_file.stem.strip()
            if stem:
                families.add(stem)
                families.add(_strip_ribbi(stem))
            if ImageFont is None:
                continue
            try:
                face = ImageFont.truetype(str(font_file), 12)
                family, style = face.getname()
            except Exception:
                continue
            if not family:
                continue
            families.add(family)
            families.add(_strip_ribbi(family))
            if style and style.strip().lower() not in {
                "regular",
                "bold",
                "italic",
                "bold italic",
            }:
                families.add(f"{family} {style.strip()}")

    return frozenset(families)


def is_available(family: str) -> bool:
    """Whether a font family looks installed."""

    if not family:
        return False
    available = installed_families()
    if family in available:
        return True
    lowered = family.lower()
    return any(item.lower() == lowered for item in available)


@dataclass
class FontResolution:
    """The outcome of resolving a requested family against the machine."""

    requested: str
    resolved: str
    substituted: bool
    reason: str = ""

    def as_dict(self) -> dict:
        return {
            "requested": self.requested,
            "resolved": self.resolved,
            "substituted": self.substituted,
            "reason": self.reason,
        }


class FontResolver:
    """Resolve requested families to installed ones, recording substitutions."""

    def __init__(self, level: str | None = None) -> None:
        self.level = level.upper() if level else None
        self.substitutions: list[FontResolution] = []

    def resolve(self, spec: FontSpec) -> FontResolution:
        candidates = (spec.family,) + tuple(spec.fallback)
        for candidate in candidates:
            if is_available(candidate):
                resolution = FontResolution(
                    requested=spec.family,
                    resolved=candidate,
                    substituted=candidate != spec.family,
                    reason="" if candidate == spec.family else "fallback",
                )
                if resolution.substituted:
                    self.substitutions.append(resolution)
                return resolution
        resolution = FontResolution(
            requested=spec.family,
            resolved="Microsoft YaHei",
            substituted=True,
            reason="no candidate installed",
        )
        self.substitutions.append(resolution)
        return resolution

    def report(self) -> list[dict]:
        return [item.as_dict() for item in self.substitutions]


def set_run_fonts(run, family: str, latin: str | None = None) -> None:
    """Set the East-Asian, Latin and complex-script typefaces of a run."""

    r_pr = run._r.get_or_add_rPr()
    for tag, typeface in (
        ("a:latin", latin or family),
        ("a:ea", family),
        ("a:cs", latin or family),
    ):
        node = r_pr.find(qn(tag))
        if node is None:
            node = r_pr.makeelement(qn(tag), {})
            r_pr.append(node)
        node.set("typeface", typeface)


def set_tracking(run, points: float) -> None:
    """Set letter spacing (tracking) in points."""

    if not points:
        return
    r_pr = run._r.get_or_add_rPr()
    r_pr.set("spc", str(int(round(points * 100))))


def set_caps(run, mode: str | None) -> None:
    """Apply all-caps / small-caps to a run."""

    if not mode:
        return
    r_pr = run._r.get_or_add_rPr()
    r_pr.set("cap", {"all": "all", "small": "small"}.get(mode, "none"))


def style_run(
    run,
    *,
    family: str,
    size_pt: float,
    color: str,
    bold: bool = False,
    italic: bool = False,
    tracking: float = 0.0,
    caps: str | None = None,
    latin: str | None = None,
) -> None:
    """Apply a complete run style in a single call."""

    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(*hex_to_rgb(color))
    set_run_fonts(run, family, latin)
    set_tracking(run, tracking)
    set_caps(run, caps)


def set_text_frame(
    shape,
    *,
    margin_left: float = 0.04,
    margin_right: float = 0.04,
    margin_top: float = 0.02,
    margin_bottom: float = 0.02,
    word_wrap: bool = True,
    vertical_anchor: str = "top",
    autofit: str | None = None,
) -> None:
    """Normalise text frame geometry and behaviour."""

    frame = shape.text_frame
    frame.word_wrap = word_wrap
    frame.margin_left = Inches_(margin_left)
    frame.margin_right = Inches_(margin_right)
    frame.margin_top = Inches_(margin_top)
    frame.margin_bottom = Inches_(margin_bottom)
    if vertical_anchor == "middle":
        frame.vertical_anchor = _MSO_ANCHOR.MIDDLE
    elif vertical_anchor == "bottom":
        frame.vertical_anchor = _MSO_ANCHOR.BOTTOM
    else:
        frame.vertical_anchor = _MSO_ANCHOR.TOP
    if autofit == "shrink":
        frame.auto_size = _MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    elif autofit == "grow":
        frame.auto_size = _MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT


def Inches_(value: float):
    from pptx.util import Inches

    return Inches(value)


from pptx.enum.text import MSO_ANCHOR as _MSO_ANCHOR  # noqa: E402
from pptx.enum.text import MSO_AUTO_SIZE as _MSO_AUTO_SIZE  # noqa: E402


def style_paragraph(
    paragraph,
    *,
    align: str = "left",
    line_spacing: float = 1.25,
    space_before_pt: float = 0.0,
    space_after_pt: float = 6.0,
    indent_level: int = 0,
) -> None:
    """Apply paragraph level typography."""

    from pptx.enum.text import PP_ALIGN

    mapping = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY,
    }
    paragraph.alignment = mapping.get(align, PP_ALIGN.LEFT)
    paragraph.line_spacing = line_spacing
    paragraph.space_before = Pt(space_before_pt)
    paragraph.space_after = Pt(space_after_pt)
    if indent_level:
        paragraph.level = indent_level


def measure_text_height(text: str, *, size_pt: float, width_inches: float, line_spacing: float = 1.25) -> float:
    """Rough text height estimate in inches (CJK aware).

    CJK glyphs occupy roughly one em; latin glyphs about half an em. The
    estimate is deliberately conservative so callers can reserve enough room.
    """

    if not text:
        return 0.0
    em_inches = size_pt / 72.0
    units = 0.0
    for char in text:
        units += 1.0 if ord(char) > 0x2E80 else 0.5
    per_line = max(1.0, (width_inches * 72.0) / max(size_pt, 1.0))
    lines = max(1, int(units / per_line) + (1 if units % per_line else 0))
    return lines * em_inches * line_spacing


def fit_font_size(
    text: str,
    *,
    size_pt: float,
    min_size_pt: float,
    width_inches: float,
    height_inches: float,
    line_spacing: float = 1.25,
) -> float:
    """Shrink a font size until the estimated text fits the given box."""

    size = size_pt
    while size > min_size_pt:
        height = measure_text_height(
            text, size_pt=size, width_inches=width_inches, line_spacing=line_spacing
        )
        if height <= height_inches:
            return round(size, 1)
        size -= 0.5
    return round(min_size_pt, 1)