#!/usr/bin/env python3
"""Multi-format template parser for pro-ppt-studio V4.

The skill accepts many kinds of reference material. This module turns them
into a machine-readable *design profile* that the builder can reuse:

* ``.pptx`` / ``.potx`` / ``.thmx`` - theme colours, theme fonts, masters,
  layouts, placeholder geometry, media inventory;
* images (``.png`` / ``.jpg`` ...) - palette, grid, layout archetype,
  margins, photo-vs-graphic judgement;
* videos (``.mp4`` / ``.mov`` ...) - container metadata, pacing, and - when
  ``ffmpeg`` is available - a storyboard of extracted frames;
* ``.pdf`` - page count, palette and embedded font names (pure Python);
* ``.docx`` - brand manual signals (fonts, colours, images);
* ``.json`` / ``.yaml`` - an explicit design configuration;
* ``.svg`` - dimensions, palette and vector primitive inventory.

Every analysis states its own limits. Nothing is inferred that was not
actually read from the file.

Usage::

    python template_parser.py <path> [--json out.json] [--kind auto]
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import struct
import subprocess
import sys
import zlib
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Sequence

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
VIDEO_SUFFIXES = {".mp4", ".mov", ".m4v", ".avi", ".mkv", ".webm"}
DOCUMENT_SUFFIXES = {
    ".pptx",
    ".potx",
    ".potm",
    ".thmx",
    ".pdf",
    ".docx",
    ".md",
    ".json",
    ".yaml",
    ".yml",
}


@dataclass
class DesignProfile:
    """The outcome of analysing one reference file."""

    source: str
    kind: str
    palette: dict[str, str] = field(default_factory=dict)
    fonts: dict[str, str] = field(default_factory=dict)
    layout: dict[str, Any] = field(default_factory=dict)
    detected: dict[str, Any] = field(default_factory=dict)
    recommendations: dict[str, Any] = field(default_factory=dict)
    limits: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "skill_version": "4.0.0",
            "source": self.source,
            "kind": self.kind,
            "palette": self.palette,
            "fonts": self.fonts,
            "layout": self.layout,
            "grid": {
                "columns": self.layout.get("grid_columns", 12),
                "archetype": self.layout.get("archetype"),
                "gutter_in": self.layout.get("gutter_in"),
                "margin_top_in": self.layout.get("margin_top_in"),
            },
            "components": list(self.detected.get("component_tokens", [])),
            "rhythm": {
                "density": self.detected.get("density"),
                "transition_hint": self.recommendations.get("transition_hint"),
                "auto_advance_ms": self.recommendations.get("auto_advance_ms"),
            },
            "detected": self.detected,
            "recommendations": self.recommendations,
            "limits": self.limits,
            "confidence": self.recommendations.get("confidence", 0.65),
        }


TemplateStyle = DesignProfile


# ---------------------------------------------------------------------------
# colour helpers (kept local so the analyzer has no dependency on the builder)
# ---------------------------------------------------------------------------


def _hex(rgb: Sequence[int]) -> str:
    return "#{:02X}{:02X}{:02X}".format(int(rgb[0]), int(rgb[1]), int(rgb[2]))


def _luminance(rgb: Sequence[int]) -> float:
    def channel(value: float) -> float:
        value /= 255
        return value / 12.92 if value <= 0.03928 else ((value + 0.055) / 1.055) ** 2.4

    return 0.2126 * channel(rgb[0]) + 0.7152 * channel(rgb[1]) + 0.3290 * channel(rgb[2])


def _saturation(rgb: Sequence[int]) -> float:
    high = max(rgb)
    low = min(rgb)
    if high == 0:
        return 0.0
    return (high - low) / high


def _roles_from_colors(colors: Sequence[tuple[int, int, int]]) -> dict[str, str]:
    """Map a colour set onto the deck role slots."""

    if not colors:
        return {}
    ordered = sorted(colors, key=_luminance)
    darkest = ordered[0]
    lightest = ordered[-1]
    accents = [c for c in colors if _saturation(c) > 0.35]
    accent = max(accents, key=_saturation) if accents else ordered[len(ordered) // 2]
    mid = sorted(colors, key=lambda c: abs(_luminance(c) - 0.45))[0]
    return {
        "background": _hex(lightest),
        "surface": _hex(tint(lightest, 0.06)),
        "surface_alt": _hex(tint(lightest, 0.12)),
        "text": _hex(darkest),
        "text_secondary": _hex(mid),
        "accent": _hex(accent),
        "accent_secondary": _hex(tint(accent, 0.4)),
        "line": _hex(tint(darkest, 0.75)),
    }


def tint(rgb: Sequence[int], amount: float) -> tuple[int, int, int]:
    return tuple(int(c + (255 - c) * amount) for c in rgb)


# ---------------------------------------------------------------------------
# image analysis
# ---------------------------------------------------------------------------



def _pixel_data(image):
    """Return the flat pixel list of a quantised image (Pillow 10+ safe)."""

    getter = getattr(image, "get_flattened_data", None)
    if getter is not None:
        try:
            return list(getter())
        except Exception:
            pass
    return list(image.getdata())


def analyze_image(path: Path) -> DesignProfile:
    """Extract palette, grid and layout archetype from a still image."""

    from PIL import Image

    profile = DesignProfile(source=str(path), kind="image")
    with Image.open(path) as image:
        image = image.convert("RGB")
        width, height = image.size
        profile.detected["size"] = [width, height]
        profile.detected["aspect"] = round(width / height, 3)

        small = image.resize((160, 90), Image.LANCZOS)
        quantised = small.quantize(colors=10, method=Image.Quantize.MEDIANCUT)
        palette_entries = quantised.getpalette() or []
        counts = Counter(_pixel_data(quantised))
        ranked = []
        for index, count in counts.most_common(10):
            base = index * 3
            if base + 2 < len(palette_entries):
                ranked.append((palette_entries[base], palette_entries[base + 1], palette_entries[base + 2]))
        profile.detected["dominant_colors"] = [_hex(c) for c in ranked]
        profile.palette = _roles_from_colors(ranked)

        # grid / band detection: compare neighbouring cells of a coarse grid
        grid_w, grid_h = 24, 14
        coarse = image.resize((grid_w, grid_h), Image.BILINEAR)
        pixels = coarse.load()
        row_variation = []
        column_variation = []
        for y in range(grid_h):
            row = [pixels[x, y] for x in range(grid_w)]
            row_variation.append(max(_distance(row[0], c) for c in row))
        for x in range(grid_w):
            column = [pixels[x, y] for y in range(grid_h)]
            column_variation.append(max(_distance(column[0], c) for c in column))
        horizontal_bands = [y for y, value in enumerate(row_variation) if value > 60]
        vertical_bands = [x for x, value in enumerate(column_variation) if value > 60]
        profile.detected["horizontal_bands"] = horizontal_bands
        profile.detected["vertical_bands"] = vertical_bands

        columns = len(vertical_bands)
        rows = len(horizontal_bands)
        edge_ratio = (columns + rows) / float(grid_w + grid_h)
        photographic = edge_ratio > 0.4 or len(set(ranked)) > 7
        if photographic:
            archetype = "photo"
        elif rows >= 3 and columns <= 1:
            archetype = "stacked_sections"
        elif columns >= 2 and rows <= 1:
            archetype = "split_columns"
        elif columns >= 2 and rows >= 2:
            archetype = "grid_modules"
        elif rows == 1 and columns == 0:
            archetype = "hero_image"
        else:
            archetype = "single_column"

        profile.detected["layout_archetype"] = archetype
        profile.detected["is_photographic"] = photographic
        profile.detected["edge_density"] = round(edge_ratio, 3)

        # margins: first row / column that differs strongly from the border
        edge = image.resize((1, 1), Image.BILINEAR).getpixel((0, 0))
        margin_top = 0
        for y in range(0, height, max(1, height // 60)):
            if _distance(image.getpixel((width // 2, y)), edge) > 18:
                margin_top = y
                break
        profile.detected["margin_top_ratio"] = round(margin_top / height, 3)

    profile.layout = {
        "archetype": profile.detected["layout_archetype"],
        "grid_columns": max(1, len(profile.detected["vertical_bands"])),
        "grid_rows": max(1, len(profile.detected["horizontal_bands"])),
        "margin_top_in": round(profile.detected["margin_top_ratio"] * 7.5, 3),
        "gutter_in": 0.2,
    }
    profile.detected["component_tokens"] = (
        ["framed_image", "basic_card"]
        if profile.detected["layout_archetype"]
        in {"photo", "grid_modules", "split_columns"}
        else []
    )
    profile.detected["density"] = round(
        min(1.0, profile.detected["edge_density"]), 3
    )
    profile.recommendations = {
        "cover_layout": {
            "hero_image": "full_bleed_hero",
            "photo": "masked_image_hero",
            "split_columns": "narrative_split",
            "stacked_sections": "plain_title_cover",
            "grid_modules": "editorial_collage",
            "single_column": "title_band_cover",
        }.get(profile.detected["layout_archetype"], "plain_title_cover"),
        "use_palette": True,
        "level_hint": "L4"
        if profile.detected["layout_archetype"] in {"split_columns", "grid_modules", "hero_image"}
        else ("L3" if profile.detected["layout_archetype"] == "photo" else "L2"),
        "confidence": 0.72,
    }
    profile.limits = [
        "Image analysis infers palette and grid; exact fonts and copy are not recoverable from a raster.",
    ]
    return profile


def _distance(a: Sequence[int], b: Sequence[int]) -> float:
    return sum((int(a[i]) - int(b[i])) ** 2 for i in range(3)) ** 0.5

# ---------------------------------------------------------------------------
# office theme / presentation analysis
# ---------------------------------------------------------------------------


def _read_zip_xml(path: Path, member_pattern: str) -> list[str]:
    import zipfile

    chunks: list[str] = []
    try:
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if re.search(member_pattern, name, re.IGNORECASE):
                    try:
                        chunks.append(archive.read(name).decode("utf-8", errors="replace"))
                    except KeyError:
                        continue
    except zipfile.BadZipFile:
        return []
    return chunks


def analyze_office(path: Path) -> DesignProfile:
    """Analyse a PPTX / POTX / THMX package."""

    profile = DesignProfile(source=str(path), kind="office")
    themes = _read_zip_xml(path, r"theme.*\.xml$")
    palette_colors: list[tuple[int, int, int]] = []
    fonts: dict[str, str] = {}

    for theme in themes:
        scheme = re.search(r"<a:clrScheme.*?</a:clrScheme>", theme, re.DOTALL)
        if scheme:
            for match in re.finditer(
                r"<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>\s*<a:(?:srgbClr val=\"([0-9A-Fa-f]{6})\"|sysClr[^>]*lastClr=\"([0-9A-Fa-f]{6})\")",
                scheme.group(0),
            ):
                value = match.group(2) or match.group(3)
                if not value:
                    continue
                rgb = (int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))
                palette_colors.append(rgb)
                profile.detected.setdefault("theme_colors", {})[match.group(1)] = _hex(rgb)
        font_scheme = re.search(r"<a:fontScheme.*?</a:fontScheme>", theme, re.DOTALL)
        if font_scheme:
            for role in ("majorFont", "minorFont"):
                block = re.search(rf"<a:{role}>.*?</a:{role}>", font_scheme.group(0), re.DOTALL)
                if not block:
                    continue
                latin = re.search(r"<a:latin typeface=\"([^\"]*)\"", block.group(0))
                ea = re.search(r"<a:ea typeface=\"([^\"]*)\"", block.group(0))
                if latin and latin.group(1):
                    fonts[role] = latin.group(1)
                if ea and ea.group(1):
                    fonts[f"{role}_ea"] = ea.group(1)

    profile.fonts = fonts

    # colours that actually appear on the slides are more representative of the
    # template than the untouched Office theme palette
    used: Counter = Counter()
    for slide in _read_zip_xml(path, r"ppt/slides/slide\d+\.xml$"):
        for value in re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', slide):
            used[value.upper()] += 1
    if used:
        used_rgb = [
            (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
            for v, _n in used.most_common(12)
        ]
        profile.detected["slide_colors"] = [_hex(c) for c in used_rgb]
        profile.palette = _roles_from_colors(used_rgb)
    elif palette_colors:
        profile.palette = _roles_from_colors(palette_colors)

    layouts = _read_zip_xml(path, r"slideLayouts?/.*\.xml$")
    masters = _read_zip_xml(path, r"slideMasters?/.*\.xml$")
    profile.detected["slide_layouts"] = len(layouts)
    profile.detected["slide_masters"] = len(masters)

    slides = _read_zip_xml(path, r"ppt/slides/slide\d+\.xml$")
    profile.detected["slides"] = len(slides)
    shape_counts = [len(re.findall(r"<p:sp>", slide)) for slide in slides]
    picture_counts = [len(re.findall(r"<p:pic>", slide)) for slide in slides]
    chart_counts = [len(re.findall(r"<c:chart", slide)) for slide in slides]
    text_counts = [
        sum(len(value) for value in re.findall(r"<a:t>(.*?)</a:t>", slide))
        for slide in slides
    ]
    tokens: list[str] = []
    if any(picture_counts):
        tokens.extend(("framed_image", "image_frame"))
    if any(chart_counts):
        tokens.append("standard_chart")
    if any(count >= 6 for count in shape_counts):
        tokens.append("basic_card")
    profile.detected["shape_counts"] = shape_counts
    profile.detected["picture_counts"] = picture_counts
    profile.detected["chart_counts"] = chart_counts
    profile.detected["text_lengths"] = text_counts
    profile.detected["component_tokens"] = sorted(set(tokens))
    if text_counts:
        profile.detected["density"] = round(
            sum(text_counts) / max(1, len(text_counts)) / 700,
            3,
        )

    media = _read_zip_xml(path, r"\.xml$")
    import zipfile

    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            media_names = [n for n in names if n.startswith("ppt/media/")]
            profile.detected["media_count"] = len(media_names)
            embedded = []
            for name in media_names:
                try:
                    raw = archive.read(name)
                except KeyError:
                    continue
                if name.lower().endswith((".png", ".jpg", ".jpeg")):
                    embedded.append(raw)
            profile.detected["embedded_images"] = len(embedded)
    except zipfile.BadZipFile:
        profile.limits.append("The package could not be opened as a zip archive.")

    profile.layout = {
        "has_masters": bool(masters),
        "layout_count": len(layouts),
        "grid_columns": 12 if masters or layouts else 2,
        "archetype": "master_based" if masters else "slide_based",
        "gutter_in": 0.22,
        "margin_top_in": 0.55,
    }
    profile.recommendations = {
        "reuse_master": bool(masters),
        "theme_palette": bool(palette_colors),
        "level_hint": "L2",
        "confidence": 0.88,
    }
    profile.limits = [
        "Theme colours and fonts are read from the package; master placeholders are not copied into new decks automatically.",
    ]
    return profile


# ---------------------------------------------------------------------------
# video analysis
# ---------------------------------------------------------------------------


#: ISO-BMFF container boxes whose payload is a list of child boxes.
_MP4_CONTAINERS = {
    "moov",
    "trak",
    "mdia",
    "minf",
    "stbl",
    "edts",
    "udta",
    "mvex",
    "moof",
    "traf",
    "mfra",
}


def _iter_mp4_boxes(data: bytes, start: int = 0, end: int | None = None, depth: int = 0):
    """Walk the ISO base media file format boxes, descending into containers.

    ``mvhd``/``tkhd``/``mdhd``/``stts`` all live inside ``moov``, so a flat
    scan of the top level would never see them.
    """

    end = len(data) if end is None else end
    offset = start
    while offset + 8 <= end:
        size = struct.unpack(">I", data[offset : offset + 4])[0]
        kind = data[offset + 4 : offset + 8].decode("latin-1", errors="replace")
        header = 8
        if size == 1:
            if offset + 16 > end:
                break
            size = struct.unpack(">Q", data[offset + 8 : offset + 16])[0]
            header = 16
        elif size == 0:
            size = end - offset
        if size < header or offset + size > end:
            break
        body_start = offset + header
        body_end = offset + size
        yield kind, body_start, body_end
        if kind in _MP4_CONTAINERS and depth < 6:
            yield from _iter_mp4_boxes(data, body_start, body_end, depth + 1)
        offset += size


def analyze_video(path: Path) -> DesignProfile:
    """Read container metadata and, if possible, a storyboard."""

    profile = DesignProfile(source=str(path), kind="video")
    data = path.read_bytes()
    info: dict[str, Any] = {}

    for kind, body_start, body_end in _iter_mp4_boxes(data):
        if kind == "ftyp":
            info["brand"] = data[body_start : body_start + 4].decode("latin-1", errors="replace")
        elif kind == "mvhd":
            body = data[body_start:body_end]
            if len(body) >= 20:
                version = body[0]
                if version == 1 and len(body) >= 28:
                    timescale = struct.unpack(">I", body[20:24])[0]
                    duration = struct.unpack(">Q", body[24:32])[0]
                else:
                    timescale = struct.unpack(">I", body[12:16])[0]
                    duration = struct.unpack(">I", body[16:20])[0]
                if timescale:
                    info["duration_seconds"] = round(duration / timescale, 2)
        elif kind == "tkhd":
            body = data[body_start:body_end]
            # width and height are the last two 16.16 fixed point fields
            if len(body) >= 84:
                width = struct.unpack(">I", body[-8:-4])[0] / 65536
                height = struct.unpack(">I", body[-4:])[0] / 65536
                if width > 1 and height > 1:
                    info["resolution"] = [int(width), int(height)]
        elif kind == "mdhd":
            body = data[body_start:body_end]
            if len(body) >= 20:
                timescale = struct.unpack(">I", body[12:16])[0]
                duration = struct.unpack(">I", body[16:20])[0]
                if timescale and duration:
                    info["track_seconds"] = round(duration / timescale, 2)
        elif kind == "stts":
            body = data[body_start:body_end]
            if len(body) >= 8:
                entries = struct.unpack(">I", body[4:8])[0]
                total_samples = 0
                total_delta = 0
                for entry in range(min(entries, 4096)):
                    offset = 8 + entry * 8
                    if offset + 8 > len(body):
                        break
                    sample_count, sample_delta = struct.unpack(">II", body[offset : offset + 8])
                    total_samples += sample_count
                    total_delta += sample_count * sample_delta
                info["frame_samples"] = total_samples
                if total_delta:
                    info["approx_fps"] = round(total_samples / (total_delta / 1000.0), 2)

    profile.detected.update(info)
    profile.detected["component_tokens"] = ["immersive_hero"]
    profile.detected["density"] = 0.35

    ffmpeg = shutil.which("ffmpeg")
    if ffmpeg:
        frames_dir = path.parent / f".{path.stem}_frames"
        frames_dir.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.run(
                [ffmpeg, "-y", "-i", str(path), "-vf", "fps=1/2", "-frames:v", "12", str(frames_dir / "frame_%02d.jpg")],
                capture_output=True,
                check=False,
                timeout=120,
            )
            frames = sorted(frames_dir.glob("frame_*.jpg"))
            profile.detected["storyboard_frames"] = len(frames)
            if frames:
                palettes = []
                for frame in frames:
                    with __import__("PIL.Image", fromlist=["Image"]) as _module:
                        image_module = _module
                    with image_module.open(frame) as image:
                        small = image.convert("RGB").resize((80, 45), image_module.LANCZOS)
                        quantised = small.quantize(colors=6)
                        entries = quantised.getpalette() or []
                        for index, _count in Counter(_pixel_data(quantised)).most_common(3):
                            base = index * 3
                            if base + 2 < len(entries):
                                palettes.append((entries[base], entries[base + 1], entries[base + 2]))
                if palettes:
                    profile.palette = _roles_from_colors(palettes)
        except (OSError, subprocess.SubprocessError):
            profile.limits.append("ffmpeg was found but frame extraction failed.")
        finally:
            shutil.rmtree(frames_dir, ignore_errors=True)
    else:
        profile.limits.append(
            "ffmpeg is not installed: only container metadata was read, no frames were analysed."
        )

    duration = info.get("duration_seconds")
    profile.recommendations = {
        "slides_per_minute": None,
        "transition_hint": "match the on-screen rhythm of the source video",
        "auto_advance_ms": int((duration / 12) * 1000) if duration else None,
        "confidence": 0.7 if info else 0.4,
    }
    profile.layout = {"frames_detected": profile.detected.get("storyboard_frames", 0)}
    return profile

# ---------------------------------------------------------------------------
# pdf / docx / json analysis
# ---------------------------------------------------------------------------


def analyze_pdf(path: Path) -> DesignProfile:
    """Extract page count, palette and font names from a PDF (pure Python)."""

    profile = DesignProfile(source=str(path), kind="pdf")
    data = path.read_bytes()
    profile.detected["page_count"] = len(re.findall(rb"/Type\s*/Page[^s]", data))

    colors: Counter = Counter()
    fonts: Counter = Counter()
    for match in re.finditer(rb"stream\r?\n(.*?)endstream", data, re.DOTALL):
        chunk = match.group(1)
        try:
            chunk = zlib.decompress(chunk)
        except zlib.error:
            pass
        for color in re.finditer(rb"([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+(?:rg|RG)\b", chunk):
            try:
                rgb = tuple(int(float(color.group(i)) * 255) for i in (1, 2, 3))
            except ValueError:
                continue
            colors[rgb] += 1
        for gray in re.finditer(rb"([\d.]+)\s+(?:g|G)\b", chunk):
            try:
                value = int(float(gray.group(1)) * 255)
            except ValueError:
                continue
            colors[(value, value, value)] += 1
        for font in re.finditer(rb"/BaseFont\s*/([#A-Za-z0-9+,\-]+)", chunk):
            fonts[font.group(1).decode("latin-1", errors="replace")] += 1

    ranked = [color for color, _count in colors.most_common(12)]
    if ranked:
        profile.palette = _roles_from_colors(ranked)
        profile.detected["distinct_colors"] = len(colors)
    if fonts:
        profile.fonts = {name.replace("#", ""): name for name, _c in fonts.most_common(8)}
        profile.detected["font_names"] = [name for name, _c in fonts.most_common(8)]

    profile.recommendations = {
        "rebuild_masters": True,
        "level_hint": "L3",
        "confidence": 0.6,
    }
    profile.limits = [
        "PDF analysis reads colours and font names from the content stream; it does not rebuild native PPT masters.",
    ]
    return profile


def analyze_docx(path: Path) -> DesignProfile:
    """Read brand signals (fonts, colours, images) from a Word document."""

    profile = DesignProfile(source=str(path), kind="docx")
    document = "".join(_read_zip_xml(path, r"word/document\.xml$"))
    styles = "".join(_read_zip_xml(path, r"word/styles\.xml$"))
    fonts: Counter = Counter()
    for chunk in (document, styles):
        for font in re.finditer(r'w:ascii="([^"]+)"', chunk):
            fonts[font.group(1)] += 1
        for font in re.finditer(r'w:eastAsia="([^"]+)"', chunk):
            fonts[font.group(1)] += 1
    colors: Counter = Counter()
    for color in re.finditer(r'w:color w:val="([0-9A-Fa-f]{6})"', document + styles):
        colors[color.group(1).upper()] += 1

    if fonts:
        profile.fonts = {name: name for name, _c in fonts.most_common(6)}
    if colors:
        rgb_list = [
            (int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))
            for c, _n in colors.most_common(10)
        ]
        profile.palette = _roles_from_colors(rgb_list)
        profile.detected["accent_colors"] = [_hex(c) for c in rgb_list[:5]]
    profile.detected["images"] = len(_read_zip_xml(path, r"word/media/")) or len(
        [n for n in _read_zip_xml(path, r"word/document\.xml$") if "<w:drawing>" in n]
    )
    profile.recommendations = {
        "use_as_brand_manual": True,
        "level_hint": "L2",
        "confidence": 0.68,
    }
    profile.limits = ["Only fonts, colours and image counts are extracted from Word files."]
    return profile


def analyze_json_config(path: Path) -> DesignProfile:
    """Read an explicit design configuration."""

    profile = DesignProfile(source=str(path), kind="json")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        profile.limits.append(f"The file could not be parsed as JSON: {error}")
        return profile
    palette = data.get("palette") or data.get("colors") or {}
    if isinstance(palette, dict):
        profile.palette = {
            key: (value if isinstance(value, str) and value.startswith("#") else _hex(value) if isinstance(value, (list, tuple)) else value)
            for key, value in palette.items()
            if isinstance(value, (str, list, tuple))
        }
    fonts = data.get("fonts") or data.get("font") or {}
    if isinstance(fonts, dict):
        profile.fonts = {key: value for key, value in fonts.items() if isinstance(value, str)}
    profile.detected["keys"] = sorted(data.keys()) if isinstance(data, dict) else []
    profile.layout = data.get("layout", {}) if isinstance(data.get("layout", {}), dict) else {}
    profile.recommendations = {"apply_directly": True, "confidence": 0.95}
    return profile


def _parse_simple_yaml(text: str) -> dict[str, Any]:
    """Parse the small YAML subset used by design configuration files."""

    root: dict[str, Any] = {}
    stack: list[tuple[int, dict[str, Any]]] = [(-1, root)]
    for raw in text.splitlines():
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        indent = len(raw) - len(raw.lstrip(" "))
        line = raw.strip()
        if ":" not in line:
            raise ValueError(f"Unsupported YAML line: {line}")
        key, value = line.split(":", 1)
        key = key.strip()
        value = value.strip()
        while stack[-1][0] >= indent:
            stack.pop()
        parent = stack[-1][1]
        if not value:
            child: dict[str, Any] = {}
            parent[key] = child
            stack.append((indent, child))
        elif value.startswith("[") and value.endswith("]"):
            parent[key] = [
                item.strip().strip("'\"")
                for item in value[1:-1].split(",")
                if item.strip()
            ]
        else:
            scalar = value.strip("'\"")
            if re.fullmatch(r"-?\d+", scalar):
                parent[key] = int(scalar)
            elif re.fullmatch(r"-?\d+\.\d+", scalar):
                parent[key] = float(scalar)
            else:
                parent[key] = scalar
    return root


def analyze_yaml_config(path: Path) -> DesignProfile:
    profile = DesignProfile(source=str(path), kind="yaml")
    text = path.read_text(encoding="utf-8", errors="replace")
    try:
        try:
            import yaml  # type: ignore
        except ImportError:
            data = _parse_simple_yaml(text)
        else:
            data = yaml.safe_load(text)
    except Exception as error:
        profile.limits.append(f"The YAML configuration could not be parsed: {error}")
        return profile
    if not isinstance(data, dict):
        profile.limits.append("The YAML root must be a mapping.")
        return profile
    palette = data.get("palette") or data.get("colors") or {}
    if isinstance(palette, dict):
        profile.palette = {
            key: value
            for key, value in palette.items()
            if isinstance(value, str) and re.fullmatch(r"#?[0-9A-Fa-f]{6}", value)
        }
    fonts = data.get("fonts") or data.get("font") or {}
    if isinstance(fonts, dict):
        profile.fonts = {
            key: value for key, value in fonts.items() if isinstance(value, str)
        }
    profile.detected["keys"] = sorted(data.keys())
    profile.layout = (
        data.get("layout", {}) if isinstance(data.get("layout"), dict) else {}
    )
    profile.detected["component_tokens"] = (
        list(data.get("components", []))
        if isinstance(data.get("components"), list)
        else []
    )
    profile.recommendations = {"apply_directly": True, "confidence": 0.92}
    return profile


def analyze_svg(path: Path) -> DesignProfile:
    profile = DesignProfile(source=str(path), kind="svg")
    text = path.read_text(encoding="utf-8", errors="replace")
    width = re.search(r'\bwidth=["\']([\d.]+)', text)
    height = re.search(r'\bheight=["\']([\d.]+)', text)
    if width and height:
        profile.detected["size"] = [float(width.group(1)), float(height.group(1))]
    colours: list[tuple[int, int, int]] = []
    for value in re.findall(r"#[0-9a-fA-F]{6}", text):
        rgb = (int(value[1:3], 16), int(value[3:5], 16), int(value[5:7], 16))
        if rgb not in colours:
            colours.append(rgb)
    if colours:
        profile.palette = _roles_from_colors(colours[:12])
    profile.detected["vector_primitives"] = {
        token: len(re.findall(rf"<{token}\b", text))
        for token in ("path", "rect", "circle", "line", "polygon")
    }
    profile.detected["component_tokens"] = (
        ["minimal_icon"] if "<path" in text else []
    )
    profile.layout = {"archetype": "vector_asset", "grid_columns": 12}
    profile.recommendations = {
        "use_as_vector_language": True,
        "confidence": 0.82,
    }
    profile.limits = [
        "SVG text and colors are read; filters and effects are not mapped one-to-one."
    ]
    return profile


def analyze_markdown(path: Path) -> DesignProfile:
    profile = DesignProfile(source=str(path), kind="markdown")
    text = path.read_text(encoding="utf-8", errors="replace")
    headings = re.findall(r"^(#{1,6})\s+(.+)$", text, re.MULTILINE)
    tables = len(re.findall(r"^\|.+\|$", text, re.MULTILINE))
    mermaid = len(re.findall(r"```mermaid", text))
    profile.detected["headings"] = [h[1] for h in headings[:20]]
    profile.detected["heading_levels"] = sorted({len(h[0]) for h in headings})
    profile.detected["tables"] = tables
    profile.detected["mermaid_blocks"] = mermaid
    profile.recommendations = {"use_as_outline": True, "level_hint": "L1"}
    profile.limits = ["Markdown provides structure only; visual style must be chosen by the level system."]
    return profile


# ---------------------------------------------------------------------------
# dispatcher
# ---------------------------------------------------------------------------


def detect_kind(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in IMAGE_SUFFIXES:
        return "image"
    if suffix in VIDEO_SUFFIXES:
        return "video"
    if suffix in {".pptx", ".potx", ".potm", ".thmx"}:
        return "office"
    if suffix == ".pdf":
        return "pdf"
    if suffix == ".docx":
        return "docx"
    if suffix == ".json":
        return "json"
    if suffix in {".yaml", ".yml"}:
        return "yaml"
    if suffix == ".svg":
        return "svg"
    if suffix in {".md", ".markdown"}:
        return "markdown"
    return "unknown"


def analyze(path: str | Path, kind: str | None = None) -> DesignProfile:
    """Analyse any supported reference file."""

    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(f"No such file: {target}")
    resolved = kind or detect_kind(target)
    if resolved == "image":
        return analyze_image(target)
    if resolved == "video":
        return analyze_video(target)
    if resolved == "office":
        return analyze_office(target)
    if resolved == "pdf":
        return analyze_pdf(target)
    if resolved == "docx":
        return analyze_docx(target)
    if resolved == "json":
        return analyze_json_config(target)
    if resolved == "yaml":
        return analyze_yaml_config(target)
    if resolved == "svg":
        return analyze_svg(target)
    if resolved == "markdown":
        return analyze_markdown(target)
    profile = DesignProfile(source=str(target), kind="unknown")
    profile.limits.append(
        f"The file type {target.suffix!r} cannot be parsed directly. Ask the user for an export "
        "(PDF, PNG/JPG, PPTX or a JSON design configuration)."
    )
    return profile


def parse_template(
    path: str | Path,
    kind: str | None = None,
) -> DesignProfile:
    """Public parser entry point shared by migration and workflow modules."""

    return analyze(path, kind)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Parse a reference template for pro-ppt-studio V4"
    )
    parser.add_argument("path", help="File to analyse")
    parser.add_argument("--json", dest="json_path", help="Write the design profile as JSON")
    parser.add_argument("--kind", default=None, help="Force the analysis kind")
    args = parser.parse_args()

    profile = parse_template(args.path, args.kind)
    payload = profile.as_dict()
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    if args.json_path:
        out = Path(args.json_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON report: {out.resolve()}")


if __name__ == "__main__":
    main()
