#!/usr/bin/env python3
"""
pro-ppt-studio V4.0.0 质检脚本（结构 / 内容 / 几何 / 模板 / 动效 / 策略）

用法：
    python check_pptx.py <file.pptx> [--level L1|L2|L3|L4]
        [--mode auto|static|dynamic] [--template-style style.json]
        [--json report.json] [--min-font 10]
    python check_pptx.py --compare L1.pptx L2.pptx L3.pptx L4.pptx

对单个 PPTX 做自动化静态检查：
  - 文件存在、可解析、大小、页数、页面比例
  - 每页元素统计（形状/图片/表格/图表/文本）
  - 残留占位文本、调试文本、空白页、重复页
  - 文本框溢出（启发式估算）、元素越界、元素重叠、字号过小
  - 字体清单、异常图片
  - 网格对齐、字体层级、策略字体上限和版式重复度
  - 模板配色/字体匹配度
  - 动画、媒体与运动路径的等级权限
  - 四级相邻样例的视觉相似度

说明：
  - 溢出为启发式估算（未实际排版），结果标注 "estimate"，需配合 render_preview.py 渲染确认。
  - 越界/重叠可能是刻意的出血或装饰叠层，判定结果仅供人工复核。

依赖：python-pptx
"""

import argparse
import json
import math
import re
import sys
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    print("错误：未安装 python-pptx。请先执行：python -m pip install python-pptx")
    sys.exit(2)

from level_profiles import SKILL_VERSION, get_profile
import dynamic_pptx

EMU_PER_PT = 12700.0
EMU_PER_IN = 914400.0

# 常见占位/调试文本模式（小写匹配）
PLACEHOLDER_PATTERNS = [
    r"click to add",
    r"lorem ipsum",
    r"\btodo\b",
    r"\btbd\b",
    r"\bplaceholder\b",
    r"your text here",
    r"sample text",
    r"请在此输入",
    r"请在此处输入",
    r"此处添加",
    r"单击此处",
    r"请替换",
    r"占位",
    r"示例文字",
    r"示例文本",
    r"待填写",
    r"待补充",
    r"【\s*[^】]{0,20}\s*】",  # 占位括号，如【此处填写】
]


def emu_to_pt(emu):
    if emu is None:
        return None
    return int(emu) / EMU_PER_PT


def size_to_pt(size):
    """将 run.font.size 转为点数；兼容不同返回类型。"""
    if size is None:
        return None
    pt = getattr(size, "pt", None)
    if pt is not None:
        return float(pt)
    return int(size) / EMU_PER_PT


def has_meaningful_text(shape):
    try:
        if not shape.has_text_frame:
            return False
        return bool(shape.text_frame.text and shape.text_frame.text.strip())
    except Exception:
        return False


def shape_kind(shape):
    try:
        st = shape.shape_type
        if st == MSO_SHAPE_TYPE.PICTURE:
            return "picture"
        if st == MSO_SHAPE_TYPE.TABLE:
            return "table"
        if st == MSO_SHAPE_TYPE.CHART:
            return "chart"
        if st == MSO_SHAPE_TYPE.GROUP:
            return "group"
        if st == MSO_SHAPE_TYPE.PLACEHOLDER:
            return "placeholder"
        if st == MSO_SHAPE_TYPE.AUTO_SHAPE:
            return "autoshape"
        if st == MSO_SHAPE_TYPE.TEXT_BOX:
            return "textbox"
        return str(st)
    except Exception:
        return "unknown"


def collect_shapes(shapes, acc):
    """递归收集形状（含 group）。"""
    for sh in shapes:
        acc.append(sh)
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                collect_shapes(sh.shapes, acc)
            except Exception:
                pass


def estimate_text_height(shape, width_pt):
    """粗略估算文本框所需高度（点）。CJK 按 1.0 em、拉丁按 0.5 em 估算。"""
    tf = shape.text_frame
    total = 0.0
    try:
        for para in tf.paragraphs:
            runs = para.runs
            if not runs:
                text = para.text or ""
                size = 18.0
            else:
                text = "".join(r.text or "" for r in runs)
                sizes = [size_to_pt(r.font.size) for r in runs if size_to_pt(r.font.size)]
                size = max(sizes) if sizes else 18.0
            if not text:
                total += size * 1.2
                continue
            # 估算整段字符总宽
            char_w = 0.0
            for ch in text:
                char_w += size if ord(ch) > 0x2E7F else size * 0.52
            lines = max(1, int(char_w / max(width_pt, 1)) + 1)
            total += lines * size * 1.25
    except Exception:
        return None
    return total


def check_slide_overflow(shape, slide_w, slide_h, issues):
    """文本框溢出（启发式）与越界检查。"""
    try:
        if not shape.has_text_frame:
            return
        if not (shape.text_frame.text and shape.text_frame.text.strip()):
            return  # 空文本框（装饰色块/圆点等）不参与溢出估算
        w = emu_to_pt(shape.width)
        h = emu_to_pt(shape.height)
        if not w or not h:
            return
        est_h = estimate_text_height(shape, w)
        if est_h is not None:
            if est_h > h * 1.05:
                issues.append({
                    "type": "text_overflow",
                    "level": "warn",
                    "detail": f"文本可能溢出文本框（估算高度 {est_h:.0f}pt > 形状高度 {h:.0f}pt）",
                    "estimate": True,
                    "text": shape.text_frame.text[:40],
                })
        # 越界
        l = emu_to_pt(shape.left)
        t = emu_to_pt(shape.top)
        if l is None or t is None:
            return
        tol = 2.0
        if l < -tol or t < -tol or (l + w) > slide_w + tol or (t + h) > slide_h + tol:
            issues.append({
                "type": "out_of_bounds",
                "level": "warn",
                "detail": f"元素越出页面边界 (left={l:.0f}, top={t:.0f}, right={l + w:.0f}, bottom={t + h:.0f}; 页面 {slide_w:.0f}x{slide_h:.0f})",
            })
    except Exception:
        pass


def rect_intersection_area(a, b):
    x1 = max(a[0], b[0])
    y1 = max(a[1], b[1])
    x2 = min(a[2], b[2])
    y2 = min(a[3], b[3])
    if x2 <= x1 or y2 <= y1:
        return 0.0
    return (x2 - x1) * (y2 - y1)


def rect_area(a):
    return max(0.0, a[2] - a[0]) * max(0.0, a[3] - a[1])


def check_overlaps(shapes, issues, slide_w, slide_h):
    """检测文本/图片元素之间的明显重叠（装饰元素忽略）。"""
    def rect(sh):
        try:
            if sh.left is None or sh.top is None or sh.width is None or sh.height is None:
                return None
            l = emu_to_pt(sh.left)
            t = emu_to_pt(sh.top)
            return (l, t, l + emu_to_pt(sh.width), t + emu_to_pt(sh.height))
        except Exception:
            return None

    def significant(sh):
        # 只有带文本或为图片/表格/图表的元素参与重叠检测
        if has_meaningful_text(sh):
            return True
        return shape_kind(sh) in ("picture", "table", "chart")

    def full_slide_picture(sh):
        if shape_kind(sh) != "picture":
            return False
        r = rect(sh)
        if not r:
            return False
        return rect_area(r) >= slide_w * slide_h * 0.75

    sig = [s for s in shapes if significant(s)]
    for i in range(len(sig)):
        for j in range(i + 1, len(sig)):
            r1, r2 = rect(sig[i]), rect(sig[j])
            if not r1 or not r2:
                continue
            inter = rect_intersection_area(r1, r2)
            smaller = min(rect_area(r1), rect_area(r2))
            if smaller <= 0:
                continue
            # A full-bleed image under text or data is a common intentional design.
            if full_slide_picture(sig[i]) or full_slide_picture(sig[j]):
                continue
            text_shape = sig[i] if has_meaningful_text(sig[i]) else sig[j]
            other_shape = sig[j] if text_shape is sig[i] else sig[i]
            short_text = (
                text_shape.text_frame.text.strip()
                if has_meaningful_text(text_shape)
                else ""
            )
            if (
                shape_kind(other_shape) == "picture"
                and short_text.isdigit()
                and smaller / (slide_w * slide_h) < 0.02
            ):
                continue
            if inter / smaller > 0.15:
                t1 = sig[i].text_frame.text[:20] if has_meaningful_text(sig[i]) else shape_kind(sig[i])
                t2 = sig[j].text_frame.text[:20] if has_meaningful_text(sig[j]) else shape_kind(sig[j])
                issues.append({
                    "type": "overlap",
                    "level": "warn",
                    "detail": f"元素明显重叠（{t1!r} 与 {t2!r}，重叠占比 {inter / smaller:.0%}）",
                })


def check_fonts_and_sizes(shapes, min_font, issues, font_set):
    for sh in shapes:
        try:
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    fname = run.font.name
                    if fname:
                        font_set.add(fname)
                    sz = size_to_pt(run.font.size)
                    if sz is not None and sz < min_font:
                        issues.append({
                            "type": "tiny_font",
                            "level": "warn",
                            "detail": f"字号过小 {sz:.1f}pt：{run.text[:30]!r}",
                        })
        except Exception:
            continue


def collect_run_typography(shapes, font_set, size_set):
    for sh in shapes:
        try:
            if not sh.has_text_frame:
                continue
            for paragraph in sh.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.font.name:
                        font_set.add(run.font.name)
                    size = size_to_pt(run.font.size)
                    if size is not None:
                        size_set.add(round(size, 1))
        except Exception:
            continue


def check_grid_alignment(shapes, profile, issues):
    """Measure shape-edge alignment against the strategy's vertical grid."""

    tolerance = profile.grid.alignment_tolerance_in
    lines = profile.grid.vertical_lines()
    candidates = 0
    total_edges = 0
    aligned_edges = 0
    for shape in shapes:
        try:
            if shape.left is None or shape.width is None:
                continue
            if shape_kind(shape) not in ("autoshape", "picture", "table", "chart"):
                continue
            width_in = float(shape.width) / EMU_PER_IN
            height_in = float(shape.height or 0) / EMU_PER_IN
            if width_in < 1.0 or height_in < 0.55:
                continue
            if height_in >= 6.8:
                continue
            if width_in >= 12.7:
                continue
            left = float(shape.left) / EMU_PER_IN
            right = left + width_in
            candidates += 1
            for edge in (left, right):
                total_edges += 1
                if any(abs(edge - line) <= tolerance for line in lines):
                    aligned_edges += 1
        except Exception:
            continue
    ratio = aligned_edges / total_edges if total_edges else 1.0
    if candidates and ratio < 0.45:
        issues.append(
            {
                "type": "grid_alignment",
                "level": "warn",
                "detail": (
                    f"{profile.level} 网格对齐覆盖偏低：{ratio:.0%}，"
                    f"允差 {tolerance:.2f}in，{profile.grid.columns} 列"
                ),
            }
        )
    return round(ratio, 4)


def check_font_hierarchy(shapes, profile, issues):
    sizes = set()
    families = set()
    for sh in shapes:
        try:
            if not sh.has_text_frame:
                continue
            for paragraph in sh.text_frame.paragraphs:
                for run in paragraph.runs:
                    size = size_to_pt(run.font.size)
                    if size is not None:
                        sizes.add(round(size, 1))
                    if run.font.name:
                        families.add(run.font.name)
        except Exception:
            continue
    maximum_sizes = profile.font.hierarchy_levels + 1
    if len(sizes) > maximum_sizes:
        issues.append(
            {
                "type": "font_hierarchy",
                "level": "warn",
                "detail": (
                    f"检测到 {len(sizes)} 个字号，超过策略上限 "
                    f"{maximum_sizes}：{sorted(sizes)}"
                ),
            }
        )
    if len(families) > profile.font.max_font_families:
        issues.append(
            {
                "type": "strategy_font_limit",
                "level": "warn",
                "detail": (
                    f"检测到 {len(families)} 个字体家族，策略上限 "
                    f"{profile.font.max_font_families}"
                ),
            }
        )
    return {
        "font_sizes": sorted(sizes),
        "font_size_count": len(sizes),
        "font_families": sorted(families),
        "font_family_count": len(families),
    }


def check_broken_pictures(shapes, issues):
    for sh in shapes:
        if shape_kind(sh) == "picture":
            try:
                _ = sh.image
            except Exception as e:
                issues.append({
                    "type": "broken_picture",
                    "level": "error",
                    "detail": f"图片资源异常：{e}",
                })


def normalize(text):
    return re.sub(r"\s+", "", text).lower()


def _shape_area(shape):
    try:
        return max(0, int(shape.width or 0)) * max(0, int(shape.height or 0))
    except Exception:
        return 0


def check_animation_compliance(slide, profile, mode, issues):
    transition = dynamic_pptx.transition_summary(slide)
    animation = dynamic_pptx.animation_summary(slide)
    rule = profile.animation
    transition_name = transition.get("transition")
    if transition_name and transition_name not in rule.transition_types:
        issues.append(
            {
                "type": "animation_permission",
                "level": "error",
                "detail": (
                    f"{profile.level} 不允许转场 {transition_name!r}；"
                    f"允许 {list(rule.transition_types)}"
                ),
            }
        )
    maximum = rule.max_element_animations
    if maximum is not None and animation["steps"] > maximum:
        issues.append(
            {
                "type": "animation_density",
                "level": "error",
                "detail": (
                    f"对象动画 {animation['steps']} 个，超过 {profile.level} "
                    f"上限 {maximum}"
                ),
            }
        )
    media_nodes = [
        node
        for node in slide._element.iter()
        if node.tag.rsplit("}", 1)[-1] == "cMediaNode"
    ]
    if media_nodes and not rule.allow_media:
        issues.append(
            {
                "type": "media_permission",
                "level": "error",
                "detail": f"{profile.level} 不允许音视频媒体",
            }
        )
    if (
        "path" in animation.get("categories", [])
        and not rule.allow_motion_path
    ):
        issues.append(
            {
                "type": "motion_path_permission",
                "level": "error",
                "detail": f"{profile.level} 不允许运动路径",
            }
        )
    if mode == "dynamic" and not transition_name:
        issues.append(
            {
                "type": "missing_transition",
                "level": "warn",
                "detail": "动态模式页面未检测到转场",
            }
        )
    if mode == "static" and (transition_name or animation["animated"]):
        issues.append(
            {
                "type": "static_motion_leak",
                "level": "warn",
                "detail": "静态模式页面包含转场或动画",
            }
        )
    return {
        "transition": transition,
        "animation": animation,
        "media_count": len(media_nodes),
    }


def check_design_gene(slide, profile, issues):
    if profile.level == "L1":
        return {"visual_center": "not_required"}
    shapes = [
        shape
        for shape in slide.shapes
        if _shape_area(shape) > 0
        and (
            shape_kind(shape) in {"picture", "chart", "table"}
            or has_meaningful_text(shape)
        )
        and not (
            getattr(shape, "has_text_frame", False)
            and shape.text_frame.text.strip().isdigit()
        )
    ]
    ranked = sorted((_shape_area(shape) for shape in shapes), reverse=True)
    if not ranked:
        return {"visual_center": "empty"}
    dominance = ranked[0] / max(1, ranked[1]) if len(ranked) > 1 else 9.9
    if profile.level == "L3" and dominance < 1.05 and len(ranked) > 4:
        issues.append(
            {
                "type": "visual_center",
                "level": "warn",
                "detail": f"主视觉对象面积优势不足（{dominance:.2f}×）",
            }
        )
    return {
        "visual_center": "single_dominant" if dominance >= 1.05 else "distributed",
        "dominance_ratio": round(dominance, 3),
    }


def _load_template_style(template_style):
    if template_style is None:
        return None
    if isinstance(template_style, dict):
        return template_style
    path = Path(template_style)
    return json.loads(path.read_text(encoding="utf-8"))


def _palette_cosine(deck_colors, expected_colors):
    universe = sorted(set(deck_colors) | set(expected_colors))
    if not universe:
        return 1.0
    left = [1.0 if value in deck_colors else 0.0 for value in universe]
    right = [1.0 if value in expected_colors else 0.0 for value in universe]
    dot = sum(a * b for a, b in zip(left, right))
    denominator = math.sqrt(sum(a * a for a in left) * sum(b * b for b in right))
    return dot / denominator if denominator else 0.0


def _deck_colors(prs):
    colors = set()
    for slide in prs.slides:
        for shape in slide.shapes:
            try:
                if shape.fill.type is not None and shape.fill.fore_color.rgb:
                    colors.add(str(shape.fill.fore_color.rgb).upper())
            except Exception:
                pass
    return colors


def check_template_match(prs, template_style, issues):
    style = _load_template_style(template_style)
    if not style:
        return None
    expected = {
        value.lstrip("#").upper()
        for value in (style.get("palette") or {}).values()
        if isinstance(value, str) and re.fullmatch(r"#?[0-9A-Fa-f]{6}", value)
    }
    actual = _deck_colors(prs)
    palette_score = _palette_cosine(actual, expected)
    expected_fonts = {
        value for value in (style.get("fonts") or {}).values() if isinstance(value, str)
    }
    actual_fonts = set()
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.font.name:
                        actual_fonts.add(run.font.name)
    font_score = (
        len(expected_fonts & actual_fonts) / len(expected_fonts)
        if expected_fonts
        else 1.0
    )
    overall = round(0.65 * palette_score + 0.35 * font_score, 3)
    if overall < 0.55:
        issues.append(
            {
                "type": "template_match",
                "level": "warn",
                "detail": f"模板匹配度偏低：{overall:.0%}",
            }
        )
    return {
        "palette_cosine": round(palette_score, 3),
        "font_match": round(font_score, 3),
        "overall": overall,
    }


def deck_visual_metrics(path):
    """Return a compact visual vector for adjacent-level similarity checks."""

    prs = Presentation(str(path))
    fonts = set()
    colors = _deck_colors(prs)
    text_lengths = []
    shape_counts = []
    for slide in prs.slides:
        shapes = []
        collect_shapes(slide.shapes, shapes)
        shape_counts.append(len(shapes))
        text = "\n".join(
            shape.text_frame.text for shape in shapes if has_meaningful_text(shape)
        )
        text_lengths.append(len(text))
        for shape in shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.font.name:
                        fonts.add(run.font.name)
    return {
        "fonts": sorted(fonts),
        "colors": sorted(colors),
        "average_text_length": sum(text_lengths) / max(1, len(text_lengths)),
        "average_shape_count": sum(shape_counts) / max(1, len(shape_counts)),
    }


def visual_similarity(first, second):
    first = Path(first)
    second = Path(second)
    a = deck_visual_metrics(first)
    b = deck_visual_metrics(second)
    palette = _palette_cosine(set(a["colors"]), set(b["colors"]))
    fonts = len(set(a["fonts"]) & set(b["fonts"])) / max(
        1, len(set(a["fonts"]) | set(b["fonts"]))
    )
    density = 1.0 - min(
        1.0,
        abs(a["average_text_length"] - b["average_text_length"])
        / max(1.0, a["average_text_length"], b["average_text_length"]),
    )
    structural = 1.0 - min(
        1.0,
        abs(a["average_shape_count"] - b["average_shape_count"])
        / max(1.0, a["average_shape_count"], b["average_shape_count"]),
    )
    score = round(
        0.50 * palette
        + 0.15 * fonts
        + 0.15 * density
        + 0.20 * structural,
        3,
    )
    return {"score": score, "palette": palette, "fonts": fonts, "density": density, "structure": structural, "first": a, "second": b}


def compare_decks(paths, thresholds=None):
    thresholds = thresholds or {"L1-L2": 0.7, "L2-L3": 0.6, "L3-L4": 0.5}
    pairs = {}
    errors = []
    for index in range(len(paths) - 1):
        key = f"L{index + 1}-L{index + 2}"
        similarity = visual_similarity(paths[index], paths[index + 1])
        limit = thresholds.get(key, 0.65)
        similarity["threshold"] = limit
        similarity["pass"] = similarity["score"] < limit
        if not similarity["pass"]:
            errors.append(
                f"{key} visual similarity {similarity['score']:.2f} >= {limit:.2f}"
            )
        pairs[key] = similarity
    return {"valid": not errors, "pairs": pairs, "errors": errors}


def check_pptx(
    path,
    min_font=10.0,
    level=None,
    mode=None,
    template_style=None,
):
    p = Path(path)
    report = {
        "skill_version": SKILL_VERSION,
        "file": str(path),
        "errors": [],
        "warnings": [],
        "slides": [],
        "font_names": [],
        "font_sizes_pt": [],
    }
    if not p.exists():
        report["errors"].append("文件不存在")
        return report
    size = p.stat().st_size
    if size <= 0:
        report["errors"].append("文件大小为 0")
        return report
    report["file_size_bytes"] = size

    try:
        prs = Presentation(str(path))
    except Exception as e:
        report["errors"].append(f"无法解析 PPTX：{e}")
        return report

    slide_w = emu_to_pt(prs.slide_width)
    slide_h = emu_to_pt(prs.slide_height)
    report["core_properties"] = {
        "title": prs.core_properties.title or "",
        "subject": prs.core_properties.subject or "",
        "keywords": prs.core_properties.keywords or "",
        "comments": prs.core_properties.comments or "",
    }
    aspect = slide_w / slide_h if slide_h else 0
    report["slide_size_pt"] = [round(slide_w, 1), round(slide_h, 1)]
    report["aspect_ratio"] = round(aspect, 4)
    if abs(aspect - 16 / 9) > 0.02 and abs(aspect - 4 / 3) > 0.02:
        report["warnings"].append(f"非标准页面比例 {aspect:.3f}（16:9=1.778，4:3=1.333）")

    slides = list(prs.slides)
    report["slide_count"] = len(slides)
    seen_norm = {}
    font_set = set()
    size_set = set()
    profile = get_profile(level) if level else None
    if profile:
        report["expression_level"] = profile.level
        report["strategy"] = profile.strategy.key
        report["strategy_name"] = profile.strategy.name
        report["design_principle"] = profile.design_gene.design_principle
        report["motion_intensity"] = profile.animation.motion_intensity
    for idx, slide in enumerate(slides, 1):
        shapes = []
        collect_shapes(slide.shapes, shapes)
        texts = [sh.text_frame.text.strip() for sh in shapes if has_meaningful_text(sh)]
        all_text = "\n".join(texts)
        stat = {
            "index": idx,
            "shape_count": len(shapes),
            "picture_count": sum(1 for s in shapes if shape_kind(s) == "picture"),
            "table_count": sum(1 for s in shapes if shape_kind(s) == "table"),
            "chart_count": sum(1 for s in shapes if shape_kind(s) == "chart"),
            "autoshape_count": sum(1 for s in shapes if shape_kind(s) == "autoshape"),
            "text_length": len(all_text),
            "paragraph_count": sum(
                len(s.text_frame.paragraphs)
                for s in shapes
                if has_meaningful_text(s)
            ),
        }
        stat["visual_asset_count"] = (
            stat["picture_count"] + stat["table_count"] + stat["chart_count"]
        )
        picture_area_ratio = 0.0
        slide_area = slide_w * slide_h
        for shape in shapes:
            if shape_kind(shape) != "picture":
                continue
            try:
                picture_area_ratio += (
                    float(shape.width) * float(shape.height)
                ) / (EMU_PER_PT * EMU_PER_PT)
            except Exception:
                pass
        stat["picture_area_ratio"] = round(
            picture_area_ratio / slide_area if slide_area else 0.0,
            3,
        )
        stat["text_dominant_candidate"] = (
            stat["visual_asset_count"] == 0
            and stat["text_length"] >= 180
            and stat["paragraph_count"] >= 4
        )
        stat["layout_signature"] = (
            f"p{stat['picture_count']}-c{stat['chart_count']}-"
            f"t{stat['table_count']}-a{min(stat['autoshape_count'], 9)}-"
            f"img{round(stat['picture_area_ratio'] * 10)}"
        )
        issues = []
        if profile:
            stat["motion"] = check_animation_compliance(
                slide, profile, mode, issues
            )
            stat["design_gene"] = check_design_gene(slide, profile, issues)

        # 空页
        if not all_text and stat["picture_count"] == 0 and stat["chart_count"] == 0 and stat["table_count"] == 0:
            issues.append({"type": "empty_slide", "level": "warn", "detail": "空白页（无文本且无图片/图表/表格）"})

        # 占位文本
        for m in PLACEHOLDER_PATTERNS:
            for t in texts:
                if re.search(m, t, re.IGNORECASE):
                    issues.append({"type": "placeholder_text", "level": "error", "detail": f"疑似占位文本 {t[:40]!r}（命中 {m}）"})
                    break

        # 几何
        for sh in shapes:
            check_slide_overflow(sh, slide_w, slide_h, issues)
        check_overlaps(shapes, issues, slide_w, slide_h)
        check_fonts_and_sizes(shapes, min_font, issues, font_set)
        collect_run_typography(shapes, font_set, size_set)
        check_broken_pictures(shapes, issues)
        if profile:
            # L4 custom covers intentionally use art-directed composition rather
            # than the content grid; their typography is still checked.
            if profile.level == "L4" and idx == 1:
                stat["grid_alignment"] = 1.0
            else:
                stat["grid_alignment"] = check_grid_alignment(shapes, profile, issues)
            stat["typography"] = check_font_hierarchy(shapes, profile, issues)

        # 重复页
        norm = normalize(all_text)
        if norm:
            if norm in seen_norm:
                issues.append({"type": "duplicate_slide", "level": "warn", "detail": f"与第 {seen_norm[norm]} 页内容重复"})
            else:
                seen_norm[norm] = idx

        stat["issues"] = issues
        for it in issues:
            target = report["errors"] if it["level"] == "error" else report["warnings"]
            target.append(f"[第{idx}页] {it['type']}: {it['detail']}")
        report["slides"].append(stat)

    report["font_names"] = sorted(font_set)
    report["font_sizes_pt"] = sorted(size_set)
    content_slides = report["slides"][1:-1] if len(report["slides"]) > 2 else report["slides"]
    text_dominant = [s["index"] for s in content_slides if s["text_dominant_candidate"]]
    total_visual_assets = sum(s["visual_asset_count"] for s in report["slides"])
    report["palette_colors"] = sorted(_deck_colors(prs))
    repeated_runs = []
    repetition_threshold = profile.layout.max_adjacent_same_layout if profile else 3
    run_start = 0
    for i in range(1, len(report["slides"]) + 1):
        previous = report["slides"][run_start]["layout_signature"]
        if i == len(report["slides"]) or report["slides"][i]["layout_signature"] != previous:
            if i - run_start > repetition_threshold:
                repeated_runs.append({
                    "start": report["slides"][run_start]["index"],
                    "end": report["slides"][i - 1]["index"],
                    "layout_signature": previous,
                    "run_length": i - run_start,
                })
            run_start = i

    report["visual_summary"] = {
        "total_visual_assets": total_visual_assets,
        "text_dominant_slides": text_dominant,
        "repeated_layout_runs": repeated_runs,
    }
    if profile:
        grid_values = [
            slide["grid_alignment"]
            for slide in report["slides"]
            if "grid_alignment" in slide
        ]
        dominant_sizes = []
        for slide in report["slides"]:
            sizes = slide.get("typography", {}).get("font_sizes", [])
            if sizes:
                dominant_sizes.append(max(sizes))
        report["strategy_summary"] = {
            "level": profile.level,
            "strategy": profile.strategy.key,
            "strategy_name": profile.strategy.name,
            "grid_columns": profile.grid.columns,
            "minimum_grid_alignment": min(grid_values) if grid_values else None,
            "font_family_count": len(font_set),
            "font_family_limit": profile.font.max_font_families,
            "font_size_count": len(size_set),
            "font_size_limit": profile.font.hierarchy_levels + 1,
            "dominant_title_sizes_pt": sorted(set(dominant_sizes)),
            "max_adjacent_same_layout": profile.layout.max_adjacent_same_layout,
            "design_principle": profile.design_gene.design_principle,
            "palette_declared_colors": profile.palette_count(),
            "palette_color_count": len(report["palette_colors"]),
            "motion_intensity": profile.animation.motion_intensity,
            "animated_slides": sum(
                1
                for slide in report["slides"]
                if slide.get("motion", {}).get("animation", {}).get("animated")
            ),
            "transitioned_slides": sum(
                1
                for slide in report["slides"]
                if slide.get("motion", {}).get("transition", {}).get("transition")
            ),
        }
        template_issues = []
        template_match = check_template_match(
            prs, template_style, template_issues
        )
        if template_match:
            report["template_match"] = template_match
            report["warnings"].extend(
                f"{item['type']}: {item['detail']}" for item in template_issues
            )
        if len(font_set) > profile.font.max_font_families:
            report["warnings"].append(
                f"字体家族 {len(font_set)} 个，超过 {profile.level} 策略上限 "
                f"{profile.font.max_font_families}。"
            )
        if len(size_set) > profile.font.hierarchy_levels + 1:
            report["warnings"].append(
                f"字号 {len(size_set)} 个，超过 {profile.level} 层级允许范围。"
            )
        if len(set(dominant_sizes)) > 1:
            report["warnings"].append(
                "各页主标题字号不一致："
                + ", ".join(f"{value:g}pt" for value in sorted(set(dominant_sizes)))
            )
        if profile.level == "L1":
            rounded = 0
            for slide in slides:
                for shape in slide.shapes:
                    try:
                        if "ROUNDED_RECTANGLE" in str(shape.auto_shape_type):
                            rounded += 1
                    except Exception:
                        pass
            if rounded:
                report["warnings"].append(
                    f"L1 信息优先策略检测到 {rounded} 个圆角组件。"
                )
        warning_text = "\n".join(report["warnings"])
        report["strategy_summary"]["checks"] = {
            "font_hierarchy_check": (
                "warn"
                if "字号" in warning_text
                or "主标题字号" in warning_text
                or "字体家族" in warning_text
                else "pass"
            ),
            "grid_alignment_check": (
                "warn"
                if report["strategy_summary"]["minimum_grid_alignment"] is not None
                and report["strategy_summary"]["minimum_grid_alignment"] < 0.45
                else "pass"
            ),
            "layout_diversity_check": "warn" if repeated_runs else "pass",
            "strategy_boundary_check": (
                "warn"
                if "L1 信息优先策略检测到" in warning_text
                else "pass"
            ),
            "animation_permission_check": (
                "warn"
                if any("animation_" in item for item in report["warnings"])
                else "pass"
            ),
            "design_gene_check": (
                "warn"
                if any("visual_center" in item for item in report["warnings"])
                else "pass"
            ),
        }
    if len(report["slides"]) >= 4 and total_visual_assets == 0:
        report["warnings"].append(
            "整套 PPT 未检测到图片、原生图表或表格；请人工确认是否缺少必要的视觉表达。"
        )
    if len(text_dominant) >= max(2, (len(content_slides) + 2) // 3):
        report["warnings"].append(
            f"检测到 {len(text_dominant)} 页文字主导内容（第 {', '.join(map(str, text_dominant))} 页）；请检查是否连续使用“标题 + 项目符号”版式。"
        )
    for run in repeated_runs:
        report["warnings"].append(
            f"第 {run['start']}–{run['end']} 页连续使用相近布局签名 {run['layout_signature']}；请人工复核版式是否机械重复。"
        )
    return report


def print_report(report):
    if report.get("skill_version"):
        print(f"Skill 版本：{report['skill_version']}")
    if report.get("strategy_name"):
        print(f"表达策略：{report['expression_level']} / {report['strategy_name']}")
    print(f"文件：{report['file']}")
    if "file_size_bytes" in report:
        print(f"大小：{report['file_size_bytes']} 字节")
    if report.get("slide_count") is not None:
        print(f"页数：{report['slide_count']}")
        print(f"页面尺寸：{report['slide_size_pt'][0]} x {report['slide_size_pt'][1]} pt（比例 {report['aspect_ratio']}）")
    if report["font_names"]:
        print(f"字体：{', '.join(report['font_names'])}")

    for sl in report["slides"]:
        s = sl
        flags = " ".join(f"{i['type']}" for i in s["issues"]) if s["issues"] else "OK"
        print(f"  第{s['index']:>2}页  形状{s['shape_count']:>2} 图{s['picture_count']} 表{s['table_count']} 图表{s['chart_count']} 文本{s['text_length']:>4}字  -> {flags}")

    if report["errors"]:
        print(f"\n错误 {len(report['errors'])} 条：")
        for e in report["errors"]:
            print(f"  [ERROR] {e}")
    if report["warnings"]:
        print(f"\n警告 {len(report['warnings'])} 条：")
        for w in report["warnings"]:
            print(f"  [WARN ] {w}")
    if not report["errors"] and not report["warnings"]:
        print("\n未发现明显问题。")
    if "visual_summary" in report:
        summary = report["visual_summary"]
        print(f"\n视觉结构提示：视觉资产 {summary['total_visual_assets']} 个；"
              f"文字主导候选页 {len(summary['text_dominant_slides'])} 页；"
              f"连续相近布局 {len(summary['repeated_layout_runs'])} 段。")
    if "strategy_summary" in report:
        strategy = report["strategy_summary"]
        print(
            f"策略质检：{strategy['grid_columns']} 列网格；"
            f"最低网格覆盖 {strategy['minimum_grid_alignment']:.0%}；"
            f"字体 {strategy['font_family_count']}/{strategy['font_family_limit']}；"
            f"字号 {strategy['font_size_count']}/{strategy['font_size_limit']}；"
            f"动效 {strategy['motion_intensity']}。"
        )
    if "template_match" in report:
        match = report["template_match"]
        print(
            f"模板匹配：{match['overall']:.0%} "
            f"（配色 {match['palette_cosine']:.0%}，字体 {match['font_match']:.0%}）"
        )
    print(f"\n汇总：错误 {len(report['errors'])}，警告 {len(report['warnings'])}。")


def main():
    ap = argparse.ArgumentParser(
        description="pro-ppt-studio V4.0.0 PPTX 结构、模板、动效与策略质检"
    )
    ap.add_argument("file", nargs="?", help="PPTX 文件路径")
    ap.add_argument("--json", dest="json_out", help="将结果导出为 JSON 文件")
    ap.add_argument("--min-font", type=float, default=10.0, help="最小字号阈值（pt），默认 10")
    ap.add_argument("--level", choices=["L1", "L2", "L3", "L4"], help="按表达策略执行附加质检")
    ap.add_argument(
        "--mode",
        choices=["auto", "static", "dynamic"],
        default="auto",
        help="动效交付模式；auto 只检查权限边界",
    )
    ap.add_argument(
        "--template-style",
        help="模板解析或迁移后的 JSON，用于模板匹配质检",
    )
    ap.add_argument(
        "--compare",
        nargs=4,
        metavar=("L1_PPTX", "L2_PPTX", "L3_PPTX", "L4_PPTX"),
        help="比较四级样例的视觉相似度",
    )
    args = ap.parse_args()

    if args.compare:
        report = compare_decks(args.compare)
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        if not args.file:
            ap.error("file is required unless --compare is used")
        report = check_pptx(
            args.file,
            min_font=args.min_font,
            level=args.level,
            mode=None if args.mode == "auto" else args.mode,
            template_style=args.template_style,
        )
        print_report(report)

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        print(f"\n已写入 JSON 报告：{args.json_out}")

    sys.exit(1 if report.get("errors") else 0)


if __name__ == "__main__":
    main()
