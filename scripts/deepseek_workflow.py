#!/usr/bin/env python3
"""Deterministic orchestration layer for Codex and DeepSeek V4 workflows.

The model plans content; these functions enforce sequencing, level selection,
template migration, editable generation, quality checks and render outputs.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

from check_pptx import check_pptx
from level_profiles import (
    LEVEL_PROFILES,
    SKILL_VERSION,
    get_profile,
    match_content_layout,
    normalize_level,
)
from leveled_pptx import LeveledDeckBuilder
from style_migrator import migrate_template
from template_parser import parse_template


DEEPSEEK_SYSTEM_PROMPT = """你是 pro-ppt-studio V4 的演示文稿设计代理。

硬约束：
1. 必须先确认内容表达等级。用户未指定时，解释 L1-L4 的场景差异并请求确认；不得默认 L3/L4。
2. 先选版式与视觉任务，再填充内容；每页必须写页面角色、内容类型、版式、视觉中心、组件和证据。
3. 严格遵守 level_profiles.py 的白名单和设计基因。低等级不得使用高等级组件、蒙版、拼贴、纹理或动画。
4. L1 完整保留原文；L2 保持章节化对称结构；L3 结论先行且每页只有一个视觉中心；L4 必须有跨页主题系统和叙事节奏。
5. 模板只迁移目标等级允许的色彩、字体、网格和组件，越级元素必须过滤并记录原因。
6. 不得虚构数据、来源、案例、Logo、产品图或实验结果。
7. 生成后必须执行结构、模板、动画与策略质检，并实际渲染后逐页审查。
"""


LEVEL_TERMS = {
    "L1": ("L1", "基础信息表达", "信息优先", "一般效果"),
    "L2": ("L2", "规范商务表达", "结构优先", "标准效果"),
    "L3": ("L3", "专业视觉设计", "重点强化", "精美效果"),
    "L4": ("L4", "高端定制设计", "主题叙事", "高端效果"),
}

SCENE_TERMS = {
    "课堂作业": ("课堂", "作业", "课程"),
    "内部草稿": ("草稿", "初稿", "留档"),
    "department_report": ("部门汇报", "工作汇报", "内部评审"),
    "project_review": ("项目汇报", "方案评审", "咨询交付", "经营分析"),
    "launch": ("科技发布", "发布会", "产品发布"),
    "brand": ("轻奢", "品牌", "高定"),
    "culture": ("文化", "文创", "非遗"),
    "roadshow": ("路演", "商业计划", "融资"),
}

CONTENT_ROLES = {
    "cover": ("title", "title"),
    "conclusion": ("narrative", "key_conclusion"),
    "metrics": ("metrics", "metric_cards"),
    "data": ("data", "data_dashboard"),
    "process": ("process", "process_steps"),
    "timeline": ("timeline", "timeline"),
    "comparison": ("comparison", "architecture_compare"),
    "case": ("case", "case_visual"),
    "evidence": ("image", "image_text_split"),
    "section": ("section", "section_transition"),
    "action": ("closing", "action_summary"),
}


def parse_requirement(user_input: str | Mapping[str, Any]) -> dict[str, Any]:
    """Extract structured requirements without inventing a missing level."""

    if isinstance(user_input, Mapping):
        data = dict(user_input)
        text = str(data.get("prompt") or data.get("text") or "")
    else:
        data = {}
        text = str(user_input)

    lower = text.lower()
    level = None
    specified_level = data.get("level")
    if specified_level:
        level = normalize_level(specified_level)
    else:
        for candidate, terms in LEVEL_TERMS.items():
            if any(re.search(re.escape(term), text, re.IGNORECASE) for term in terms):
                if re.search(rf"\b{candidate.lower()}\b", lower) or candidate in text:
                    level = candidate
                    break
                if any(term in text for term in terms[1:]):
                    level = candidate
                    break

    page_count = data.get("page_count")
    if page_count is None:
        matches = [
            int(value)
            for value in re.findall(r"(\d{1,3})\s*(?:页|张|slides?)", text, re.IGNORECASE)
        ]
        page_count = matches[-1] if matches else None

    dynamic = data.get("dynamic")
    if dynamic is None:
        dynamic = bool(re.search(r"动态|动画|视频|动效|animated", text, re.IGNORECASE))

    template_match = re.search(
        r"(?:模板|参考)[：:\s]*([^\s，。；;]+\.(?:pptx|potx|potm|thmx|pdf|png|jpe?g|webp|mp4|mov|svg|docx|json|ya?ml))",
        text,
        re.IGNORECASE,
    )
    template_path = data.get("template_path") or (
        template_match.group(1) if template_match else None
    )
    scene = data.get("scene")
    if not scene:
        scene = next(
            (
                key
                for key, terms in SCENE_TERMS.items()
                if any(term.lower() in lower for term in terms)
            ),
            "unspecified",
        )
    return {
        "topic": data.get("topic") or text[:80] or "未命名演示",
        "language": data.get("language") or ("zh-CN" if re.search(r"[\u4e00-\u9fff]", text) else "en-US"),
        "audience": data.get("audience"),
        "scene": scene,
        "page_count": page_count,
        "level": level,
        "requires_level_confirmation": level is None,
        "template_path": template_path,
        "dynamic": bool(dynamic),
        "must_keep": data.get("must_keep") or [],
        "forbidden": data.get("forbidden") or [],
        "output_path": data.get("output_path"),
        "raw_prompt": text,
    }


def select_strategy(level: str, scene: str | None = None) -> dict[str, Any]:
    profile = get_profile(level)
    return {
        "skill_version": SKILL_VERSION,
        "level": profile.level,
        "strategy": profile.strategy.asdict()
        if hasattr(profile.strategy, "asdict")
        else profile.as_dict()["strategy"],
        "design_gene": profile.as_dict()["design_gene"],
        "animation": profile.as_dict()["animation"],
        "scene": scene,
        "scene_match": scene in profile.strategy.scene_tags if scene else None,
        "layout_whitelist": list(profile.layout.allowed_layouts),
        "component_whitelist": list(profile.components.allowed_components),
        "content_rules": {
            "density_coefficient": profile.information.density_coefficient,
            "max_body_chars": profile.information.max_body_chars,
            "max_bullets": profile.information.max_bullets,
            "evidence_policy": profile.information.evidence_policy,
            "text_policy": profile.information.text_policy,
        },
    }


def parse_template_file(template_path: str | Path) -> dict[str, Any]:
    return parse_template(template_path).as_dict()


def _default_role_sequence(page_count: int) -> list[str]:
    if page_count <= 1:
        return ["cover"]
    if page_count == 2:
        return ["cover", "action"]
    middle = [
        "conclusion",
        "metrics",
        "data",
        "process",
        "timeline",
        "comparison",
        "case",
        "evidence",
    ]
    roles = ["cover"] + middle + ["section", "action"]
    return roles[:page_count]


def generate_outline(
    topic: str,
    level: str,
    page_count: int,
    *,
    points: Sequence[str] | None = None,
    model_callable: Callable[[str], Any] | None = None,
) -> dict[str, Any]:
    """Create a level-valid page plan.

    When a model callable is supplied, its JSON response is validated against
    the level whitelist. The fallback outline is deterministic and leaves
    factual claims to user-provided ``points``.
    """

    profile = get_profile(level)
    if page_count <= 0:
        raise ValueError("page_count must be positive")
    prompt = (
        DEEPSEEK_SYSTEM_PROMPT
        + f"\n主题：{topic}\n等级：{profile.level}\n页数：{page_count}\n"
        + f"允许版式：{list(profile.layout.allowed_layouts)}\n"
        + "返回 JSON：{\"slides\":[{\"role\":\"...\",\"title\":\"...\","
        + "\"content_type\":\"...\",\"layout\":\"...\",\"visual_center\":\"...\","
        + "\"components\":[\"...\"],\"evidence\":\"...\"}]}\n"
        + "不得虚构事实。"
    )
    if model_callable is not None:
        response = model_callable(prompt)
        payload = json.loads(response) if isinstance(response, str) else response
        slides = payload.get("slides", []) if isinstance(payload, dict) else []
        if len(slides) != page_count:
            raise ValueError("Model outline must match the requested page count")
        validated = []
        errors = []
        for index, slide in enumerate(slides, 1):
            layout = slide.get("layout") or match_content_layout(
                profile.level, slide.get("content_type", "narrative")
            )
            if layout not in profile.layout.allowed_layouts:
                errors.append(f"slide {index}: forbidden layout {layout!r}")
            components = [
                component
                for component in slide.get("components", [])
                if component in profile.components.allowed_components
            ]
            validated.append(
                {
                    **slide,
                    "page": index,
                    "layout": layout,
                    "components": components,
                }
            )
        if errors:
            raise ValueError("; ".join(errors))
        return {
            "level": profile.level,
            "design_gene": profile.as_dict()["design_gene"],
            "slides": validated,
            "source": "model",
        }

    roles = _default_role_sequence(page_count)
    user_points = list(points or [])
    slides = []
    for index, role in enumerate(roles, 1):
        content_type = CONTENT_ROLES.get(role, ("narrative", ""))[0]
        layout = match_content_layout(profile.level, content_type)
        evidence = user_points[index - 2] if 1 < index <= len(user_points) + 1 else ""
        slides.append(
            {
                "page": index,
                "role": role,
                "title": topic if role == "cover" else f"{topic} / {role}",
                "content_type": content_type,
                "layout": layout,
                "visual_center": (
                    "主题标题" if role in {"cover", "section"} else "主结论与证据"
                ),
                "components": [layout],
                "evidence": evidence,
            }
        )
    return {
        "level": profile.level,
        "design_gene": profile.as_dict()["design_gene"],
        "slides": slides,
        "source": "deterministic_plan",
    }


def write_design_plan(plan: Mapping[str, Any], output_path: str | Path) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    profile = get_profile(plan["level"])
    lines = [
        "# design_plan",
        "",
        f"- skill_version: {SKILL_VERSION}",
        f"- selected_expression_level: {profile.level}",
        f"- expression_strategy: {profile.strategy.name}",
        f"- strategy_objective: {profile.strategy.objective}",
        f"- design_principle: {profile.design_gene.design_principle}",
        f"- narrative_rhythm: {profile.design_gene.narrative_rhythm}",
        f"- visual_center_rule: {profile.design_gene.visual_center_rule}",
        "",
        "## 逐页规划",
        "",
        "| 页码 | 角色 | 标题 | 内容类型 | 版式 | 视觉中心 | 组件 | 证据 |",
        "|---:|---|---|---|---|---|---|---|",
    ]
    for slide in plan["slides"]:
        values = [
            slide.get("page"),
            slide.get("role"),
            slide.get("title"),
            slide.get("content_type"),
            slide.get("layout"),
            slide.get("visual_center"),
            ", ".join(slide.get("components", [])),
            slide.get("evidence"),
        ]
        escaped = [str(value or "").replace("|", "\\|") for value in values]
        lines.append("| " + " | ".join(escaped) + " |")
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return output


def build_pptx(
    outline: Mapping[str, Any],
    template_config: Mapping[str, Any] | None,
    output_path: str | Path,
    *,
    level: str | None = None,
    dynamic: bool = False,
    style_branch: str | None = None,
    audit_path: str | Path | None = None,
) -> Path:
    selected_level = normalize_level(level or outline["level"])
    builder = LeveledDeckBuilder(
        selected_level,
        style_branch=style_branch,
        mode="dynamic" if dynamic else "static",
        template_style=template_config,
    )
    slides = outline.get("slides", [])
    for index, slide in enumerate(slides, 1):
        role = slide.get("role", "conclusion")
        title = slide.get("title", "")
        evidence = str(slide.get("evidence") or "")
        if role == "cover":
            builder.add_cover(title, evidence or "演示文稿")
            continue
        if role == "section":
            builder.add_section(title, evidence)
            continue
        if role == "metrics":
            builder.add_metric_cards(
                title,
                [("指标", "待提供", evidence or "需提供口径")],
                insight=evidence,
            )
            continue
        if (
            role == "case"
            and "case_visual" in builder.profile.layout.allowed_layouts
            and hasattr(builder, "add_case_study")
        ):
            builder.add_case_study(
                title,
                evidence or "案例背景待提供",
                "设计任务待定义",
                "结果证据待提供",
            )
            continue
        if (
            role == "process"
            and "process_steps" in builder.profile.layout.allowed_layouts
        ):
            builder.add_process_steps(
                title,
                [("步骤 1", evidence or "待提供"), ("步骤 2", "待提供")],
            )
            continue
        if (
            role == "timeline"
            and "timeline" in builder.profile.layout.allowed_layouts
        ):
            builder.add_timeline(
                title,
                [("阶段 1", "开始", evidence or "待提供"), ("阶段 2", "完成", "待提供")],
            )
            continue
        if role == "action":
            builder.add_closing(title, [evidence] if evidence else ["请补充行动项。"])
            continue
        builder.add_content(
            title,
            [evidence] if evidence else ["请补充该页证据或说明。"],
            content_type=slide.get("content_type", "narrative"),
            requested_layout=slide.get("layout", ""),
        )
    return builder.save(output_path, audit_path=audit_path)


def full_quality_check(
    pptx_path: str | Path,
    level: str,
    *,
    template_style: Mapping[str, Any] | str | None = None,
    mode: str = "auto",
) -> dict[str, Any]:
    return check_pptx(
        pptx_path,
        level=level,
        mode=None if mode == "auto" else mode,
        template_style=template_style,
    )


def render_result(
    pptx_path: str | Path,
    out_dir: str | Path,
    *,
    video: bool = False,
    gif: bool = False,
) -> dict[str, Any]:
    script = Path(__file__).with_name("render_preview.py")
    command = [sys.executable, str(script), str(pptx_path), str(out_dir)]
    if video:
        command.append("--video")
    if gif:
        command.append("--gif")
    process = subprocess.run(command, capture_output=True, text=True)
    manifest_path = Path(out_dir) / "render_manifest.json"
    manifest = (
        json.loads(manifest_path.read_text(encoding="utf-8"))
        if manifest_path.is_file()
        else {}
    )
    return {
        "returncode": process.returncode,
        "stdout": process.stdout,
        "stderr": process.stderr,
        "manifest_path": str(manifest_path),
        "manifest": manifest,
    }


def run_workflow(
    user_input: str | Mapping[str, Any],
    *,
    output_dir: str | Path,
    model_callable: Callable[[str], Any] | None = None,
) -> dict[str, Any]:
    requirement = parse_requirement(user_input)
    if not requirement["level"]:
        return {
            "status": "level_confirmation_required",
            "requirement": requirement,
            "choices": {
                level: {
                    "name": profile.strategy.name,
                    "scenes": list(profile.strategy.scene_tags),
                    "design_principle": profile.design_gene.design_principle,
                }
                for level, profile in LEVEL_PROFILES.items()
            },
        }
    level = requirement["level"]
    if level == "L1" and requirement["dynamic"]:
        requirement["dynamic"] = False
        requirement["motion_note"] = "L1 不支持动画，已降级为静态交付。"
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    template_raw = (
        parse_template_file(requirement["template_path"])
        if requirement.get("template_path")
        else None
    )
    template_style = (
        migrate_template(template_raw, level) if template_raw else None
    )
    page_count = requirement.get("page_count") or 8
    outline = generate_outline(
        requirement["topic"],
        level,
        page_count,
        points=requirement.get("must_keep", []),
        model_callable=model_callable,
    )
    plan_path = write_design_plan(outline, output / "design_plan.md")
    pptx_path = build_pptx(
        outline,
        template_style,
        output / "presentation.pptx",
        dynamic=requirement["dynamic"],
        audit_path=output / "builder_audit.json",
    )
    quality = full_quality_check(
        pptx_path,
        level,
        template_style=template_style,
        mode="dynamic" if requirement["dynamic"] else "static",
    )
    (output / "quality_report.json").write_text(
        json.dumps(quality, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return {
        "status": "generated",
        "requirement": requirement,
        "strategy": select_strategy(level, requirement.get("scene")),
        "template": template_style,
        "outline": outline,
        "design_plan": str(plan_path),
        "pptx": str(pptx_path),
        "quality_report": quality,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run the deterministic pro-ppt-studio V4 workflow"
    )
    parser.add_argument("requirement", help="Prompt text or JSON requirement file")
    parser.add_argument("out_dir", help="Workflow output directory")
    args = parser.parse_args()
    source = Path(args.requirement)
    if source.is_file():
        payload: str | dict[str, Any] = json.loads(
            source.read_text(encoding="utf-8")
        )
    else:
        payload = args.requirement
    result = run_workflow(payload, output_dir=args.out_dir)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["status"] in {"generated", "level_confirmation_required"} else 1)


if __name__ == "__main__":
    main()
