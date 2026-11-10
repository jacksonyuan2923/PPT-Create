#!/usr/bin/env python3
"""Canonical V4.0.0 expression-strategy profiles for pro-ppt-studio.

The four levels are peer strategies, not a visual quality ladder. This module
is the machine-readable source of truth for strategy intent, typography,
layout permissions, component permissions, grids, density, design genes,
palette systems, motion permissions, and L4 branches.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field, replace
from hashlib import sha256
import json
from typing import Any, Iterable


SKILL_VERSION = "4.0.0"
STRATEGY_SCHEMA_VERSION = "2.0"

LEVEL_ALIASES = {
    "L1": "L1",
    "1": "L1",
    "基础信息表达": "L1",
    "信息优先": "L1",
    "信息优先策略": "L1",
    "一般": "L1",
    "一般效果": "L1",
    "L2": "L2",
    "2": "L2",
    "规范商务表达": "L2",
    "结构优先": "L2",
    "结构优先策略": "L2",
    "标准": "L2",
    "标准效果": "L2",
    "L3": "L3",
    "3": "L3",
    "专业视觉设计": "L3",
    "重点强化": "L3",
    "重点强化策略": "L3",
    "精美": "L3",
    "精美效果": "L3",
    "L4": "L4",
    "4": "L4",
    "高端定制设计": "L4",
    "主题叙事": "L4",
    "主题叙事策略": "L4",
    "高端": "L4",
    "高端效果": "L4",
}


@dataclass(frozen=True)
class ScoreDimension:
    key: str
    label: str
    max_points: int


@dataclass(frozen=True)
class StrategyRule:
    key: str
    name: str
    objective: str
    scene_tags: tuple[str, ...]
    acceptance: tuple[str, ...]
    acceptance_min_score: int
    score_dimensions: tuple[ScoreDimension, ...]


@dataclass(frozen=True)
class TypeLevel:
    key: str
    label: str
    size_pt: float
    weight: str
    line_spacing: float
    font_role: str


@dataclass(frozen=True)
class FontRule:
    title_font: str
    body_font: str
    accent_font: str
    numeric_font: str
    max_font_families: int
    hierarchy_levels: int
    levels: tuple[TypeLevel, ...]
    metadata_size_pt: float
    weight_policy: str
    allow_gradient_text: bool
    allow_shadow: bool
    allow_outline: bool
    allow_texture: bool
    allow_vertical_text: bool

    def size_for(self, role: str) -> float:
        aliases = {
            "subtitle": "module" if "module" in {item.key for item in self.levels} else "body",
            "section": "module" if "module" in {item.key for item in self.levels} else "title",
            "accent": "module" if "module" in {item.key for item in self.levels} else "title",
            "caption": "caption",
        }
        key = aliases.get(role, role)
        for item in self.levels:
            if item.key == key:
                return item.size_pt
        if key == "caption":
            return self.metadata_size_pt
        return self.levels[-1].size_pt

    def line_spacing_for(self, role: str) -> float:
        aliases = {
            "subtitle": "module",
            "section": "module",
            "accent": "module",
            "caption": "caption",
        }
        key = aliases.get(role, role)
        for item in self.levels:
            if item.key == key:
                return item.line_spacing
        return 1.2


@dataclass(frozen=True)
class GridRule:
    columns: int
    margin_left_in: float
    margin_right_in: float
    margin_top_in: float
    margin_bottom_in: float
    gutter_in: float
    baseline_pt: float
    component_radius_in: float
    shadow_policy: str
    shadow_opacity: float
    shadow_blur_pt: float
    alignment_tolerance_in: float

    @property
    def content_width_in(self) -> float:
        return 13.333 - self.margin_left_in - self.margin_right_in

    @property
    def column_width_in(self) -> float:
        return (
            self.content_width_in - self.gutter_in * (self.columns - 1)
        ) / self.columns

    def vertical_lines(self) -> tuple[float, ...]:
        lines = [self.margin_left_in]
        step = self.column_width_in + self.gutter_in
        for index in range(1, self.columns):
            lines.append(self.margin_left_in + index * step - self.gutter_in / 2)
            lines.append(self.margin_left_in + index * step)
        lines.append(13.333 - self.margin_right_in)
        return tuple(lines)


@dataclass(frozen=True)
class ComponentRule:
    allowed_components: tuple[str, ...]
    forbidden_components: tuple[str, ...]
    max_cards_per_slide: int
    allow_data_components: bool
    allow_icon_modules: bool
    allow_collage_components: bool


@dataclass(frozen=True)
class LayoutRule:
    allowed_layouts: tuple[str, ...]
    forbidden_layouts: tuple[str, ...]
    content_layout_map: dict[str, str]
    cover_layout: str
    content_layout: str
    section_layout: str
    closing_layout: str
    background_mode: str
    allow_background_image: bool
    allow_cards: bool
    allow_rounded_cards: bool
    allow_image_mask: bool
    allow_collage: bool
    allow_icons: bool
    allow_decorative_shapes: bool
    allow_asymmetric: bool
    fixed_header: bool
    fixed_footer: bool
    whitespace_level: str
    max_adjacent_same_layout: int


@dataclass(frozen=True)
class ImageRule:
    default_fit: str
    allow_mask: bool
    allow_rounding: bool
    allow_border: bool
    allow_collage: bool
    max_images_per_slide: int
    treatment: str


@dataclass(frozen=True)
class ChartRule:
    preferred_types: tuple[str, ...]
    show_data_labels: bool
    show_legend: bool
    show_gridlines: bool
    allow_gradient: bool
    allow_glow: bool
    max_series: int
    axis_zero_based: bool
    style_note: str


@dataclass(frozen=True)
class InformationRule:
    density_label: str
    density_coefficient: float
    max_body_chars: int | None
    max_bullets: int | None
    target_visual_ratio: tuple[float, float]
    evidence_policy: str
    text_policy: str


@dataclass(frozen=True)
class StyleBranch:
    key: str
    name: str
    summary: str
    palette_overrides: dict[str, str]
    font_overrides: dict[str, str]
    composition_rules: tuple[str, ...]
    image_policy: str
    visual_devices: tuple[str, ...]
    forbidden: tuple[str, ...]


@dataclass(frozen=True)
class DesignGene:
    """The visual logic that makes each level unmistakably different.

    Permissions alone only prevent illegal calls. These fields describe how a
    valid page is organised: reading path, rhythm, visual weight, and the
    recurring design moves that define the level.
    """

    design_principle: str
    narrative_rhythm: str
    information_density_curve: tuple[float, ...]
    composition_signature: str
    visual_center_rule: str
    hierarchy_strategy: str
    object_character: str
    spacing_behavior: str
    signature_moves: tuple[str, ...]
    anti_patterns: tuple[str, ...]


@dataclass(frozen=True)
class AnimationRule:
    """Machine-enforced motion contract for one content level."""

    transition_types: tuple[str, ...]
    entrance_types: tuple[str, ...]
    emphasis_types: tuple[str, ...]
    allow_exit: bool
    allow_motion_path: bool
    allow_trigger: bool
    allow_media: bool
    max_element_animations: int | None
    duration_range_ms: tuple[int, int]
    stagger_ms: int
    default_transition: str
    motion_intensity: str


@dataclass(frozen=True)
class PaletteSystem:
    """Palette inventory plus controlled-usage rules."""

    primary: tuple[str, ...]
    secondary: tuple[str, ...]
    neutrals: tuple[str, ...]
    accents: tuple[str, ...]
    functional: dict[str, str]
    max_colors: int
    usage_ratio: dict[str, float]
    contrast_policy: str


_EMPTY_DESIGN_GENE = DesignGene(
    design_principle="",
    narrative_rhythm="",
    information_density_curve=(),
    composition_signature="",
    visual_center_rule="",
    hierarchy_strategy="",
    object_character="",
    spacing_behavior="",
    signature_moves=(),
    anti_patterns=(),
)
_EMPTY_ANIMATION = AnimationRule(
    transition_types=(),
    entrance_types=(),
    emphasis_types=(),
    allow_exit=False,
    allow_motion_path=False,
    allow_trigger=False,
    allow_media=False,
    max_element_animations=0,
    duration_range_ms=(0, 0),
    stagger_ms=0,
    default_transition="",
    motion_intensity="none",
)
_EMPTY_PALETTE_SYSTEM = PaletteSystem(
    primary=(),
    secondary=(),
    neutrals=(),
    accents=(),
    functional={},
    max_colors=0,
    usage_ratio={},
    contrast_policy="",
)


@dataclass(frozen=True)
class LevelProfile:
    level: str
    strategy: StrategyRule
    font: FontRule
    grid: GridRule
    components: ComponentRule
    layout: LayoutRule
    images: ImageRule
    charts: ChartRule
    information: InformationRule
    palette: dict[str, str]
    style_branches: tuple[StyleBranch, ...]
    design_gene: DesignGene = field(default=_EMPTY_DESIGN_GENE)
    animation: AnimationRule = field(default=_EMPTY_ANIMATION)
    palette_system: PaletteSystem = field(default=_EMPTY_PALETTE_SYSTEM)

    @property
    def name(self) -> str:
        return self.strategy.name

    @property
    def acceptance_range(self) -> tuple[int, int]:
        return (self.strategy.acceptance_min_score, 100)

    def as_dict(self) -> dict[str, Any]:
        return _json_ready(asdict(self))

    def fingerprint(self) -> str:
        payload = json.dumps(self.as_dict(), sort_keys=True, ensure_ascii=False)
        return sha256(payload.encode("utf-8")).hexdigest()[:16]

    def style_branch(self, key: str | None = None) -> StyleBranch | None:
        if self.level != "L4":
            return None
        if not self.style_branches:
            raise ValueError("L4 profile does not define style branches")
        selected = key or self.style_branches[0].key
        for branch in self.style_branches:
            if branch.key == selected:
                return branch
        raise ValueError(
            f"Unsupported L4 style branch {selected!r}; "
            f"expected one of {[branch.key for branch in self.style_branches]}"
        )

    def palette_count(self) -> int:
        """Return the declared, unique palette size used by visual QA."""

        values = {
            value.upper()
            for value in (
                *self.palette_system.primary,
                *self.palette_system.secondary,
                *self.palette_system.neutrals,
                *self.palette_system.accents,
            )
        }
        return len(values)


def _json_ready(value: Any) -> Any:
    if isinstance(value, tuple):
        return [_json_ready(item) for item in value]
    if isinstance(value, list):
        return [_json_ready(item) for item in value]
    if isinstance(value, dict):
        return {key: _json_ready(item) for key, item in value.items()}
    return value


def _levels(*items: tuple[str, str, float, str, float, str]) -> tuple[TypeLevel, ...]:
    return tuple(TypeLevel(*item) for item in items)


def _dimensions(*items: tuple[str, str, int]) -> tuple[ScoreDimension, ...]:
    return tuple(ScoreDimension(*item) for item in items)


LEVEL_PROFILES: dict[str, LevelProfile] = {
    "L1": LevelProfile(
        level="L1",
        strategy=StrategyRule(
            key="information_first",
            name="基础信息表达 · 信息优先策略",
            objective="以最低成本完成准确、高效的信息传递。",
            scene_tags=("课堂作业", "内部草稿", "原始资料快速转写", "信息留档"),
            acceptance=("信息完整", "文字清晰可读", "无意外排版错误"),
            acceptance_min_score=80,
            score_dimensions=_dimensions(
                ("content_completeness", "信息完整度", 30),
                ("readability", "可读性", 25),
                ("structure", "基础结构", 15),
                ("layout_basics", "基础版式", 10),
                ("strategy_boundary", "策略边界", 10),
                ("execution", "执行准确性", 10),
            ),
        ),
        font=FontRule(
            title_font="Microsoft YaHei",
            body_font="Microsoft YaHei",
            accent_font="Microsoft YaHei",
            numeric_font="Microsoft YaHei",
            max_font_families=1,
            hierarchy_levels=2,
            levels=_levels(
                ("title", "页面标题", 30, "bold", 1.15, "title"),
                ("body", "正文", 17, "regular", 1.3, "body"),
            ),
            metadata_size_pt=12,
            weight_policy="single_regular_sans_with_bold_titles",
            allow_gradient_text=False,
            allow_shadow=False,
            allow_outline=False,
            allow_texture=False,
            allow_vertical_text=False,
        ),
        grid=GridRule(
            columns=2,
            margin_left_in=0.72,
            margin_right_in=0.72,
            margin_top_in=0.65,
            margin_bottom_in=0.55,
            gutter_in=0.4,
            baseline_pt=0,
            component_radius_in=0.0,
            shadow_policy="none",
            shadow_opacity=0.0,
            shadow_blur_pt=0,
            alignment_tolerance_in=0.2,
        ),
        components=ComponentRule(
            allowed_components=("title", "bullet_list", "divider_line"),
            forbidden_components=(
                "card",
                "metric_card",
                "process_step",
                "timeline",
                "icon_module",
                "image_mask",
                "collage",
            ),
            max_cards_per_slide=0,
            allow_data_components=False,
            allow_icon_modules=False,
            allow_collage_components=False,
        ),
        layout=LayoutRule(
            allowed_layouts=("title_body", "simple_split"),
            forbidden_layouts=(
                "structured_cards",
                "image_mask_hero",
                "metric_cards",
                "data_dashboard",
                "timeline",
                "process_steps",
                "editorial_collage",
            ),
            content_layout_map={
                "title": "title_body",
                "bullets": "title_body",
                "narrative": "title_body",
                "section": "simple_split",
                "closing": "simple_split",
                "image": "simple_split",
                "data": "title_body",
            },
            cover_layout="title_body",
            content_layout="title_body",
            section_layout="simple_split",
            closing_layout="simple_split",
            background_mode="solid_white",
            allow_background_image=False,
            allow_cards=False,
            allow_rounded_cards=False,
            allow_image_mask=False,
            allow_collage=False,
            allow_icons=False,
            allow_decorative_shapes=False,
            allow_asymmetric=False,
            fixed_header=False,
            fixed_footer=False,
            whitespace_level="functional",
            max_adjacent_same_layout=20,
        ),
        images=ImageRule(
            default_fit="contain",
            allow_mask=False,
            allow_rounding=False,
            allow_border=False,
            allow_collage=False,
            max_images_per_slide=1,
            treatment="direct_unstyled",
        ),
        charts=ChartRule(
            preferred_types=("column", "line", "table"),
            show_data_labels=False,
            show_legend=True,
            show_gridlines=True,
            allow_gradient=False,
            allow_glow=False,
            max_series=1,
            axis_zero_based=True,
            style_note="基础可读图表；保留单位和必要图例，不添加装饰。",
        ),
        information=InformationRule(
            density_label="complete",
            density_coefficient=1.0,
            max_body_chars=None,
            max_bullets=None,
            target_visual_ratio=(0.0, 0.10),
            evidence_policy="source_text_verbatim",
            text_policy="完整保留原始信息；不为了视觉空间删除、概括或碎片化原文。",
        ),
        palette={
            "background": "#FFFFFF",
            "surface": "#FFFFFF",
            "text": "#111111",
            "secondary": "#555555",
            "accent": "#666666",
            "line": "#B8B8B8",
            "contrast": "#000000",
        },
        style_branches=(),
    ),
    "L2": LevelProfile(
        level="L2",
        strategy=StrategyRule(
            key="structure_first",
            name="规范商务表达 · 结构优先策略",
            objective="通过标准化结构建立正式感，提升结构化阅读效率。",
            scene_tags=("部门汇报", "项目申报", "迎检材料", "内部评审"),
            acceptance=("正式统一", "章节清晰", "符合商务规范"),
            acceptance_min_score=80,
            score_dimensions=_dimensions(
                ("structure_fit", "结构匹配", 20),
                ("content_organization", "内容组织", 20),
                ("formal_consistency", "正式一致性", 20),
                ("hierarchy", "层级清晰度", 15),
                ("components", "商务组件", 15),
                ("execution", "执行准确性", 10),
            ),
        ),
        font=FontRule(
            title_font="Microsoft YaHei",
            body_font="Microsoft YaHei",
            accent_font="Arial",
            numeric_font="Arial",
            max_font_families=2,
            hierarchy_levels=3,
            levels=_levels(
                ("title", "页面标题", 34, "bold", 1.12, "title"),
                ("module", "模块标题", 20, "semibold", 1.2, "body"),
                ("body", "正文", 16, "regular", 1.28, "body"),
            ),
            metadata_size_pt=10,
            weight_policy="formal_sans_pair_bold_title",
            allow_gradient_text=False,
            allow_shadow=False,
            allow_outline=False,
            allow_texture=False,
            allow_vertical_text=False,
        ),
        grid=GridRule(
            columns=12,
            margin_left_in=0.68,
            margin_right_in=0.68,
            margin_top_in=0.55,
            margin_bottom_in=0.5,
            gutter_in=0.22,
            baseline_pt=4,
            component_radius_in=0.04,
            shadow_policy="none",
            shadow_opacity=0.0,
            shadow_blur_pt=0,
            alignment_tolerance_in=0.12,
        ),
        components=ComponentRule(
            allowed_components=(
                "title_bar",
                "section_number",
                "header",
                "footer",
                "basic_card",
                "framed_image",
                "minimal_icon",
                "comparison",
                "standard_chart",
                "table",
            ),
            forbidden_components=(
                "metric_dashboard",
                "process_timeline",
                "image_mask",
                "collage",
                "custom_infographic",
            ),
            max_cards_per_slide=4,
            allow_data_components=True,
            allow_icon_modules=True,
            allow_collage_components=False,
        ),
        layout=LayoutRule(
            allowed_layouts=(
                "title_band",
                "structured_title_bar",
                "symmetric_columns",
                "structured_cards",
                "framed_image_text",
                "structured_table",
                "numbered_section",
                "structured_summary",
            ),
            forbidden_layouts=(
                "image_mask_hero",
                "metric_cards",
                "data_dashboard",
                "process_steps",
                "timeline",
                "editorial_collage",
            ),
            content_layout_map={
                "title": "title_band",
                "bullets": "structured_title_bar",
                "narrative": "symmetric_columns",
                "comparison": "symmetric_columns",
                "metrics": "structured_cards",
                "cards": "structured_cards",
                "image": "framed_image_text",
                "case": "framed_image_text",
                "data": "structured_table",
                "table": "structured_table",
                "section": "numbered_section",
                "closing": "structured_summary",
            },
            cover_layout="title_band",
            content_layout="symmetric_columns",
            section_layout="numbered_section",
            closing_layout="structured_summary",
            background_mode="plain_with_band",
            allow_background_image=False,
            allow_cards=True,
            allow_rounded_cards=False,
            allow_image_mask=False,
            allow_collage=False,
            allow_icons=True,
            allow_decorative_shapes=False,
            allow_asymmetric=False,
            fixed_header=True,
            fixed_footer=True,
            whitespace_level="medium",
            max_adjacent_same_layout=3,
        ),
        images=ImageRule(
            default_fit="cover",
            allow_mask=False,
            allow_rounding=False,
            allow_border=True,
            allow_collage=False,
            max_images_per_slide=2,
            treatment="framed_realistic",
        ),
        charts=ChartRule(
            preferred_types=("column", "bar", "line", "pie", "table"),
            show_data_labels=False,
            show_legend=True,
            show_gridlines=False,
            allow_gradient=False,
            allow_glow=False,
            max_series=2,
            axis_zero_based=True,
            style_note="正式商务配色；保留单位、图例和比较口径。",
        ),
        information=InformationRule(
            density_label="structured",
            density_coefficient=0.78,
            max_body_chars=650,
            max_bullets=8,
            target_visual_ratio=(0.15, 0.30),
            evidence_policy="structured_source_points",
            text_policy="按章节与小标题组织内容，保留完整论述，不把长文压缩成口号。",
        ),
        palette={
            "background": "#FFFFFF",
            "surface": "#EEF3F8",
            "text": "#10263B",
            "secondary": "#4A647A",
            "accent": "#1F5FA8",
            "line": "#B7CADB",
            "contrast": "#0A1B2A",
        },
        style_branches=(),
    ),
    "L3": LevelProfile(
        level="L3",
        strategy=StrategyRule(
            key="emphasis_first",
            name="专业视觉设计 · 重点强化策略",
            objective="用专业可视化手段强化核心论点，兼顾正式感与设计感。",
            scene_tags=(
                "项目汇报",
                "方案评审",
                "咨询交付",
                "经营分析",
                "产品介绍",
                "客户提案",
            ),
            acceptance=("重点突出", "逻辑可视化", "现代专业", "商务可读性"),
            acceptance_min_score=80,
            score_dimensions=_dimensions(
                ("key_emphasis", "重点突出度", 25),
                ("visual_logic", "逻辑可视化", 20),
                ("layout_diversity", "版式丰富度", 15),
                ("hierarchy", "信息层级", 15),
                ("business_professionalism", "商务专业度", 15),
                ("execution", "执行准确性", 10),
            ),
        ),
        font=FontRule(
            title_font="Microsoft YaHei",
            body_font="Microsoft YaHei",
            accent_font="Consolas",
            numeric_font="Consolas",
            max_font_families=3,
            hierarchy_levels=4,
            levels=_levels(
                ("title", "主标题", 36, "bold", 1.12, "title"),
                ("module", "模块标题", 22, "semibold", 1.18, "body"),
                ("body", "正文", 16, "regular", 1.26, "body"),
                ("caption", "注释", 11, "regular", 1.2, "body"),
            ),
            metadata_size_pt=11,
            weight_policy="title_body_and_monospace_numeric",
            allow_gradient_text=False,
            allow_shadow=False,
            allow_outline=False,
            allow_texture=False,
            allow_vertical_text=False,
        ),
        grid=GridRule(
            columns=12,
            margin_left_in=0.72,
            margin_right_in=0.72,
            margin_top_in=0.5,
            margin_bottom_in=0.48,
            gutter_in=0.2,
            baseline_pt=6,
            component_radius_in=0.12,
            shadow_policy="single_soft_ambient",
            shadow_opacity=0.12,
            shadow_blur_pt=12,
            alignment_tolerance_in=0.08,
        ),
        components=ComponentRule(
            allowed_components=(
                "key_conclusion",
                "metric_card",
                "data_dashboard",
                "process_step",
                "timeline",
                "architecture_compare",
                "case_visual",
                "section_transition",
                "icon_text_module",
                "standard_chart",
                "standard_table",
            ),
            forbidden_components=(
                "texture_background",
                "editorial_collage",
                "custom_narrative_collage",
                "multiple_shadow_styles",
            ),
            max_cards_per_slide=4,
            allow_data_components=True,
            allow_icon_modules=True,
            allow_collage_components=False,
        ),
        layout=LayoutRule(
            allowed_layouts=(
                "key_conclusion",
                "metric_cards",
                "data_dashboard",
                "process_steps",
                "timeline",
                "architecture_compare",
                "case_visual",
                "section_transition",
                "icon_text_modules",
                "image_text_split",
                "chart_analysis",
                "standard_table",
                "action_summary",
            ),
            forbidden_layouts=(
                "texture_collage",
                "editorial_collage",
                "decorated_narrative",
            ),
            content_layout_map={
                "title": "key_conclusion",
                "bullets": "key_conclusion",
                "narrative": "image_text_split",
                "comparison": "architecture_compare",
                "metrics": "metric_cards",
                "data": "data_dashboard",
                "process": "process_steps",
                "timeline": "timeline",
                "architecture": "architecture_compare",
                "case": "case_visual",
                "image": "image_text_split",
                "table": "standard_table",
                "section": "section_transition",
                "closing": "action_summary",
            },
            cover_layout="case_visual",
            content_layout="key_conclusion",
            section_layout="section_transition",
            closing_layout="action_summary",
            background_mode="clean_neutral_with_local_surfaces",
            allow_background_image=True,
            allow_cards=True,
            allow_rounded_cards=True,
            allow_image_mask=True,
            allow_collage=False,
            allow_icons=True,
            allow_decorative_shapes=True,
            allow_asymmetric=True,
            fixed_header=True,
            fixed_footer=True,
            whitespace_level="medium_high",
            max_adjacent_same_layout=2,
        ),
        images=ImageRule(
            default_fit="cover",
            allow_mask=True,
            allow_rounding=True,
            allow_border=False,
            allow_collage=False,
            max_images_per_slide=2,
            treatment="business_cover_or_rounded_case_visual",
        ),
        charts=ChartRule(
            preferred_types=("bar", "line", "donut", "bullet", "table"),
            show_data_labels=True,
            show_legend=False,
            show_gridlines=True,
            allow_gradient=False,
            allow_glow=False,
            max_series=3,
            axis_zero_based=True,
            style_note="专业商务图表；单一主结论、清晰单位与来源、无装饰性截断轴。",
        ),
        information=InformationRule(
            density_label="balanced",
            density_coefficient=0.58,
            max_body_chars=320,
            max_bullets=5,
            target_visual_ratio=(0.40, 0.60),
            evidence_policy="conclusion_then_evidence",
            text_policy="结论先行，正文提炼为关键词与短句，证据用图表、流程、案例或图示承载。",
        ),
        palette={
            "background": "#F5F7FA",
            "surface": "#FFFFFF",
            "text": "#122033",
            "secondary": "#5D6B7A",
            "accent": "#0B7A75",
            "line": "#C8D2DC",
            "contrast": "#08131F",
            "highlight": "#DCEFEB",
        },
        style_branches=(),
    ),
    "L4": LevelProfile(
        level="L4",
        strategy=StrategyRule(
            key="theme_narrative",
            name="高端定制设计 · 主题叙事策略",
            objective="围绕核心主题建立沉浸式视觉叙事与强识别度视觉系统。",
            scene_tags=("科技发布", "轻奢品牌", "文化文创", "产品路演"),
            acceptance=("主题统一贯穿", "风格独特完整", "明确艺术指导"),
            acceptance_min_score=80,
            score_dimensions=_dimensions(
                ("theme_narrative", "主题叙事", 25),
                ("visual_system", "视觉系统", 20),
                ("composition", "构图创新", 20),
                ("consistency", "跨页一致性", 15),
                ("information_support", "信息支撑", 10),
                ("execution", "执行准确性", 10),
            ),
        ),
        font=FontRule(
            title_font="STZhongsong",
            body_font="Microsoft YaHei",
            accent_font="Arial",
            numeric_font="Arial",
            max_font_families=4,
            hierarchy_levels=5,
            levels=_levels(
                ("title", "主题标题", 44, "bold", 1.08, "title"),
                ("lead", "导语", 24, "medium", 1.2, "body"),
                ("section", "章节标题", 20, "semibold", 1.2, "body"),
                ("body", "正文", 16, "regular", 1.28, "body"),
                ("caption", "注释", 10, "regular", 1.2, "body"),
            ),
            metadata_size_pt=10,
            weight_policy="theme_typography_with_visual_expression",
            allow_gradient_text=True,
            allow_shadow=True,
            allow_outline=True,
            allow_texture=True,
            allow_vertical_text=True,
        ),
        grid=GridRule(
            columns=12,
            margin_left_in=0.5,
            margin_right_in=0.5,
            margin_top_in=0.45,
            margin_bottom_in=0.42,
            gutter_in=0.24,
            baseline_pt=8,
            component_radius_in=0.18,
            shadow_policy="theme_controlled",
            shadow_opacity=0.2,
            shadow_blur_pt=24,
            alignment_tolerance_in=0.16,
        ),
        components=ComponentRule(
            allowed_components=(
                "immersive_hero",
                "editorial_collage",
                "cross_slide_motif",
                "custom_infographic",
                "metric_card",
                "timeline",
                "process_step",
                "case_visual",
                "thematic_texture",
                "custom_chart",
            ),
            forbidden_components=(),
            max_cards_per_slide=5,
            allow_data_components=True,
            allow_icon_modules=True,
            allow_collage_components=True,
        ),
        layout=LayoutRule(
            allowed_layouts=(
                "custom_cover",
                "full_bleed_story",
                "editorial_collage",
                "asymmetric_narrative",
                "texture_collage",
                "immersive_section",
                "custom_infographic",
                "narrative_timeline",
                "product_roadshow",
                "narrative_closing",
            ),
            forbidden_layouts=(),
            content_layout_map={
                "title": "custom_cover",
                "bullets": "asymmetric_narrative",
                "narrative": "editorial_collage",
                "comparison": "asymmetric_narrative",
                "metrics": "custom_infographic",
                "data": "custom_infographic",
                "process": "narrative_timeline",
                "timeline": "narrative_timeline",
                "architecture": "asymmetric_narrative",
                "case": "full_bleed_story",
                "image": "full_bleed_story",
                "table": "asymmetric_narrative",
                "section": "immersive_section",
                "closing": "narrative_closing",
                "product": "product_roadshow",
            },
            cover_layout="custom_cover",
            content_layout="editorial_collage",
            section_layout="immersive_section",
            closing_layout="narrative_closing",
            background_mode="theme_controlled_layers",
            allow_background_image=True,
            allow_cards=True,
            allow_rounded_cards=True,
            allow_image_mask=True,
            allow_collage=True,
            allow_icons=True,
            allow_decorative_shapes=True,
            allow_asymmetric=True,
            fixed_header=False,
            fixed_footer=False,
            whitespace_level="high",
            max_adjacent_same_layout=1,
        ),
        images=ImageRule(
            default_fit="cover",
            allow_mask=True,
            allow_rounding=True,
            allow_border=True,
            allow_collage=True,
            max_images_per_slide=5,
            treatment="theme_narrative_masks_and_collage",
        ),
        charts=ChartRule(
            preferred_types=(
                "illustrated_bar",
                "editorial_line",
                "donut",
                "custom_infographic",
            ),
            show_data_labels=True,
            show_legend=False,
            show_gridlines=False,
            allow_gradient=True,
            allow_glow=True,
            max_series=4,
            axis_zero_based=True,
            style_note="图表服从主题艺术指导，但不得牺牲单位、范围与比较口径。",
        ),
        information=InformationRule(
            density_label="narrative",
            density_coefficient=0.40,
            max_body_chars=180,
            max_bullets=3,
            target_visual_ratio=(0.60, 0.85),
            evidence_policy="theme_claims_with_visual_evidence",
            text_policy="以场景、图像和点睛文案承载叙事；长信息拆分到辅助页面或附录。",
        ),
        palette={
            "background": "#071016",
            "surface": "#11232B",
            "text": "#F5F8F7",
            "secondary": "#B7C4C2",
            "accent": "#35D3E2",
            "line": "#45636B",
            "contrast": "#000000",
            "highlight": "#D7F5F7",
        },
        style_branches=(
            StyleBranch(
                key="technology_launch",
                name="科技发布",
                summary="深色空间、精确线路、产品层级与发布节奏。",
                palette_overrides={
                    "background": "#071016",
                    "surface": "#10252B",
                    "text": "#F3FAFC",
                    "secondary": "#9AB3BB",
                    "accent": "#22D3EE",
                    "line": "#28515C",
                    "contrast": "#000000",
                    "highlight": "#A7F3FA",
                },
                font_overrides={
                    "title_font": "Microsoft YaHei",
                    "body_font": "Microsoft YaHei",
                    "accent_font": "Consolas",
                    "numeric_font": "Consolas",
                },
                composition_rules=("深色全幅主视觉", "水平信号线", "产品细节特写", "节点式信息层级"),
                image_policy="产品渲染、界面细节、抽象系统图；避免廉价粒子背景。",
                visual_devices=("fine_lines", "signal_nodes", "product_detail"),
                forbidden=("通用蓝紫渐变模板", "随机粒子", "无关机器人插画"),
            ),
            StyleBranch(
                key="luxury_brand",
                name="轻奢品牌",
                summary="暖中性色、精确留白、材质特写与克制金色点缀。",
                palette_overrides={
                    "background": "#F4EFE6",
                    "surface": "#FFFDF8",
                    "text": "#1F1A17",
                    "secondary": "#74695F",
                    "accent": "#B48A4A",
                    "line": "#D9CCBB",
                    "contrast": "#14110F",
                    "highlight": "#E8D8B9",
                },
                font_overrides={
                    "title_font": "STZhongsong",
                    "body_font": "Microsoft YaHei",
                    "accent_font": "Arial",
                    "numeric_font": "Arial",
                },
                composition_rules=("大留白", "非对称静物主视觉", "细字距", "跨页材质呼应"),
                image_policy="品牌产品、材质、工艺和真实空间；禁止把生成图伪装成真实产品。",
                visual_devices=("material_closeup", "hairline", "precise_crop"),
                forbidden=("金色大面积铺底", "厚重欧式花纹", "拥挤卖点墙"),
            ),
            StyleBranch(
                key="cultural_creative",
                name="文化文创",
                summary="纸张质感、现代编辑构图、文化符号提炼与内容化纹样。",
                palette_overrides={
                    "background": "#F3EDE2",
                    "surface": "#E8DDCC",
                    "text": "#2B2520",
                    "secondary": "#705E4B",
                    "accent": "#A44A32",
                    "line": "#CBBDA7",
                    "contrast": "#1B1714",
                    "highlight": "#E7C9B2",
                },
                font_overrides={
                    "title_font": "STZhongsong",
                    "body_font": "Microsoft YaHei",
                    "accent_font": "Arial",
                    "numeric_font": "Arial",
                },
                composition_rules=("现代横排编辑结构", "局部竖排点缀", "纹样从内容中生成", "章节式节奏"),
                image_policy="实物、工艺、人物与场域；尊重文化语境和版权。",
                visual_devices=("paper_grain", "cultural_motif", "editorial_rhythm"),
                forbidden=("默认国风模板", "无来源文物图", "伪装手写历史文献"),
            ),
            StyleBranch(
                key="product_roadshow",
                name="产品路演",
                summary="高对比产品主视觉、清晰价值层级、演示节奏与跨页记忆色。",
                palette_overrides={
                    "background": "#F7F5F2",
                    "surface": "#FFFFFF",
                    "text": "#172033",
                    "secondary": "#596476",
                    "accent": "#E84855",
                    "line": "#D8D3CB",
                    "contrast": "#0D1420",
                    "highlight": "#F6CCD1",
                },
                font_overrides={
                    "title_font": "Microsoft YaHei",
                    "body_font": "Microsoft YaHei",
                    "accent_font": "Arial",
                    "numeric_font": "Arial",
                },
                composition_rules=("产品大图", "价值主张单点强化", "跨页色块呼应", "问题到方案的清晰切换"),
                image_policy="产品实拍、界面、使用场景和关键细节；保持透视与比例正确。",
                visual_devices=("product_crop", "memory_color", "demo_sequence"),
                forbidden=("虚构功能界面", "通用 Startup 模板", "无序功能卡片墙"),
            ),
        ),
    ),
}


DESIGN_GENES: dict[str, DesignGene] = {
    "L1": DesignGene(
        design_principle="纯信息转写，保留原文与阅读顺序，零装饰。",
        narrative_rhythm="平稳线性；标题、正文、图片依次出现，不人为制造高潮。",
        information_density_curve=(1.0, 1.0, 0.95, 1.0),
        composition_signature="文档式文本流与简单双栏，所有页面共享克制阅读骨架。",
        visual_center_rule="不强制视觉中心；以标题和首段提供阅读入口。",
        hierarchy_strategy="仅标题与正文两级，靠字号和字重区分。",
        object_character="无卡片、无蒙版、无装饰图形；图片和图表保持原始形态。",
        spacing_behavior="固定页边距和自然行距，不追求夸张留白。",
        signature_moves=("连续文段", "简单分栏", "原文完整呈现"),
        anti_patterns=("为了美观删减原文", "卡片墙", "图标堆叠", "艺术字"),
    ),
    "L2": DesignGene(
        design_principle="结构化商务表达，用稳定网格建立正式感与可信度。",
        narrative_rhythm="章节化、对称推进；每节以编号和页眉建立秩序。",
        information_density_curve=(0.78, 0.82, 0.74, 0.8),
        composition_signature="12 列对称网格、固定页眉页脚、直角色块分区。",
        visual_center_rule="每页一个主模块或一组对称比较，中心由结构而非装饰建立。",
        hierarchy_strategy="标题、模块标题、正文三级，统一编号与对齐。",
        object_character="直角基础卡片、标准图文框、标准图表和表格。",
        spacing_behavior="均匀栏距、固定页边距、重复结构形成商务秩序。",
        signature_moves=("章节编号", "对称分栏", "页眉页脚", "标准化卡片"),
        anti_patterns=("非对称艺术构图", "全屏蒙版", "数据看板", "自由拼贴"),
    ),
    "L3": DesignGene(
        design_principle="结论先行，用单一视觉中心组织证据与行动。",
        narrative_rhythm="结论、证据、含义、行动形成短节奏；相邻页变换构图。",
        information_density_curve=(0.58, 0.52, 0.66, 0.48),
        composition_signature="12 列专业网格中建立清晰主次，局部非对称但不失商务秩序。",
        visual_center_rule="每页必须只有一个主视觉中心，由主结论、主数字或主图表承担。",
        hierarchy_strategy="主标题说明结论，模块标题组织证据，正文只保留限定条件。",
        object_character="克制圆角、单一柔和阴影、数据组件与逻辑图示。",
        spacing_behavior="依据信息任务切换疏密，关键对象周围获得更多留白。",
        signature_moves=("结论先行", "视觉中心", "证据可视化", "版式随任务变化"),
        anti_patterns=("放大版 L2", "卡片墙", "所有页面标题加列表", "无结论数据页"),
    ),
    "L4": DesignGene(
        design_principle="以主题叙事建立沉浸式视觉系统，让每页成为同一故事的不同场景。",
        narrative_rhythm="建立情境、制造转折、展开证据、完成收束；密度与明暗有起伏。",
        information_density_curve=(0.28, 0.44, 0.36, 0.5),
        composition_signature="编辑式自由构图、非对称图文关系、沉浸大图与跨页视觉呼应。",
        visual_center_rule="每页由主题视觉锤与一个核心叙事对象共同建立焦点。",
        hierarchy_strategy="标题、导语、章节标题、正文、注释形成主题化层级。",
        object_character="主题化异形、纹理、渐变、蒙版与定制信息图。",
        spacing_behavior="大留白与密集信息场景交替，形成电影式节奏。",
        signature_moves=("跨页视觉锤", "编辑式构图", "主题材质", "叙事节点"),
        anti_patterns=("只换深色背景", "大字撑场面", "每页不同风格", "伪造事实素材"),
    ),
}


ANIMATION_RULES: dict[str, AnimationRule] = {
    "L1": AnimationRule(
        transition_types=(),
        entrance_types=(),
        emphasis_types=(),
        allow_exit=False,
        allow_motion_path=False,
        allow_trigger=False,
        allow_media=False,
        max_element_animations=0,
        duration_range_ms=(0, 0),
        stagger_ms=0,
        default_transition="",
        motion_intensity="none",
    ),
    "L2": AnimationRule(
        transition_types=("fade", "push_left", "push_up"),
        entrance_types=("fade",),
        emphasis_types=(),
        allow_exit=False,
        allow_motion_path=False,
        allow_trigger=False,
        allow_media=True,
        max_element_animations=1,
        duration_range_ms=(300, 700),
        stagger_ms=0,
        default_transition="fade",
        motion_intensity="minimal",
    ),
    "L3": AnimationRule(
        transition_types=(
            "fade",
            "push_left",
            "push_up",
            "wipe_right",
            "morph",
        ),
        entrance_types=("fade", "fly_up", "wipe_up", "zoom"),
        emphasis_types=("pulse", "grow", "color_change"),
        allow_exit=False,
        allow_motion_path=False,
        allow_trigger=False,
        allow_media=True,
        max_element_animations=3,
        duration_range_ms=(350, 900),
        stagger_ms=120,
        default_transition="fade",
        motion_intensity="structured",
    ),
    "L4": AnimationRule(
        transition_types=(
            "fade",
            "push_left",
            "push_up",
            "wipe_right",
            "morph",
            "ripple",
            "prism",
            "vortex",
            "gallery",
            "reveal",
            "zoom_in",
            "zoom_out",
        ),
        entrance_types=(
            "fade",
            "fly_up",
            "fly_down",
            "fly_left",
            "fly_right",
            "float_up",
            "wipe_up",
            "zoom",
            "circle",
            "box",
        ),
        emphasis_types=("pulse", "grow", "shrink", "spin", "color_change"),
        allow_exit=True,
        allow_motion_path=True,
        allow_trigger=True,
        allow_media=True,
        max_element_animations=None,
        duration_range_ms=(500, 1500),
        stagger_ms=180,
        default_transition="morph",
        motion_intensity="cinematic",
    ),
}


PALETTE_SYSTEMS: dict[str, PaletteSystem] = {
    "L1": PaletteSystem(
        primary=("#000000", "#FFFFFF"),
        secondary=(),
        neutrals=("#000000", "#FFFFFF"),
        accents=(),
        functional={"text": "#111111", "background": "#FFFFFF"},
        max_colors=2,
        usage_ratio={"neutral": 1.0, "accent": 0.0},
        contrast_policy="WCAG AA for body text; no decorative colour.",
    ),
    "L2": PaletteSystem(
        primary=("#10263B", "#1F5FA8"),
        secondary=("#4A647A", "#B7CADB"),
        neutrals=("#FFFFFF",),
        accents=(),
        functional={"surface": "#EEF3F8", "line": "#B7CADB"},
        max_colors=5,
        usage_ratio={"neutral": 0.72, "primary": 0.22, "accent": 0.06},
        contrast_policy="WCAG AA; formality comes from restrained contrast.",
    ),
    "L3": PaletteSystem(
        primary=("#122033", "#0B7A75"),
        secondary=("#5D6B7A", "#C8D2DC"),
        neutrals=("#F5F7FA", "#FFFFFF", "#08131F"),
        accents=("#DCEFEB",),
        functional={"line": "#C8D2DC", "highlight": "#DCEFEB"},
        max_colors=8,
        usage_ratio={"neutral": 0.58, "primary": 0.27, "accent": 0.15},
        contrast_policy="Single accent system; body contrast at least 4.5:1.",
    ),
    "L4": PaletteSystem(
        primary=("#071016", "#22D3EE"),
        secondary=("#11232B", "#45636B"),
        neutrals=("#F5F8F7", "#B7C4C2", "#000000"),
        accents=("#A7F3FA", "#F4EFE6", "#B48A4A", "#A44A32", "#E84855"),
        functional={
            "surface": "#10252B",
            "line": "#28515C",
            "highlight": "#A7F3FA",
        },
        max_colors=12,
        usage_ratio={"neutral": 0.35, "primary": 0.35, "accent": 0.30},
        contrast_policy="Theme-controlled contrast; verify text over imagery.",
    ),
}


for _profile_level, _profile_value in tuple(LEVEL_PROFILES.items()):
    LEVEL_PROFILES[_profile_level] = replace(
        _profile_value,
        design_gene=DESIGN_GENES[_profile_level],
        animation=ANIMATION_RULES[_profile_level],
        palette_system=PALETTE_SYSTEMS[_profile_level],
    )


def normalize_level(level: str | int) -> str:
    key = str(level).strip()
    normalized = LEVEL_ALIASES.get(key.upper(), LEVEL_ALIASES.get(key))
    if normalized is None:
        raise ValueError(
            f"Unsupported expression level {level!r}; expected one of {list(LEVEL_PROFILES)}"
        )
    return normalized


def get_profile(
    level: str | int,
    style_branch: str | None = None,
) -> LevelProfile:
    profile = LEVEL_PROFILES[normalize_level(level)]
    if style_branch is None:
        return profile
    branch = profile.style_branch(style_branch)
    if branch is None:
        raise ValueError(f"Style branches are only available for L4, got {profile.level}")
    font = replace(profile.font, **branch.font_overrides)
    palette = {**profile.palette, **branch.palette_overrides}
    return replace(profile, font=font, palette=palette)


def resolve_style_branch(
    level: str | int,
    requested_style: str | None = None,
) -> StyleBranch | None:
    profile = get_profile(level)
    return profile.style_branch(requested_style)


def resolve_layout(
    level: str | int,
    requested_layout: str = "",
    *,
    content_type: str = "bullets",
    has_image: bool = False,
    has_data: bool = False,
) -> str:
    """Resolve a requested/content type to an allowed strategy layout."""

    profile = get_profile(level)
    allowed = profile.layout.allowed_layouts
    if requested_layout and requested_layout in allowed:
        return requested_layout

    mapped = profile.layout.content_layout_map.get(content_type)
    if mapped in allowed:
        return mapped

    if has_data:
        for candidate in ("data_dashboard", "structured_table", "title_body"):
            if candidate in allowed:
                return candidate
    if has_image:
        for candidate in ("case_visual", "framed_image_text", "simple_split"):
            if candidate in allowed:
                return candidate
    if profile.layout.content_layout in allowed:
        return profile.layout.content_layout
    return allowed[0]


def match_content_layout(
    level: str | int,
    content_type: str,
    **flags: bool,
) -> str:
    """Public content-type matcher used by the builder."""

    return resolve_layout(level, content_type=content_type, **flags)


def font_stack(
    level: str | int,
    style_branch: str | None = None,
) -> dict[str, str]:
    profile = get_profile(level, style_branch)
    return {
        "title": profile.font.title_font,
        "subtitle": profile.font.body_font,
        "module": profile.font.body_font,
        "section": profile.font.body_font,
        "body": profile.font.body_font,
        "accent": profile.font.accent_font,
        "numeric": profile.font.numeric_font,
        "caption": profile.font.body_font,
    }


def validate_profile_matrix() -> dict:
    """Validate strategy uniqueness, boundaries, and required extensions."""

    errors: list[str] = []
    fingerprints = {
        level: profile.fingerprint() for level, profile in LEVEL_PROFILES.items()
    }
    if len(fingerprints) != 4 or len(set(fingerprints.values())) != 4:
        errors.append("L1-L4 profiles must contain four unique fingerprints.")

    strategy_keys = {profile.strategy.key for profile in LEVEL_PROFILES.values()}
    if len(strategy_keys) != 4:
        errors.append("Each level must define a distinct strategy key.")

    for level, profile in LEVEL_PROFILES.items():
        total = sum(item.max_points for item in profile.strategy.score_dimensions)
        if total != 100:
            errors.append(f"{level} score dimensions must total 100, got {total}.")
        for layout, mapped in profile.layout.content_layout_map.items():
            if mapped not in profile.layout.allowed_layouts:
                errors.append(
                    f"{level} maps {layout!r} to disallowed layout {mapped!r}."
                )
        if len({item.key for item in profile.font.levels}) != profile.font.hierarchy_levels:
            errors.append(f"{level} typography hierarchy does not match hierarchy_levels.")
        if profile.font.max_font_families < len(
            {
                profile.font.title_font,
                profile.font.body_font,
                profile.font.accent_font,
                profile.font.numeric_font,
            }
        ):
            errors.append(f"{level} font stack exceeds max_font_families.")

    l1 = LEVEL_PROFILES["L1"]
    if l1.font.max_font_families != 1 or l1.font.hierarchy_levels != 2:
        errors.append("L1 must use one font family and two hierarchy levels.")
    if set(l1.layout.allowed_layouts) != {"title_body", "simple_split"}:
        errors.append("L1 layout whitelist must contain only title_body and simple_split.")
    if l1.palette["background"] != "#FFFFFF":
        errors.append("L1 must keep a pure white background.")
    if l1.layout.allow_cards or l1.components.max_cards_per_slide != 0:
        errors.append("L1 must not allow cards.")
    if l1.information.max_body_chars is not None:
        errors.append("L1 must not impose a body-character truncation budget.")

    l2 = LEVEL_PROFILES["L2"]
    if l2.font.max_font_families > 2 or l2.font.hierarchy_levels != 3:
        errors.append("L2 must stay within two typefaces and three primary levels.")
    if not l2.layout.allow_cards or not l2.layout.fixed_header or not l2.layout.fixed_footer:
        errors.append("L2 must provide basic cards and fixed header/footer structure.")
    if l2.layout.allow_image_mask or l2.layout.allow_asymmetric:
        errors.append("L2 must not use image masks or asymmetric layouts.")

    l3 = LEVEL_PROFILES["L3"]
    required_l3 = {
        "metric_cards",
        "data_dashboard",
        "process_steps",
        "timeline",
        "architecture_compare",
        "case_visual",
        "section_transition",
        "icon_text_modules",
    }
    if not required_l3.issubset(l3.layout.allowed_layouts):
        errors.append("L3 is missing one or more required professional layouts.")
    if l3.grid.columns != 12:
        errors.append("L3 must use a 12-column grid.")
    if not (0.08 <= l3.grid.component_radius_in <= 0.15):
        errors.append("L3 component radius must stay within 0.08-0.15 inches.")
    if l3.grid.shadow_policy != "single_soft_ambient":
        errors.append("L3 must use one restrained ambient shadow style.")
    if l3.font.max_font_families > 3 or l3.font.hierarchy_levels != 4:
        errors.append("L3 must use at most three typefaces and four hierarchy levels.")
    if l3.font.numeric_font != "Consolas":
        errors.append("L3 requires a monospace numeric typeface.")
    if l3.layout.allow_collage:
        errors.append("L3 must not use L4 collage layouts.")

    l4 = LEVEL_PROFILES["L4"]
    expected_branches = {
        "technology_launch",
        "luxury_brand",
        "cultural_creative",
        "product_roadshow",
    }
    branch_keys = {branch.key for branch in l4.style_branches}
    if branch_keys != expected_branches:
        errors.append("L4 must define four required style branches.")
    if not l4.images.allow_collage or not l4.font.allow_vertical_text:
        errors.append("L4 must enable narrative collage and vertical editorial text.")

    order = ["L1", "L2", "L3", "L4"]
    coefficients = [
        LEVEL_PROFILES[level].information.density_coefficient for level in order
    ]
    if any(lower <= upper for lower, upper in zip(coefficients, coefficients[1:])):
        errors.append("Information density coefficients must decrease from L1 to L4.")

    design_principles = {
        profile.design_gene.design_principle for profile in LEVEL_PROFILES.values()
    }
    if len(design_principles) != 4:
        errors.append("Every level must define a distinct design gene.")

    expected_palette_sizes = {"L1": 2, "L2": 5, "L3": 8, "L4": 12}
    for level, expected_size in expected_palette_sizes.items():
        profile = LEVEL_PROFILES[level]
        if profile.palette_system.max_colors != expected_size:
            errors.append(
                f"{level} palette system must declare exactly {expected_size} colours."
            )
        if profile.palette_count() > expected_size:
            errors.append(
                f"{level} palette system contains {profile.palette_count()} "
                f"unique colours; cap is {expected_size}."
            )

    l1_motion = LEVEL_PROFILES["L1"].animation
    l2_motion = LEVEL_PROFILES["L2"].animation
    l3_motion = LEVEL_PROFILES["L3"].animation
    l4_motion = LEVEL_PROFILES["L4"].animation
    if any(
        (
            l1_motion.transition_types,
            l1_motion.entrance_types,
            l1_motion.emphasis_types,
            l1_motion.max_element_animations,
        )
    ):
        errors.append("L1 must remain completely static.")
    if not l2_motion.transition_types or l2_motion.max_element_animations != 1:
        errors.append("L2 must allow basic transitions and one page-level entrance.")
    if l3_motion.allow_exit or l3_motion.max_element_animations != 3:
        errors.append("L3 motion contract must permit up to three element entrances.")
    if not l4_motion.allow_trigger or not l4_motion.allow_motion_path:
        errors.append("L4 motion contract must permit triggers and motion paths.")
    if not any(char in l3.design_gene.visual_center_rule for char in ("只有一个", "单一")):
        errors.append("L3 design gene must require a single visual centre.")
    if not LEVEL_PROFILES["L4"].design_gene.signature_moves:
        errors.append("L4 design gene must define cross-page signature moves.")

    return {
        "valid": not errors,
        "version": SKILL_VERSION,
        "schema_version": STRATEGY_SCHEMA_VERSION,
        "errors": errors,
        "fingerprints": fingerprints,
        "design_genes": {
            level: profile.design_gene.design_principle
            for level, profile in LEVEL_PROFILES.items()
        },
        "palette_matrix": {
            level: {
                "max_colors": profile.palette_system.max_colors,
                "declared_colors": profile.palette_count(),
                "usage_ratio": profile.palette_system.usage_ratio,
            }
            for level, profile in LEVEL_PROFILES.items()
        },
        "animation_matrix": {
            level: {
                "intensity": profile.animation.motion_intensity,
                "max_element_animations": profile.animation.max_element_animations,
                "transitions": list(profile.animation.transition_types),
            }
            for level, profile in LEVEL_PROFILES.items()
        },
        "strategies": {
            level: {
                "key": profile.strategy.key,
                "name": profile.strategy.name,
                "density_coefficient": profile.information.density_coefficient,
                "acceptance_range": list(profile.acceptance_range),
                "style_branches": [branch.key for branch in profile.style_branches],
            }
            for level, profile in LEVEL_PROFILES.items()
        },
        "levels": order,
    }


def profiles_as_dict(levels: Iterable[str] | None = None) -> dict[str, dict]:
    selected = levels or LEVEL_PROFILES.keys()
    return {
        normalize_level(level): get_profile(level).as_dict() for level in selected
    }


if __name__ == "__main__":
    result = validate_profile_matrix()
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["valid"] else 1)
