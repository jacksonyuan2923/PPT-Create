#!/usr/bin/env python3
"""Design tokens for pro-ppt-studio V4.

Level differentiation used to be carried by a single hard-coded palette and a
single font family per role, so every deck looked the same. V4 splits the
tokens into two orthogonal axes:

* ``LEVEL_FONT_STACKS`` - the *typographic* contract of L1-L4 (how many
  families, which families, which effects are allowed).
* ``PALETTES`` - named colour systems that a level is allowed to use.

A level therefore constrains *which* tokens are legal, while the visual
direction (科技未来 / 商业路演 / 杂志编辑 / 国风 ...) selects *which* legal
token set the deck actually uses. That is what makes L3 and L4 decks look
materially different from each other and from L1/L2.

Only fonts that ship with Windows / Microsoft Office CJK are referenced so the
declared stack can actually be rendered on a normal Chinese desktop.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Iterable


# ---------------------------------------------------------------------------
# colour utilities
# ---------------------------------------------------------------------------


def hex_to_rgb(value: str) -> tuple[int, int, int]:
    """Convert ``#RRGGBB`` (or ``RRGGBB``) to an RGB triplet."""

    value = value.strip().lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    if len(value) != 6:
        raise ValueError(f"Expected a 6-digit hex colour, got {value!r}")
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)


def rgb_to_hex(rgb: Iterable[int]) -> str:
    r, g, b = list(rgb)[:3]
    return "#{:02X}{:02X}{:02X}".format(int(r), int(g), int(b))


def _channel_luminance(channel: float) -> float:
    if channel <= 0.03928:
        return channel / 12.92
    return ((channel + 0.055) / 1.055) ** 2.4


def relative_luminance(value: str) -> float:
    r, g, b = hex_to_rgb(value)
    luminance = 0.2126 * _channel_luminance(r / 255)
    luminance += 0.7152 * _channel_luminance(g / 255)
    luminance += 0.3290 * _channel_luminance(b / 255)
    return luminance


def contrast_ratio(fg: str, bg: str) -> float:
    """WCAG contrast ratio between two hex colours (1.0 - 21.0)."""

    light = relative_luminance(fg)
    dark = relative_luminance(bg)
    if light < dark:
        light, dark = dark, light
    return round((light + 0.05) / (dark + 0.05), 2)


def mix(colour_a: str, colour_b: str, ratio: float) -> str:
    """Blend two hex colours; ``ratio`` 0.0 keeps A, 1.0 keeps B."""

    ratio = max(0.0, min(1.0, ratio))
    a = hex_to_rgb(colour_a)
    b = hex_to_rgb(colour_b)
    return rgb_to_hex([a[i] + (b[i] - a[i]) * ratio for i in range(3)])


def tint(colour: str, amount: float = 0.15) -> str:
    """Mix a colour toward white."""

    return mix(colour, "#FFFFFF", amount)


def shade(colour: str, amount: float = 0.15) -> str:
    """Mix a colour toward black."""

    return mix(colour, "#000000", amount)


# ---------------------------------------------------------------------------
# palettes
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Palette:
    """A complete colour system for one deck."""

    name: str
    background: str
    surface: str
    surface_alt: str
    text: str
    text_secondary: str
    accent: str
    accent_secondary: str
    on_accent: str
    line: str
    dark: bool = False
    mood: str = ""
    tags: tuple[str, ...] = field(default_factory=tuple)

    def as_dict(self) -> dict:
        data = asdict(self)
        data["tags"] = list(self.tags)
        return data

    def slot(self, name: str, fallback: str | None = None) -> str:
        """Return a colour slot, tolerating unknown names."""

        value = getattr(self, name, None)
        if isinstance(value, str):
            return value
        if fallback is not None:
            return fallback
        raise KeyError(f"Palette {self.name!r} has no slot {name!r}")


def _p(**kwargs) -> Palette:
    return Palette(**kwargs)


PALETTES: dict[str, Palette] = {
    # --- L1: one accent, maximum restraint -------------------------------
    "mono_warm": _p(
        name="mono_warm",
        background="#FFFFFF",
        surface="#F5F5F4",
        surface_alt="#EBEBE9",
        text="#1C1C1C",
        text_secondary="#5C5C5C",
        accent="#2F5D8C",
        accent_secondary="#8C97A3",
        on_accent="#FFFFFF",
        line="#D8D8D6",
        mood="朴素克制",
        tags=("l1", "neutral", "report"),
    ),
    "mono_ink": _p(
        name="mono_ink",
        background="#FFFFFF",
        surface="#F2F2F0",
        surface_alt="#E4E4E1",
        text="#111111",
        text_secondary="#555555",
        accent="#31414F",
        accent_secondary="#8A949E",
        on_accent="#FFFFFF",
        line="#D2D2D0",
        mood="灰阶中性",
        tags=("l1", "neutral"),
    ),
    # --- L2: formal business ---------------------------------------------
    "formal_blue": _p(
        name="formal_blue",
        background="#FFFFFF",
        surface="#EDF2F8",
        surface_alt="#DCE6F2",
        text="#12263A",
        text_secondary="#4A647A",
        accent="#1F5FA8",
        accent_secondary="#5B87C0",
        on_accent="#FFFFFF",
        line="#B7CADB",
        mood="政务蓝白",
        tags=("l2", "formal", "business"),
    ),
    "formal_indigo": _p(
        name="formal_indigo",
        background="#FFFFFF",
        surface="#EEF1F8",
        surface_alt="#DCE1F0",
        text="#16213E",
        text_secondary="#4B5576",
        accent="#2A3F87",
        accent_secondary="#6D7CB8",
        on_accent="#FFFFFF",
        line="#C3C9DC",
        mood="稳重靛蓝",
        tags=("l2", "formal"),
    ),
    "formal_teal": _p(
        name="formal_teal",
        background="#FFFFFF",
        surface="#EAF3F3",
        surface_alt="#D7E7E7",
        text="#12302F",
        text_secondary="#4A6A69",
        accent="#1B6B68",
        accent_secondary="#5E9C99",
        on_accent="#FFFFFF",
        line="#BBD3D2",
        mood="国企青绿",
        tags=("l2", "formal"),
    ),
    # --- L3: modern, depth, data-forward ----------------------------------
    "modern_deep": _p(
        name="modern_deep",
        background="#0B1220",
        surface="#131C2E",
        surface_alt="#1B2740",
        text="#F2F6FC",
        text_secondary="#9FB0CA",
        accent="#3D8BFD",
        accent_secondary="#22D3EE",
        on_accent="#061020",
        line="#2A3A57",
        dark=True,
        mood="现代深空",
        tags=("l3", "dark", "tech", "modern"),
    ),
    "modern_ivory": _p(
        name="modern_ivory",
        background="#F6F5F1",
        surface="#FFFFFF",
        surface_alt="#EAE7DF",
        text="#16181D",
        text_secondary="#5D616B",
        accent="#1D4ED8",
        accent_secondary="#C2410C",
        on_accent="#FFFFFF",
        line="#DDD9D0",
        mood="高级灰白",
        tags=("l3", "light", "modern", "editorial"),
    ),
    "modern_forest": _p(
        name="modern_forest",
        background="#0E1A16",
        surface="#15251F",
        surface_alt="#1D3129",
        text="#EFF6F1",
        text_secondary="#9DB6AA",
        accent="#4ADE80",
        accent_secondary="#F59E0B",
        on_accent="#06120D",
        line="#2C453B",
        dark=True,
        mood="墨绿自然",
        tags=("l3", "dark", "nature"),
    ),
    "modern_slate": _p(
        name="modern_slate",
        background="#F4F6F8",
        surface="#FFFFFF",
        surface_alt="#E4E9EF",
        text="#10161F",
        text_secondary="#54606E",
        accent="#0F766E",
        accent_secondary="#B45309",
        on_accent="#FFFFFF",
        line="#D3DBE4",
        mood="清冷灰蓝",
        tags=("l3", "light", "modern"),
    ),
    # --- L4: cultural, crafted, editorial ---------------------------------
    "guofeng_daiqing": _p(
        name="guofeng_daiqing",
        background="#F5F0E6",
        surface="#EFE7D7",
        surface_alt="#E4D8C2",
        text="#2B2A26",
        text_secondary="#6B6558",
        accent="#2E4E5E",
        accent_secondary="#C3272B",
        on_accent="#F5F0E6",
        line="#D6C7AC",
        mood="黛青朱砂",
        tags=("l4", "guofeng", "cultural"),
    ),
    "luxury_gold": _p(
        name="luxury_gold",
        background="#0F0F10",
        surface="#1A1A1C",
        surface_alt="#26262A",
        text="#F5F0E6",
        text_secondary="#B6AC96",
        accent="#C8A96A",
        accent_secondary="#8C7A55",
        on_accent="#14140F",
        line="#3A3A3E",
        dark=True,
        mood="夜墨鎏金",
        tags=("l4", "dark", "luxury", "premium"),
    ),
    "editorial_cream": _p(
        name="editorial_cream",
        background="#FBF7F0",
        surface="#FFFFFF",
        surface_alt="#F1E9DC",
        text="#1B1A17",
        text_secondary="#6A6558",
        accent="#9A3412",
        accent_secondary="#1F3A5F",
        on_accent="#FFFFFF",
        line="#E3D9C8",
        mood="杂志暖白",
        tags=("l4", "light", "editorial", "cultural"),
    ),
    "celadon_blue": _p(
        name="celadon_blue",
        background="#EEF3F2",
        surface="#F8FBFA",
        surface_alt="#DCE8E6",
        text="#1E2E2C",
        text_secondary="#5C706D",
        accent="#37666A",
        accent_secondary="#A6633C",
        on_accent="#FFFFFF",
        line="#C9DBD8",
        mood="青瓷雅白",
        tags=("l4", "light", "cultural", "minimal"),
    ),
    "ink_purple": _p(
        name="ink_purple",
        background="#14101C",
        surface="#1E1830",
        surface_alt="#2A2142",
        text="#F3EFFA",
        text_secondary="#A99FC4",
        accent="#A78BFA",
        accent_secondary="#F0ABFC",
        on_accent="#140F1E",
        line="#3A2F58",
        dark=True,
        mood="墨紫幻境",
        tags=("l4", "dark", "dream", "tech"),
    ),
}


#: Which palettes each visual level is allowed to use.
LEVEL_PALETTES: dict[str, tuple[str, ...]] = {
    "L1": ("mono_warm", "mono_ink"),
    "L2": ("formal_blue", "formal_indigo", "formal_teal"),
    "L3": ("modern_deep", "modern_ivory", "modern_forest", "modern_slate"),
    "L4": (
        "guofeng_daiqing",
        "luxury_gold",
        "editorial_cream",
        "celadon_blue",
        "ink_purple",
    ),
}


def palettes_for_level(level: str) -> list[Palette]:
    return [PALETTES[name] for name in LEVEL_PALETTES.get(level.upper(), ())]


def resolve_palette(level: str, name: str | None = None) -> Palette:
    """Resolve a palette for a level.

    ``name`` may be a palette key, a mood keyword, or a direction keyword. When
    it cannot be resolved, the level's first legal palette is returned so the
    deck always has a legal colour system.
    """

    level = level.upper()
    legal = LEVEL_PALETTES.get(level, ())
    if name:
        key = name.strip()
        if key in PALETTES and key in legal:
            return PALETTES[key]
        lowered = key.lower()
        for palette_name in legal:
            palette = PALETTES[palette_name]
            if lowered == palette_name.lower():
                return palette
            if lowered == palette.mood.lower():
                return palette
            if lowered in {tag.lower() for tag in palette.tags}:
                return palette
    return PALETTES[legal[0]]


# ---------------------------------------------------------------------------
# typography
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class FontSpec:
    """A resolved typeface assignment for one semantic role."""

    family: str
    fallback: tuple[str, ...] = ()
    latin: str | None = None
    weight: str = "regular"  # regular | bold | light
    tracking: float = 0.0     # letter spacing in points
    note: str = ""

    def as_dict(self) -> dict:
        data = asdict(self)
        data["fallback"] = list(self.fallback)
        return data


#: The typographic contract of each level.
#:
#: The *number of distinct families* is the single most visible signal of the
#: level, therefore it is validated at runtime by ``validate_profile_matrix``.
LEVEL_FONT_STACKS: dict[str, dict[str, FontSpec]] = {
    "L1": {
        "title": FontSpec("Microsoft YaHei", ("Microsoft YaHei UI",), note="单一无衬线字体"),
        "body": FontSpec("Microsoft YaHei", ("Microsoft YaHei UI",)),
        "accent": FontSpec("Microsoft YaHei", ("Microsoft YaHei UI",)),
        "display": FontSpec("Microsoft YaHei", ("Microsoft YaHei UI",)),
    },
    "L2": {
        "title": FontSpec("Microsoft YaHei", ("Microsoft YaHei UI",), weight="bold"),
        "body": FontSpec("Microsoft YaHei", ("Microsoft YaHei UI",)),
        "accent": FontSpec("Arial", ("Segoe UI",), note="拉丁字母与数字"),
        "display": FontSpec("Microsoft YaHei", ("Microsoft YaHei UI",), weight="bold"),
    },
    "L3": {
        "title": FontSpec(
            "Microsoft YaHei Light",
            ("Microsoft YaHei", "Noto Sans SC"),
            weight="light",
            tracking=1.5,
            note="细体大字 + 字距，塑造编辑感标题",
        ),
        "body": FontSpec("Microsoft YaHei", ("Microsoft YaHei UI",)),
        "accent": FontSpec("Bahnschrift", ("Arial",), tracking=0.8, note="几何压缩体数字"),
    },
    "L4": {
        "title": FontSpec(
            "STZhongsong",
            ("STSong", "SimSun"),
            weight="bold",
            tracking=2.0,
            note="华文中宋文化展示体",
        ),
        "body": FontSpec("STSong", ("SimSun", "Noto Serif SC")),
        "accent": FontSpec("Georgia", ("Times New Roman",), note="拉丁数字与外文"),
        "display": FontSpec("STXingkai", ("STXinwei",), note="书法题字，仅用于题字/引言"),
    },
}

#: Alternative title families offered per level (design direction may pick one).
LEVEL_TITLE_ALTERNATIVES: dict[str, dict[str, str]] = {
    "L1": {
        "sans": "Microsoft YaHei",
        "hei": "SimHei",
        "song": "SimSun",
    },
    "L2": {
        "sans": "Microsoft YaHei",
        "hei": "SimHei",
        "song": "SimSun",
        "kai": "KaiTi",
    },
    "L3": {
        "light": "Microsoft YaHei Light",
        "sans": "Microsoft YaHei",
        "serif": "STZhongsong",
        "modern": "Noto Sans SC",
    },
    "L4": {
        "song": "STZhongsong",
        "wei": "STXinwei",
        "xing": "STXingkai",
        "li": "STLiti",
        "fangsong": "STFangsong",
        "hei": "FZCHeiSong",
    },
}


def font_stack(level: str) -> dict[str, FontSpec]:
    """Return the font specs for a level, keyed by semantic role."""

    key = level.upper()
    if key not in LEVEL_FONT_STACKS:
        raise ValueError(f"Unknown visual level {level!r}")
    return dict(LEVEL_FONT_STACKS[key])


def resolve_font(level: str, role: str, family: str | None = None) -> FontSpec:
    """Resolve one role, honouring an explicit family override."""

    stack = font_stack(level)
    base = stack.get(role) or stack["body"]
    if not family:
        return base
    return FontSpec(
        family=family,
        fallback=base.fallback,
        latin=base.latin,
        weight=base.weight,
        tracking=base.tracking,
        note="override",
    )


def palette_contrast_report(palette: Palette) -> dict:
    """Contrast audit used by the quality gate."""

    return {
        "palette": palette.name,
        "text_on_background": contrast_ratio(palette.text, palette.background),
        "text_secondary_on_background": contrast_ratio(palette.text_secondary, palette.background),
        "accent_on_background": contrast_ratio(palette.accent, palette.background),
        "on_accent_on_accent": contrast_ratio(palette.on_accent, palette.accent),
        "text_on_surface": contrast_ratio(palette.text, palette.surface),
    }


if __name__ == "__main__":
    import json

    report = {}
    for level in ("L1", "L2", "L3", "L4"):
        stack = font_stack(level)
        families = sorted({spec.family for spec in stack.values()})
        report[level] = {
            "families": families,
            "family_count": len(families),
            "palettes": [p.name for p in palettes_for_level(level)],
            "contrast": {
                palette.name: palette_contrast_report(palette)
                for palette in palettes_for_level(level)
            },
        }
    print(json.dumps(report, ensure_ascii=False, indent=2))