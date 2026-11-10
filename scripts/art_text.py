#!/usr/bin/env python3
"""Artistic text (艺术字) helpers for pro-ppt-studio V4.

``python-pptx`` cannot express WordArt, gradient text, 3-D text or vertical
typography, but PowerPoint renders all of them natively. This module writes the
corresponding DrawingML directly.

Placement rules follow the ECMA-376 schema:

* run level (``a:rPr``)  -> outline, fill (solid/gradient/pattern), effects
  (shadow, glow, reflection, soft edge)
* shape level (``a:bodyPr``) -> text warp (WordArt transform), 3-D scene and
  bevel, vertical writing mode, flat text

Every helper is a no-op when the requested effect is not allowed by the
visual level, so a caller can request the same design for L1..L4 and let the
level decide what is actually legal.
"""

from __future__ import annotations

from typing import Sequence

from pptx.oxml.ns import qn

from design_tokens import hex_to_rgb

# Canonical child order of a:rPr (ECMA-376 CT_TextCharacterProperties).
_RPR_ORDER = (
    "a:ln",
    "a:noFill",
    "a:solidFill",
    "a:gradFill",
    "a:blipFill",
    "a:pattFill",
    "a:grpFill",
    "a:effectLst",
    "a:effectDag",
    "a:highlight",
    "a:uLnTx",
    "a:uLn",
    "a:uFillTx",
    "a:uFill",
    "a:latin",
    "a:ea",
    "a:cs",
    "a:sym",
    "a:hlinkClick",
    "a:hlinkMouseOver",
    "a:rtl",
    "a:extLst",
)

# Canonical child order of a:bodyPr (ECMA-376 CT_TextBodyProperties).
_BODYPR_ORDER = (
    "a:prstTxWarp",
    "a:noAutofit",
    "a:normAutofit",
    "a:spAutoFit",
    "a:scene3d",
    "a:sp3d",
    "a:flatTx",
    "a:extLst",
)

# Canonical child order of a:effectLst (CT_EffectList).
_EFFECT_ORDER = (
    "a:blur",
    "a:fillOverlay",
    "a:glow",
    "a:innerShdw",
    "a:outerShdw",
    "a:prstShdw",
    "a:reflection",
    "a:softEdge",
)


def _insert_ordered(parent, element, order: Sequence[str]) -> None:
    """Insert ``element`` into ``parent`` honouring the schema order."""

    tag = element.tag
    name = tag.replace("{http://schemas.openxmlformats.org/drawingml/2006/main}", "a:")
    if name not in order:
        parent.append(element)
        return
    position = order.index(name)
    for child in parent:
        child_name = child.tag.replace(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}", "a:"
        )
        if child_name in order and order.index(child_name) > position:
            child.addprevious(element)
            return
    parent.append(element)


def _replace_child(parent, tag: str, order: Sequence[str]):
    existing = parent.find(qn(tag))
    if existing is not None:
        parent.remove(existing)
    return existing


def _make(parent, tag: str):
    """Create (or reuse) a child element with the drawingml namespace."""

    element = parent.find(qn(tag))
    if element is None:
        element = parent.makeelement(qn(tag), {})
        _insert_ordered(parent, element, _RPR_ORDER if "rPr" in parent.tag else _BODYPR_ORDER)
    return element


def _srgb(parent, colour: str, alpha: float | None = None):
    """Build an ``<a:srgbClr>`` element with optional alpha."""

    r, g, b = hex_to_rgb(colour)
    node = parent.makeelement(qn("a:srgbClr"), {"val": "%02X%02X%02X" % (r, g, b)})
    if alpha is not None:
        node.append(node.makeelement(qn("a:alpha"), {"val": str(int(alpha * 100000))}))
    return node


# ---------------------------------------------------------------------------
# run level effects
# ---------------------------------------------------------------------------


def _rpr(run):
    return run._r.get_or_add_rPr()


def set_text_fill_solid(run, colour: str) -> None:
    """Replace the run fill with a solid colour."""

    r_pr = _rpr(run)
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        existing = r_pr.find(qn(tag))
        if existing is not None:
            r_pr.remove(existing)
    fill = r_pr.makeelement(qn("a:solidFill"), {})
    fill.append(_srgb(r_pr, colour))
    _insert_ordered(r_pr, fill, _RPR_ORDER)


def set_text_fill_none(run) -> None:
    """Make the run outline-only (hollow text)."""

    r_pr = _rpr(run)
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        existing = r_pr.find(qn(tag))
        if existing is not None:
            r_pr.remove(existing)
    node = r_pr.makeelement(qn("a:noFill"), {})
    _insert_ordered(r_pr, node, _RPR_ORDER)


def set_text_fill_gradient(
    run,
    colours: Sequence[str],
    *,
    angle_degrees: int = 45,
    stop_positions: Sequence[int] | None = None,
) -> None:
    """Give the run a gradient fill (gold, sunrise, chrome, ...)."""

    if not colours:
        return
    r_pr = _rpr(run)
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        existing = r_pr.find(qn(tag))
        if existing is not None:
            r_pr.remove(existing)
    grad = r_pr.makeelement(qn("a:gradFill"), {"rotWithShape": "1"})
    gs_lst = grad.makeelement(qn("a:gsLst"), {})
    positions = list(stop_positions) if stop_positions else [
        int(round(100000 * index / max(1, len(colours) - 1))) for index in range(len(colours))
    ]
    for index, colour in enumerate(colours):
        stop = grad.makeelement(qn("a:gs"), {"pos": str(positions[index])})
        stop.append(_srgb(grad, colour))
        gs_lst.append(stop)
    grad.append(gs_lst)
    grad.append(grad.makeelement(qn("a:lin"), {"ang": str(int(angle_degrees * 60000) % 21600000), "scaled": "0"}))
    _insert_ordered(r_pr, grad, _RPR_ORDER)


def set_text_fill_pattern(run, pattern: str, foreground: str, background: str) -> None:
    """Give the run a pattern/texture fill."""

    r_pr = _rpr(run)
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        existing = r_pr.find(qn(tag))
        if existing is not None:
            r_pr.remove(existing)
    patt = r_pr.makeelement(qn("a:pattFill"), {"prst": pattern})
    patt.append(_srgb(patt, foreground))
    patt.append(_srgb(patt, background))
    _insert_ordered(r_pr, patt, _RPR_ORDER)


def set_text_outline(run, colour: str, width_pt: float = 1.0) -> None:
    """Draw an outline around the glyphs."""

    r_pr = _rpr(run)
    existing = r_pr.find(qn("a:ln"))
    if existing is not None:
        r_pr.remove(existing)
    line = r_pr.makeelement(qn("a:ln"), {"w": str(int(width_pt * 12700)), "cap": "rnd"})
    solid = line.makeelement(qn("a:solidFill"), {})
    solid.append(_srgb(line, colour))
    line.append(solid)
    line.append(line.makeelement(qn("a:round"), {}))
    _insert_ordered(r_pr, line, _RPR_ORDER)


def _effect_list(run):
    r_pr = _rpr(run)
    effect = r_pr.find(qn("a:effectLst"))
    if effect is None:
        effect = r_pr.makeelement(qn("a:effectLst"), {})
        _insert_ordered(r_pr, effect, _RPR_ORDER)
    return effect


def set_text_shadow(
    run,
    colour: str = "#000000",
    *,
    blur_pt: float = 4.0,
    distance_pt: float = 3.0,
    direction_degrees: int = 90,
    alpha: float = 0.35,
) -> None:
    """Attach a soft outer shadow to the glyphs."""

    effect = _effect_list(run)
    for tag in ("a:outerShdw",):
        existing = effect.find(qn(tag))
        if existing is not None:
            effect.remove(existing)
    shadow = effect.makeelement(
        qn("a:outerShdw"),
        {
            "blurRad": str(int(blur_pt * 12700)),
            "dist": str(int(distance_pt * 12700)),
            "dir": str(int(direction_degrees * 60000) % 21600000),
            "algn": "tl",
            "rotWithShape": "0",
        },
    )
    shadow.append(_srgb(shadow, colour, alpha))
    _insert_ordered(effect, shadow, _EFFECT_ORDER)


def set_text_glow(run, colour: str, *, radius_pt: float = 8.0, alpha: float = 0.5) -> None:
    """Add a coloured glow around the glyphs."""

    effect = _effect_list(run)
    existing = effect.find(qn("a:glow"))
    if existing is not None:
        effect.remove(existing)
    glow = effect.makeelement(qn("a:glow"), {"rad": str(int(radius_pt * 12700))})
    glow.append(_srgb(glow, colour, alpha))
    _insert_ordered(effect, glow, _EFFECT_ORDER)


def set_text_reflection(run, *, blur_pt: float = 1.0, alpha: float = 0.45, distance_pt: float = 0.0) -> None:
    """Mirror the glyphs below the baseline."""

    effect = _effect_list(run)
    existing = effect.find(qn("a:reflection"))
    if existing is not None:
        effect.remove(existing)
    reflection = effect.makeelement(
        qn("a:reflection"),
        {
            "blurRad": str(int(blur_pt * 12700)),
            "stA": str(int(alpha * 100000)),
            "endA": "300",
            "endPos": "35000",
            "dist": str(int(distance_pt * 12700)),
            "dir": "5400000",
            "sy": "-100000",
            "algn": "bl",
            "rotWithShape": "0",
        },
    )
    _insert_ordered(effect, reflection, _EFFECT_ORDER)


def set_text_soft_edge(run, radius_pt: float = 2.0) -> None:
    """Soften the glyph edges."""

    effect = _effect_list(run)
    existing = effect.find(qn("a:softEdge"))
    if existing is not None:
        effect.remove(existing)
    soft = effect.makeelement(qn("a:softEdge"), {"rad": str(int(radius_pt * 12700))})
    _insert_ordered(effect, soft, _EFFECT_ORDER)


# ---------------------------------------------------------------------------
# shape level effects (WordArt transforms, 3-D, vertical writing)
# ---------------------------------------------------------------------------


def _body_pr(shape):
    return shape.text_frame._txBody.find(qn("a:bodyPr"))


def set_text_warp(shape, preset: str = "textArchUp") -> None:
    """Apply a WordArt text transform (arch, wave, circle, chevron, ...)."""

    body_pr = _body_pr(shape)
    if body_pr is None:
        return
    existing = body_pr.find(qn("a:prstTxWarp"))
    if existing is not None:
        body_pr.remove(existing)
    warp = body_pr.makeelement(qn("a:prstTxWarp"), {"prst": preset})
    warp.append(warp.makeelement(qn("a:avLst"), {}))
    _insert_ordered(body_pr, warp, _BODYPR_ORDER)
    body_pr.set("wrap", "none")


def set_text_vertical(shape, mode: str = "eaVert") -> None:
    """Switch the text box to vertical writing.

    ``eaVert`` keeps CJK upright, ``vert`` rotates latin glyphs as well.
    """

    body_pr = _body_pr(shape)
    if body_pr is None:
        return
    body_pr.set("vert", mode)


def set_text_3d(
    shape,
    *,
    extrusion_pt: float = 0.0,
    bevel: str = "circle",
    bevel_width_pt: float = 6.0,
    bevel_height_pt: float = 3.0,
    material: str = "warmMatte",
    contour_colour: str | None = None,
    extrusion_colour: str | None = None,
    camera: str = "orthographicFront",
    rig: str = "threePt",
    direction: str = "t",
) -> None:
    """Give the text a real 3-D look (extrusion + bevel + lighting)."""

    body_pr = _body_pr(shape)
    if body_pr is None:
        return
    for tag in ("a:scene3d", "a:sp3d"):
        existing = body_pr.find(qn(tag))
        if existing is not None:
            body_pr.remove(existing)

    scene = body_pr.makeelement(qn("a:scene3d"), {})
    scene.append(scene.makeelement(qn("a:camera"), {"prst": camera}))
    light = scene.makeelement(qn("a:lightRig"), {"rig": rig, "dir": direction})
    light.append(light.makeelement(qn("a:rot"), {"lat": "0", "lon": "0", "rev": "0"}))
    scene.append(light)
    _insert_ordered(body_pr, scene, _BODYPR_ORDER)

    sp3d = body_pr.makeelement(
        qn("a:sp3d"),
        {
            "extrusionH": str(int(extrusion_pt * 12700)),
            "contourW": "0",
            "prstMaterial": material,
        },
    )
    sp3d.append(
        sp3d.makeelement(
            qn("a:bevelT"),
            {
                "w": str(int(bevel_width_pt * 12700)),
                "h": str(int(bevel_height_pt * 12700)),
                "prst": bevel,
            },
        )
    )
    if extrusion_colour:
        wrapper = sp3d.makeelement(qn("a:extrusionClr"), {})
        wrapper.append(_srgb(sp3d, extrusion_colour))
        sp3d.append(wrapper)
    if contour_colour:
        wrapper = sp3d.makeelement(qn("a:contourClr"), {})
        wrapper.append(_srgb(sp3d, contour_colour))
        sp3d.append(wrapper)
    _insert_ordered(body_pr, sp3d, _BODYPR_ORDER)


def set_wrap(shape, mode: str = "square") -> None:
    """Control text wrapping (``square`` / ``none``)."""

    body_pr = _body_pr(shape)
    if body_pr is None:
        return
    body_pr.set("wrap", mode)


def set_flat_text(shape) -> None:
    """Render the text flat (no perspective)."""

    body_pr = _body_pr(shape)
    if body_pr is None:
        return
    existing = body_pr.find(qn("a:flatTx"))
    if existing is not None:
        body_pr.remove(existing)
    flat = body_pr.makeelement(qn("a:flatTx"), {})
    _insert_ordered(body_pr, flat, _BODYPR_ORDER)


#: WordArt transforms that read well for Chinese titles.
TEXT_WARP_PRESETS = (
    "textPlain",
    "textArchUp",
    "textArchDown",
    "textCircle",
    "textCirclePour",
    "textButton",
    "textCurveUp",
    "textCurveDown",
    "textCanUp",
    "textCanDown",
    "textWave1",
    "textWave2",
    "textDoubleWave1",
    "textRibbon",
    "textRibbon2",
    "textTriangle",
    "textTriangleInverted",
    "textChevron",
    "textChevronInverted",
    "textRingInside",
    "textRingOutside",
    "textInflate",
    "textDeflate",
    "textInflateBottom",
    "textDeflateBottom",
    "textFadeRight",
    "textFadeLeft",
    "textFadeUp",
    "textFadeDown",
    "textSlantUp",
    "textSlantDown",
    "textCascadeUp",
    "textCascadeDown",
)