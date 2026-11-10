#!/usr/bin/env python3
"""Motion engine for pro-ppt-studio V4 (static + dynamic decks).

``python-pptx`` cannot express slide transitions or animation timing, yet a
"dynamic" deck is expected to use them properly: a transition language that
matches the narrative, and entrance / emphasis / exit choreography rather than
one effect pasted on every slide.

Two independent features are provided:

``set_transition``
    Writes ``p:transition`` (or the ``mc:AlternateContent`` form needed for
    modern PowerPoint 2010+ transitions such as Ripple, Morph, Prism, Vortex).

``animate``
    Writes a ``p:timing`` tree describing a click-grouped animation sequence
    with entrance, emphasis and exit behaviours.

Both are written straight into the slide XML. The generated markup follows the
ECMA-376 / Microsoft extension schemas so PowerPoint opens the file without a
repair prompt.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable, Sequence
from pathlib import Path

from pptx.util import Inches

from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P159_NS = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"

_NSDECL = (
    f'xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:mc="{MC_NS}" '
    f'xmlns:p14="{P14_NS}" xmlns:p159="{P159_NS}"'
)


# ---------------------------------------------------------------------------
# transitions
# ---------------------------------------------------------------------------

#: Classic transitions: name -> (ooxml child, attributes)
CLASSIC_TRANSITIONS: dict[str, tuple[str, dict[str, str]]] = {
    "fade": ("fade", {}),
    "cut": ("cut", {}),
    "dissolve": ("dissolve", {}),
    "newsflash": ("newsflash", {}),
    "wipe_up": ("wipe", {"dir": "u"}),
    "wipe_down": ("wipe", {"dir": "d"}),
    "wipe_left": ("wipe", {"dir": "l"}),
    "wipe_right": ("wipe", {"dir": "r"}),
    "push_up": ("push", {"dir": "u"}),
    "push_down": ("push", {"dir": "d"}),
    "push_left": ("push", {"dir": "l"}),
    "push_right": ("push", {"dir": "r"}),
    "cover_up": ("cover", {"dir": "u"}),
    "cover_down": ("cover", {"dir": "d"}),
    "cover_left": ("cover", {"dir": "l"}),
    "cover_right": ("cover", {"dir": "r"}),
    "pull_up": ("pull", {"dir": "u"}),
    "pull_down": ("pull", {"dir": "d"}),
    "pull_left": ("pull", {"dir": "l"}),
    "pull_right": ("pull", {"dir": "r"}),
    "split_horizontal": ("split", {"orient": "horz", "dir": "out"}),
    "split_vertical": ("split", {"orient": "vert", "dir": "out"}),
    "blinds_horizontal": ("blinds", {"dir": "horz"}),
    "blinds_vertical": ("blinds", {"dir": "vert"}),
    "checkerboard": ("checker", {"dir": "horz"}),
    "comb_horizontal": ("comb", {"dir": "horz"}),
    "circle": ("circle", {}),
    "diamond": ("diamond", {}),
    "plus": ("plus", {}),
    "random": ("random", {}),
    "random_bars": ("randomBar", {"dir": "horz"}),
    "strips": ("strips", {"dir": "rd"}),
    "wedge": ("wedge", {}),
    "wheel": ("wheel", {"spokes": "4"}),
    "zoom_in": ("zoom", {"dir": "in"}),
    "zoom_out": ("zoom", {"dir": "out"}),
}

#: Modern transitions (PowerPoint 2010+): name -> (p14 child, attributes, fallback)
MODERN_TRANSITIONS: dict[str, tuple[str, dict[str, str], str]] = {
    "ripple": ("ripple", {"dir": "center"}, "circle"),
    "honeycomb": ("honeycomb", {}, "dissolve"),
    "glitter": ("glitter", {"dir": "l", "pattern": "hexagon"}, "dissolve"),
    "vortex": ("vortex", {"dir": "l"}, "wheel"),
    "prism": ("prism", {"isInverted": "0", "isContent": "1"}, "dissolve"),
    "doors": ("doors", {"dir": "horz"}, "split_vertical"),
    "window": ("window", {"dir": "horz"}, "split_horizontal"),
    "ferris": ("ferris", {"dir": "l"}, "wheel"),
    "gallery": ("gallery", {"dir": "l"}, "push_left"),
    "conveyor": ("conveyor", {"dir": "l"}, "push_left"),
    "pan": ("pan", {"dir": "d"}, "push_up"),
    "flash": ("flash", {}, "fade"),
    "shred": ("shred", {"style": "strip", "dir": "in"}, "blinds_vertical"),
    "reveal": ("reveal", {"thruBlk": "0", "dir": "l"}, "push_left"),
    "wheel_reverse": ("wheelReverse", {"spokes": "4"}, "wheel"),
    "switch": ("switch", {"dir": "l"}, "push_left"),
    "flip": ("flip", {"dir": "l"}, "dissolve"),
    "curtains": ("curtains", {}, "split_horizontal"),
    "window_blinds": ("windowBlinds", {"dir": "horz"}, "blinds_horizontal"),
    "window_checkerboard": ("windowCheckerboard", {"dir": "horz"}, "checkerboard"),
    "airplane": ("airplane", {"dir": "l"}, "push_left"),
    "origami": ("origami", {"dir": "l"}, "dissolve"),
    "crush": ("crush", {}, "dissolve"),
    "morph": ("morph", {"option": "byObject"}, "fade"),
}


def _transition_xml(
    kind: str,
    *,
    duration_ms: int,
    advance_ms: int | None,
    advance_click: bool,
) -> str:
    """Build the transition XML fragment for a slide."""

    speed = "slow" if duration_ms >= 1200 else ("med" if duration_ms >= 600 else "fast")
    advance_attrs = ""
    if advance_ms is not None:
        advance_attrs += f' advTm="{int(advance_ms)}"'
    if not advance_click:
        advance_attrs += ' advClick="0"'

    if kind in MODERN_TRANSITIONS:
        child, attrs, fallback = MODERN_TRANSITIONS[kind]
        attr_text = " ".join(f'{key}="{value}"' for key, value in attrs.items())
        inner = f'<p14:{child}{(" " + attr_text) if attr_text else ""}/>'
        requires = "p159" if child == "morph" else "p14"
        fallback_xml = _classic_transition_xml(fallback, speed=speed, advance_attrs=advance_attrs)
        return (
            f"<mc:AlternateContent {_NSDECL}>"
            f'<mc:Choice xmlns:p14="{P14_NS}" xmlns:p159="{P159_NS}" Requires="{requires}">'
            f'<p:transition spd="{speed}" p14:dur="{int(duration_ms)}"{advance_attrs}>{inner}</p:transition>'
            f"</mc:Choice>"
            f"<mc:Fallback>{fallback_xml}</mc:Fallback>"
            f"</mc:AlternateContent>"
        )

    return _classic_transition_xml(kind, speed=speed, advance_attrs=advance_attrs)


def _classic_transition_xml(kind: str, *, speed: str, advance_attrs: str) -> str:
    if kind not in CLASSIC_TRANSITIONS:
        kind = "fade"
    child, attrs = CLASSIC_TRANSITIONS[kind]
    attr_text = " ".join(f'{key}="{value}"' for key, value in attrs.items())
    inner = f'<p:{child}{(" " + attr_text) if attr_text else ""}/>'
    return (
        f'<p:transition {_NSDECL} spd="{speed}"{advance_attrs}>{inner}</p:transition>'
    )


def set_transition(
    slide,
    kind: str = "fade",
    *,
    duration_ms: int = 800,
    advance_ms: int | None = None,
    advance_click: bool = True,
) -> None:
    """Attach a slide transition.

    ``kind`` may be any key of :data:`CLASSIC_TRANSITIONS` or
    :data:`MODERN_TRANSITIONS` (``fade``, ``push_left``, ``ripple``,
    ``morph``, ``vortex`` ...).
    """

    kind = kind.strip().lower()
    if kind not in CLASSIC_TRANSITIONS and kind not in MODERN_TRANSITIONS:
        raise ValueError(f"Unknown transition {kind!r}")

    sld = slide._element
    for tag in ("mc:AlternateContent", "p:transition"):
        for node in sld.findall(
            "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
            if tag.startswith("mc")
            else qn(tag)
        ):
            sld.remove(node)

    xml = _transition_xml(
        kind, duration_ms=duration_ms, advance_ms=advance_ms, advance_click=advance_click
    )
    fragment = parse_xml(xml)
    clr_map = sld.find(qn("p:clrMapOvr"))
    if clr_map is not None:
        clr_map.addnext(fragment)
    else:
        sld.append(fragment)


def has_transition(slide) -> bool:
    """Whether the slide carries any transition."""

    sld = slide._element
    if sld.find(qn("p:transition")) is not None:
        return True
    for node in sld.findall("{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"):
        if node.find(qn("p:transition")) is not None:
            return True
    return False


# ---------------------------------------------------------------------------
# animation
# ---------------------------------------------------------------------------


@dataclass
class AnimationStep:
    """One animated shape in a slide's choreography."""

    shape_id: int
    effect: str = "fade"
    duration_ms: int = 600
    delay_ms: int = 0
    trigger: str = "click"  # click | with | after
    kind: str = "entrance"   # entrance | emphasis | exit
    amount: float = 0.25     # travel distance for motion effects
    emphasis_scale: float = 1.12
    rotation_degrees: float = 0.0


#: entrance effects: name -> (presetID, presetSubtype, filter, motion)
_ENTRANCE = {
    "fade": (10, 0, "fade", None),
    "dissolve": (9, 0, "dissolve", None),
    "wipe_up": (22, 4, "wipe(up)", None),
    "wipe_down": (22, 1, "wipe(down)", None),
    "wipe_left": (22, 8, "wipe(left)", None),
    "wipe_right": (22, 2, "wipe(right)", None),
    "circle": (3, 16, "circle", None),
    "diamond": (5, 16, "diamond", None),
    "plus": (13, 16, "plus", None),
    "wheel": (18, 4, "wheel(4)", None),
    "checkerboard": (5, 10, "checkerboard(across)", None),
    "blinds": (3, 10, "blinds(horizontal)", None),
    "box": (2, 16, "box(in)", None),
    "split": (13, 13, "split(horizontal,in)", None),
    "strips": (18, 10, "strips(downLeft)", None),
    "random_bars": (14, 10, "randombar(horizontal)", None),
    "wedge": (18, 16, "wedge", None),
    "zoom": (23, 16, "fade", "zoom"),
    "fly_up": (2, 4, "fade", "fly_up"),
    "fly_down": (2, 1, "fade", "fly_down"),
    "fly_left": (2, 8, "fade", "fly_left"),
    "fly_right": (2, 2, "fade", "fly_right"),
    "float_up": (42, 4, "fade", "float_up"),
    "float_down": (42, 1, "fade", "float_down"),
    "rise_up": (42, 4, "fade", "float_up"),
}

#: emphasis effects: name -> (presetID, kind)
_EMPHASIS = {
    "pulse": (1, "scale"),
    "grow": (6, "scale"),
    "shrink": (6, "scale"),
    "spin": (8, "rotation"),
    "teeter": (11, "rotation"),
    "color_change": (9, "color"),
    "bold_flash": (2, "style"),
}

def _effect_fragment(step: AnimationStep, node_id: int) -> tuple[int, str]:
    """Return the behaviour XML and the next free node id."""

    effect = step.effect.strip().lower()
    shape = step.shape_id
    dur = int(step.duration_ms)

    if effect == "motion_path":
        distance = max(0.01, min(1.0, step.amount))
        path = f"M 0 0 L {distance:.3f} {distance:.3f} E"
        body = (
            f'<p:animMotion origin="layout" path="{path}" '
            f'pathEditMode="relative" rAng="0" ptsTypes="">'
            f'<p:cBhvr><p:cTn id="{node_id}" dur="{dur}" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl></p:cBhvr>'
            f"</p:animMotion>"
        )
        return node_id + 1, body

    if effect in _EMPHASIS:
        preset_id, kind = _EMPHASIS[effect]
        preset_class = "emph"
        if kind == "scale":
            scale = step.emphasis_scale if effect != "shrink" else 1 / step.emphasis_scale
            body = (
                f'<p:animScale><p:cBhvr><p:cTn id="{node_id}" dur="{dur}" fill="hold"/>'
                f'<p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl></p:cBhvr>'
                f'<p:from x="100000" y="100000"/><p:to x="{int(scale * 100000)}" y="{int(scale * 100000)}"/>'
                f"</p:animScale>"
            )
            node_id += 1
        elif kind == "rotation":
            body = (
                f'<p:animRot by="{int(step.rotation_degrees * 60000)}"><p:cBhvr>'
                f'<p:cTn id="{node_id}" dur="{dur}" fill="hold"/>'
                f'<p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl>'
                f'<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst>'
                f"</p:cBhvr></p:animRot>"
            )
            node_id += 1
        elif kind == "color":
            body = (
                f'<p:animClr clrSpc="rgb"><p:cBhvr><p:cTn id="{node_id}" dur="{dur}" fill="hold"/>'
                f'<p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl>'
                f'<p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst>'
                f"</p:cBhvr><p:to><a:srgbClr val=\"FF0000\"/></p:to></p:animClr>"
            )
            node_id += 1
        else:
            body = (
                f'<p:animEffect transition="none" filter="fade"><p:cBhvr>'
                f'<p:cTn id="{node_id}" dur="{dur}" fill="hold"/>'
                f'<p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl>'
                f"</p:cBhvr></p:animEffect>"
            )
            node_id += 1
        return node_id, body

    is_exit = step.kind == "exit"
    key = effect
    if key not in _ENTRANCE:
        key = "fade"
    preset_id, preset_subtype, filter_name, motion = _ENTRANCE[key]
    preset_class = "exit" if is_exit else "entr"
    transition = "out" if is_exit else "in"

    parts = [
        f'<p:set><p:cBhvr><p:cTn id="{node_id}" dur="1" fill="hold">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        f'<p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl>'
        f"<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>"
        f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
    ]
    node_id += 1

    parts.append(
        f'<p:animEffect transition="{transition}" filter="{filter_name}"><p:cBhvr>'
        f'<p:cTn id="{node_id}" dur="{dur}"/>'
        f'<p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl></p:cBhvr></p:animEffect>'
    )
    node_id += 1

    if motion == "zoom":
        parts.append(
            f'<p:animScale><p:cBhvr><p:cTn id="{node_id}" dur="{dur}" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl></p:cBhvr>'
            f'<p:from x="20000" y="20000"/><p:to x="100000" y="100000"/></p:animScale>'
        )
        node_id += 1
    elif motion in {"fly_up", "fly_down", "fly_left", "fly_right", "float_up", "float_down"}:
        amount = step.amount
        if motion.endswith("up"):
            axis, formula = "ppt_y", f"#ppt_y+{amount}"
        elif motion.endswith("down"):
            axis, formula = "ppt_y", f"#ppt_y-{amount}"
        elif motion.endswith("left"):
            axis, formula = "ppt_x", f"#ppt_x+{amount}"
        else:
            axis, formula = "ppt_x", f"#ppt_x-{amount}"
        parts.append(
            f'<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            f'<p:cTn id="{node_id}" dur="{dur}" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl>'
            f"<p:attrNameLst><p:attrName>{axis}</p:attrName></p:attrNameLst>"
            f"</p:cBhvr><p:tavLst>"
            f'<p:tav tm="0"><p:val><p:strVal val="{formula}"/></p:val></p:tav>'
            f'<p:tav tm="100000"><p:val><p:strVal val="#{axis}"/></p:val></p:tav>'
            f"</p:tavLst></p:anim>"
        )
        node_id += 1

    return node_id, "".join(parts)


def _step_xml(step: AnimationStep, node_id: int) -> tuple[int, str]:
    """Build one ``p:par`` describing an animated shape."""

    node_id, body = _effect_fragment(step, node_id)
    node_type = {"click": "clickEffect", "with": "withEffect", "after": "afterEffect"}[step.trigger]
    delay = int(step.delay_ms)
    preset_class = (
        "emph"
        if step.kind == "emphasis" or step.effect.strip().lower() in _EMPHASIS
        else (
            "path"
            if step.kind == "path"
            else ("exit" if step.kind == "exit" else "entr")
        )
    )
    xml = (
        f'<p:par><p:cTn id="{node_id}" presetID="0" presetClass="{preset_class}" presetSubtype="0" '
        f'fill="hold" grpId="0" nodeType="{node_type}">'
        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>'
        f"<p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>"
    )
    return node_id + 1, xml


def animate(slide, steps: Sequence[AnimationStep | dict]) -> bool:
    """Attach an animation sequence to a slide.

    Each step is either an :class:`AnimationStep` or a dict with the same
    fields. Steps with ``trigger='click'`` start a new click group; ``with``
    and ``after`` continue the current group.
    """

    normalised: list[AnimationStep] = []
    for step in steps:
        if isinstance(step, AnimationStep):
            normalised.append(step)
        else:
            normalised.append(AnimationStep(**step))
    if not normalised:
        return False

    node_id = 5
    groups: list[list[str]] = []
    current: list[str] = []
    for step in normalised:
        node_id, xml = _step_xml(step, node_id)
        if step.trigger == "click" and current:
            groups.append(current)
            current = []
        current.append(xml)
    if current:
        groups.append(current)

    group_xml = []
    node_id = 3
    for group in groups:
        inner = []
        for member in group:
            inner.append(member)
        group_xml.append(
            f'<p:par><p:cTn id="{node_id}" fill="hold">'
            f'<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
            f'<p:childTnLst><p:par><p:cTn id="{node_id + 1}" fill="hold">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f"<p:childTnLst>{''.join(inner)}</p:childTnLst>"
            f"</p:cTn></p:par></p:childTnLst></p:cTn></p:par>"
        )
        node_id += 2
        # the shape level ids were generated before the group ids; keep them
        # unique by shifting the group ids above the shape level range.
    offset = node_id
    shape_xml = []
    node_id = 100
    for step in normalised:
        node_id, xml = _step_xml(step, node_id)
        shape_xml.append(xml)

    rebuilt_groups: list[list[str]] = []
    current = []
    for step, xml in zip(normalised, shape_xml):
        if step.trigger == "click" and current:
            rebuilt_groups.append(current)
            current = []
        current.append(xml)
    if current:
        rebuilt_groups.append(current)

    group_xml = []
    gid = 3
    for group in rebuilt_groups:
        group_xml.append(
            f'<p:par><p:cTn id="{gid}" fill="hold">'
            f'<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
            f'<p:childTnLst><p:par><p:cTn id="{gid + 1}" fill="hold">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f"<p:childTnLst>{''.join(group)}</p:childTnLst>"
            f"</p:cTn></p:par></p:childTnLst></p:cTn></p:par>"
        )
        gid += 2

    xml = (
        f"<p:timing {_NSDECL}><p:tnLst><p:par>"
        f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        f"<p:childTnLst>{''.join(group_xml)}</p:childTnLst></p:cTn>"
        f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        f"</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"
    )
    fragment = parse_xml(xml)
    sld = slide._element
    existing = sld.find(qn("p:timing"))
    if existing is not None:
        sld.remove(existing)
    sld.append(fragment)
    return True


def has_animation(slide) -> bool:
    """Whether the slide carries an animation timeline."""

    return slide._element.find(qn("p:timing")) is not None


def animation_summary(slide) -> dict:
    """Summarise the animation timeline of a slide for the QA report."""

    timing = slide._element.find(qn("p:timing"))
    if timing is None:
        return {
            "animated": False,
            "steps": 0,
            "targets": [],
            "effects": [],
            "categories": [],
        }
    target_ids: list[str] = []
    for node in timing.iter(qn("p:spTgt")):
        target = node.get("spid")
        if target and target not in target_ids:
            target_ids.append(target)
    sequences = [
        node
        for node in timing.iter(qn("p:cTn"))
        if node.get("nodeType") in {"clickEffect", "withEffect", "afterEffect"}
    ]
    effects: list[str] = []
    for node in sequences:
        for child in node.iter():
            tag = child.tag.rsplit("}", 1)[-1]
            if tag not in {
                "cTn",
                "stCondLst",
                "cond",
                "childTnLst",
                "cBhvr",
                "tgtEl",
                "spTgt",
                "attrNameLst",
                "attrName",
                "to",
                "strVal",
                "from",
                "tavLst",
                "tav",
                "val",
                "srgbClr",
            }:
                effects.append(tag)
    return {
        "animated": True,
        "steps": len(sequences),
        "targets": target_ids,
        "effects": sorted(set(effects)),
        "categories": sorted(
            {
                node.get("presetClass", "entr")
                for node in sequences
            }
        ),
    }


def transition_summary(slide) -> dict:
    """Summarise the transition of a slide for the QA report."""

    sld = slide._element
    node = sld.find(qn("p:transition"))
    kind = None
    child_name = None
    child_attrs: dict[str, str] = {}
    duration = None
    if node is not None:
        for child in node:
            child_name = child.tag.split("}")[-1]
            child_attrs = dict(child.attrib)
        duration = node.get("{http://schemas.microsoft.com/office/powerpoint/2010/main}dur")
    if kind is None:
        for alternate in sld.findall(
            "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
        ):
            choice = alternate.find(
                "{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice"
            )
            if choice is None:
                continue
            transition = choice.find(qn("p:transition"))
            if transition is None:
                continue
            for child in transition:
                child_name = child.tag.split("}")[-1]
                child_attrs = dict(child.attrib)
            duration = transition.get(
                "{http://schemas.microsoft.com/office/powerpoint/2010/main}dur"
            )
            break
    if child_name:
        for candidate, (tag, attrs) in CLASSIC_TRANSITIONS.items():
            if tag == child_name and attrs == child_attrs:
                kind = candidate
                break
        if kind is None:
            for candidate, (tag, attrs, _fallback) in MODERN_TRANSITIONS.items():
                if tag == child_name and attrs == child_attrs:
                    kind = candidate
                    break
        if kind is None:
            kind = child_name
    return {"transition": kind, "duration_ms": duration}


def add_media(
    slide,
    media_path: str | Path,
    *,
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    poster_path: str | Path | None = None,
    play_mode: str = "click",
    loop: bool = False,
    volume: float = 0.8,
):
    """Embed video or audio and configure click/automatic playback.

    ``python-pptx`` already creates the media relationship and poster shape.
    This helper adjusts the native media timing node for auto-start and loop.
    Audio uses the same relationship container with an ``a:audioFile`` tag when
    a playable audio file is supplied.
    """

    if play_mode not in {"click", "auto"}:
        raise ValueError("play_mode must be 'click' or 'auto'")
    media_path = Path(media_path)
    if not media_path.is_file():
        raise FileNotFoundError(media_path)
    suffix = media_path.suffix.lower()
    is_audio = suffix in {".mp3", ".wav", ".m4a", ".aac", ".wma", ".flac"}
    mime = (
        "audio/mpeg"
        if suffix == ".mp3"
        else "audio/wav"
        if suffix == ".wav"
        else "video/mp4"
        if suffix in {".mp4", ".m4v"}
        else "video/unknown"
    )
    shape = slide.shapes.add_movie(
        str(media_path),
        Inches(left_in),
        Inches(top_in),
        Inches(width_in),
        Inches(height_in),
        poster_frame_image=str(poster_path) if poster_path else None,
        mime_type=mime,
    )
    if is_audio:
        for node in shape._element.iter(qn("a:videoFile")):
            node.tag = qn("a:audioFile")
    media_nodes = list(slide._element.iter(qn("p:cMediaNode")))
    if media_nodes:
        node = media_nodes[-1]
        node.set("vol", str(int(max(0.0, min(1.0, volume)) * 100000)))
        control = node.find(qn("p:cTn"))
        if control is not None:
            if play_mode == "auto":
                condition = control.find(qn("p:stCondLst"))
                if condition is None:
                    condition = control.makeelement(qn("p:stCondLst"), {})
                    control.insert(0, condition)
                for child in list(condition):
                    condition.remove(child)
                delay = condition.makeelement(qn("p:cond"), {"delay": "0"})
                condition.append(delay)
            if loop:
                control.set("repeatCount", "indefinite")
    return shape
