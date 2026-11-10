#!/usr/bin/env python3
"""V4.0.0 capability probe for generation, templates, motion, and rendering.

Usage:
    python probe_environment.py [--json environment.json]

The probe is intentionally non-blocking: optional tools that are missing are
reported as unavailable instead of being installed or invented.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import platform
import shutil
import sys
from datetime import datetime
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

from level_profiles import LEVEL_PROFILES, SKILL_VERSION


POWERPOINT_COM_ERROR_CODE = "0x80070520"

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

MODULES = {
    "python_pptx": "pptx",
    "pillow": "PIL",
    "matplotlib": "matplotlib",
    "numpy": "numpy",
    "pandas": "pandas",
    "cairosvg": "cairosvg",
    "pymupdf": "fitz",
    "pypdf": "pypdf",
    "pywin32": "win32com",
    "opencv": "cv2",
    "fonttools": "fontTools",
    "scikit_learn": "sklearn",
    "moviepy": "moviepy",
    "colorthief": "colorthief",
}

PACKAGES = {
    "python_pptx": "python-pptx",
    "pillow": "Pillow",
    "matplotlib": "matplotlib",
    "numpy": "numpy",
    "pandas": "pandas",
    "cairosvg": "cairosvg",
    "pymupdf": "PyMuPDF",
    "pypdf": "pypdf",
    "pywin32": "pywin32",
    "opencv": "opencv-python",
    "fonttools": "fonttools",
    "scikit_learn": "scikit-learn",
    "moviepy": "moviepy",
    "colorthief": "colorthief-py",
}


def _module_status(module_name: str) -> tuple[bool, str | None]:
    available = importlib.util.find_spec(module_name) is not None
    package = next((p for k, p in PACKAGES.items() if MODULES[k] == module_name), module_name)
    if not available:
        return False, None
    try:
        return True, version(package)
    except PackageNotFoundError:
        return True, None


def _command_status(command: str) -> str | None:
    return shutil.which(command)


def _powerpoint_status() -> dict:
    if platform.system() != "Windows":
        return {"prog_id_registered": False, "executable": None}

    prog_id_registered = False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application"):
            prog_id_registered = True
    except (ImportError, FileNotFoundError, OSError):
        pass

    candidates = []
    for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
        root = os.environ.get(env_name)
        if root:
            candidates.append(Path(root) / "Microsoft Office" / "root" / "Office16" / "POWERPNT.EXE")
            candidates.append(Path(root) / "Microsoft Office" / "Office16" / "POWERPNT.EXE")
    executable = next((str(p) for p in candidates if p.is_file()), None)
    return {"prog_id_registered": prog_id_registered, "executable": executable}


def _font_status() -> dict[str, bool]:
    windows = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    return {
        "Microsoft YaHei": (windows / "msyh.ttc").is_file(),
        "Arial": (windows / "arial.ttf").is_file(),
        "Consolas": any(windows.glob("consola*.ttf")),
        "STZhongsong": any(windows.glob("STZHONGS*.TTF")),
        "Segoe UI": (windows / "segoeui.ttf").is_file(),
        "Aptos": any(windows.glob("Aptos*.ttf")),
        "Noto Sans SC": any(windows.glob("NotoSansSC*.ttf")),
    }


def _artifact_tool_status() -> dict:
    node = os.environ.get("CODEX_ARTIFACT_NODE")
    modules = os.environ.get("CODEX_ARTIFACT_NODE_MODULES")
    package_file = (
        Path(modules) / "@oai" / "artifact-tool" / "package.json"
        if modules
        else None
    )
    available = bool(
        node
        and modules
        and Path(node).is_file()
        and package_file is not None
        and package_file.is_file()
    )
    return {
        "available": available,
        "node_executable": node,
        "node_modules": modules,
        "package_file": str(package_file) if package_file else None,
        "requirement": "Set CODEX_ARTIFACT_NODE and CODEX_ARTIFACT_NODE_MODULES explicitly",
    }


def _load_previous_report(path: str | None) -> dict:
    if not path:
        return {}
    output = Path(path)
    if not output.is_file():
        return {}
    try:
        return json.loads(output.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _powerpoint_blocked(previous_state: dict | None) -> bool:
    if not previous_state:
        return False
    return bool(
        previous_state.get("auto_retry_allowed") is False
        and previous_state.get("retry_policy") == "skip_powerpoint_com"
    )


def probe(previous_powerpoint_com: dict | None = None) -> dict:
    modules = {key: _module_status(name) for key, name in MODULES.items()}
    module_report = {
        key: {"available": status[0], "version": status[1]}
        for key, status in modules.items()
    }

    commands = {
        name: _command_status(name)
        for name in ("soffice", "libreoffice", "pdftoppm", "magick", "ffmpeg", "powershell", "pwsh")
    }
    powerpoint = _powerpoint_status()
    artifact_tool = _artifact_tool_status()
    fonts = _font_status() if platform.system() == "Windows" else {}

    powerpoint_blocked = _powerpoint_blocked(previous_powerpoint_com)
    powerpoint["runtime_state"] = (
        "permanently_limited_for_current_environment"
        if powerpoint_blocked
        else "runtime_unverified"
    )
    powerpoint["error_code"] = (
        previous_powerpoint_com.get("error_code")
        if powerpoint_blocked and previous_powerpoint_com
        else None
    )
    powerpoint["auto_retry_allowed"] = not powerpoint_blocked
    renderers = []
    if powerpoint["prog_id_registered"] and not powerpoint_blocked:
        if modules["pywin32"][0]:
            renderers.append("PowerPoint COM via pywin32")
        if commands["powershell"] or commands["pwsh"]:
            renderers.append("PowerPoint COM via PowerShell")
    if commands["soffice"] or commands["libreoffice"]:
        if commands["pdftoppm"]:
            renderers.append("LibreOffice PDF to PNG via pdftoppm")
        else:
            renderers.append("LibreOffice PNG (slide coverage may be limited)")
    if artifact_tool["available"]:
        renderers.append("Artifact Tool (conditional, explicitly configured)")

    generation = []
    if modules["python_pptx"][0]:
        generation.append("python-pptx")
    image_processing = ["Pillow"] if modules["pillow"][0] else []
    charting = []
    if modules["python_pptx"][0]:
        charting.append("native PowerPoint charts via python-pptx")
    if modules["matplotlib"][0]:
        charting.append("Matplotlib raster charts")
    template_parsing = [
        "PPTX/POTX/THMX themes and layouts",
        "PDF colours and font names",
        "DOCX brand signals",
        "raster images via Pillow",
        "SVG colour and primitive inventory",
        "JSON/YAML design configs",
        "MP4/MOV container metadata",
        "Markdown outline",
    ]
    if commands["ffmpeg"]:
        template_parsing.append("video keyframes via ffmpeg")
    else:
        template_parsing.append("video frames unavailable without ffmpeg")
    dynamic = {
        "transition_xml": True,
        "object_animation_xml": True,
        "media_embedding": bool(modules["python_pptx"][0]),
        "native_video_export": bool(
            powerpoint["prog_id_registered"] and commands["powershell"]
        ),
        "ffmpeg_video_fallback": bool(commands["ffmpeg"]),
        "gif_transition_preview": bool(modules["pillow"][0]),
    }

    limitations = []
    if not image_processing:
        limitations.append("No supported image processing library detected; validate images with another tool.")
    if not renderers:
        limitations.append("No PowerPoint-compatible renderer detected; visual QA cannot be completed.")
    else:
        limitations.append(
            "Renderer registration does not guarantee automation in the current session; run smoke_test.py before claiming rendering works."
        )
    if powerpoint_blocked:
        limitations.append(
            "PowerPoint COM is disabled for the recorded environment after "
            f"{POWERPOINT_COM_ERROR_CODE}; retry_policy=skip_powerpoint_com. "
            "Reset only after the execution session/logon environment changes."
        )
    if not modules["matplotlib"][0]:
        limitations.append("Matplotlib is unavailable; use native PowerPoint charts or another installed charting tool.")
    if not commands["ffmpeg"]:
        limitations.append(
            "ffmpeg is unavailable: video templates use container metadata only, "
            "and MP4 fallback export is disabled; GIF transition preview remains available."
        )
    limitations.append(
        "The probe cannot verify network image search or image generation; confirm those capabilities at task time."
    )

    return {
        "skill_version": SKILL_VERSION,
        "probed_at": datetime.now().astimezone().isoformat(),
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
        },
        "python": {
            "version": platform.python_version(),
            "executable": sys.executable,
            "utf8_mode": bool(sys.flags.utf8_mode),
        },
        "capabilities": {
            "ppt_generation": generation,
            "image_processing": image_processing,
            "charting": charting,
            "template_parsing": template_parsing,
            "dynamic_ppt": dynamic,
            "rendering_candidates": renderers,
            "expression_strategies": {
                level: {
                    "strategy": profile.strategy.key,
                    "style_branches": [
                        branch.key for branch in profile.style_branches
                    ],
                }
                for level, profile in LEVEL_PROFILES.items()
            },
        },
        "runtime_verification_required": {
            "smoke_test": "python scripts/smoke_test.py <out_dir>",
            "rendering_runtime_verified": False,
        },
        "modules": module_report,
        "commands": commands,
        "powerpoint": powerpoint,
        "powerpoint_com": previous_powerpoint_com,
        "artifact_tool": artifact_tool,
        "fonts": fonts,
        "limitations": limitations,
    }


def _print_summary(report: dict) -> None:
    capabilities = report["capabilities"]
    print("PPT generation:", ", ".join(capabilities["ppt_generation"]) or "unavailable")
    print("Image processing:", ", ".join(capabilities["image_processing"]) or "unavailable")
    print("Charting:", ", ".join(capabilities["charting"]) or "unavailable")
    print("Rendering candidates:", ", ".join(capabilities["rendering_candidates"]) or "unavailable")
    powerpoint_com = report.get("powerpoint_com")
    if powerpoint_com:
        print(
            "PowerPoint COM state:",
            powerpoint_com.get("status"),
            powerpoint_com.get("error_code"),
            powerpoint_com.get("retry_policy"),
        )
    print("Limitations:")
    for item in report["limitations"]:
        print(f"- {item}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Probe pro-ppt-studio tool capabilities")
    parser.add_argument("--json", dest="json_path", help="Write the complete report to this JSON file")
    args = parser.parse_args()

    previous_report = _load_previous_report(args.json_path)
    report = probe(previous_report.get("powerpoint_com"))
    if args.json_path:
        report["environment_file"] = str(Path(args.json_path).resolve())
    _print_summary(report)
    if args.json_path:
        output = Path(args.json_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON report: {output.resolve()}")


if __name__ == "__main__":
    main()
