#!/usr/bin/env python3
"""
pro-ppt-studio V4.0.0 渲染脚本（实际渲染、动效预览）

用法：
    python render_preview.py <file.pptx> <out_dir> [--environment environment.json]
        [--width 1920] [--height 1080] [--no-contact-sheet] [--video] [--gif]

将 PPTX 逐页渲染为 PNG，并生成缩略图总览（contact sheet）。

渲染后端按优先级自动选择：
  1. pywin32（win32com，PowerPoint COM）
  2. Windows PowerShell + PowerPoint COM（无需 pywin32）
  3. LibreOffice（soffice）通过 PDF 转逐页 PNG
  4. Artifact Tool（仅在显式提供 Node 与 @oai/artifact-tool 路径时）

每个后端只尝试一次，并在独立临时目录中导出。只有导出页数与 PPTX
实际页数一致时，才把该尝试标记为渲染成功。

PowerPoint COM 返回 0x80070520 时，将其记录为本次临时故障和当前环境的
永久性限制；同一环境后续跳过整个 PowerPoint COM 后端族。只有显式使用
--reset-powerpoint-com-state 才清除该状态。

依赖：
  - 渲染：Microsoft PowerPoint（Windows）或 LibreOffice
  - 缩略图总览：Pillow（可选）

输出：
  - <out_dir>/slide_01.png ... slide_NN.png
  - <out_dir>/contact_sheet.png（可选）
  - <out_dir>/preview.mp4（可选，优先 Native PowerPoint）
  - <out_dir>/preview.gif（可选，交叉淡入的页面节奏预览）
  - <out_dir>/render_manifest.json
"""

import argparse
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

SKILL_VERSION = "4.0.0"
POWERPOINT_COM_ERROR_CODE = "0x80070520"
POWERPOINT_COM_HINT = (
    "PowerPoint COM 在当前执行会话或登录上下文中无法启动。"
    "该错误按本次临时运行故障记录，并将 PowerPoint COM 标记为当前环境的"
    "永久性限制；后续自动渲染跳过 PowerPoint COM，直到显式重置或环境改变。"
)
BACKEND_KIND_VALUES = {
    "native_powerpoint",
    "compatible_libreoffice",
    "alternative_artifact_tool",
}
VERIFICATION_BY_BACKEND_KIND = {
    "native_powerpoint": "rendered_native_powerpoint",
    "compatible_libreoffice": "rendered_compatible_libreoffice",
    "alternative_artifact_tool": "rendered_alternative_artifact_tool",
}


class RenderBackendError(RuntimeError):
    """Structured backend failure with code, hint, full detail, and stack."""

    def __init__(
        self,
        backend: str,
        *,
        error_code: str | None,
        hint: str,
        detail: str,
        stack: str | None = None,
    ) -> None:
        super().__init__(hint or detail or "Render backend failed")
        self.backend = backend
        self.error_code = error_code
        self.hint = hint
        self.detail = detail
        self.stack = stack or detail


def extract_error_code(text: str) -> str | None:
    if not text:
        return None
    match = re.search(r"(?i)(?:HRESULT[: ]*)?(?:0x)?(80070520)\b", text)
    if match:
        return "0x" + match.group(1).lower()
    # pywin32 may expose the signed HRESULT instead of the hex form.
    match = re.search(r"-2147023584\b", text)
    if match:
        return POWERPOINT_COM_ERROR_CODE
    return None


def load_environment(path: Path) -> dict:
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def write_environment(path: Path, report: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def powerpoint_com_blocked(environment: dict) -> bool:
    state = environment.get("powerpoint_com") or {}
    return bool(
        state.get("auto_retry_allowed") is False
        and state.get("retry_policy") == "skip_powerpoint_com"
    )


def reset_powerpoint_com_state(environment: dict) -> dict:
    updated = dict(environment)
    updated.pop("powerpoint_com", None)
    return updated


def record_powerpoint_com_fault(environment_path: Path, failure: dict) -> dict:
    environment = load_environment(environment_path)
    environment["skill_version"] = SKILL_VERSION
    environment["environment_file"] = str(environment_path.resolve())
    environment["powerpoint_com"] = {
        "status": "temporary_fault_recorded",
        "display_status": "临时故障",
        "fault_class": "transient_runtime_startup_failure",
        "environment_constraint": "permanent_for_current_environment",
        "environment_constraint_display": "当前环境的永久性限制",
        "retry_policy": "skip_powerpoint_com",
        "auto_retry_allowed": False,
        "manual_recovery_required": True,
        "reset_argument": "--reset-powerpoint-com-state",
        "error_code": failure.get("error_code") or POWERPOINT_COM_ERROR_CODE,
        "backend": failure.get("backend"),
        "backend_kind": failure.get("backend_kind"),
        "hint": failure.get("hint") or POWERPOINT_COM_HINT,
        "detail": failure.get("detail") or failure.get("error"),
        "stack": failure.get("stack") or failure.get("detail"),
        "detected_at": datetime.now().astimezone().isoformat(),
    }
    write_environment(environment_path, environment)
    return environment


def backend_pywin32():
    if not powerpoint_installed():
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def powerpoint_installed():
    if platform.system() != "Windows":
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application"):
            return True
    except (ImportError, FileNotFoundError, OSError):
        pass
    for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
        root = os.environ.get(env_name)
        if root:
            for candidate in (
                Path(root) / "Microsoft Office" / "root" / "Office16" / "POWERPNT.EXE",
                Path(root) / "Microsoft Office" / "Office16" / "POWERPNT.EXE",
            ):
                if candidate.is_file():
                    return True
    return False


def backend_powershell():
    return (
        sys.platform.startswith("win")
        and shutil.which("powershell") is not None
        and powerpoint_installed()
    )


def backend_soffice():
    return shutil.which("soffice") is not None or shutil.which("libreoffice") is not None


def artifact_tool_ready(node_executable, node_modules, helper_script):
    if not node_executable or not node_modules:
        return False
    node = Path(node_executable)
    modules = Path(node_modules)
    helper = Path(helper_script)
    package = modules / "@oai" / "artifact-tool" / "package.json"
    return node.is_file() and modules.is_dir() and helper.is_file() and package.is_file()


def expected_slide_count(path):
    """Read the expected page count without treating renderer registration as proof."""
    from pptx import Presentation

    return len(Presentation(str(path)).slides)


def _command_error(proc):
    detail = (proc.stderr or proc.stdout or "").strip()
    if not detail:
        detail = f"命令返回状态 {proc.returncode}"
    return detail.splitlines()[-1][:500]


def _render_success(count, expected_count, rendered_files):
    if not rendered_files:
        raise RuntimeError("未导出任何页面图片")
    if count != expected_count:
        raise RuntimeError(f"导出页数不完整：{count}/{expected_count}")
    if len(rendered_files) != expected_count:
        raise RuntimeError(
            f"图片数量与声明页数不一致：{len(rendered_files)} 张，预期 {expected_count} 张"
        )
    empty = [item.name for item in rendered_files if item.stat().st_size == 0]
    if empty:
        raise RuntimeError("存在空图片：" + ", ".join(empty))
    return rendered_files


def verification_status_for(backend_kind: str) -> str:
    if backend_kind not in BACKEND_KIND_VALUES:
        raise ValueError(f"Unknown render backend kind: {backend_kind!r}")
    return VERIFICATION_BY_BACKEND_KIND[backend_kind]


def render_pywin32(path, out_dir, width, height):
    import win32com.client
    out_dir.mkdir(parents=True, exist_ok=True)
    app = None
    pres = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        pres = app.Presentations.Open(str(path), True, False, False)
        count = 0
        for i, slide in enumerate(pres.Slides, 1):
            out = out_dir / f"slide_{i:02d}.png"
            slide.Export(str(out), "PNG", width, height)
            count += 1
        return count
    except Exception as exc:
        detail = traceback.format_exc()
        code = extract_error_code(detail)
        raise RenderBackendError(
            "pywin32(PowerPoint COM)",
            error_code=code,
            hint=POWERPOINT_COM_HINT if code == POWERPOINT_COM_ERROR_CODE else str(exc),
            detail=detail,
            stack=detail,
        ) from exc
    finally:
        if pres is not None:
            pres.Close()
        if app is not None:
            app.Quit()


def render_powershell(path, out_dir, width, height):
    out_dir.mkdir(parents=True, exist_ok=True)
    ps = f"""
$ErrorActionPreference = "Stop"
$pptxPath = {_ps_quote(str(path))}
$outDir = {_ps_quote(str(out_dir))}
$width = {int(width)}
$height = {int(height)}
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$pp = New-Object -ComObject PowerPoint.Application
$pres = $null
try {{
    try {{
        $pres = $pp.Presentations.Open($pptxPath, $true, $false, $false)
    }} catch {{
        $pres = $pp.Presentations.Open($pptxPath)
    }}
    $i = 1
    foreach ($slide in $pres.Slides) {{
        $out = Join-Path $outDir ("slide_{{0:D2}}.png" -f $i)
        $slide.Export($out, "PNG", $width, $height)
        $i++
    }}
    Write-Output ("RENDERED:" + ($i - 1))
}} finally {{
    if ($pres -ne $null) {{ $pres.Close() }}
    $pp.Quit()
}}
"""
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8") as f:
        f.write(ps)
        ps_path = f.name
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_path],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=120,
        )
        out = proc.stdout
        count = 0
        for line in out.splitlines():
            if line.startswith("RENDERED:"):
                count = int(line.split(":", 1)[1])
        if proc.returncode != 0:
            detail = (proc.stderr or out).strip()
            print("PowerShell 渲染失败：")
            print(detail)
            code = extract_error_code(detail)
            hint = (
                POWERPOINT_COM_HINT
                if code == POWERPOINT_COM_ERROR_CODE
                else (detail.splitlines()[-1] if detail else "PowerShell 返回非零状态")
            )
            raise RenderBackendError(
                "PowerShell(PowerPoint COM)",
                error_code=code,
                hint=hint,
                detail=detail,
                stack=detail,
            )
        return count
    finally:
        Path(ps_path).unlink(missing_ok=True)


def render_soffice(path, out_dir, width, height):
    """Render all slides through LibreOffice PDF and pdftoppm when available."""
    out_dir.mkdir(parents=True, exist_ok=True)
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    pdftoppm = shutil.which("pdftoppm")
    with tempfile.TemporaryDirectory(prefix="pro_ppt_render_") as tmp:
        tmp_dir = Path(tmp)
        convert = subprocess.run(
            [exe, "--headless", "--convert-to", "pdf", "--outdir", str(tmp_dir), str(path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )
        pdfs = list(tmp_dir.glob("*.pdf"))
        if not pdfs:
            raise RuntimeError("LibreOffice 未生成 PDF：" + _command_error(convert))

        if pdftoppm:
            dpi = max(72, round(int(width) / 13.333))
            prefix = tmp_dir / "slide"
            raster = subprocess.run(
                [pdftoppm, "-png", "-r", str(dpi), str(pdfs[0]), str(prefix)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=180,
            )
            pngs = sorted(tmp_dir.glob("slide-*.png"))
            if raster.returncode != 0 or not pngs:
                raise RuntimeError("PDF 转 PNG 失败：" + _command_error(raster))
            for index, png in enumerate(pngs, 1):
                shutil.copy2(png, out_dir / f"slide_{index:02d}.png")
            return len(pngs)

        # Render every PDF page with PyMuPDF when pdftoppm is unavailable.
        try:
            import fitz
        except ImportError as exc:
            raise RuntimeError(
                "LibreOffice 已生成 PDF，但缺少 pdftoppm 和 PyMuPDF，无法逐页转图"
            ) from exc

        dpi = max(72, round(int(width) / 13.333))
        document = fitz.open(str(pdfs[0]))
        try:
            for index, page in enumerate(document, 1):
                pixmap = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
                pixmap.save(str(out_dir / f"slide_{index:02d}.png"))
            return len(document)
        finally:
            document.close()


def render_artifact_tool(
    path,
    out_dir,
    width,
    height,
    node_executable,
    node_modules,
    helper_script,
):
    """Render through an explicitly supplied Artifact Tool runtime."""
    out_dir.mkdir(parents=True, exist_ok=True)
    proc = subprocess.run(
        [
            str(node_executable),
            str(helper_script),
            "--input",
            str(path),
            "--output-dir",
            str(out_dir),
            "--node-modules",
            str(node_modules),
            "--scale",
            str(max(0.5, int(width) / 1280)),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=300,
    )
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout).strip()
        print("Artifact Tool 渲染失败：")
        print(detail)
        summary = detail.splitlines()[0] if detail else "Artifact Tool 返回非零状态"
        raise RuntimeError(summary)
    return len(sorted(out_dir.glob("slide_*.png")))


def failure_record(candidate_name, candidate_kind, exc):
    if isinstance(exc, RenderBackendError):
        return {
            "backend": candidate_name,
            "backend_kind": candidate_kind,
            "error_code": exc.error_code,
            "error": exc.hint,
            "hint": exc.hint,
            "detail": exc.detail,
            "stack": exc.stack,
            "attempted": True,
            "skipped_retry": False,
        }
    detail = traceback.format_exc()
    code = extract_error_code(detail) or extract_error_code(str(exc))
    return {
        "backend": candidate_name,
        "backend_kind": candidate_kind,
        "error_code": code,
        "error": str(exc) or exc.__class__.__name__,
        "hint": POWERPOINT_COM_HINT if code == POWERPOINT_COM_ERROR_CODE else "",
        "detail": detail,
        "stack": detail,
        "attempted": True,
        "skipped_retry": False,
    }


def render_with_fallback(
    out_dir,
    candidates,
    expected_count,
    *,
    initial_failures=None,
    on_failure=None,
):
    """Try each backend family once; PowerPoint COM faults short-circuit retries."""
    failures = list(initial_failures or [])
    blocked_families = set()
    for candidate_name, candidate_kind, render in candidates:
        if candidate_kind in blocked_families:
            skipped_failure = {
                "backend": candidate_name,
                "backend_kind": candidate_kind,
                "error_code": POWERPOINT_COM_ERROR_CODE,
                "error": "同一环境已记录 PowerPoint COM 永久性限制，跳过重复尝试",
                "hint": POWERPOINT_COM_HINT,
                "detail": "PowerPoint COM retry suppressed after 0x80070520 in this environment.",
                "stack": "",
                "attempted": False,
                "skipped_retry": True,
            }
            failures.append(skipped_failure)
            if on_failure:
                on_failure(skipped_failure)
            continue
        with tempfile.TemporaryDirectory(prefix="pro_ppt_attempt_") as tmp:
            attempt_dir = Path(tmp)
            try:
                candidate_count = int(render(attempt_dir))
                rendered_files = sorted(attempt_dir.glob("slide_*.png"))
                _render_success(candidate_count, expected_count, rendered_files)
            except Exception as exc:
                failure = failure_record(candidate_name, candidate_kind, exc)
                failures.append(failure)
                if on_failure:
                    on_failure(failure)
                if (
                    failure.get("error_code") == POWERPOINT_COM_ERROR_CODE
                    and candidate_kind == "native_powerpoint"
                ):
                    blocked_families.add("native_powerpoint")
                continue

            out_dir.mkdir(parents=True, exist_ok=True)
            for stale in out_dir.glob("slide_*.png"):
                stale.unlink()
            for index, source in enumerate(rendered_files, 1):
                shutil.copy2(source, out_dir / f"slide_{index:02d}.png")
            return candidate_name, candidate_kind, len(rendered_files), failures

    return None, None, 0, failures


def clear_previous_preview(out_dir):
    """Remove only prior generated preview files, never unrelated directory content."""
    out_dir.mkdir(parents=True, exist_ok=True)
    for stale in out_dir.glob("slide_*.png"):
        stale.unlink()
    (out_dir / "contact_sheet.png").unlink(missing_ok=True)


def _ps_quote(s):
    return "'" + s.replace("'", "''") + "'"


def make_contact_sheet(pngs, out_dir, thumb_w=360, cols=4):
    if not HAS_PIL:
        print("提示：未安装 Pillow，跳过缩略图总览。可执行 python -m pip install Pillow")
        return None
    thumbs = []
    for p in pngs:
        im = Image.open(p).convert("RGB")
        ratio = thumb_w / im.width
        im = im.resize((thumb_w, int(im.height * ratio)))
        thumbs.append(im)
    if not thumbs:
        return None
    cols = min(cols, len(thumbs))
    rows = (len(thumbs) + cols - 1) // cols
    h = max(im.height for im in thumbs)
    pad = 16
    canvas = Image.new("RGB", (cols * thumb_w + (cols + 1) * pad, rows * h + (rows + 1) * pad), (235, 235, 235))
    for idx, im in enumerate(thumbs):
        r, c = divmod(idx, cols)
        canvas.paste(im, (pad + c * (thumb_w + pad), pad + r * (h + pad)))
    out = out_dir / "contact_sheet.png"
    canvas.save(out)
    return out


def export_transition_gif(pngs, output, *, fps=4, hold_seconds=0.8):
    """Create a deterministic transition preview without external codecs."""

    if not HAS_PIL:
        raise RenderBackendError(
            "Pillow GIF",
            error_code=None,
            hint="GIF export requires Pillow.",
            detail="Pillow could not be imported.",
        )
    frames = []
    images = [Image.open(path).convert("RGB") for path in pngs]
    try:
        width = min(image.width for image in images)
        height = min(image.height for image in images)
        images = [
            image.resize((width, height), Image.LANCZOS)
            for image in images
        ]
        hold_frames = max(1, int(fps * hold_seconds))
        blend_frames = max(2, int(fps * 0.25))
        for index, image in enumerate(images):
            for _ in range(hold_frames):
                frames.append(image.copy())
            if index + 1 < len(images):
                next_image = images[index + 1]
                for step in range(1, blend_frames + 1):
                    frames.append(
                        Image.blend(image, next_image, step / (blend_frames + 1))
                    )
        duration_ms = max(40, int(1000 / fps))
        frames[0].save(
            output,
            save_all=True,
            append_images=frames[1:],
            duration=duration_ms,
            loop=0,
            optimize=False,
        )
    finally:
        for image in images:
            image.close()
    return output


def export_ffmpeg_video(pngs, output, *, fps=2):
    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        raise RenderBackendError(
            "ffmpeg",
            error_code=None,
            hint="ffmpeg is not available for fallback video export.",
            detail="No ffmpeg command found in PATH.",
        )
    pattern = pngs[0].with_name("slide_%02d.png")
    command = [
        ffmpeg,
        "-y",
        "-framerate",
        str(fps),
        "-i",
        str(pattern),
        "-c:v",
        "libx264",
        "-pix_fmt",
        "yuv420p",
        "-movflags",
        "+faststart",
        str(output),
    ]
    process = subprocess.run(command, capture_output=True, text=True)
    if process.returncode != 0 or not output.exists():
        raise RenderBackendError(
            "ffmpeg",
            error_code=None,
            hint="ffmpeg fallback video export failed.",
            detail=process.stderr[-4000:],
        )
    return output


def export_powerpoint_video(path, output, *, duration_seconds=5):
    """Export native transition/object animation through PowerPoint COM."""

    if not backend_powershell():
        raise RenderBackendError(
            "PowerShell(PowerPoint video)",
            error_code=None,
            hint="PowerPoint COM is unavailable for native video export.",
            detail="PowerPoint COM backend was not detected.",
        )
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    script = f"""
$ErrorActionPreference = 'Stop'
$app = $null
$presentation = $null
try {{
    $app = New-Object -ComObject PowerPoint.Application
    $presentation = $app.Presentations.Open('{str(path).replace("'", "''")}', $false, $false, $false)
    $presentation.CreateVideo('{str(output).replace("'", "''")}', $true, {int(duration_seconds)}, 1080, 30, 90)
    $deadline = (Get-Date).AddMinutes(10)
    while ($presentation.CreateVideoStatus -eq 1 -and (Get-Date) -lt $deadline) {{
        Start-Sleep -Milliseconds 500
    }}
    if ($presentation.CreateVideoStatus -ne 3) {{
        throw "PowerPoint video export failed with status $($presentation.CreateVideoStatus)"
    }}
}}
finally {{
    if ($presentation -ne $null) {{ $presentation.Close() }}
    if ($app -ne $null) {{ $app.Quit() }}
}}
"""
    with tempfile.TemporaryDirectory(prefix="ppt_video_") as directory:
        script_path = Path(directory) / "export_video.ps1"
        script_path.write_text(script, encoding="utf-8-sig")
        process = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
            ],
            capture_output=True,
            text=True,
            timeout=660,
        )
    if process.returncode != 0 or not output.exists():
        raise RenderBackendError(
            "PowerShell(PowerPoint video)",
            error_code=extract_error_code(process.stderr),
            hint="PowerPoint 原生视频导出失败。",
            detail=(process.stderr or process.stdout)[-6000:],
        )
    return output


def export_dynamic_preview(path, pngs, out_dir, *, want_video=False, want_gif=False):
    results = {}
    if want_gif:
        try:
            gif = export_transition_gif(pngs, out_dir / "preview.gif")
            results["gif"] = {
                "status": "rendered",
                "path": str(gif),
                "note": "Static slide frames with crossfade; object timing is illustrative.",
            }
        except Exception as error:
            results["gif"] = {"status": "failed", "error": str(error)}
    if want_video:
        video_path = out_dir / "preview.mp4"
        try:
            video = export_powerpoint_video(path, video_path)
            results["video"] = {
                "status": "rendered_native_powerpoint",
                "path": str(video),
                "note": "Includes PowerPoint transitions and object animation.",
            }
        except Exception as native_error:
            try:
                video = export_ffmpeg_video(pngs, video_path)
                results["video"] = {
                    "status": "rendered_static_frames",
                    "path": str(video),
                    "note": "ffmpeg fallback shows slide timing; native object effects are not preserved.",
                }
            except Exception as fallback_error:
                results["video"] = {
                    "status": "failed",
                    "native_error": str(native_error),
                    "fallback_error": str(fallback_error),
                }
    return results


def main():
    ap = argparse.ArgumentParser(description="pro-ppt-studio PPTX 逐页渲染与缩略图总览")
    ap.add_argument("file", help="PPTX 文件路径")
    ap.add_argument("out_dir", help="输出目录")
    ap.add_argument("--width", type=int, default=1920)
    ap.add_argument("--height", type=int, default=1080)
    ap.add_argument("--no-contact-sheet", action="store_true")
    ap.add_argument(
        "--artifact-node",
        default=os.environ.get("CODEX_ARTIFACT_NODE"),
        help="Optional Node executable for Artifact Tool fallback",
    )
    ap.add_argument(
        "--artifact-node-modules",
        default=os.environ.get("CODEX_ARTIFACT_NODE_MODULES"),
        help="Optional node_modules directory containing @oai/artifact-tool",
    )
    ap.add_argument(
        "--artifact-helper",
        default=str(Path(__file__).with_name("artifact_tool_render.mjs")),
        help="Path to artifact_tool_render.mjs",
    )
    ap.add_argument(
        "--environment",
        default=os.environ.get("CODEX_PPT_ENVIRONMENT"),
        help="environment.json path used to persist the PowerPoint COM circuit breaker",
    )
    ap.add_argument(
        "--reset-powerpoint-com-state",
        action="store_true",
        help="Explicitly clear a persisted PowerPoint COM limitation after the environment changes",
    )
    ap.add_argument(
        "--video",
        action="store_true",
        help="Export preview.mp4; uses native PowerPoint timing first, then ffmpeg fallback",
    )
    ap.add_argument(
        "--gif",
        action="store_true",
        help="Export a slide-transition GIF with crossfade",
    )
    args = ap.parse_args()

    path = Path(args.file)
    out_dir = Path(args.out_dir)
    if not path.exists():
        print(f"错误：文件不存在 {path}")
        sys.exit(2)
    try:
        expected_count = expected_slide_count(path)
    except Exception as exc:
        print(f"错误：无法读取 PPTX 页数：{exc}")
        sys.exit(2)
    if expected_count <= 0:
        print("错误：PPTX 不包含任何页面")
        sys.exit(2)
    clear_previous_preview(out_dir)
    environment_path = (
        Path(args.environment)
        if args.environment
        else out_dir.parent / "environment.json"
    )
    environment = load_environment(environment_path)
    if args.reset_powerpoint_com_state:
        environment = reset_powerpoint_com_state(environment)
        environment["skill_version"] = SKILL_VERSION
        write_environment(environment_path, environment)
    powerpoint_state = environment.get("powerpoint_com")
    powerpoint_blocked = powerpoint_com_blocked(environment)
    initial_failures = []
    if powerpoint_blocked and powerpoint_state:
        initial_failures.append(
            {
                "backend": powerpoint_state.get("backend")
                or "PowerShell(PowerPoint COM)",
                "backend_kind": powerpoint_state.get("backend_kind")
                or "native_powerpoint",
                "error_code": powerpoint_state.get("error_code")
                or POWERPOINT_COM_ERROR_CODE,
                "error": powerpoint_state.get("hint") or POWERPOINT_COM_HINT,
                "hint": powerpoint_state.get("hint") or POWERPOINT_COM_HINT,
                "detail": powerpoint_state.get("detail") or "",
                "stack": powerpoint_state.get("stack") or "",
                "attempted": False,
                "skipped_retry": True,
                "source": "environment.json",
            }
        )

    candidates = []
    skipped = []
    if powerpoint_blocked:
        skipped.extend(
            [
                {
                    "backend": "pywin32(PowerPoint COM)",
                    "backend_kind": "native_powerpoint",
                    "reason": (
                        f"environment.json retry_policy=skip_powerpoint_com; "
                        f"error_code={POWERPOINT_COM_ERROR_CODE}"
                    ),
                },
                {
                    "backend": "PowerShell(PowerPoint COM)",
                    "backend_kind": "native_powerpoint",
                    "reason": (
                        f"environment.json retry_policy=skip_powerpoint_com; "
                        f"error_code={POWERPOINT_COM_ERROR_CODE}"
                    ),
                },
            ]
        )
    elif backend_pywin32():
        candidates.append(
            (
                "pywin32(PowerPoint COM)",
                "native_powerpoint",
                lambda target: render_pywin32(path, target, args.width, args.height),
            )
        )
    else:
        skipped.append(
            {
                "backend": "pywin32(PowerPoint COM)",
                "backend_kind": "native_powerpoint",
                "reason": "PowerPoint 未安装或 pywin32 不可用",
            }
        )
    if not powerpoint_blocked and backend_powershell():
        candidates.append(
            (
                "PowerShell(PowerPoint COM)",
                "native_powerpoint",
                lambda target: render_powershell(path, target, args.width, args.height),
            )
        )
    elif not powerpoint_blocked:
        skipped.append(
            {
                "backend": "PowerShell(PowerPoint COM)",
                "backend_kind": "native_powerpoint",
                "reason": "PowerPoint 未安装或 PowerShell 不可用",
            }
        )
    if backend_soffice():
        candidates.append(
            (
                "LibreOffice(PDF→PNG)",
                "compatible_libreoffice",
                lambda target: render_soffice(path, target, args.width, args.height),
            )
        )
    else:
        skipped.append(
            {
                "backend": "LibreOffice(PDF→PNG)",
                "backend_kind": "compatible_libreoffice",
                "reason": "未找到 soffice 或 libreoffice 命令",
            }
        )
    if artifact_tool_ready(
        args.artifact_node,
        args.artifact_node_modules,
        args.artifact_helper,
    ):
        candidates.append(
            (
                "Artifact Tool",
                "alternative_artifact_tool",
                lambda target: render_artifact_tool(
                    path,
                    target,
                    args.width,
                    args.height,
                    args.artifact_node,
                    args.artifact_node_modules,
                    args.artifact_helper,
                ),
            )
        )
    else:
        skipped.append(
            {
                "backend": "Artifact Tool",
                "backend_kind": "alternative_artifact_tool",
                "reason": "未显式提供可用的 Node、@oai/artifact-tool node_modules 和 helper",
            }
        )

    if not candidates:
        print("错误：未找到可用渲染后端（PowerPoint、LibreOffice 和 Artifact Tool 均不可用）。")
        out_dir.mkdir(parents=True, exist_ok=True)
        (out_dir / "render_manifest.json").write_text(
            json.dumps(
                {
                    "skill_version": SKILL_VERSION,
                    "file": str(path),
                    "environment_file": str(environment_path),
                    "backend": None,
                    "backend_kind": None,
                    "render_backend_kind": None,
                    "native_powerpoint_verified": False,
                    "verification_status": "unverified",
                    "expected_count": expected_count,
                    "rendered_count": 0,
                    "pngs": [],
                    "contact_sheet": None,
                    "backend_failures": initial_failures,
                    "skipped_backends": skipped,
                    "powerpoint_com_state": powerpoint_state,
                    "powerpoint_com_attempted": False,
                    "visual_quality_status": "not_assessed",
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print("无法完成第四层渲染检查。")
        sys.exit(3)

    def persist_fault(failure):
        nonlocal environment, powerpoint_state
        if (
            failure.get("error_code") == POWERPOINT_COM_ERROR_CODE
            and failure.get("attempted")
            and not failure.get("skipped_retry")
        ):
            environment = record_powerpoint_com_fault(environment_path, failure)
            powerpoint_state = environment.get("powerpoint_com")

    backend, backend_kind, count, failures = render_with_fallback(
        out_dir,
        candidates,
        expected_count,
        initial_failures=initial_failures,
        on_failure=persist_fault,
    )
    pngs = sorted(out_dir.glob("slide_*.png"))
    if not backend or count != expected_count or len(pngs) != expected_count:
        detail = "; ".join(f"{item['backend']}: {item['error']}" for item in failures)
        out_dir.mkdir(parents=True, exist_ok=True)
        (out_dir / "render_manifest.json").write_text(
            json.dumps(
                {
                    "skill_version": SKILL_VERSION,
                    "file": str(path),
                    "environment_file": str(environment_path),
                    "backend": None,
                    "backend_kind": None,
                    "render_backend_kind": None,
                    "native_powerpoint_verified": False,
                    "verification_status": "unverified",
                    "expected_count": expected_count,
                    "rendered_count": 0,
                    "pngs": [],
                    "contact_sheet": None,
                    "backend_failures": failures,
                    "skipped_backends": skipped,
                    "powerpoint_com_state": powerpoint_state,
                    "powerpoint_com_attempted": any(
                        item.get("backend_kind") == "native_powerpoint"
                        and item.get("attempted")
                        for item in failures
                    ),
                    "visual_quality_status": "not_assessed",
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"错误：所有渲染后端均未生成完整页面集。{detail}")
        sys.exit(4)

    contact = None
    if not args.no_contact_sheet and pngs:
        contact = make_contact_sheet(pngs, out_dir)
    dynamic_exports = export_dynamic_preview(
        path,
        pngs,
        out_dir,
        want_video=args.video,
        want_gif=args.gif,
    )

    manifest = {
        "skill_version": SKILL_VERSION,
        "file": str(path),
        "environment_file": str(environment_path),
        "backend": backend,
        "backend_kind": backend_kind,
        "render_backend_kind": backend_kind,
        "native_powerpoint_verified": backend_kind == "native_powerpoint",
        "verification_status": verification_status_for(backend_kind),
        "width": args.width,
        "height": args.height,
        "expected_count": expected_count,
        "rendered_count": count,
        "pngs": [p.name for p in pngs],
        "contact_sheet": str(contact) if contact else None,
        "backend_failures": failures,
        "skipped_backends": skipped,
        "powerpoint_com_state": powerpoint_state,
        "powerpoint_com_attempted": any(
            item.get("backend_kind") == "native_powerpoint"
            and item.get("attempted")
            for item in failures
        ),
        "visual_quality_status": "not_assessed",
        "dynamic_exports": dynamic_exports,
    }
    (out_dir / "render_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(f"渲染后端：{backend}")
    print(f"验证状态：{manifest['verification_status']}")
    print(f"导出图片：{len(pngs)} 张 -> {out_dir}")
    if contact:
        print(f"缩略图总览：{contact}")
    if backend.startswith("LibreOffice"):
        print("提示：请核对导出图片数量与原 PPT 页数；LibreOffice 路径可能受本机转换器限制。")


if __name__ == "__main__":
    main()
