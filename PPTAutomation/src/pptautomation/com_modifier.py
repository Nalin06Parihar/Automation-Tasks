from __future__ import annotations

import importlib
import math
import sys
import tempfile
from pathlib import Path
from typing import Any


# Values from PowerPoint's PpEntryEffect and PpSaveAsFileType enums.
_TRANSITION_CONSTANTS = {
    "Fade": 1793,
    "Push": 3853,
    "Wipe": 2817,
}
_MIN_TRANSITION_DURATION = 0.1
_MAX_TRANSITION_DURATION = 10.0
_PPTX_FILE_FORMAT = 24


def _load_com_api() -> tuple[Any, Any]:
    try:
        pythoncom = importlib.import_module("pythoncom")
        client = importlib.import_module("win32com.client")
    except ImportError as exc:
        raise RuntimeError(
            "PowerPoint COM features require the optional pywin32 dependency. "
            "Install it with `uv sync --extra windows-com` on Windows."
        ) from exc
    return pythoncom, client


def _cleanup_com(
    presentation: Any,
    application: Any,
    pythoncom: Any,
    initialized: bool,
) -> list[Exception]:
    errors: list[Exception] = []
    if presentation is not None:
        try:
            presentation.Close()
        except Exception as exc:
            errors.append(exc)
    if application is not None:
        try:
            application.Quit()
        except Exception as exc:
            errors.append(exc)
    if initialized:
        try:
            pythoncom.CoUninitialize()
        except Exception as exc:
            errors.append(exc)
    return errors


def get_com_unavailable_reason() -> str | None:
    if sys.platform != "win32":
        return "Advanced PowerPoint features require Windows and desktop PowerPoint."

    try:
        pythoncom, client = _load_com_api()
    except RuntimeError as exc:
        return str(exc)

    initialized = False
    application = None
    reason = None
    try:
        pythoncom.CoInitialize()
        initialized = True
        application = client.DispatchEx("PowerPoint.Application")
        getattr(application, "Version")
    except Exception as exc:
        reason = f"PowerPoint COM automation is unavailable: {exc}"

    cleanup_errors = _cleanup_com(None, application, pythoncom, initialized)
    if cleanup_errors:
        cleanup_detail = "; ".join(str(error) for error in cleanup_errors)
        if reason is None:
            reason = f"PowerPoint COM cleanup failed: {cleanup_detail}"
        else:
            reason = f"{reason} PowerPoint cleanup also failed: {cleanup_detail}"
    return reason


def apply_slide_transitions(
    presentation_bytes: bytes,
    effect: str,
    duration: float,
) -> bytes:
    if sys.platform != "win32":
        raise RuntimeError("PowerPoint COM editing is only available on Windows.")
    if effect not in _TRANSITION_CONSTANTS:
        raise ValueError(f"Unsupported slide transition: {effect}")

    try:
        transition_duration = float(duration)
    except (TypeError, ValueError) as exc:
        raise ValueError("Transition duration must be a number between 0.1 and 10 seconds.") from exc
    if not math.isfinite(transition_duration) or not (
        _MIN_TRANSITION_DURATION <= transition_duration <= _MAX_TRANSITION_DURATION
    ):
        raise ValueError("Transition duration must be between 0.1 and 10 seconds.")

    pythoncom, client = _load_com_api()
    initialized = False
    application = None
    presentation = None
    output_bytes = None
    operation_error = None

    try:
        with tempfile.TemporaryDirectory(prefix="pptautomation-") as temp_dir:
            try:
                pythoncom.CoInitialize()
                initialized = True
                application = client.DispatchEx("PowerPoint.Application")
                effect_constant = _TRANSITION_CONSTANTS[effect]

                input_path = Path(temp_dir) / "working_copy.pptx"
                output_path = Path(temp_dir) / "modified_copy.pptx"
                input_path.write_bytes(presentation_bytes)

                presentation = application.Presentations.Open(
                    str(input_path),
                    ReadOnly=False,
                    Untitled=False,
                    WithWindow=False,
                )
                for slide_index in range(1, presentation.Slides.Count + 1):
                    transition = presentation.Slides(slide_index).SlideShowTransition
                    transition.EntryEffect = effect_constant
                    transition.Duration = transition_duration

                presentation.SaveAs(str(output_path), _PPTX_FILE_FORMAT)
                output_bytes = output_path.read_bytes()
            except Exception as exc:
                operation_error = exc
            finally:
                cleanup_errors = _cleanup_com(presentation, application, pythoncom, initialized)
                presentation = None
                application = None
    except Exception as exc:
        if operation_error is None:
            operation_error = exc
        cleanup_errors = []

    if operation_error is not None:
        message = f"PowerPoint COM transition editing failed: {operation_error}"
        if cleanup_errors:
            message += " PowerPoint cleanup also failed: " + "; ".join(
                str(error) for error in cleanup_errors
            )
        raise RuntimeError(message) from operation_error
    if cleanup_errors:
        raise RuntimeError(
            "PowerPoint COM cleanup failed: "
            + "; ".join(str(error) for error in cleanup_errors)
        ) from cleanup_errors[0]
    if output_bytes is None:
        raise RuntimeError("PowerPoint COM did not produce a modified presentation.")
    return output_bytes
