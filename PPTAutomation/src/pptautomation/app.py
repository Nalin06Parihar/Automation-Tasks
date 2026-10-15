from __future__ import annotations

import json

import streamlit as st

from pptautomation.com_modifier import apply_slide_transitions, get_com_unavailable_reason
from pptautomation.inspector import inspect_presentation
from pptautomation.json_generator import FIELD_LABELS, supported_editable_fields
from pptautomation.modifier import apply_changes_to_presentation
from pptautomation.validator import validate_numeric, validate_uploaded_pptx, validate_color


def _label_for(field_name: str) -> str:
    return FIELD_LABELS.get(field_name, field_name.replace("_", " ").title())


def _build_widget_key(slide_number: int, shape_name: str, field_name: str) -> str:
    return f"slide_{slide_number}_{shape_name}_{field_name}"


def _render_shape_form(slide_number: int, shape: dict) -> None:
    shape_name = shape.get("name", f"Shape {shape.get('id')}")
    existing_properties = shape.get("properties", {})
    editable_fields = supported_editable_fields(shape)

    if not editable_fields:
        return

    st.markdown(f"#### {shape.get('display_name', shape_name)}")
    for field_name in editable_fields:
        values = existing_properties.get(field_name)
        key = _build_widget_key(slide_number, shape_name, field_name)

        if field_name == "text":
            st.text_area(_label_for(field_name), value=str(values) if values is not None else "", key=key)
        elif field_name == "font_size":
            default_value = float(values) if values is not None else 12.0
            st.number_input(_label_for(field_name), min_value=6.0, max_value=200.0, value=default_value, step=1.0, key=key)
        elif field_name in {"font_color", "fill_color", "line_color"}:
            default_value = values if values is not None else "#000000"
            st.color_picker(_label_for(field_name), value=str(default_value), key=key)
        elif field_name in {"bold", "italic"}:
            st.checkbox(_label_for(field_name), value=bool(values), key=key)
        else:
            continue



def _collect_form_data(template: dict) -> dict:
    input_data: dict[str, dict[str, dict[str, object]]] = {}

    for slide in template.get("slides", []):
        slide_number = int(slide.get("slide", 0))
        slide_key = f"slide_{slide_number}"
        input_data[slide_key] = {}

        for shape in slide.get("shapes", []):
            shape_name = shape.get("name", f"Shape {shape.get('id')}")
            editable_fields = supported_editable_fields(shape)
            if not editable_fields:
                continue

            shape_updates: dict[str, object] = {}
            for field_name in editable_fields:
                widget_key = _build_widget_key(slide_number, shape_name, field_name)
                value = st.session_state.get(widget_key)
                if value is None:
                    continue

                if field_name in {"font_size"}:
                    try:
                        validate_numeric(value, "Font Size")
                    except ValueError as exc:
                        st.warning(str(exc))
                        continue

                if field_name in {"font_color", "fill_color", "line_color"}:
                    try:
                        value = validate_color(str(value))
                    except ValueError as exc:
                        st.warning(str(exc))
                        continue

                shape_updates[field_name] = value

            if shape_updates:
                input_data[slide_key][shape_name] = shape_updates

    return input_data


def _render_slide_editor(template: dict) -> None:
    for slide in template.get("slides", []):
        slide_number = int(slide.get("slide", 0))
        with st.expander(f"Slide {slide_number}", expanded=True):
            shape_entries = slide.get("shapes", [])
            if not shape_entries:
                st.caption("No supported editable shapes were found on this slide.")
                continue

            for shape in shape_entries:
                if supported_editable_fields(shape):
                    _render_shape_form(slide_number, shape)


def main() -> None:
    st.set_page_config(page_title="PowerPoint Template Modifier", layout="wide")
    st.title("PowerPoint Template Modifier")

    uploaded_file = st.file_uploader("Upload PowerPoint Template", type=["pptx"])
    if uploaded_file is None:
        return

    st.session_state.pop("output_bytes", None)
    st.session_state.pop("output_name", None)

    try:
        validate_uploaded_pptx(uploaded_file, uploaded_file.name)
        file_bytes = uploaded_file.getvalue()
        template = inspect_presentation(file_bytes, uploaded_file.name)
        template_json = json.dumps(template, indent=2)
        st.session_state["template"] = template
        st.session_state["template_json"] = template_json
    except ValueError as exc:
        st.error(str(exc))
        return
    except Exception:
        st.error("Could not inspect this PowerPoint file. The upload may be invalid or corrupted.")
        return

    slide_count = len(template.get("slides", []))
    st.success(f"Template successfully inspected. {slide_count} slide(s) detected.")

    if "com_unavailable_reason" not in st.session_state:
        st.session_state["com_unavailable_reason"] = get_com_unavailable_reason()
    com_unavailable_reason = st.session_state["com_unavailable_reason"]

    with st.form("ppt_modifier_form"):
        _render_slide_editor(template)
        transition_effect = "No change"
        transition_duration = 1.0
        if com_unavailable_reason is None:
            with st.expander("Advanced PowerPoint Features (Windows)"):
                transition_effect = st.selectbox(
                    "Slide transition for every slide",
                    ["No change", "Fade", "Wipe", "Push"],
                )
                transition_duration = st.number_input(
                    "Transition duration (seconds)",
                    min_value=0.1,
                    max_value=10.0,
                    value=1.0,
                    step=0.1,
                )
        else:
            st.caption(f"Advanced PowerPoint features are unavailable: {com_unavailable_reason}")
        submitted = st.form_submit_button("Generate PowerPoint")

    if submitted:
        input_data = _collect_form_data(template)
        try:
            output_bytes = apply_changes_to_presentation(file_bytes, template, input_data)
            if transition_effect != "No change":
                output_bytes = apply_slide_transitions(
                    output_bytes,
                    transition_effect,
                    transition_duration,
                )
            st.session_state["output_bytes"] = output_bytes
            st.session_state["output_name"] = uploaded_file.name.replace(".pptx", "") + "_modified.pptx"
            st.success("PowerPoint generated successfully.")
        except (ValueError, RuntimeError) as exc:
            st.error(str(exc))
            return
        except Exception:
            st.error("The PowerPoint could not be generated. Please review the selected values and try again.")
            return

    if "output_bytes" in st.session_state:
        st.download_button(
            label="Download Modified PPT",
            data=st.session_state["output_bytes"],
            file_name=st.session_state.get("output_name", "modified_template.pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )


if __name__ == "__main__":
    main()
