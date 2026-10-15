# PowerPoint Template Modifier

This project is a local Streamlit application for inspecting and editing
supported content in a user-provided PowerPoint `.pptx` presentation. It lets
you revise text and selected formatting through a browser form, then download
a new presentation. It does not use AI, cloud services, external APIs, or a
database.

## End-to-end flow

1. **Start the app.** Streamlit launches the interface from `app.py`.
2. **Upload a presentation.** The file picker accepts `.pptx` files. The app
   checks that a file was provided and that its name ends with `.pptx`.
3. **Inspect its slides.** The uploaded bytes are passed to
   `inspector.inspect_presentation`. The inspector opens the deck in memory,
   visits slides in order, and records supported text-bearing shapes and
   ordinary shapes. It captures each shape's ID, name, slide, position and
   dimensions, plus the properties that can be edited. Picture shapes and
   other unsupported shape types are omitted.
4. **Build the editor.** The app uses the inspection result as an internal
   template. It displays a form grouped by slide and shape, prefilled from the
   current content. Display names use existing text or a readable shape type
   where possible, rather than exposing technical shape names as headings.
5. **Enter changes.** For supported fields, the form provides text areas,
   font-size inputs, color pickers, and bold/italic checkboxes. Form values are
   validated as the user submits them.
6. **Generate a new deck.** On submission, the modifier reopens the original
   uploaded bytes, locates shapes by their slide and shape ID (falling back to
   shape name), and applies the submitted properties.
7. **Optional native editing.** If Windows PowerPoint COM is available and a
   slide transition is selected, the app applies it to the generated copy.
   Otherwise, generation uses only `python-pptx`.
8. **Download.** The modified presentation is offered as
   `<original-name>_modified.pptx`. The uploaded original is never overwritten.

## Editing architecture

`python-pptx` remains the default, cross-platform engine for common edits.
PowerPoint COM is a small, optional Windows extension used only for selected
native features that are not conveniently handled by the existing
`python-pptx` layer. It does not replace the common editing backend or expose
the full PowerPoint object model.

```text
                 Streamlit
                     |
              Inspection / UI
                     |
          -----------------------
          |                     |
      python-pptx          PowerPoint COM
       (default)            (optional)
          |                     |
      Common edits         Advanced edits
          |                     |
          -----------------------
                    |
               Final PPTX
```

The current COM-backed feature applies a Fade, Wipe, or Push slide transition
with a configurable duration to every slide. The app first applies standard
edits with `python-pptx`, then gives PowerPoint COM a temporary copy of that
generated presentation. The original upload is not modified.

## Supported editing

- Text on supported text-frame shapes.
- Text font size, font color, bold, and italic.
- Fill and border color for supported ordinary shapes.
- Hex colors in three- or six-digit form (for example `#09C` or `#0099CC`).
- Font-size values in the form are limited to 6 through 200 points.

Not every PowerPoint object exposes editable fields through this MVP. Elements
without supported properties do not get form controls; pictures and unsupported
objects are skipped. The inspector reads color values when they are explicit
RGB colors; theme colors may not be represented as a hex value.

## Text formatting behavior and limitations

Replacing text clears the existing text frame and writes the replacement into
one paragraph/run. If the original text used multiple runs with different
formatting, that run-level formatting is not retained for the replacement.
When a text field is edited, the selected font size, color, bold, or italic
setting is applied across the resulting runs. This favors predictable edits
over complex preservation of PowerPoint's full XML formatting model.

The current color/font controls are based on properties the inspector can
read; they do not provide controls for every possible PowerPoint style such as
theme colors, gradients, transparency, font family, or paragraph layout.

## Requirements

- Python 3.13 or newer (as declared in `pyproject.toml`).
- `uv` for the recommended setup.
- Dependencies: Streamlit and `python-pptx`, installed from project metadata.
- A browser and a `.pptx` file that `python-pptx` can open.
- Optional advanced features: Windows, the Microsoft PowerPoint desktop
  application, and the optional `pywin32` dependency. COM automation is
  detected at runtime; without it, the normal `python-pptx` workflow remains
  available.

## Install and run

From this folder:

```powershell
uv sync
uv run streamlit run app.py
```

To enable the optional PowerPoint COM features on Windows, install the extra:

```powershell
uv sync --extra windows-com
```

With `pip`, install the optional project extra on Windows:

```powershell
python -m pip install ".[windows-com]"
```

PowerPoint COM requires the desktop PowerPoint application to be installed and
registered for the current Windows environment. Administrator privileges are
not required. On non-Windows systems, or when the COM dependency or application
is unavailable, the advanced controls are omitted and standard editing
continues to work. The extension intentionally covers only selected native
features; it is not a replacement for PowerPoint or a full object-model editor.

Streamlit prints a local URL in the terminal. Open that URL in a browser,
upload a presentation, edit the displayed fields, select **Generate
PowerPoint**, and use **Download Modified PPT** when generation succeeds.

The installed command-line entry point can also start the app:

```powershell
uv run pptautomation
```

To install dependencies without `uv`, use the active Python environment:

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Project structure

| Path | Responsibility |
|---|---|
| `app.py` | Top-level Streamlit launcher. |
| `src/pptautomation/app.py` | UI, upload handling, widget rendering, validation, and download. |
| `src/pptautomation/inspector.py` | Reads the presentation and extracts slide/shape metadata and supported properties. |
| `src/pptautomation/json_generator.py` | Defines editable field labels/order and builds default values from inspected content. |
| `src/pptautomation/validator.py` | Checks the `.pptx` extension, color values, and numeric values. |
| `src/pptautomation/modifier.py` | Applies submitted values and serializes the new `.pptx`. |
| `src/pptautomation/com_modifier.py` | Optional Windows-only PowerPoint COM detection and slide-transition editing. |
| `pyproject.toml` / `uv.lock` | Project metadata, Python requirement, dependencies, and locked dependency versions. |

## Data handling

The presentation is read from its uploaded bytes, and the modifier serializes
the result through in-memory buffers. The project has no database or external
service dependency. Streamlit runs locally by default; do not expose the
development server to an untrusted network unless you have intentionally
configured access controls.

## Troubleshooting

- **The upload is rejected:** Use a valid `.pptx` file. Legacy `.ppt` files are
  not accepted.
- **The file is accepted but inspection fails:** The deck may be corrupt or
  contain content unsupported by `python-pptx`. Try opening and resaving it
  from PowerPoint, then upload the saved `.pptx`.
- **A shape is missing from the editor:** The inspector only exposes supported
  text frames and shapes; unsupported objects are intentionally skipped.
- **The generated text looks differently formatted:** Replacing text
  normalizes it to one run, as described above. Existing per-run formatting is
  not preserved when replacing the text.
- **Generation reports an invalid color/size:** Choose a valid hex color and
  keep font size in the displayed 6-200 point range.
