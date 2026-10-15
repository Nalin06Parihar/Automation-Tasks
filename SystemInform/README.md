# SystemInform

SystemInform is a local desktop utility that collects system details from the
machine on which it runs and displays them in a Tkinter window. The Python
application has separate Windows and Linux collectors. This folder also
contains a standalone Windows-only PowerShell script for a smaller,
console-based report.

## What the Python application reports

- **System:** host name, operating system, version/build, uptime, architecture,
  manufacturer, and model where available.
- **Hardware:** CPU name and core counts, total/used/available RAM and swap or
  page-file space where available, GPU, and BIOS version.
- **Storage:** devices or mounted filesystems with total, used, and free space
  and usage percentage; Windows volume labels are included where available.
- **Installed software:** names and versions discoverable from the local
  Windows uninstall registry or a supported Linux package manager.
- **Network:** selected adapter, IPv4 address, and MAC address where available.

Some fields depend on platform tools, hardware interfaces, permissions, and
their availability. A field the collector cannot resolve is displayed as
`Unavailable`; installed-software results are a local inventory, not a
complete audit of every application or component.

## Complete application flow

1. `main.py` starts the `systeminform` package's `main()` function.
2. `collect_system_details()` checks the current OS using Python's platform
   information and chooses the Windows or Linux collector.
3. The selected collector reads local system information:
   - On **Windows**, it invokes PowerShell non-interactively and queries CIM
     classes for operating system, computer, processor, BIOS, video,
     memory, network, and logical disk data. It also reads installed
     application names/versions from the machine-wide uninstall registry
     locations.
   - On **Linux**, it reads local system files such as `/etc/os-release`,
     `/proc/cpuinfo`, `/proc/meminfo`, and `/sys/class/dmi`, and uses local
     commands such as `df`, `ip`, and available package managers to obtain
     storage, network, and software data.
4. Both collectors return the same structured data groups: `system`,
   `hardware`, `storage`, `software`, and `network`. Helpers normalize values
   and format byte quantities for display.
5. Normally, the data is passed to the Tkinter GUI. The GUI presents Overview,
   Storage, Installed Software, and Network tabs; **Refresh** recollects
   current values and **Export TXT** writes the current displayed snapshot.
6. If `SYSTEMINFORM_HEADLESS` is set to `1`, the app instead prints the
   structured data as indented JSON and exits without opening the GUI.

The collector runs on startup and again when **Refresh** is selected. Export
does not trigger a new scan: it saves the current GUI snapshot.

## Requirements and installation

- Python 3.10 or newer.
- Tkinter, which is included with many Python distributions. On Linux, install
  the distribution's Tk package if it is not available with Python.
- On Windows, PowerShell and CIM cmdlets are used by the Python collector.
- On Linux, individual commands (`ip`, `df`, `dpkg-query`, `rpm`, `pacman`,
  `apk`, `lspci`, or `lshw`) may be unavailable. The collector tries available
  sources where applicable; fields can remain unavailable.
- `uv` is recommended for installing project/build dependencies.

From this folder:

```powershell
uv sync
uv run python main.py
```

Or, with an already prepared Python environment:

```powershell
python -m pip install -r requirements.txt
python main.py
```

## Headless JSON output

This is useful for checking collector output or piping it to another local
tool:

```powershell
$env:SYSTEMINFORM_HEADLESS = "1"
python main.py
Remove-Item Env:SYSTEMINFORM_HEADLESS
```

In Command Prompt, set the environment variable for the command with:

```bat
set SYSTEMINFORM_HEADLESS=1
python main.py
```

## Windows console script

`system-info.ps1` is a separate, Windows-only PowerShell automation. It prints
operating-system and uptime, manufacturer/model, RAM and page-file usage, CPU,
fixed-disk capacity/usage/volume labels, and installed software to the console
using CIM and registry data. Run it from PowerShell:

```powershell
.\system-info.ps1
```

Unlike the Python application, this script does not launch a GUI or export a
file. It is useful for a quick terminal report. It uses the current PowerShell
session's access to CIM and the uninstall registry locations it queries.

## Exporting and packaging

In the GUI, select **Export TXT**, choose a destination, and save the current
snapshot as a UTF-8 text file. The export contains separate system, hardware,
storage, installed-software, and network sections.

To build a windowed Windows executable with PyInstaller, run from this folder:

```powershell
uv run pyinstaller --onefile --windowed --name SystemDetails main.py
```

The executable is placed under `dist`. The repository also includes
`SystemDetails.spec`, which describes a windowed executable named
`SystemDetails`; it can be used with:

```powershell
uv run pyinstaller SystemDetails.spec
```

Test a packaged build on a clean Windows machine. The app still depends on
local Windows facilities such as PowerShell/CIM and may display unavailable
fields if those facilities or data are restricted.

## Project structure

| Path | Responsibility |
|---|---|
| `main.py` | Root launcher for the Python application. |
| `src/systeminform/main.py` | Collects details, selects GUI versus headless JSON output. |
| `src/systeminform/collectors/__init__.py` | Chooses the collector for the current OS and defines fallback data. |
| `src/systeminform/collectors/windows_collector.py` | Windows CIM/PowerShell and registry collection. |
| `src/systeminform/collectors/linux_collector.py` | Linux local-file and command-based collection. |
| `src/systeminform/gui/main_window.py` | Tkinter tabs, refresh, rendering, and TXT export. |
| `src/systeminform/utils/helpers.py` | Value normalization, byte formatting, and safe local command/file helpers. |
| `system-info.ps1` | Independent Windows console report. |
| `SystemDetails.spec` | PyInstaller configuration for a windowed executable. |
| `pyproject.toml` / `uv.lock` | Project metadata, supported Python version, and build dependency lock. |

## Privacy and troubleshooting

The collectors read local machine information. The application does not need
to upload that data to a service. Treat exported reports as sensitive: host
names, network addresses, hardware details, and installed software can reveal
information about the machine. Store and share reports accordingly.

- **GUI does not open:** Check that Python includes Tkinter. On Linux, install
  the OS package that provides Tkinter, then run again.
- **Some values say `Unavailable`:** The platform does not expose the value,
  the corresponding command is not installed, the result could not be parsed,
  or the current user lacks access. Many collector operations intentionally
  fall back to `Unavailable` when local sources are missing.
- **Software list is empty or incomplete:** Windows queries machine-wide
  uninstall registry locations. Linux support depends on a recognized
  package manager. Portable, per-user, or manually installed apps may not
  appear.
- **Headless mode still opens a window:** Confirm the environment variable is
  exactly `SYSTEMINFORM_HEADLESS=1` in the same process environment used to
  start Python.
- **PowerShell report fails:** Run it on Windows in PowerShell and confirm
  that CIM queries are available in the current environment.
