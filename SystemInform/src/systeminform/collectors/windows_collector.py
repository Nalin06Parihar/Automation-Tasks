import json
import platform
import socket
import subprocess

from systeminform.utils.helpers import first_available, format_bytes, safe_text, windows_reg_value


WINDOWS_POWER_SHELL = ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command"]


def _run_windows_powershell(script: str) -> str:
    try:
        completed = subprocess.run(
            WINDOWS_POWER_SHELL + [script],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError:
        return ""
    return completed.stdout.strip() or completed.stderr.strip()


def _load_object(script: str):
    output = _run_windows_powershell(script)
    if not output:
        return []
    try:
        return json.loads(output)
    except json.JSONDecodeError:
        return []


def _pick_value(container, *keys):
    if not isinstance(container, dict):
        return "Unavailable"
    for key in keys:
        value = container.get(key)
        if value not in (None, "", "Unavailable"):
            return value
    return "Unavailable"


def _pick_record(records, keys):
    if records is None:
        return {}
    if isinstance(records, dict):
        records = [records]
    if not isinstance(records, list):
        return {}
    for record in records:
        if isinstance(record, dict):
            for key in keys:
                if record.get(key) not in (None, "", "Unavailable"):
                    return record
    return {}


def collect_windows_system_details():
    system_data = {
        "hostname": socket.gethostname(),
        "os_name": "Windows",
        "os_version": platform.version(),
        "os_build": "Unavailable",
        "architecture": platform.machine() or platform.architecture()[0],
        "manufacturer": "Unavailable",
        "model": "Unavailable",
        "uptime": "Unavailable",
    }

    computer_system = _load_object(
        r"""
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
        $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue | Select-Object -First 1
        $gpu = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue | Select-Object -First 1
        $mem = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue | Select-Object -First 1
        $pageFileUsage = @(Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue)
        $pageFileTotalKb = ($pageFileUsage | Measure-Object -Property AllocatedBaseSize -Sum).Sum
        $pageFileUsedKb = ($pageFileUsage | Measure-Object -Property CurrentUsage -Sum).Sum
        $net = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True" -ErrorAction SilentlyContinue | Select-Object -First 1
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DriveType = 3" -ErrorAction SilentlyContinue | Select-Object DeviceID, VolumeName, FileSystem, Size, FreeSpace
        $software = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall","HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue | ForEach-Object { $item = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue; if ($item.DisplayName) { [pscustomobject]@{ Name = $item.DisplayName; Version = $item.DisplayVersion } } } | Sort-Object Name -Unique
        $uptime = "Unavailable"
        if ($os -and $os.LastBootUpTime) {
            $elapsed = (Get-Date) - $os.LastBootUpTime
            $uptime = "{0}d {1}h {2}m" -f [int]$elapsed.TotalDays, $elapsed.Hours, $elapsed.Minutes
        }

        [pscustomobject]@{
            ComputerSystem = $cs
            OperatingSystem = $os
            Uptime = $uptime
            Processor = $cpu
            BIOS = $bios
            GPU = $gpu
            Memory = $mem
            PageFile = [pscustomobject]@{ TotalSize = $pageFileTotalKb; UsedSize = $pageFileUsedKb }
            Network = $net
            Disks = $disk
            Software = $software
        } | ConvertTo-Json -Depth 8
        """
    )

    if isinstance(computer_system, dict):
        computer_system_record = computer_system.get("ComputerSystem") or {}
        if isinstance(computer_system_record, dict):
            system_data["manufacturer"] = safe_text(_pick_value(computer_system_record, "Manufacturer"), "Unavailable")
            system_data["model"] = safe_text(_pick_value(computer_system_record, "Model"), "Unavailable")
        else:
            if isinstance(computer_system_record, list):
                first_item = computer_system_record[0] if computer_system_record else {}
                system_data["manufacturer"] = safe_text(_pick_value(first_item, "Manufacturer"), "Unavailable")
                system_data["model"] = safe_text(_pick_value(first_item, "Model"), "Unavailable")

        os_record = computer_system.get("OperatingSystem") or {}
        if isinstance(os_record, list):
            os_record = os_record[0] if os_record else {}
        system_data["os_build"] = safe_text(_pick_value(os_record, "BuildNumber"), "Unavailable")
        system_data["os_version"] = safe_text(_pick_value(os_record, "Caption", "Version"), system_data["os_version"])
        system_data["uptime"] = safe_text(computer_system.get("Uptime"), "Unavailable")

    hardware = {
        "cpu": {
            "name": "Unavailable",
            "physical_cores": "Unavailable",
            "logical_processors": "Unavailable",
        },
        "memory": {
            "total_ram": "Unavailable",
            "used_ram": "Unavailable",
            "available_ram": "Unavailable",
            "usage_percentage": "Unavailable",
            "total_swap": "Unavailable",
            "used_swap": "Unavailable",
            "available_swap": "Unavailable",
            "swap_usage_percentage": "Unavailable",
        },
        "gpu": {"name": "Unavailable"},
        "bios": {"version": "Unavailable"},
    }

    cpu_record = _pick_record(computer_system.get("Processor") if isinstance(computer_system, dict) else [], ["Name", "Caption", "DeviceID"])
    if cpu_record:
        hardware["cpu"]["name"] = safe_text(_pick_value(cpu_record, "Name", "Caption"), "Unavailable")
        hardware["cpu"]["physical_cores"] = safe_text(_pick_value(cpu_record, "NumberOfCores"), "Unavailable")
        hardware["cpu"]["logical_processors"] = safe_text(_pick_value(cpu_record, "NumberOfLogicalProcessors"), "Unavailable")

    memory_record = _pick_record(computer_system.get("Memory") if isinstance(computer_system, dict) else [], ["TotalVisibleMemorySize", "FreePhysicalMemory"])
    if memory_record:
        total_visible_kb = memory_record.get("TotalVisibleMemorySize")
        free_physical_kb = memory_record.get("FreePhysicalMemory")
        if total_visible_kb is not None:
            total_ram_bytes = float(total_visible_kb) * 1024
            hardware["memory"]["total_ram"] = format_bytes(total_ram_bytes)
        else:
            total_ram_bytes = None
        if free_physical_kb is not None:
            available_ram_bytes = float(free_physical_kb) * 1024
            hardware["memory"]["available_ram"] = format_bytes(available_ram_bytes)
        else:
            available_ram_bytes = None
        if total_ram_bytes is not None and available_ram_bytes is not None:
            used_ram_bytes = max(0, total_ram_bytes - available_ram_bytes)
            hardware["memory"]["used_ram"] = format_bytes(used_ram_bytes)
            if total_ram_bytes:
                hardware["memory"]["usage_percentage"] = f"{(used_ram_bytes / total_ram_bytes) * 100:.1f}%"

    page_file_record = _pick_record(
        computer_system.get("PageFile") if isinstance(computer_system, dict) else [],
        ["TotalSize", "UsedSize"],
    )
    if page_file_record:
        total_swap_kb = page_file_record.get("TotalSize")
        used_swap_kb = page_file_record.get("UsedSize")
        if total_swap_kb is not None:
            total_swap_bytes = float(total_swap_kb) * 1024
            hardware["memory"]["total_swap"] = format_bytes(total_swap_bytes)
        else:
            total_swap_bytes = None
        if used_swap_kb is not None:
            used_swap_bytes = float(used_swap_kb) * 1024
            hardware["memory"]["used_swap"] = format_bytes(used_swap_bytes)
        else:
            used_swap_bytes = None
        if total_swap_bytes is not None and used_swap_bytes is not None:
            available_swap_bytes = max(0, total_swap_bytes - used_swap_bytes)
            hardware["memory"]["available_swap"] = format_bytes(available_swap_bytes)
            if total_swap_bytes:
                hardware["memory"]["swap_usage_percentage"] = f"{(used_swap_bytes / total_swap_bytes) * 100:.1f}%"

    gpu_record = _pick_record(computer_system.get("GPU") if isinstance(computer_system, dict) else [], ["Name", "Caption"])
    if gpu_record:
        hardware["gpu"]["name"] = safe_text(_pick_value(gpu_record, "Name", "Caption"), "Unavailable")

    bios_record = _pick_record(computer_system.get("BIOS") if isinstance(computer_system, dict) else [], ["Version", "Caption"])
    if bios_record:
        hardware["bios"]["version"] = safe_text(_pick_value(bios_record, "Version", "Caption"), "Unavailable")

    disks = []
    raw_disks = (computer_system.get("Disks") if isinstance(computer_system, dict) else [])
    if isinstance(raw_disks, dict):
        raw_disks = [raw_disks]
    if isinstance(raw_disks, list):
        for disk in raw_disks:
            if not isinstance(disk, dict):
                continue
            device = disk.get("DeviceID") or "Unavailable"
            total_size = disk.get("Size")
            free_space = disk.get("FreeSpace")
            used_space = None
            if total_size not in (None, "") and free_space not in (None, ""):
                try:
                    used_space = float(total_size) - float(free_space)
                except (TypeError, ValueError):
                    used_space = None
            try:
                total_value = float(total_size) if total_size not in (None, "") else None
                free_value = float(free_space) if free_space not in (None, "") else None
                total_display = format_bytes(total_value) if total_value is not None else "Unavailable"
                free_display = format_bytes(free_value) if free_value is not None else "Unavailable"
                used_display = format_bytes(used_space) if used_space is not None else "Unavailable"
                percent_used = "Unavailable"
                if total_value not in (None, 0):
                    percent_used = f"{((used_space / total_value) * 100) if used_space is not None else 0:.1f}%"
            except (TypeError, ValueError):
                total_display = "Unavailable"
                free_display = "Unavailable"
                used_display = "Unavailable"
                percent_used = "Unavailable"
            disks.append(
                {
                    "drive_letter": safe_text(device,
                    "Unavailable"),
                    "volume_name": safe_text(disk.get("VolumeName"), "Unavailable"),
                    "mount_point": safe_text(device, "Unavailable"),
                    "file_system": safe_text(disk.get("FileSystem"), "Unavailable"),
                    "total_capacity": total_display,
                    "used_space": used_display,
                    "free_space": free_display,
                    "usage_percentage": percent_used,
                }
            )
    if not disks:
        disks = [{
            "drive_letter": "Unavailable",
            "volume_name": "Unavailable",
            "mount_point": "Unavailable",
            "file_system": "Unavailable",
            "total_capacity": "Unavailable",
            "used_space": "Unavailable",
            "free_space": "Unavailable",
            "usage_percentage": "Unavailable",
        }]

    software_records = computer_system.get("Software") if isinstance(computer_system, dict) else []
    software = []
    if isinstance(software_records, list):
        for item in software_records:
            if not isinstance(item, dict):
                continue
            name = item.get("Name")
            version = item.get("Version")
            if name:
                software.append({"name": safe_text(name), "version": safe_text(version, "Unavailable")})

    network = {
        "hostname": socket.gethostname(),
        "adapter": "Unavailable",
        "ipv4_address": "Unavailable",
        "mac_address": "Unavailable",
    }
    network_record = _pick_record(computer_system.get("Network") if isinstance(computer_system, dict) else [], ["IPAddress", "IPEnabled", "MacAddress", "Description"])
    if network_record:
        network["adapter"] = safe_text(_pick_value(network_record, "Description", "Caption"), "Unavailable")
        ip_values = network_record.get("IPAddress")
        if isinstance(ip_values, list):
            for value in ip_values:
                if isinstance(value, str) and "." in value:
                    network["ipv4_address"] = value
                    break
        elif isinstance(ip_values, str):
            network["ipv4_address"] = ip_values
        network["mac_address"] = safe_text(_pick_value(network_record, "MACAddress"), "Unavailable")

    return {
        "system": system_data,
        "hardware": hardware,
        "storage": disks,
        "software": software,
        "network": network,
    }
