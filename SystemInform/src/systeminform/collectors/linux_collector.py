import os
import platform
import socket
import subprocess

from systeminform.utils.helpers import format_bytes, read_text_file, run_command, safe_text


def _read_os_release():
    content = read_text_file("/etc/os-release")
    values = {}
    for line in content.splitlines():
        if "=" not in line:
            continue
        key, raw_value = line.split("=", 1)
        values[key.strip()] = raw_value.strip().strip('"')
    return values


def _get_cpu_details():
    cpu_info = read_text_file("/proc/cpuinfo")
    processor_name = "Unavailable"
    physical_cores = "Unavailable"
    logical_processors = str(os.cpu_count() or "Unavailable")
    current_processor = "Unavailable"
    for line in cpu_info.splitlines():
        if line.startswith("model name") or line.startswith("Processor"):
            current_processor = line.split(":", 1)[1].strip()
            processor_name = current_processor
            break
    if "physical id" in cpu_info:
        ids = set()
        for line in cpu_info.splitlines():
            if line.startswith("physical id"):
                ids.add(line.split(":", 1)[1].strip())
        if ids:
            physical_cores = str(len(ids))
    return {
        "name": processor_name,
        "physical_cores": physical_cores or "Unavailable",
        "logical_processors": logical_processors,
    }


def _get_meminfo():
    meminfo_text = read_text_file("/proc/meminfo")
    values = {}
    for line in meminfo_text.splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        parts = value.strip().split()
        if parts:
            values[key] = float(parts[0]) * 1024

    total_ram_bytes = values.get("MemTotal")
    available_ram_bytes = values.get("MemAvailable")
    used_ram_bytes = (
        max(0, total_ram_bytes - available_ram_bytes)
        if total_ram_bytes is not None and available_ram_bytes is not None
        else None
    )
    total_swap_bytes = values.get("SwapTotal")
    available_swap_bytes = (
        max(0, total_swap_bytes - values["SwapFree"])
        if total_swap_bytes is not None and "SwapFree" in values
        else None
    )

    return {
        "total_ram": format_bytes(total_ram_bytes),
        "used_ram": format_bytes(used_ram_bytes),
        "available_ram": format_bytes(available_ram_bytes),
        "usage_percentage": (
            f"{(used_ram_bytes / total_ram_bytes) * 100:.1f}%"
            if total_ram_bytes and used_ram_bytes is not None
            else "Unavailable"
        ),
        "total_swap": format_bytes(total_swap_bytes),
        "used_swap": format_bytes(available_swap_bytes),
        "available_swap": format_bytes(values.get("SwapFree")),
        "swap_usage_percentage": (
            f"{(available_swap_bytes / total_swap_bytes) * 100:.1f}%"
            if total_swap_bytes and available_swap_bytes is not None
            else "Unavailable"
        ),
    }


def _get_uptime():
    uptime_text = read_text_file("/proc/uptime")
    try:
        uptime_seconds = int(float(uptime_text.split()[0]))
    except (IndexError, ValueError):
        return "Unavailable"
    days, remainder = divmod(uptime_seconds, 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes = remainder // 60
    return f"{days}d {hours}h {minutes}m"


def _get_gpu_name():
    output = run_command(["lspci", "-mm"]) + "\n" + run_command(["lshw", "-class", "display"])
    lines = output.splitlines()
    for line in lines:
        if "VGA" in line or "3D" in line or "Display" in line or "Graphics" in line:
            if ":" in line:
                return line.split(":", 1)[1].strip()
    return "Unavailable"


def _get_bios_version():
    for path in ["/sys/class/dmi/id/bios_version", "/sys/class/dmi/id/bios_vendor"]:
        value = read_text_file(path)
        if value:
            return value.strip()
    return "Unavailable"


def _list_linux_disks():
    output = run_command(["df", "-kP"])
    rows = []
    for line in output.splitlines()[1:]:
        parts = line.split()
        if len(parts) < 6:
            continue
        filesystem, total_kb, used_kb, available_kb, percent, mount_point = parts[:6]
        if filesystem.startswith("tmpfs") or mount_point in ["/boot/efi", "/dev", "/proc", "/sys"]:
            continue
        total_bytes = float(total_kb) * 1024
        used_bytes = float(used_kb) * 1024
        free_bytes = float(available_kb) * 1024
        rows.append(
            {
                "drive_letter": filesystem,
                "volume_name": "Unavailable",
                "mount_point": mount_point,
                "file_system": filesystem,
                "total_capacity": format_bytes(total_bytes),
                "used_space": format_bytes(used_bytes),
                "free_space": format_bytes(free_bytes),
                "usage_percentage": percent,
            }
        )
    return rows or [{
        "drive_letter": "Unavailable",
        "volume_name": "Unavailable",
        "mount_point": "Unavailable",
        "file_system": "Unavailable",
        "total_capacity": "Unavailable",
        "used_space": "Unavailable",
        "free_space": "Unavailable",
        "usage_percentage": "Unavailable",
    }]


def _list_installed_software():
    package_managers = [
        (["dpkg-query", "-W", "-f=${Package}\t${Version}\n"], "dpkg"),
        (["rpm", "-qa", "--queryformat", "%{NAME}\t%{VERSION}\\n"], "rpm"),
        (["pacman", "-Q"], "pacman"),
        (["apk", "info"], "apk"),
    ]
    for command, manager_name in package_managers:
        output = run_command(command)
        if not output or output == "Permission denied" or output == "command not found":
            continue
        software = []
        for line in output.splitlines():
            if not line.strip():
                continue
            parts = line.split("\t")
            if len(parts) == 2:
                name, version = parts
            elif len(parts) == 1:
                name, version = parts[0], "Unknown"
            else:
                name, version = parts[0], parts[1]
            software.append({"name": safe_text(name), "version": safe_text(version, "Unavailable")})
        if software:
            return software[:200]
    return [{"name": "Unavailable", "version": "Unavailable"}]


def _read_network_details():
    hostname = socket.gethostname()
    adapter = "Unavailable"
    ipv4 = "Unavailable"
    mac = "Unavailable"
    if os.path.exists("/proc/net/route"):
        route_lines = read_text_file("/proc/net/route").splitlines()[1:]
        for line in route_lines:
            fields = line.split()
            if len(fields) >= 4 and fields[1] != "00000000":
                adapter = fields[0]
                break
    output = run_command(["ip", "-o", "addr", "show", "up"])
    for line in output.splitlines():
        if "inet " in line:
            parts = line.split()
            if parts[1] != "lo":
                adapter = parts[1]
                ipv4 = parts[3].split("/", 1)[0]
                break
    iface_path = "/sys/class/net/%s/address" % adapter if adapter != "Unavailable" else ""
    if iface_path and os.path.exists(iface_path):
        mac = read_text_file(iface_path).strip() or "Unavailable"
    return {
        "hostname": hostname,
        "adapter": adapter,
        "ipv4_address": ipv4,
        "mac_address": mac,
    }


def collect_linux_system_details():
    release = _read_os_release()
    architecture = platform.machine() or platform.architecture()[0]
    system_data = {
        "hostname": socket.gethostname(),
        "os_name": release.get("NAME", "Linux"),
        "os_version": release.get("VERSION", release.get("PRETTY_NAME", "Unavailable")),
        "os_build": release.get("BUILD_ID", "Unavailable"),
        "architecture": architecture,
        "manufacturer": "Unavailable",
        "model": read_text_file("/sys/class/dmi/id/product_name").strip() or "Unavailable",
        "uptime": _get_uptime(),
    }
    hardware = {
        "cpu": _get_cpu_details(),
        "memory": _get_meminfo(),
        "gpu": {"name": _get_gpu_name()},
        "bios": {"version": _get_bios_version()},
    }
    return {
        "system": system_data,
        "hardware": hardware,
        "storage": _list_linux_disks(),
        "software": _list_installed_software(),
        "network": _read_network_details(),
    }
