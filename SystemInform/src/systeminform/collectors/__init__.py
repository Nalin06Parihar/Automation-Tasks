import platform

from .linux_collector import collect_linux_system_details
from .windows_collector import collect_windows_system_details


def collect_system_details():
    system_name = platform.system()
    if system_name == "Windows":
        return collect_windows_system_details()
    if system_name == "Linux":
        return collect_linux_system_details()

    return {
        "system": {
            "hostname": "Unavailable",
            "os_name": system_name or "Unknown",
            "os_version": "Unavailable",
            "os_build": "Unavailable",
            "architecture": "Unavailable",
            "manufacturer": "Unavailable",
            "model": "Unavailable",
            "uptime": "Unavailable",
        },
        "hardware": {
            "cpu": {"name": "Unavailable", "physical_cores": "Unavailable", "logical_processors": "Unavailable"},
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
        },
        "storage": [],
        "software": [],
        "network": {
            "hostname": "Unavailable",
            "adapter": "Unavailable",
            "ipv4_address": "Unavailable",
            "mac_address": "Unavailable",
        },
    }


__all__ = ["collect_system_details", "collect_windows_system_details", "collect_linux_system_details"]
