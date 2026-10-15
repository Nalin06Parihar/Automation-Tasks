def main() -> None:
    from .main import main as _main

    _main()


def collect_system_details():
    from .collectors import collect_system_details as _collect_system_details

    return _collect_system_details()


__all__ = ["main", "collect_system_details"]
