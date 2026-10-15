import json
import os

from .collectors import collect_system_details
from .gui.main_window import launch_gui


def main() -> None:
    details = collect_system_details()
    if os.environ.get("SYSTEMINFORM_HEADLESS") == "1":
        print(json.dumps(details, indent=2, default=str))
        return
    launch_gui(details)


if __name__ == "__main__":
    main()
