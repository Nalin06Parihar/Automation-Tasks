import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from systeminform.collectors import collect_system_details


class SystemDetailsWindow(tk.Tk):
    def __init__(self, initial_data=None):
        super().__init__()
        self.title("System Details")
        self.geometry("1080x760")
        self.minsize(820, 580)
        self.details = initial_data if initial_data is not None else {}
        self._configure_style()
        self._build_ui()
        if self.details:
            self.render_details()
        else:
            self.refresh_details()

    def _configure_style(self):
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        self.configure(background="#f3f5f8")
        style.configure("TFrame", background="#f3f5f8")
        style.configure("Header.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#ffffff", foreground="#172b4d", font=("Segoe UI", 19, "bold"))
        style.configure("Subtitle.TLabel", background="#ffffff", foreground="#64748b", font=("Segoe UI", 9))
        style.configure("Section.TLabelframe", background="#ffffff", bordercolor="#dce2ea", relief="solid")
        style.configure("Section.TLabelframe.Label", background="#ffffff", foreground="#334155", font=("Segoe UI", 10, "bold"))
        style.configure("Value.TLabel", background="#ffffff", foreground="#172b4d", font=("Segoe UI", 9))
        style.configure("Key.TLabel", background="#ffffff", foreground="#64748b", font=("Segoe UI", 9))
        style.configure("TNotebook", background="#f3f5f8", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 9), font=("Segoe UI", 9))
        style.map("TNotebook.Tab", background=[("selected", "#ffffff")], foreground=[("selected", "#172b4d")])
        style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground="#243247", rowheight=27, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background="#eaf0f6", foreground="#334155", font=("Segoe UI", 9, "bold"), padding=(8, 7))
        style.map("Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", "#172b4d")])
        style.configure("Status.TLabel", background="#f3f5f8", foreground="#64748b", font=("Segoe UI", 9))
        style.configure("Accent.TButton", font=("Segoe UI", 9, "bold"), padding=(12, 7))

    def _build_ui(self):
        header = ttk.Frame(self, style="Header.TFrame", padding=(22, 17, 22, 15))
        header.pack(fill="x")

        heading = ttk.Frame(header, style="Header.TFrame")
        heading.pack(side="left", fill="x", expand=True)
        ttk.Label(heading, text="System Details", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            heading,
            text="Hardware, operating system, storage, software, and network information",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        actions = ttk.Frame(header, style="Header.TFrame")
        actions.pack(side="right", anchor="center")
        ttk.Button(actions, text="Refresh", command=self.refresh_details).pack(side="left")
        ttk.Button(actions, text="Export TXT", command=self.export_details, style="Accent.TButton").pack(side="left", padx=(8, 0))

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=18, pady=(14, 8))

        self.overview_tab = ttk.Frame(self.notebook, padding=15)
        self.storage_tab = ttk.Frame(self.notebook, padding=15)
        self.software_tab = ttk.Frame(self.notebook, padding=15)
        self.network_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.overview_tab, text="Overview")
        self.notebook.add(self.storage_tab, text="Storage")
        self.notebook.add(self.software_tab, text="Installed Software")
        self.notebook.add(self.network_tab, text="Network")

        self._build_overview_tab()
        self._build_storage_tab()
        self._build_software_tab()
        self._build_network_tab()

        self.status = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status, style="Status.TLabel", padding=(20, 5, 20, 10)).pack(fill="x")

    def _build_overview_tab(self):
        self.overview_tab.columnconfigure((0, 1), weight=1, uniform="overview")
        self.overview_tab.rowconfigure(0, weight=1)

        self.system_frame = ttk.LabelFrame(self.overview_tab, text="System", style="Section.TLabelframe", padding=14)
        self.hardware_frame = ttk.LabelFrame(self.overview_tab, text="Hardware", style="Section.TLabelframe", padding=14)
        self.system_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 10))
        self.hardware_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 10))

        self.system_fields = self._add_fields(
            self.system_frame,
            ["hostname", "os_name", "os_version", "os_build", "uptime", "architecture", "manufacturer", "model"],
        )
        self.hardware_fields = {}
        for section, keys in (
            ("CPU", ["name", "physical_cores", "logical_processors"]),
            (
                "Memory",
                [
                    "total_ram",
                    "used_ram",
                    "available_ram",
                    "usage_percentage",
                    "total_swap",
                    "used_swap",
                    "available_swap",
                    "swap_usage_percentage",
                ],
            ),
            ("GPU", ["name"]),
            ("BIOS", ["version"]),
        ):
            subframe = ttk.Frame(self.hardware_frame, style="Header.TFrame")
            subframe.pack(fill="x", anchor="n", pady=(0, 10))
            ttk.Label(subframe, text=section, background="#ffffff", foreground="#334155", font=("Segoe UI", 9, "bold")).grid(row=0,column=0, sticky="w", pady=(0, 3))
            self.hardware_fields[section.lower()] = self._add_fields(subframe, keys,start_row=1)

    def _build_storage_tab(self):
        self.storage_tab.rowconfigure(0, weight=1)
        self.storage_tab.columnconfigure(0, weight=1)
        columns = ("drive", "volume", "mount", "filesystem", "total", "used", "free", "usage")
        self.storage_tree = ttk.Treeview(self.storage_tab, columns=columns, show="headings")
        headings = {
            "drive": ("Device", 105),
            "volume": ("Volume label", 115),
            "mount": ("Mount point", 130),
            "filesystem": ("File system", 95),
            "total": ("Total", 90),
            "used": ("Used", 90),
            "free": ("Free", 90),
            "usage": ("Usage", 75),
        }
        for column, (label, width) in headings.items():
            self.storage_tree.heading(column, text=label)
            self.storage_tree.column(column, width=width, minwidth=65, anchor="w", stretch=True)
        self._add_tree_scrollbar(self.storage_tab, self.storage_tree)

    def _build_software_tab(self):
        self.software_tab.rowconfigure(0, weight=1)
        self.software_tab.columnconfigure(0, weight=1)
        self.software_tree = ttk.Treeview(self.software_tab, columns=("name", "version"), show="headings")
        self.software_tree.heading("name", text="Application")
        self.software_tree.heading("version", text="Version")
        self.software_tree.column("name", width=650, minwidth=250, anchor="w", stretch=True)
        self.software_tree.column("version", width=220, minwidth=120, anchor="w", stretch=False)
        self._add_tree_scrollbar(self.software_tab, self.software_tree)

    def _build_network_tab(self):
        self.network_frame = ttk.LabelFrame(self.network_tab, text="Network adapter", style="Section.TLabelframe", padding=18)
        self.network_frame.pack(fill="x", anchor="n")
        self.network_fields = self._add_fields(
            self.network_frame,
            ["hostname", "adapter", "ipv4_address", "mac_address"],
        )

    @staticmethod
    def _add_tree_scrollbar(parent, tree):
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

    @staticmethod
    def _add_fields(parent, keys, start_row=0):
        fields = {}

        parent.columnconfigure(1, weight=1)

        for index, key in enumerate(keys):
            row = start_row + index

            ttk.Label(
                parent,
                text=key.replace("_", " ").title(),
                style="Key.TLabel",
            ).grid(
                row=row,
                column=0,
                sticky="nw",
                padx=(0, 18),
                pady=5,
            )

            value = ttk.Label(
                parent,
                text="Unavailable",
                style="Value.TLabel",
                wraplength=390,
                justify="left",
            )
            value.grid(
                row=row,
                column=1,
                sticky="nw",
                pady=5,
            )
            fields[key] = value

        return fields

    @staticmethod
    def _display_value(value):
        if value is None or value == "":
            return "Unavailable"
        if isinstance(value, (dict, list)):
            return ", ".join(str(item) for item in value) if value else "Unavailable"
        return str(value)

    def _set_fields(self, labels, data):
        for key, label in labels.items():
            label.configure(text=self._display_value(data.get(key, "Unavailable")))

    def refresh_details(self):
        self.status.set("Collecting system information...")
        self.update_idletasks()
        try:
            self.details = collect_system_details()
        except Exception as exc:
            self.status.set("Refresh failed")
            messagebox.showerror("Refresh failed", f"Could not collect system information.\n\n{exc}", parent=self)
            return
        self.render_details()

    def render_details(self):
        system = self.details.get("system") or {}
        hardware = self.details.get("hardware") or {}
        self._set_fields(self.system_fields, system)
        for section, labels in self.hardware_fields.items():
            self._set_fields(labels, hardware.get(section) or {})
        self._set_fields(self.network_fields, self.details.get("network") or {})

        storage = self.details.get("storage") or []
        self._fill_tree(
            self.storage_tree,
            storage,
            (
                "drive_letter",
                "volume_name",
                "mount_point",
                "file_system",
                "total_capacity",
                "used_space",
                "free_space",
                "usage_percentage",
            ),
        )
        software = self.details.get("software") or []
        self._fill_tree(self.software_tree, software, ("name", "version"))
        self.status.set(
            f"Updated · {len(storage)} storage item(s) · {len(software)} installed application(s)"
        )

    @staticmethod
    def _fill_tree(tree, records, keys):
        tree.delete(*tree.get_children())
        for record in records:
            if not isinstance(record, dict):
                continue
            values = tuple(SystemDetailsWindow._display_value(record.get(key)) for key in keys)
            tree.insert("", "end", values=values)
        if not tree.get_children():
            tree.insert("", "end", values=tuple("Unavailable" for _ in keys))

    @staticmethod
    def _format_export_value(value, indent=0):
        prefix = " " * indent
        if isinstance(value, dict):
            lines = []
            for key, item in value.items():
                label = key.replace("_", " ").title()
                if isinstance(item, (dict, list)):
                    lines.append(f"{prefix}{label}:")
                    lines.append(SystemDetailsWindow._format_export_value(item, indent + 2))
                else:
                    lines.append(f"{prefix}{label}: {SystemDetailsWindow._display_value(item)}")
            return "\n".join(lines)
        if isinstance(value, list):
            if not value:
                return f"{prefix}Unavailable"
            lines = []
            for item in value:
                if isinstance(item, dict):
                    lines.append(f"{prefix}- " + " | ".join(
                        f"{key.replace('_', ' ').title()}: {SystemDetailsWindow._display_value(item.get(key))}"
                        for key in item
                    ))
                else:
                    lines.append(f"{prefix}- {SystemDetailsWindow._display_value(item)}")
            return "\n".join(lines)
        return f"{prefix}{SystemDetailsWindow._display_value(value)}"

    def _export_text(self):
        sections = (
            ("System", "system"),
            ("Hardware", "hardware"),
            ("Storage", "storage"),
            ("Installed Software", "software"),
            ("Network", "network"),
        )
        result = ["SYSTEM DETAILS", "=" * 60, ""]
        for title, key in sections:
            result.extend((title.upper(), "-" * len(title)))
            result.append(self._format_export_value(self.details.get(key) or {}))
            result.append("")
        return "\n".join(result)

    def export_details(self):
        target = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile="system_details.txt",
        )
        if not target:
            return
        try:
            with open(target, "w", encoding="utf-8") as handle:
                handle.write(self._export_text())
        except OSError as exc:
            messagebox.showerror("Export failed", f"Could not save the file.\n\n{exc}", parent=self)
            return
        self.status.set(f"Exported system details to {target}")


def launch_gui(data=None):
    app = SystemDetailsWindow(initial_data=data)
    app.mainloop()


if __name__ == "__main__":
    launch_gui()
