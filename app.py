# =============================================================================
#  Power BI Metadata Extractor — Desktop GUI (Production)
# =============================================================================

import os
import sys
import threading
import subprocess
from pathlib import Path

try:
    import customtkinter as ctk
    from tkinter import filedialog, messagebox
except ImportError:
    print("Missing dependency: customtkinter")
    sys.exit(1)

if getattr(sys, "frozen", False):
    base_path = Path(sys._MEIPASS)
else:
    base_path = Path(__file__).parent

sys.path.insert(0, str(base_path))
from core_extractor import run_extraction, get_default_output_dir

CREATE_NO_WINDOW = 0x08000000

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


def _silent_kwargs():
    kwargs = {}
    if os.name == "nt":
        kwargs["creationflags"] = CREATE_NO_WINDOW
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = 0
        kwargs["startupinfo"] = si
    return kwargs


def check_powershell_module() -> tuple:
    """Strong check: actually try Import-Module. Returns (ok, detail)."""
    try:
        list_cmd = (
            "Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt | "
            "Select-Object -First 1 -ExpandProperty Name"
        )
        r1 = subprocess.run(
            ["powershell", "-NoProfile", "-Command", list_cmd],
            capture_output=True, text=True, timeout=15, **_silent_kwargs()
        )
        listed = "MicrosoftPowerBIMgmt" in (r1.stdout or "")
        if not listed:
            return False, "not_listed"

        import_cmd = (
            "$ErrorActionPreference='Stop'; "
            "Import-Module MicrosoftPowerBIMgmt -ErrorAction Stop; "
            "Write-Output 'IMPORT_OK'"
        )
        r2 = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", import_cmd],
            capture_output=True, text=True, timeout=20, **_silent_kwargs()
        )
        if "IMPORT_OK" in (r2.stdout or ""):
            return True, "ok"

        err = (r2.stderr or r2.stdout or "").strip()
        return False, f"import_failed: {err[:200]}"
    except Exception as e:
        return False, f"check_error: {e}"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Power BI Metadata Extractor")
        self.geometry("880x720")
        self.minsize(780, 620)
        self.output_dir = str(get_default_output_dir())
        self.is_running = False
        self._build_ui()
        self.after(400, self._refresh_prerequisites)

    def _build_ui(self):
        header = ctk.CTkFrame(self, corner_radius=0, fg_color=("#1F497D", "#163a5f"))
        header.pack(fill="x")
        ctk.CTkLabel(header, text="Power BI Metadata Extractor",
                     font=ctk.CTkFont(size=20, weight="bold"), text_color="white").pack(pady=(16, 2))
        ctk.CTkLabel(header, text="Architecture map  ·  Data sources  ·  M code  ·  Relationships  ·  RLS",
                     font=ctk.CTkFont(size=12), text_color="#b0c4de").pack(pady=(0, 14))

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=14)

        # Prerequisites
        self.prereq_frame = ctk.CTkFrame(main, border_width=1)
        self.prereq_frame.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(self.prereq_frame, text="Prerequisites",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=14, pady=(10, 2))
        self.prereq_label = ctk.CTkLabel(self.prereq_frame, text="Checking system…",
                                         font=ctk.CTkFont(size=12), justify="left", anchor="w")
        self.prereq_label.pack(anchor="w", padx=14, pady=(0, 12))

        # Output folder
        folder_card = ctk.CTkFrame(main)
        folder_card.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(folder_card, text="Output Folder",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=14, pady=(10, 4))
        path_row = ctk.CTkFrame(folder_card, fg_color="transparent")
        path_row.pack(fill="x", padx=14, pady=(0, 12))
        self.path_entry = ctk.CTkEntry(path_row, height=34)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.path_entry.insert(0, self.output_dir)
        ctk.CTkButton(path_row, text="Browse…", width=100, height=34,
                      command=self._browse_folder).pack(side="right")

        # Options
        opts = ctk.CTkFrame(main)
        opts.pack(fill="x", pady=(0, 10))
        self.open_excel_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(opts, text="Open Excel file when finished",
                        variable=self.open_excel_var, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=14, pady=10)

        # Action
        action = ctk.CTkFrame(main, fg_color="transparent")
        action.pack(fill="x", pady=(0, 8))
        self.run_btn = ctk.CTkButton(action, text="Start Extraction", height=42,
                                     font=ctk.CTkFont(size=14, weight="bold"),
                                     command=self._start_extraction)
        self.run_btn.pack(fill="x")
        self.progress = ctk.CTkProgressBar(action, height=8)
        self.progress.pack(fill="x", pady=(10, 4))
        self.progress.set(0)
        self.status_lbl = ctk.CTkLabel(action, text="Ready", font=ctk.CTkFont(size=12),
                                       text_color=("gray40", "gray60"))
        self.status_lbl.pack(anchor="w")

        # Log
        log_card = ctk.CTkFrame(main)
        log_card.pack(fill="both", expand=True)
        ctk.CTkLabel(log_card, text="Activity Log",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=14, pady=(10, 4))
        self.log_box = ctk.CTkTextbox(log_card, font=ctk.CTkFont(family="Consolas", size=12), wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=14, pady=(0, 12))
        self.log_box.configure(state="disabled")

        footer = ctk.CTkFrame(self, corner_radius=0, height=30)
        footer.pack(fill="x", side="bottom")
        ctk.CTkLabel(footer, text="Requires a Power BI Administrator account",
                     font=ctk.CTkFont(size=11), text_color="gray50").pack(pady=6)

    def _refresh_prerequisites(self):
        ok, _ = check_powershell_module()
        if ok:
            text = (
                "✓  PowerShell module is installed and working\n\n"
                "Also required (checked at runtime):\n"
                "• Sign in with a Power BI Administrator account\n"
                "• Tenant setting “Enhance admin API responses with detailed metadata” must be Enabled"
            )
            self.prereq_frame.configure(border_color="#3a9a5c")
            self.status_lbl.configure(text="Ready")
        else:
            text = (
                "✗  PowerShell module is missing or cannot be loaded\n\n"
                "Please open PowerShell and run this command once:\n\n"
                "Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force\n\n"
                "If already installed, also run:\n"
                "Import-Module MicrosoftPowerBIMgmt\n\n"
                "Then close and restart this application."
            )
            self.prereq_frame.configure(border_color="#c04040")
            self.status_lbl.configure(text="Please fix the prerequisite above")
        self.prereq_label.configure(text=text)

    def _browse_folder(self):
        folder = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.path_entry.get() or str(Path.home()),
        )
        if folder:
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, folder)

    def _log(self, msg: str):
        def append():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, append)

    def _set_progress(self, pct: float, status: str = ""):
        def update():
            self.progress.set(max(0.0, min(1.0, pct)))
            if status:
                self.status_lbl.configure(text=status)
        self.after(0, update)

    def _set_running(self, running: bool):
        def update():
            self.is_running = running
            self.run_btn.configure(
                state="disabled" if running else "normal",
                text="Extraction in progress…" if running else "Start Extraction",
            )
        self.after(0, update)

    def _start_extraction(self):
        if self.is_running:
            return

        ok, _ = check_powershell_module()
        if not ok:
            messagebox.showerror(
                "Missing Prerequisite",
                "The PowerShell module is missing or cannot be loaded.\n\n"
                "Open PowerShell and run:\n\n"
                "Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force\n\n"
                "Then restart this application."
            )
            self._refresh_prerequisites()
            return

        out_dir = self.path_entry.get().strip()
        if not out_dir:
            messagebox.showwarning("Output Folder Required", "Please select an output folder.")
            return

        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        self._set_running(True)
        self._set_progress(0.0, "Starting…")

        def worker():
            result = run_extraction(
                output_dir=out_dir,
                open_excel_when_done=self.open_excel_var.get(),
                log_callback=self._log,
                progress_callback=self._set_progress,
            )

            def finish():
                self._set_running(False)
                if result["success"]:
                    self.status_lbl.configure(text="Completed successfully")
                    messagebox.showinfo(
                        "Extraction Complete",
                        "The report was created successfully.\n\n"
                        f"Excel:\n{result['excel_path']}\n\n"
                        f"JSON:\n{result['json_path']}"
                    )
                else:
                    self.status_lbl.configure(text="Failed")
                    messagebox.showerror("Extraction Failed", result["message"])

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()