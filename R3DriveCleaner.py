import os
import sys
import time
import shutil
import tempfile
import ctypes
import subprocess
import logging
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

APP_NAME = "R3Disk Cleaner"
APP_ID = "R3DiskCleaner.App"
LOG_FILENAME = "cleaner.log"


def get_resource_path(filename: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / filename
    return Path(__file__).parent / filename


def set_windows_app_id() -> None:
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


def setup_logger() -> logging.Logger:
    logger = logging.getLogger("r3disk_cleaner")
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        log_path = Path.cwd() / LOG_FILENAME
        handler = logging.FileHandler(log_path, encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger


logger = setup_logger()


def format_bytes(num: int) -> str:
    value = float(num)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if value < 1024:
            return f"{value:.2f} {unit}"
        value /= 1024
    return f"{value:.2f} PB"


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin() -> None:
    try:
        if getattr(sys, "frozen", False):
            executable = sys.executable
            params = ""
        else:
            executable = sys.executable
            script_path = str(Path(__file__).resolve())
            params = f'"{script_path}"'

        result = ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            executable,
            params,
            None,
            1,
        )

        if result <= 32:
            raise RuntimeError(f"ShellExecuteW failed with code {result}")

        sys.exit(0)

    except Exception as exc:
        try:
            temp_root = tk.Tk()
            temp_root.withdraw()
            messagebox.showerror(
                APP_NAME,
                f"This app needs administrator rights to continue.\n\nError: {exc}",
            )
            temp_root.destroy()
        except Exception:
            pass
        sys.exit(1)


def ensure_admin() -> None:
    if not is_admin():
        relaunch_as_admin()


def get_dir_size(folder: Path) -> int:
    total = 0
    try:
        for root, _, files in os.walk(folder, onerror=lambda e: None):
            for file_name in files:
                try:
                    total += (Path(root) / file_name).stat().st_size
                except Exception:
                    pass
    except Exception:
        pass
    return total


def is_path_excluded(path: Path, exclusions: list[Path]) -> bool:
    try:
        resolved_path = path.resolve(strict=False)
    except Exception:
        resolved_path = path

    for exclusion in exclusions:
        try:
            resolved_exclusion = exclusion.resolve(strict=False)
        except Exception:
            resolved_exclusion = exclusion

        try:
            if resolved_path == resolved_exclusion:
                return True
            resolved_path.relative_to(resolved_exclusion)
            return True
        except Exception:
            continue

    return False


def safe_unlink(path: Path, exclusions: list[Path], dry_run: bool = False) -> int:
    try:
        if is_path_excluded(path, exclusions):
            logger.info(f"SKIP excluded: {path}")
            return 0

        if path.is_symlink():
            logger.info(f"DELETE symlink: {path}")
            if not dry_run:
                path.unlink(missing_ok=True)
            return 0

        if path.is_file():
            try:
                size = path.stat().st_size
            except Exception:
                size = 0
            logger.info(f"DELETE file: {path} ({format_bytes(size)})")
            if not dry_run:
                path.unlink(missing_ok=True)
            return size

        if path.is_dir():
            size = get_dir_size(path)
            logger.info(f"DELETE folder: {path} ({format_bytes(size)})")
            if not dry_run:
                shutil.rmtree(path, ignore_errors=True)
            return size

    except Exception as exc:
        logger.exception(f"FAILED delete: {path} | {exc}")

    return 0


def clear_folder_contents(folder: Path, exclusions: list[Path], dry_run: bool = False) -> int:
    removed = 0

    if not folder.exists() or not folder.is_dir():
        logger.info(f"SKIP missing folder: {folder}")
        return 0

    logger.info(f"Scanning folder: {folder}")

    try:
        for item in folder.iterdir():
            removed += safe_unlink(item, exclusions, dry_run=dry_run)
    except Exception as exc:
        logger.exception(f"FAILED listing folder: {folder} | {exc}")

    return removed


def get_windows_temp_locations() -> list[Path]:
    locations: list[Path] = []

    try:
        locations.append(Path(tempfile.gettempdir()))
    except Exception:
        pass

    windir = os.environ.get("WINDIR", r"C:\Windows")
    locations.append(Path(windir) / "Temp")

    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        locations.append(Path(local_app_data) / "Temp")

    unique_locations: list[Path] = []
    seen = set()

    for path in locations:
        key = str(path).lower()
        if key not in seen:
            seen.add(key)
            unique_locations.append(path)

    return unique_locations


def empty_recycle_bin() -> bool:
    try:
        SHERB_NOCONFIRMATION = 0x00000001
        SHERB_NOPROGRESSUI = 0x00000002
        SHERB_NOSOUND = 0x00000004

        result = ctypes.windll.shell32.SHEmptyRecycleBinW(
            None,
            None,
            SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND,
        )
        logger.info(f"Recycle Bin result code: {result}")
        return result == 0
    except Exception as exc:
        logger.exception(f"FAILED empty recycle bin | {exc}")
        return False


def open_browser_clear_pages() -> list[subprocess.Popen]:
    """
    Open browser clear-data pages and return the processes launched by this app.
    Only these launched processes will be closed later.
    """
    processes: list[subprocess.Popen] = []

    browser_targets = [
        (
            "Chrome",
            [
                Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
                Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
                Path(os.environ.get("LOCALAPPDATA", "")) / r"Google\Chrome\Application\chrome.exe",
            ],
            "chrome://settings/clearBrowserData",
        ),
        (
            "Edge",
            [
                Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
                Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
            ],
            "edge://settings/clearBrowserData",
        ),
        (
            "Firefox",
            [
                Path(r"C:\Program Files\Mozilla Firefox\firefox.exe"),
                Path(r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe"),
            ],
            "about:preferences#privacy",
        ),
    ]

    for browser_name, paths, target in browser_targets:
        launched = False

        for exe_path in paths:
            try:
                if exe_path.exists():
                    proc = subprocess.Popen([str(exe_path), target])
                    processes.append(proc)
                    logger.info(f"Opened {browser_name} page: {target}")
                    launched = True
                    break
            except Exception as exc:
                logger.exception(f"FAILED opening {browser_name} with {exe_path} | {exc}")

        if not launched:
            logger.info(f"{browser_name} not found; skipped.")

    return processes


def close_launched_browsers(processes: list[subprocess.Popen]) -> None:
    """
    Close only browser processes launched by this app.
    """
    for proc in processes:
        try:
            if proc.poll() is None:
                proc.terminate()
                proc.wait(timeout=5)
                logger.info(f"Closed launched browser process PID={proc.pid}")
        except Exception:
            try:
                if proc.poll() is None:
                    proc.kill()
                    logger.info(f"Force killed launched browser process PID={proc.pid}")
            except Exception as exc:
                logger.exception(f"FAILED to close browser process PID={getattr(proc, 'pid', 'unknown')} | {exc}")


def open_system_tools() -> None:
    commands = [
        ["cmd", "/c", "start", "", "ms-settings:storagesense"],
        ["cleanmgr"],
    ]

    for command in commands:
        try:
            subprocess.run(command, shell=False, check=False)
            logger.info(f"Launched tool: {' '.join(command)}")
        except Exception as exc:
            logger.exception(f"FAILED launching tool {' '.join(command)} | {exc}")


class CleanerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("860x720")
        self.root.minsize(760, 620)
        self.root.resizable(True, True)

        self.apply_window_icon()

        self.clean_temp_var = tk.BooleanVar(value=True)
        self.clean_recycle_var = tk.BooleanVar(value=True)
        self.open_browsers_var = tk.BooleanVar(value=True)
        self.close_browsers_var = tk.BooleanVar(value=True)
        self.system_tools_var = tk.BooleanVar(value=False)
        self.dry_run_var = tk.BooleanVar(value=False)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_text_var = tk.StringVar(value="0%")
        self.status_text_var = tk.StringVar(value="Ready")

        self._build_ui()

    def apply_window_icon(self) -> None:
        try:
            icon_path = get_resource_path("Logo.ico")
            if icon_path.exists():
                self.root.iconbitmap(default=str(icon_path))
        except Exception as exc:
            logger.warning(f"Could not load Logo.ico | {exc}")

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=12)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(5, weight=1)

        title_label = ttk.Label(main, text=APP_NAME, font=("Segoe UI", 16, "bold"))
        title_label.grid(row=0, column=0, sticky="w", pady=(0, 8))

        description = (
            "Safe Windows cleanup utility.\n"
            "Temp files and Recycle Bin can be cleaned directly.\n"
            "Browser data opens in each browser's own settings page.\n"
            "Browsers opened by this app can be closed automatically after cleaning.\n"
            "The registry is not modified."
        )
        ttk.Label(main, text=description, justify="left", wraplength=800).grid(
            row=1, column=0, sticky="ew", pady=(0, 12)
        )

        options_frame = ttk.LabelFrame(main, text="Cleanup Options", padding=10)
        options_frame.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        options_frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(options_frame, text="Clean Temp files", variable=self.clean_temp_var).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Checkbutton(options_frame, text="Empty Recycle Bin", variable=self.clean_recycle_var).grid(
            row=1, column=0, sticky="w"
        )
        ttk.Checkbutton(
            options_frame, text="Open browser clear-data pages", variable=self.open_browsers_var
        ).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(
            options_frame, text="Close browsers opened by this app after cleaning", variable=self.close_browsers_var
        ).grid(row=3, column=0, sticky="w")
        ttk.Checkbutton(options_frame, text="Open safe system tools", variable=self.system_tools_var).grid(
            row=4, column=0, sticky="w"
        )
        ttk.Checkbutton(
            options_frame, text="Dry run only (preview without deleting)", variable=self.dry_run_var
        ).grid(row=5, column=0, sticky="w")

        exclusions_frame = ttk.LabelFrame(main, text="Exclusions", padding=10)
        exclusions_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 12))
        exclusions_frame.columnconfigure(0, weight=1)
        exclusions_frame.rowconfigure(1, weight=1)

        ttk.Label(
            exclusions_frame,
            text="One path per line. Items inside these folders will be skipped.",
            wraplength=780,
            justify="left",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 6))

        self.exclusions_text = tk.Text(exclusions_frame, height=7, wrap="word", undo=False)
        self.exclusions_text.grid(row=1, column=0, sticky="nsew")

        exclusions_scroll = ttk.Scrollbar(
            exclusions_frame, orient="vertical", command=self.exclusions_text.yview
        )
        exclusions_scroll.grid(row=1, column=1, sticky="ns")
        self.exclusions_text.configure(yscrollcommand=exclusions_scroll.set)

        default_exclusions = [
            str(Path.home() / "Downloads"),
            str(Path.home() / "Desktop"),
        ]
        self.exclusions_text.insert("1.0", "\n".join(default_exclusions))

        exclusion_buttons = ttk.Frame(exclusions_frame)
        exclusion_buttons.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        exclusion_buttons.columnconfigure(2, weight=1)

        ttk.Button(exclusion_buttons, text="Add Folder...", command=self.add_folder_exclusion).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(exclusion_buttons, text="Open Log File", command=self.open_log_file).grid(
            row=0, column=1, sticky="w", padx=(8, 0)
        )

        progress_frame = ttk.LabelFrame(main, text="Progress", padding=10)
        progress_frame.grid(row=4, column=0, sticky="ew", pady=(0, 12))
        progress_frame.columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            variable=self.progress_var,
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew")

        progress_info_frame = ttk.Frame(progress_frame)
        progress_info_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        progress_info_frame.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(
            progress_info_frame,
            textvariable=self.status_text_var,
            wraplength=700,
            justify="left",
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        self.progress_label = ttk.Label(
            progress_info_frame,
            textvariable=self.progress_text_var,
            font=("Segoe UI", 10, "bold"),
        )
        self.progress_label.grid(row=0, column=1, sticky="e")

        output_frame = ttk.LabelFrame(main, text="Output", padding=10)
        output_frame.grid(row=5, column=0, sticky="nsew", pady=(0, 12))
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)

        self.output = tk.Text(output_frame, wrap="word", state="disabled")
        self.output.grid(row=0, column=0, sticky="nsew")

        output_scroll = ttk.Scrollbar(output_frame, orient="vertical", command=self.output.yview)
        output_scroll.grid(row=0, column=1, sticky="ns")
        self.output.configure(yscrollcommand=output_scroll.set)

        bottom_frame = ttk.Frame(main)
        bottom_frame.grid(row=6, column=0, sticky="ew")
        bottom_frame.columnconfigure(1, weight=1)

        self.start_button = ttk.Button(bottom_frame, text="Start", command=self.run_cleanup)
        self.start_button.grid(row=0, column=0, sticky="w", padx=(0, 8))

        ttk.Button(bottom_frame, text="Quit", command=self.root.destroy).grid(
            row=0, column=2, sticky="e"
        )

        self.write_output(f"{APP_NAME} ready.")
        self.write_output(f"Admin: {'Yes' if is_admin() else 'No'}")
        self.write_output(f"Log file: {Path.cwd() / LOG_FILENAME}")

    def write_output(self, message: str) -> None:
        self.output.configure(state="normal")
        self.output.insert("end", message + "\n")
        self.output.see("end")
        self.output.configure(state="disabled")
        self.root.update_idletasks()

    def set_progress(self, percent: float, status: str = "") -> None:
        percent = max(0, min(100, percent))
        self.progress_var.set(percent)
        self.progress_text_var.set(f"{int(percent)}%")
        if status:
            self.status_text_var.set(status)
        self.root.update_idletasks()

    def reset_progress(self) -> None:
        self.progress_var.set(0)
        self.progress_text_var.set("0%")
        self.status_text_var.set("Ready")
        self.root.update_idletasks()

    def progress_from_substep(self, completed_units: int, total_units: int, current_status: str = "") -> None:
        if total_units <= 0:
            percent = 0
        else:
            percent = (completed_units / total_units) * 100
        self.set_progress(percent, current_status)

    def add_folder_exclusion(self) -> None:
        folder = filedialog.askdirectory(title="Select folder to exclude")
        if folder:
            current_text = self.exclusions_text.get("1.0", "end").strip()
            new_text = current_text + ("\n" if current_text else "") + folder
            self.exclusions_text.delete("1.0", "end")
            self.exclusions_text.insert("1.0", new_text)

    def get_exclusions(self) -> list[Path]:
        raw_lines = self.exclusions_text.get("1.0", "end").splitlines()
        exclusions: list[Path] = []

        for line in raw_lines:
            cleaned = line.strip().strip('"')
            if cleaned:
                expanded = os.path.expandvars(os.path.expanduser(cleaned))
                exclusions.append(Path(expanded))

        return exclusions

    def open_log_file(self) -> None:
        log_path = Path.cwd() / LOG_FILENAME
        try:
            os.startfile(str(log_path))  # type: ignore[attr-defined]
        except Exception:
            messagebox.showinfo("Log File", f"Log file location:\n{log_path}")

    def run_cleanup(self) -> None:
        exclusions = self.get_exclusions()
        dry_run = self.dry_run_var.get()
        browser_processes: list[subprocess.Popen] = []

        temp_locations = get_windows_temp_locations() if self.clean_temp_var.get() else []

        fixed_tasks = []
        if self.clean_recycle_var.get():
            fixed_tasks.append("Recycle Bin")
        if self.open_browsers_var.get():
            fixed_tasks.append("Browser Pages")
        if self.system_tools_var.get():
            fixed_tasks.append("System Tools")
        if self.open_browsers_var.get() and self.close_browsers_var.get():
            fixed_tasks.append("Close Browsers")

        total_units = len(temp_locations) + len(fixed_tasks)
        completed_units = 0

        self.start_button.config(state="disabled")
        self.reset_progress()

        logger.info("=" * 60)
        logger.info(f"Started run | dry_run={dry_run}")
        logger.info(f"Exclusions: {[str(x) for x in exclusions]}")

        self.write_output("")
        self.write_output("Starting cleanup...")
        self.write_output(f"Dry run: {'Yes' if dry_run else 'No'}")

        if total_units == 0:
            self.write_output("No cleanup options selected.")
            messagebox.showwarning(APP_NAME, "Please select at least one cleanup option.")
            self.start_button.config(state="normal")
            return

        total_removed = 0

        try:
            if temp_locations:
                self.write_output("Cleaning temp files...")

                for index, folder in enumerate(temp_locations, start=1):
                    self.set_progress(
                        (completed_units / total_units) * 100,
                        f"Cleaning temp location {index} of {len(temp_locations)}",
                    )

                    removed = clear_folder_contents(folder, exclusions, dry_run=dry_run)
                    total_removed += removed
                    self.write_output(f"  {folder} -> {format_bytes(removed)}")

                    completed_units += 1
                    self.progress_from_substep(
                        completed_units,
                        total_units,
                        f"Finished temp location {index} of {len(temp_locations)}",
                    )

            if self.clean_recycle_var.get():
                self.set_progress((completed_units / total_units) * 100, "Emptying Recycle Bin")
                self.write_output("Emptying Recycle Bin...")

                ok = True if dry_run else empty_recycle_bin()
                self.write_output(f"  Recycle Bin: {'done' if ok else 'failed'}")
                logger.info(f"Recycle Bin {'done' if ok else 'failed'}")

                completed_units += 1
                self.progress_from_substep(completed_units, total_units, "Recycle Bin complete")

            if self.open_browsers_var.get():
                self.set_progress((completed_units / total_units) * 100, "Opening browser pages")
                self.write_output("Opening browser clear-data pages...")

                browser_processes = open_browser_clear_pages()
                if browser_processes:
                    self.write_output("  Browser pages opened.")
                else:
                    self.write_output("  No supported browsers were found.")

                completed_units += 1
                self.progress_from_substep(completed_units, total_units, "Browser pages complete")

            if self.system_tools_var.get():
                self.set_progress((completed_units / total_units) * 100, "Opening system tools")
                self.write_output("Opening safe system tools...")

                open_system_tools()
                self.write_output("  Storage Sense / Disk Cleanup launched.")
                self.write_output("  Registry was not modified.")

                completed_units += 1
                self.progress_from_substep(completed_units, total_units, "System tools complete")

            if self.open_browsers_var.get() and self.close_browsers_var.get():
                self.set_progress((completed_units / total_units) * 100, "Closing launched browsers")
                self.write_output("Closing browser windows opened by this app...")

                if browser_processes:
                    time.sleep(5)
                    close_launched_browsers(browser_processes)
                    self.write_output("  Opened browser windows closed.")
                else:
                    self.write_output("  No launched browser windows to close.")

                completed_units += 1
                self.progress_from_substep(completed_units, total_units, "Browser close complete")

            self.write_output("")
            self.write_output(f"Estimated space removed: {format_bytes(total_removed)}")
            self.write_output(f"Log saved to: {Path.cwd() / LOG_FILENAME}")

            logger.info(f"Finished run | removed={format_bytes(total_removed)}")
            logger.info("=" * 60)

            self.set_progress(100, "Cleanup complete")

            messagebox.showinfo(
                APP_NAME,
                f"Cleanup complete.\n\nEstimated space removed: {format_bytes(total_removed)}\n\nSee {LOG_FILENAME} for details.",
            )

        finally:
            self.start_button.config(state="normal")


def main() -> None:
    set_windows_app_id()
    ensure_admin()

    root = tk.Tk()

    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    CleanerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()