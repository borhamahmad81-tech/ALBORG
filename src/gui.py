"""
gui.py - simple window for running the Alborg Lab Fetcher without typing
any commands. Pick the patient Excel file, choose Edge or Chrome, click Run.

Build note: the GitHub Actions workflow builds this as the .exe entry point
so double-clicking the .exe opens this window.
"""

from __future__ import annotations

import queue
import sys
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, ttk


def run_gui():
    root = tk.Tk()
    root.title("Al Borg Lab Fetcher")
    root.geometry("640x600")

    state = {"input_path": tk.StringVar(),
             "browser": tk.StringVar(value="edge"),
             "month": tk.StringVar(value=date.today().strftime("%Y-%m")),
             "merge": tk.BooleanVar(value=False),
             "running": False}

    pad = {"padx": 12, "pady": 6}

    # --- File picker ---
    frm_file = ttk.LabelFrame(root, text="1. Patient list (Excel file)")
    frm_file.pack(fill="x", **pad)
    entry = ttk.Entry(frm_file, textvariable=state["input_path"])
    entry.pack(side="left", fill="x", expand=True, padx=8, pady=8)

    def browse():
        path = filedialog.askopenfilename(
            title="Choose your patient list",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if path:
            state["input_path"].set(path)

    ttk.Button(frm_file, text="Browse...", command=browse).pack(side="left", padx=8)

    # --- Browser choice ---
    frm_browser = ttk.LabelFrame(root, text="2. Browser to use")
    frm_browser.pack(fill="x", **pad)
    ttk.Radiobutton(frm_browser, text="Microsoft Edge", value="edge",
                    variable=state["browser"]).pack(side="left", padx=12, pady=8)
    ttk.Radiobutton(frm_browser, text="Google Chrome", value="chrome",
                    variable=state["browser"]).pack(side="left", padx=12, pady=8)

    # --- Month ---
    frm_month = ttk.LabelFrame(root, text="3. Month (leave as-is for current month)")
    frm_month.pack(fill="x", **pad)
    ttk.Entry(frm_month, textvariable=state["month"], width=12).pack(side="left", padx=8, pady=8)
    ttk.Label(frm_month, text="format: YYYY-MM (e.g. 2026-07)").pack(side="left")

    # --- Mode ---
    frm_mode = ttk.LabelFrame(root, text="4. What to collect for each patient")
    frm_mode.pack(fill="x", **pad)
    ttk.Radiobutton(frm_mode, text="Latest report only (this month)",
                    value=False, variable=state["merge"]).pack(anchor="w", padx=12, pady=(8, 2))
    ttk.Radiobutton(frm_mode,
                    text="Merge ALL reports this month (newest value per test)",
                    value=True, variable=state["merge"]).pack(anchor="w", padx=12, pady=(2, 8))

    # --- Run button + log ---
    btn_run = ttk.Button(root, text="Run")
    btn_run.pack(pady=10)

    log_box = tk.Text(root, height=14, wrap="word", state="disabled",
                      bg="#1e1e1e", fg="#e0e0e0")
    log_box.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    log_queue: queue.Queue = queue.Queue()

    class _QueueWriter:
        _is_gui_writer = True
        def write(self, text):
            log_queue.put(text)
        def flush(self):
            pass

    def drain_log():
        try:
            while True:
                text = log_queue.get_nowait()
                log_box.configure(state="normal")
                log_box.insert("end", text)
                log_box.see("end")
                log_box.configure(state="disabled")
        except queue.Empty:
            pass
        root.after(150, drain_log)

    def start_run():
        if state["running"]:
            return
        input_path = state["input_path"].get().strip()
        if not input_path or not Path(input_path).is_file():
            log_queue.put("\n>> Please choose a valid Excel file first.\n")
            return

        state["running"] = True
        btn_run.configure(text="Running...", state="disabled")

        def worker():
            # Build argv as if run from the command line, then call main().
            import main as main_module
            sys.argv = [
                "gui",
                "--input", input_path,
                "--browser", state["browser"].get(),
                "--month", state["month"].get().strip() or date.today().strftime("%Y-%m"),
            ]
            if state["merge"].get():
                sys.argv.append("--merge-month")
            old_stdout = sys.stdout
            sys.stdout = _QueueWriter()
            try:
                main_module.main()
            except SystemExit:
                pass
            except Exception as exc:
                log_queue.put(f"\n!! Error: {exc}\n")
            finally:
                sys.stdout = old_stdout
                state["running"] = False
                root.after(0, lambda: btn_run.configure(text="Run", state="normal"))
                log_queue.put("\n>> Finished. You can close this window.\n")

        threading.Thread(target=worker, daemon=True).start()

    btn_run.configure(command=start_run)
    drain_log()
    root.mainloop()


if __name__ == "__main__":
    run_gui()
