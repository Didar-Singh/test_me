import csv
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
import time
import tkinter as tk
from tkinter import filedialog, messagebox


def pick_folder(title: str) -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title=title)
    root.destroy()
    return Path(folder) if folder else None


def pick_files() -> list[Path]:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    files = filedialog.askopenfilenames(title="Select files to copy (hold Ctrl for multiple)")
    root.destroy()
    return [Path(f) for f in files]


def ask_mode() -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    choice = {"value": None}

    win = tk.Toplevel(root)
    win.title("Copy Mode")
    win.attributes("-topmost", True)
    win.resizable(False, False)
    win.grab_set()

    tk.Label(win, text="How do you want to select files?", font=("Segoe UI", 11, "bold"), pady=10).pack()

    def set_mode(m):
        choice["value"] = m
        win.destroy()

    tk.Button(win, text="Entire Folder",        width=25, command=lambda: set_mode("folder")).pack(pady=4)
    tk.Button(win, text="Pick Files (GUI)",     width=25, command=lambda: set_mode("files")).pack(pady=4)
    tk.Button(win, text="Paste List from Excel",width=25, command=lambda: set_mode("excel")).pack(pady=4)

    win.protocol("WM_DELETE_WINDOW", lambda: set_mode(None))
    root.wait_window(win)
    root.destroy()
    return choice["value"]


def ask_report_folder(default_dst: Path) -> Path:
    """Ask where to save the report. Defaults to destination folder."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    result = {"path": default_dst}

    win = tk.Toplevel(root)
    win.title("Report Output Folder")
    win.attributes("-topmost", True)
    win.resizable(False, False)
    win.grab_set()

    tk.Label(
        win,
        text="Where should the CSV report be saved?",
        font=("Segoe UI", 10, "bold"),
        pady=8,
    ).pack()

    tk.Label(win, text=f"Default: {default_dst}", fg="gray", wraplength=380).pack(padx=10)

    def use_default():
        result["path"] = default_dst
        win.destroy()

    def pick_custom():
        folder = filedialog.askdirectory(title="Select Report Output Folder", parent=win)
        if folder:
            result["path"] = Path(folder)
        win.destroy()

    tk.Button(win, text="Save report to Destination folder (default)", width=38,
              command=use_default).pack(pady=4)
    tk.Button(win, text="Choose a different folder…", width=38,
              command=pick_custom).pack(pady=4, padx=10)

    win.protocol("WM_DELETE_WINDOW", use_default)
    root.wait_window(win)
    root.destroy()
    return result["path"]


def ask_recursive() -> bool:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    result = messagebox.askyesno("Include Subfolders?", "Copy files inside subfolders too?")
    root.destroy()
    return result


def paste_excel_list() -> list[str]:
    """Open a small text box for the user to paste file names from Excel."""
    root = tk.Tk()
    root.withdraw()

    names = {"value": []}

    win = tk.Toplevel(root)
    win.title("Paste File List")
    win.attributes("-topmost", True)
    win.geometry("500x400")
    win.grab_set()

    tk.Label(
        win,
        text="Paste file names from Excel below (one per line).\nFile extensions required e.g.  report.xlsx",
        justify="left",
        pady=6
    ).pack(anchor="w", padx=10)

    text = tk.Text(win, font=("Consolas", 10), wrap="none")
    text.pack(fill="both", expand=True, padx=10, pady=4)

    def confirm():
        raw = text.get("1.0", "end").strip().splitlines()
        names["value"] = [line.strip() for line in raw if line.strip()]
        win.destroy()

    tk.Button(win, text="OK — Search for these files", command=confirm, width=30).pack(pady=8)
    win.protocol("WM_DELETE_WINDOW", win.destroy)
    root.wait_window(win)
    root.destroy()
    return names["value"]


def find_files_by_name(names: list[str], src_dir: Path, recursive: bool) -> tuple[list[Path], list[str]]:
    """Search src_dir for each file name. Returns (found, not_found)."""
    all_files = list(src_dir.rglob("*") if recursive else src_dir.iterdir())
    index = {f.name.lower(): f for f in all_files if f.is_file()}

    found, not_found = [], []
    for name in names:
        match = index.get(name.lower())
        if match:
            found.append(match)
        else:
            not_found.append(name)
    return found, not_found


def collect_files(src_dir: Path, recursive: bool) -> list[Path]:
    if recursive:
        return [f for f in src_dir.rglob("*") if f.is_file()]
    return [f for f in src_dir.iterdir() if f.is_file()]


def copy_file(src_path: Path, dst_path: Path) -> tuple[str, bool, str, float, int]:
    """Returns (name, success, error, duration_s, size_bytes)."""
    size = src_path.stat().st_size
    t0 = time.perf_counter()
    try:
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_path, dst_path)
        return src_path.name, True, "", time.perf_counter() - t0, size
    except Exception as e:
        return src_path.name, False, str(e), time.perf_counter() - t0, size


def list_files(files: list[Path], base_dir: Path | None = None):
    print(f"\n{'─' * 62}")
    print(f"  {'#':<6} {'File Name':<40} {'Size':>10}")
    print(f"{'─' * 62}")
    for i, f in enumerate(files, 1):
        size = f.stat().st_size
        size_str = (
            f"{size / 1_048_576:.2f} MB" if size >= 1_048_576
            else f"{size / 1_024:.1f} KB" if size >= 1_024
            else f"{size} B"
        )
        name = str(f.relative_to(base_dir)) if base_dir else f.name
        display = name if len(name) <= 39 else "..." + name[-36:]
        print(f"  {i:<6} {display:<40} {size_str:>10}")
    print(f"{'─' * 62}")
    print(f"  Total: {len(files)} file(s)\n")


def _progress_bar(done: int, total: int, width: int = 30) -> str:
    filled = int(width * done / total) if total else 0
    bar = "█" * filled + "░" * (width - filled)
    pct = 100 * done / total if total else 0
    return f"[{bar}] {pct:5.1f}%"


def _fmt_size(n: int) -> str:
    if n >= 1_048_576:
        return f"{n / 1_048_576:.2f} MB"
    if n >= 1_024:
        return f"{n / 1_024:.1f} KB"
    return f"{n} B"


def write_report(
    report_dir: Path,
    rows: list[dict],
    total_elapsed: float,
    dst_dir: Path,
    src_dir: Path | None,
) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = report_dir / f"copy_report_{ts}.csv"

    copied  = [r for r in rows if r["status"] == "OK"]
    failed  = [r for r in rows if r["status"] == "FAILED"]
    total_bytes = sum(r["size_bytes"] for r in copied)

    with report_path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)

        # ── Summary block ────────────────────────────────────────
        w.writerow(["=== Copy Report ==="])
        w.writerow(["Generated",    datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        if src_dir:
            w.writerow(["Source",   str(src_dir)])
        w.writerow(["Destination",  str(dst_dir)])
        w.writerow(["Total files",  len(rows)])
        w.writerow(["Copied",       len(copied)])
        w.writerow(["Failed",       len(failed)])
        w.writerow(["Total size",   _fmt_size(total_bytes)])
        w.writerow(["Total time",   f"{total_elapsed:.2f}s"])
        avg_speed = total_bytes / total_elapsed / 1_048_576 if total_elapsed > 0 else 0
        w.writerow(["Avg speed",    f"{avg_speed:.2f} MB/s"])
        w.writerow([])

        # ── Per-file detail ──────────────────────────────────────
        w.writerow(["#", "File Name", "Size", "Status", "Duration (s)", "Error"])
        for i, r in enumerate(rows, 1):
            w.writerow([
                i,
                r["name"],
                _fmt_size(r["size_bytes"]),
                r["status"],
                f"{r['duration']:.3f}",
                r["error"],
            ])

    return report_path


def run_copy(
    files: list[Path],
    dst_dir: Path,
    report_dir: Path,
    workers: int = 8,
    base_dir: Path | None = None,
):
    dst_dir.mkdir(parents=True, exist_ok=True)
    list_files(files, base_dir)

    confirm = input("Start copying? (y/n): ").strip().lower()
    if confirm != "y":
        print("Cancelled.")
        return

    total = len(files)
    print(f"\nCopying {total} file(s) with {workers} threads...\n")
    start = time.perf_counter()
    done, failed = 0, 0
    rows: list[dict] = []

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(
                copy_file, f,
                dst_dir / f.relative_to(base_dir) if base_dir else dst_dir / f.name
            ): f
            for f in files
        }
        for future in as_completed(futures):
            name, success, err, duration, size = future.result()
            rows.append({"name": name, "status": "OK" if success else "FAILED",
                         "error": err, "duration": duration, "size_bytes": size})
            if success:
                done += 1
            else:
                failed += 1

            completed = done + failed
            elapsed = time.perf_counter() - start
            speed = completed / elapsed if elapsed > 0 else 0
            remaining = (total - completed) / speed if speed > 0 else 0
            eta = f"{remaining:.0f}s" if remaining < 60 else f"{remaining / 60:.1f}m"

            bar = _progress_bar(completed, total)
            status = f"  {bar}  {completed}/{total}  ✓{done}  ✗{failed}  {speed:.1f} f/s  ETA {eta}   "
            print(status, end="\r", flush=True)

    elapsed = time.perf_counter() - start
    print(" " * 90, end="\r")

    failed_rows = [r for r in rows if r["status"] == "FAILED"]
    if failed_rows:
        print("  Failed files:")
        for r in failed_rows:
            print(f"    ✗  {r['name']} — {r['error']}")

    print(f"\n  Done.  {done} copied  |  {failed} failed  |  {elapsed:.2f}s")

    report_path = write_report(report_dir, rows, elapsed, dst_dir, base_dir)
    print(f"  Report  : {report_path}\n")


if __name__ == "__main__":
    print("=== Fast File Copy ===\n")

    mode = ask_mode()

    if mode is None:
        print("No mode selected. Exiting.")
        exit()

    # ── ENTIRE FOLDER ────────────────────────────────────────────
    if mode == "folder":
        src = pick_folder("Select SOURCE folder")
        if not src:
            print("No source folder selected. Exiting.")
            exit()
        recursive = ask_recursive()
        files = collect_files(src, recursive)
        if not files:
            print("No files found in the selected folder.")
            exit()
        print(f"\nSource      : {src}")
        dst = pick_folder("Select DESTINATION folder")
        if not dst:
            print("No destination folder selected. Exiting.")
            exit()
        print(f"Destination : {dst}")
        report_dir = ask_report_folder(dst)
        print(f"Report to   : {report_dir}")
        run_copy(files, dst, report_dir, base_dir=src)

    # ── PICK FILES VIA GUI ───────────────────────────────────────
    elif mode == "files":
        files = pick_files()
        if not files:
            print("No files selected. Exiting.")
            exit()
        print(f"\n{len(files)} file(s) selected.")
        dst = pick_folder("Select DESTINATION folder")
        if not dst:
            print("No destination folder selected. Exiting.")
            exit()
        print(f"Destination : {dst}")
        report_dir = ask_report_folder(dst)
        print(f"Report to   : {report_dir}")
        run_copy(files, dst, report_dir)

    # ── PASTE LIST FROM EXCEL ────────────────────────────────────
    elif mode == "excel":
        names = paste_excel_list()
        if not names:
            print("No file names entered. Exiting.")
            exit()

        print(f"\n{len(names)} name(s) pasted.")

        src = pick_folder("Select SOURCE folder to search in")
        if not src:
            print("No source folder selected. Exiting.")
            exit()

        recursive = ask_recursive()
        print(f"\nSearching in: {src} ...")
        found, not_found = find_files_by_name(names, src, recursive)

        if not_found:
            print(f"\n  NOT FOUND ({len(not_found)}):")
            for name in not_found:
                print(f"    - {name}")

        if not found:
            print("\nNone of the listed files were found. Exiting.")
            exit()

        dst = pick_folder("Select DESTINATION folder")
        if not dst:
            print("No destination folder selected. Exiting.")
            exit()

        print(f"\nSource      : {src}")
        print(f"Destination : {dst}")
        report_dir = ask_report_folder(dst)
        print(f"Report to   : {report_dir}")
        run_copy(found, dst, report_dir)

    input("\nPress Enter to exit...")
