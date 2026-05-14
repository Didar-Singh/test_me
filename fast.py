import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
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


def copy_file(src_path: Path, dst_path: Path) -> tuple[str, bool, str]:
    try:
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_path, dst_path)
        return src_path.name, True, ""
    except Exception as e:
        return src_path.name, False, str(e)


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


def run_copy(files: list[Path], dst_dir: Path, workers: int = 8, base_dir: Path | None = None):
    dst_dir.mkdir(parents=True, exist_ok=True)
    list_files(files, base_dir)

    confirm = input("Start copying? (y/n): ").strip().lower()
    if confirm != "y":
        print("Cancelled.")
        return

    print(f"\nCopying {len(files)} file(s) with {workers} threads...\n")
    start = time.perf_counter()
    done, failed = 0, 0

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(
                copy_file, f,
                dst_dir / f.relative_to(base_dir) if base_dir else dst_dir / f.name
            ): f.name
            for f in files
        }
        for future in as_completed(futures):
            name, success, err = future.result()
            if success:
                done += 1
            else:
                failed += 1
                print(f"  FAILED: {name} — {err}")
            print(f"  Progress: {done + failed}/{len(files)}", end="\r")

    elapsed = time.perf_counter() - start
    print(f"\n\nDone.  {done} copied  |  {failed} failed  |  {elapsed:.2f}s")


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
        run_copy(files, dst, base_dir=src)

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
        run_copy(files, dst)

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
        run_copy(found, dst)

    input("\nPress Enter to exit...")
