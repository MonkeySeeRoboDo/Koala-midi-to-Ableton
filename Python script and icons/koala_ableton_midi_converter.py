import os
import sys
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from pathlib import Path

import mido

OCTAVE_SHIFT = 36  # 3 octaves


# ----------------------------
# DPI (fix blurry text/UI on Windows)
# ----------------------------
def enable_dpi_awareness():
    """
    Prevent Windows from bitmap-scaling the app (blurry UI).
    Call BEFORE creating any Tk windows.
    """
    try:
        import ctypes
        # Win10+ per-monitor v2 (best)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            import ctypes
            # Older fallback
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


# ----------------------------
# Resource paths (PyInstaller-safe)
# ----------------------------
def resource_path(name: str) -> str:
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return str(Path(base) / name)
    return str(Path(__file__).with_name(name))


def set_app_icon(window: tk.Tk):
    """
    Sets the window/taskbar icon (Windows likes .ico).
    """
    ico = resource_path("koala_ableton_converter.ico")
    try:
        if os.path.exists(ico):
            window.iconbitmap(ico)
    except Exception:
        pass


# ----------------------------
# Simple, clean ttk styling 
# ----------------------------
THEME = {
    "bg": "#F4F4F4",
    "fg": "#111111",
    "muted": "#555555",
}

def apply_style(root: tk.Tk):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    root.configure(bg=THEME["bg"])

    style.configure("TFrame", background=THEME["bg"])
    style.configure("TLabel", background=THEME["bg"], foreground=THEME["fg"])
    style.configure("Muted.TLabel", background=THEME["bg"], foreground=THEME["muted"])
    style.configure("Title.TLabel", background=THEME["bg"], foreground=THEME["fg"], font=("Segoe UI", 10, "bold"))
    style.configure("TRadiobutton", background=THEME["bg"], foreground=THEME["fg"])
    style.configure("TCheckbutton", background=THEME["bg"], foreground=THEME["fg"])
    style.configure("TButton", padding=8)


# ----------------------------
# Splash screen (uses icon.png)
# ----------------------------
def show_splash(root: tk.Tk, duration_ms: int = 900):
    """
    Shows a minimal splash screen using icon.png, then returns.
    root should be created but withdrawn before calling this.
    """
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg=THEME["bg"])

    splash.update_idletasks()
    sw = splash.winfo_screenwidth()
    sh = splash.winfo_screenheight()
    w, h = 420, 220
    x = (sw - w) // 2
    y = (sh - h) // 2
    splash.geometry(f"{w}x{h}+{x}+{y}")

    # Try to load PNG
    png_path = resource_path("icon.png")
    icon_img = None
    try:
        if os.path.exists(png_path):
            icon_img = tk.PhotoImage(file=png_path)
            if icon_img.width() >= 512:
                factor = max(1, icon_img.width() // 128)
                icon_img = icon_img.subsample(factor, factor)
    except Exception:
        icon_img = None

    container = tk.Frame(splash, bg=THEME["bg"])
    container.place(relx=0.5, rely=0.5, anchor="center")

    if icon_img is not None:
        img_label = tk.Label(container, image=icon_img, bg=THEME["bg"])
        img_label.image = icon_img
        img_label.pack(pady=(0, 10))

    tk.Label(container,
             text="Koala ↔ Ableton MIDI Converter",
             bg=THEME["bg"],
             fg=THEME["fg"],
             font=("Segoe UI", 12, "bold")).pack()

    tk.Label(container,
             text="Loading…",
             bg=THEME["bg"],
             fg=THEME["muted"],
             font=("Segoe UI", 9)).pack(pady=(6, 0))

    splash.after(duration_ms, splash.destroy)
    splash.grab_set()
    splash.update()
    root.wait_window(splash)


# ----------------------------
# Mapping core
# ----------------------------
def remap_within_32(w):
    if 0 <= w <= 3:
        return w + 12
    if 12 <= w <= 15:
        return w - 12

    if 4 <= w <= 7:
        return w + 4
    if 8 <= w <= 11:
        return w - 4

    if 16 <= w <= 19:
        return w + 12
    if 28 <= w <= 31:
        return w - 12

    if 20 <= w <= 23:
        return w + 4
    if 24 <= w <= 27:
        return w - 4

    return w


def remap_note(note):
    if note < 0 or note > 127:
        return note, False
    base = note - (note % 32)
    w = note % 32
    new_note = base + remap_within_32(w)
    return new_note, (new_note != note)


def clamp_midi(n):
    if n < 0:
        return 0, True
    if n > 127:
        return 127, True
    return n, False


def forward_koala_to_ableton(note):
    r, changed = remap_note(note)
    if changed:
        r2 = r + OCTAVE_SHIFT
        r2, did_clamp = clamp_midi(r2)
        return r2, True, did_clamp
    return note, False, False


def inverse_ableton_to_koala(note):
    n1 = note
    n2 = None

    if 0 <= note - OCTAVE_SHIFT <= 127:
        unshifted = note - OCTAVE_SHIFT
        n2, _ = remap_note(unshifted)

    f1, _, _ = forward_koala_to_ableton(n1)
    ok1 = (f1 == note)

    ok2 = False
    if n2 is not None:
        f2, _, _ = forward_koala_to_ableton(n2)
        ok2 = (f2 == note)

    if ok2 and ok1:
        return n2, True
    if ok2:
        return n2, True
    if ok1:
        return n1, False

    return n1, False


# ----------------------------
# MIDI conversion + batch
# ----------------------------
def convert_midi(in_path: Path, out_path: Path, mode: str):
    mid = mido.MidiFile(in_path)

    total = 0
    changed = 0
    clamped = 0

    for track in mid.tracks:
        for msg in track:
            if msg.type in ("note_on", "note_off"):
                total += 1
                old = msg.note

                if mode == "K2A":
                    new, did_change, did_clamp = forward_koala_to_ableton(old)
                    if did_change:
                        msg.note = new
                        changed += 1
                    if did_clamp:
                        clamped += 1

                elif mode == "A2K":
                    new, did_change = inverse_ableton_to_koala(old)
                    if did_change:
                        msg.note = new
                        changed += 1
                else:
                    raise ValueError("Unknown mode: " + str(mode))

    mid.save(out_path)
    return total, changed, clamped


def iter_midi_files(folder: Path, recursive: bool):
    exts = {".mid", ".midi"}
    if recursive:
        for p in folder.rglob("*"):
            if p.is_file() and p.suffix.lower() in exts:
                yield p
    else:
        for p in folder.iterdir():
            if p.is_file() and p.suffix.lower() in exts:
                yield p


def safe_backup(path: Path):
    bak = path.with_suffix(path.suffix + ".bak")
    if not bak.exists():
        path.replace(bak)
        return bak

    i = 1
    while True:
        bak_i = path.with_suffix(path.suffix + f".bak{i}")
        if not bak_i.exists():
            path.replace(bak_i)
            return bak_i
        i += 1


def write_with_optional_overwrite(in_path: Path, temp_out: Path, overwrite: bool):
    if not overwrite:
        return None
    backup_path = safe_backup(in_path)
    temp_out.replace(in_path)
    return backup_path


# ----------------------------
# UI
# ----------------------------
class App:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.mode = tk.StringVar(value="K2A")
        self.batch = tk.BooleanVar(value=False)
        self.recursive = tk.BooleanVar(value=False)
        self.overwrite = tk.BooleanVar(value=False)

        master.title("Koala ↔ Ableton MIDI Converter")

        outer = ttk.Frame(master, padding=12)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Koala ↔ Ableton MIDI Converter", style="Title.TLabel").pack(anchor="w", pady=(0, 8))

        box = ttk.LabelFrame(outer, text="Conversion direction", padding=10)
        box.pack(fill="x", pady=(0, 10))

        ttk.Radiobutton(
            box,
            text="Koala → Ableton (rearrange pads and octaves)",
            variable=self.mode, value="K2A"
        ).pack(anchor="w", pady=(0, 4))

        ttk.Radiobutton(
            box,
            text="Ableton → Koala (revert a conversion)",
            variable=self.mode, value="A2K"
        ).pack(anchor="w")

        opts = ttk.LabelFrame(outer, text="Options", padding=10)
        opts.pack(fill="x", pady=(0, 10))

        ttk.Checkbutton(
            opts,
            text="Batch convert a folder (instead of a single file)",
            variable=self.batch,
            command=self._toggle_batch_ui
        ).pack(anchor="w")

        self.recursive_cb = ttk.Checkbutton(
            opts,
            text="Include subfolders (recursive)",
            variable=self.recursive
        )
        self.recursive_cb.pack(anchor="w", pady=(2, 0))

        ttk.Checkbutton(
            opts,
            text="Overwrite in place (creates .bak backups)",
            variable=self.overwrite
        ).pack(anchor="w", pady=(2, 0))

        ttk.Button(outer, text="Convert", command=self.run).pack(fill="x", pady=(0, 8))

        # Disclaimer with inline hyperlink (wraps correctly)
        github_url = "https://github.com/MonkeySeeRoboDo/Koala-midi-to-Ableton"

        disclaimer = tk.Text(
            outer,
            height=2,              # enough for wrap
            wrap="word",
            borderwidth=0,
            highlightthickness=0,
            background=THEME["bg"],
            foreground=THEME["muted"],
            font=("Segoe UI", 9),
        )
        disclaimer.pack(fill="x", pady=(6, 0))

        # Insert text + link word
        prefix = ("This tool is an unofficial workaround, which is in no way endorsed by either application. "
                  "\nThe source code is freely available on ")
        disclaimer.insert("1.0", prefix)
        start = disclaimer.index("insert")
        disclaimer.insert("insert", "Github.")
        end = disclaimer.index("insert")

        # Style/link behavior
        disclaimer.tag_add("github_link", start, end)
        disclaimer.tag_config("github_link", foreground="#2F6FED", underline=1)

        def _open_github(event=None):
            webbrowser.open(github_url)

        def _on_enter(event=None):
            disclaimer.config(cursor="hand2")

        def _on_leave(event=None):
            disclaimer.config(cursor="")

        disclaimer.tag_bind("github_link", "<Button-1>", _open_github)
        disclaimer.tag_bind("github_link", "<Enter>", _on_enter)
        disclaimer.tag_bind("github_link", "<Leave>", _on_leave)

        # Make it read-only and non-focusable (still clickable)
        disclaimer.config(state="disabled")


    def _toggle_batch_ui(self):
        state = tk.NORMAL if self.batch.get() else tk.DISABLED
        self.recursive_cb.configure(state=state)

    def run(self):
        mode = self.mode.get()
        overwrite = self.overwrite.get()

        if self.batch.get():
            folder = filedialog.askdirectory(title="Select a folder containing MIDI files")
            if not folder:
                return
            folder = Path(folder)

            midi_files = list(iter_midi_files(folder, self.recursive.get()))
            if not midi_files:
                messagebox.showinfo("No files", "No .mid/.midi files found in that folder.")
                return

            out_dir = folder / "_converted"
            if not overwrite:
                out_dir.mkdir(exist_ok=True)

            ok = messagebox.askyesno(
                "Confirm batch convert",
                f"Found {len(midi_files)} MIDI file(s).\n\n"
                + (f"Outputs will be written to:\n{out_dir}\n\n"
                   if not overwrite
                   else "Overwrite is ON: originals will be replaced and .bak backups created.\n\n")
                + "Continue?"
            )
            if not ok:
                return

            total_files = 0
            total_events = 0
            total_changed = 0
            total_clamped = 0
            total_backups = 0
            failures = []

            suffix = "_KoalaToAbleton" if mode == "K2A" else "_AbletonToKoala"

            for p in midi_files:
                try:
                    if overwrite:
                        temp_out = p.with_name(p.stem + suffix + p.suffix)
                    else:
                        temp_out = out_dir / (p.stem + suffix + p.suffix)

                    t, c, cl = convert_midi(p, temp_out, mode)
                    bak = write_with_optional_overwrite(p, temp_out, overwrite)
                    if bak is not None:
                        total_backups += 1

                    total_files += 1
                    total_events += t
                    total_changed += c
                    total_clamped += cl

                except Exception as e:
                    failures.append(f"{p.name}: {e}")

            msg = (
                f"Converted files: {total_files}/{len(midi_files)}\n"
                f"Total note events: {total_events}\n"
                f"Total events changed: {total_changed}\n"
            )
            if mode == "K2A":
                msg += f"Total clamped after +36: {total_clamped}\n"
            if overwrite:
                msg += f"Backups created: {total_backups}\n"

            if failures:
                msg += "\nSome files failed:\n" + "\n".join(failures[:10])
                if len(failures) > 10:
                    msg += "\n..."

            messagebox.showinfo("Batch done", msg)
            return

        # Single file
        p = filedialog.askopenfilename(
            title="Select MIDI file",
            filetypes=[("MIDI files", "*.mid *.midi")],
        )
        if not p:
            return

        in_path = Path(p)
        suffix = "_KoalaToAbleton" if mode == "K2A" else "_AbletonToKoala"

        if overwrite:
            temp_out = in_path.with_name(in_path.stem + suffix + in_path.suffix)
            out_path = in_path
        else:
            out_path = in_path.with_name(in_path.stem + suffix + in_path.suffix)
            temp_out = out_path

        try:
            total, changed, clamped = convert_midi(in_path, temp_out, mode)
            bak = write_with_optional_overwrite(in_path, temp_out, overwrite)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        msg = (
            f"Saved:\n{out_path}\n\n"
            f"Note events total: {total}\n"
            f"Note events changed: {changed}\n"
        )
        if mode == "K2A":
            msg += f"Clamped after +36 shift: {clamped}\n"
        if overwrite and bak is not None:
            msg += f"\nBackup created:\n{bak}\n"

        messagebox.showinfo("Done", msg)


def main():
    enable_dpi_awareness()

    root = tk.Tk()
    apply_style(root)
    set_app_icon(root)

    root.withdraw()
    show_splash(root, duration_ms=900)

    root.deiconify()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
