#!/usr/bin/env python3
"""
PPT-Compatible MP4 Converter Pro (GUI, Aesthetic + Speed Control)

- Video: H.264 (yuv420p), CFR 30fps, +faststart, even dimensions
- Audio: AAC 128k, 48kHz, stereo; adds silent track if missing
- Speed: presets (0.5x .. 4x) + custom; audio tempo preserved
- Profiles: Most Compatible (Baseline), Balanced (Main), High (High)
"""

import os
import sys
import json
import shutil
import subprocess
import threading
import queue
import math
from tkinter import Tk, filedialog, StringVar, BooleanVar, N, S, E, W, messagebox
from tkinter import ttk

# ----------------------- Utilities -----------------------

def which_ffmpeg() -> bool:
    return shutil.which("ffmpeg") is not None and shutil.which("ffprobe") is not None

def run_cmd(cmd):
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    out, err = proc.communicate()
    return proc.returncode, out, err

def has_audio_stream(path: str) -> bool:
    cmd = ["ffprobe", "-v", "error", "-select_streams", "a", "-show_entries", "stream=codec_type", "-of", "json", path]
    code, out, _ = run_cmd(cmd)
    if code != 0:
        return False
    try:
        data = json.loads(out)
        streams = data.get("streams", [])
        return any(s.get("codec_type") == "audio" for s in streams)
    except Exception:
        return False

def ensure_even_dimensions_filter():
    # Ensure H.264-safe dimensions and constant 30 fps
    return "scale=trunc(iw/2)*2:trunc(ih/2)*2,fps=30"

def build_atempo_expr(speed: float) -> str:
    """
    ffmpeg atempo supports 0.5..2.0 per filter; chain to reach arbitrary factor.
    We want audio to match video speed (tempo = speed).
    """
    if speed <= 0:
        speed = 1.0
    factors = []
    remaining = speed

    # For big speeds, keep halving/doubling into 0.5..2.0 range
    while remaining > 2.0:
        factors.append(2.0)
        remaining /= 2.0
    while remaining < 0.5:
        factors.append(0.5)
        remaining /= 0.5  # multiply by 2 effectively

    # Snap residual to 3 decimals to avoid float noise
    remaining = round(remaining, 3)
    if not math.isclose(remaining, 1.0, rel_tol=1e-3, abs_tol=1e-3):
        factors.append(remaining)

    return ",".join(f"atempo={f}" for f in factors) if factors else "atempo=1.0"

PROFILES = {
    "Most Compatible (Baseline L3.1, 30fps)": {"profile": "baseline", "level": "3.1", "preset": "veryfast", "crf": "20"},
    "Balanced (Main L4.0, 30fps)":             {"profile": "main",     "level": "4.0", "preset": "faster",   "crf": "20"},
    "High Quality (High L4.1, 30fps)":         {"profile": "high",     "level": "4.1", "preset": "fast",     "crf": "18"},
}

SPEED_PRESETS = ["0.5x", "0.75x", "1.0x", "1.25x", "1.5x", "2.0x", "2.5x", "3.0x", "4.0x", "Custom…"]

def parse_speed(preset: str, custom: str) -> float:
    if preset == "Custom…":
        try:
            v = float(custom.strip().lower().replace("x", ""))
            return v if v > 0 else 1.0
        except Exception:
            return 1.0
    else:
        return float(preset.replace("x", ""))

def suggest_output_path(in_path: str, out_dir: str) -> str:
    base = os.path.splitext(os.path.basename(in_path))[0]
    return os.path.join(out_dir, f"{base}_ppt.mp4")

# ----------------------- FFmpeg Builder -----------------------

def build_ffmpeg_cmd(inp, outp, profile_cfg, speed=1.0, loud_norm=False, add_silence=False):
    """
    Build a robust ffmpeg command:
      - Video: libx264 yuv420p, CFR 30, +faststart, setpts for speed
      - Audio: AAC 128k stereo 48kHz; atempo for speed; loudnorm optional
      - Silence: add anullsrc if no audio; 'shortest' to trim trailing silence
    """
    # Video filter chain
    vf_parts = []
    # Speed: PTS/speed (faster => divide; slower => multiply)
    if not math.isclose(speed, 1.0, rel_tol=1e-6):
        vf_parts.append(f"setpts=PTS/{speed}")
    vf_parts.append(ensure_even_dimensions_filter())
    vf = ",".join(vf_parts)

    base = [
        "ffmpeg", "-y",
        "-i", inp,
        "-map_metadata", "-1",
        "-movflags", "+faststart",
        "-vsync", "vfr",
        "-vf", vf,
        "-r", "30",
        "-c:v", "libx264",
        "-profile:v", profile_cfg["profile"],
        "-level", profile_cfg["level"],
        "-pix_fmt", "yuv420p",
        "-preset", profile_cfg["preset"],
        "-crf", profile_cfg["crf"],
    ]

    if add_silence:
        # Add silent AAC track; we don't need atempo on silence since -shortest ends at video
        base += [
            "-f", "lavfi", "-t", "99999", "-i", "anullsrc=channel_layout=stereo:sample_rate=48000",
            "-shortest",
            "-c:a", "aac", "-b:a", "128k", "-ar", "48000", "-ac", "2",
            "-map", "0:v:0", "-map", "1:a:0",
        ]
    else:
        # Real audio present: tempo must match video speed
        a_filters = []
        if loud_norm:
            a_filters.append("loudnorm=I=-16:TP=-1.5:LRA=11")
        if not math.isclose(speed, 1.0, rel_tol=1e-6):
            a_filters.append(build_atempo_expr(speed))
        if a_filters:
            base += ["-filter:a", ",".join(a_filters)]
        base += ["-c:a", "aac", "-b:a", "128k", "-ar", "48000", "-ac", "2"]

    base.append(outp)
    return base

# ----------------------- Worker -----------------------

class ConverterWorker(threading.Thread):
    def __init__(self, tasks, profile_name, speed_preset, speed_custom, normalize_audio, overwrite, log_q, progress_cb):
        super().__init__(daemon=True)
        self.tasks = tasks
        self.profile = PROFILES[profile_name]
        self.speed = parse_speed(speed_preset, speed_custom)
        self.normalize_audio = normalize_audio
        self.overwrite = overwrite
        self.log_q = log_q
        self.progress_cb = progress_cb

    def log(self, msg):
        self.log_q.put(msg)

    def run(self):
        total = len(self.tasks)
        for idx, (inp, outp) in enumerate(self.tasks, start=1):
            try:
                self.progress_cb(idx - 1, total, f"Converting: {os.path.basename(inp)}")
                if not os.path.isfile(inp):
                    self.log(f"[SKIP] Not found: {inp}")
                    continue

                os.makedirs(os.path.dirname(outp), exist_ok=True)
                if os.path.exists(outp) and not self.overwrite:
                    self.log(f"[SKIP] Exists (enable Overwrite to replace): {outp}")
                    continue

                silent = not has_audio_stream(inp)
                cmd = build_ffmpeg_cmd(
                    inp, outp,
                    profile_cfg=self.profile,
                    speed=self.speed,
                    loud_norm=self.normalize_audio,
                    add_silence=silent
                )
                self.log(f"[CMD] {' '.join(cmd)}")
                code, _, err = run_cmd(cmd)
                if code != 0:
                    self.log(f"[ERROR] {os.path.basename(inp)}:\n{err}")
                else:
                    self.log(f"[OK] {outp}")
            except Exception as e:
                self.log(f"[ERROR] {os.path.basename(inp)}: {e}")
            finally:
                self.progress_cb(idx, total, f"Done {idx}/{total}")

# ----------------------- GUI -----------------------

class App:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("PPT-Compatible MP4 Converter Pro")
        self.root.geometry("880x600")
        self.root.minsize(820, 540)

        # Aesthetic theme
        self._set_style()

        self.files = []
        self.out_dir = StringVar(value=os.path.expanduser("~"))
        self.profile = StringVar(value=list(PROFILES.keys())[0])
        self.normalize_audio = BooleanVar(value=False)
        self.overwrite = BooleanVar(value=True)
        self.speed_preset = StringVar(value="1.0x")
        self.speed_custom = StringVar(value="1.0")

        self.log_q = queue.Queue()

        container = ttk.Frame(root, padding=12, style="Card.TFrame")
        container.grid(column=0, row=0, sticky=(N, S, E, W))
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # Top: Inputs and Options
        top = ttk.Frame(container)
        top.grid(column=0, row=0, sticky=(E, W))
        top.columnconfigure(0, weight=1)
        top.columnconfigure(1, weight=1)
        top.columnconfigure(2, weight=1)
        top.columnconfigure(3, weight=1)

        # File List (left)
        files_card = ttk.LabelFrame(container, text="Input Videos", padding=10, style="Card.TLabelframe")
        files_card.grid(column=0, row=1, sticky=(N, S, E, W), pady=(10, 8))
        container.rowconfigure(1, weight=1)
        files_card.columnconfigure(0, weight=1)
        files_card.rowconfigure(1, weight=1)

        self.file_list = self._mk_listbox(files_card, height=10)
        self.file_list.grid(column=0, row=1, columnspan=4, sticky=(N, S, E, W), pady=(4, 8))

        btns = ttk.Frame(files_card)
        btns.grid(column=0, row=2, columnspan=4, sticky=(E, W))
        ttk.Button(btns, text="Add Files", command=self.add_files, style="Accent.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btns, text="Clear", command=self.clear_files).pack(side="left")
        ttk.Label(files_card, text="Output Folder:").grid(column=0, row=0, sticky=W, pady=(0, 2))
        out_row = ttk.Frame(files_card)
        out_row.grid(column=1, row=0, columnspan=3, sticky=(E, W))
        out_row.columnconfigure(0, weight=1)
        ttk.Entry(out_row, textvariable=self.out_dir).grid(column=0, row=0, sticky=(E, W))
        ttk.Button(out_row, text="Browse…", command=self.pick_out_dir).grid(column=1, row=0, padx=(6,0))

        # Options (right)
        opts = ttk.LabelFrame(container, text="Conversion Options", padding=10, style="Card.TLabelframe")
        opts.grid(column=0, row=2, sticky=(E, W))
        for i in range(4):
            opts.columnconfigure(i, weight=1)

        ttk.Label(opts, text="Profile:").grid(column=0, row=0, sticky=W, pady=(0,6))
        prof = ttk.Combobox(opts, textvariable=self.profile, values=list(PROFILES.keys()), state="readonly")
        prof.grid(column=1, row=0, sticky=(E, W), pady=(0,6))

        ttk.Label(opts, text="Speed:").grid(column=2, row=0, sticky=W, pady=(0,6))
        speed_box = ttk.Combobox(opts, textvariable=self.speed_preset, values=SPEED_PRESETS, state="readonly")
        speed_box.grid(column=3, row=0, sticky=(E, W), pady=(0,6))
        speed_box.bind("<<ComboboxSelected>>", self._on_speed_preset)

        self.custom_row = ttk.Frame(opts)
        self.custom_row.grid(column=2, row=1, columnspan=2, sticky=(E, W))
        self.custom_row.columnconfigure(1, weight=1)
        ttk.Label(self.custom_row, text="Custom speed (e.g., 1.15x):").grid(column=0, row=0, sticky=W)
        self.custom_entry = ttk.Entry(self.custom_row, textvariable=self.speed_custom, state="disabled")
        self.custom_entry.grid(column=1, row=0, sticky=(E, W))

        ttk.Checkbutton(opts, text="Normalize audio loudness", variable=self.normalize_audio).grid(column=0, row=2, sticky=W, pady=(6,0))
        ttk.Checkbutton(opts, text="Overwrite existing files", variable=self.overwrite).grid(column=1, row=2, sticky=W, pady=(6,0))

        # Progress + Action
        action = ttk.Frame(container, padding=(0, 6, 0, 6))
        action.grid(column=0, row=3, sticky=(E, W))
        self.progress = ttk.Progressbar(action, mode="determinate")
        self.progress.grid(column=0, row=0, columnspan=3, sticky=(E, W), pady=(0, 6))
        self.status = StringVar(value="Idle")
        ttk.Label(action, textvariable=self.status, style="Muted.TLabel").grid(column=0, row=1, sticky=W)
        self.convert_btn = ttk.Button(action, text="Convert to PPT-Compatible MP4", command=self.convert, style="Accent.TButton")
        self.convert_btn.grid(column=2, row=1, sticky=E)

        action.columnconfigure(0, weight=1)
        action.columnconfigure(1, weight=1)
        action.columnconfigure(2, weight=0)

        # Log
        log_card = ttk.LabelFrame(container, text="Log", padding=10, style="Card.TLabelframe")
        log_card.grid(column=0, row=4, sticky=(N, S, E, W))
        container.rowconfigure(4, weight=1)
        self.log_box = self._mk_text(log_card, height=10)
        self.log_box.grid(column=0, row=0, sticky=(N, S, E, W))
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(0, weight=1)

        # periodic log updates
        self.root.after(100, self.drain_log)

        if not which_ffmpeg():
            messagebox.showwarning(
                "ffmpeg not found",
                "ffmpeg/ffprobe not found on PATH.\n\nInstall ffmpeg and try again."
            )

    # ---- Style helpers ----
    def _set_style(self):
        style = ttk.Style()
        try:
            # Use a modern theme if available
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Card.TFrame", background="#f7f7fa")
        style.configure("Card.TLabelframe", background="#ffffff")
        style.configure("TLabelframe.Label", padding=4)
        style.configure("Accent.TButton", padding=(10,6))
        style.configure("Muted.TLabel", foreground="#666")

    def _mk_listbox(self, parent, height=10):
        import tkinter as tk
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        lb = tk.Listbox(frame, height=height, activestyle="dotbox", relief="flat", highlightthickness=1, highlightbackground="#ddd")
        lb.grid(column=0, row=0, sticky=(N, S, E, W))
        sb = ttk.Scrollbar(frame, orient="vertical", command=lb.yview)
        sb.grid(column=1, row=0, sticky=(N, S))
        lb.configure(yscrollcommand=sb.set)
        return lb

    def _mk_text(self, parent, height=10):
        import tkinter as tk
        txt = tk.Text(parent, height=height, wrap="word", relief="flat", highlightthickness=1, highlightbackground="#ddd")
        sb = ttk.Scrollbar(parent, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.grid(column=0, row=0, sticky=(N, S, E, W))
        sb.grid(column=1, row=0, sticky=(N, S))
        return txt

    # ---- UI actions ----
    def _on_speed_preset(self, _event=None):
        preset = self.speed_preset.get()
        if preset == "Custom…":
            self.custom_entry.configure(state="normal")
            self.custom_entry.focus_set()
        else:
            self.custom_entry.configure(state="disabled")

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select video files",
            filetypes=[("Video", "*.mp4 *.mov *.m4v *.mkv *.avi *.webm *.wmv"), ("All files", "*.*")]
        )
        if not paths:
            return
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.file_list.insert("end", p)

    def clear_files(self):
        self.files.clear()
        self.file_list.delete(0, "end")

    def pick_out_dir(self):
        d = filedialog.askdirectory(title="Select output folder", initialdir=self.out_dir.get())
        if d:
            self.out_dir.set(d)

    def convert(self):
        if not self.files:
            messagebox.showinfo("No files", "Please add at least one video.")
            return
        if not which_ffmpeg():
            messagebox.showerror("Missing ffmpeg", "ffmpeg/ffprobe not found on PATH.")
            return

        tasks = []
        out_root = self.out_dir.get()
        for f in self.files:
            tasks.append((f, suggest_output_path(f, out_root)))

        self.convert_btn.config(state="disabled")
        self.progress.config(value=0, maximum=len(tasks))
        self.status.set("Starting…")

        worker = ConverterWorker(
            tasks=tasks,
            profile_name=self.profile.get(),
            speed_preset=self.speed_preset.get(),
            speed_custom=self.speed_custom.get(),
            normalize_audio=self.normalize_audio.get(),
            overwrite=self.overwrite.get(),
            log_q=self.log_q,
            progress_cb=self.update_progress
        )
        worker.start()

        # Re-enable button after completion (poll progress)
        def poll_done():
            if self.progress["value"] >= self.progress["maximum"]:
                self.convert_btn.config(state="normal")
                self.status.set("All done.")
            else:
                self.root.after(300, poll_done)
        self.root.after(300, poll_done)

    def update_progress(self, current, total, msg):
        self.progress.config(value=current, maximum=total)
        self.status.set(msg)

    def drain_log(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                self.log_box.insert("end", msg + "\n")
                self.log_box.see("end")
        except queue.Empty:
            pass
        self.root.after(120, self.drain_log)

# ----------------------- Main -----------------------

def main():
    root = Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()

