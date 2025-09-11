# PPT-Compatible MP4 Converter Pro

A simple, aesthetic Python GUI application for converting videos into **PowerPoint-compatible MP4s**.  
Ensures your recordings (e.g., from screen recorders) play smoothly in Microsoft PowerPoint without codec issues.  

✅ **H.264 (yuv420p) + AAC** encoding  
✅ **Constant 30 FPS CFR** with `+faststart` (moov atom at front)  
✅ **Silent audio track injection** if missing (PPT requires AAC audio)  
✅ **Multiple speed options** (0.5× → 4× + custom) with pitch-preserving audio  
✅ **Batch conversion** with logs and progress bar  
✅ **Cross-platform** (Linux, Windows, macOS)  

---

## ✨ Features

- **PowerPoint compatibility**: Fixes playback issues by re-encoding to H.264 yuv420p + AAC audio.
- **Video speed control**: Apply slow motion or fast-forward with audio kept in sync.
- **Audio normalization**: Option to normalize loudness across clips.
- **Batch mode**: Convert multiple videos in one go.
- **User-friendly GUI**: Modern Tkinter design with progress tracking and logs.

---

## 📸 Screenshots

![App Screenshot](assets/Screenshot%20from%202025-09-11%2011-52-33.png)

---

## ⚙️ Installation

### 1. Clone the repository
```bash
git clone https://github.com/<your-username>/ppt-mp4-converter.git
cd ppt-mp4-converter

## Install dependencies

sudo apt-get install ffmpeg

### run script

python3 mp4_converter.py

