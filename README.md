# OpenCV Dot Tracker for Instron Tensile Tests

Automated displacement measurement tool for Instron tensile test videos. Tracks two Sharpie marker dots on elastic specimens frame-by-frame using OpenCV, replacing manual tracking software.

## Features

- **Automated dot detection** — finds dots on specimen using annular contrast filtering and specimen region isolation
- **Sub-pixel tracking** — adaptive blob-finding centroid refinement resists drift during specimen stretching
- **Batch processing** — process multiple videos with background threading
- **Video review** — scrub through processed videos with annotated crosshair overlays
- **Data cleaning** — outlier removal via rolling median velocity + MAD thresholds
- **Displacement output** — calibrated pixel-to-mm conversion from filename-encoded initial distance (e.g., `49.9mm`)
- **CSV export** — auto-saves time, distance, and displacement data

## Usage

1. Place `.MOV` video files in `input_videos/` (filenames must contain initial dot distance, e.g., `49.9mm`)
2. Double-click `Launch Tracker.bat` or run `python app.py`
3. Select videos, click Process
4. Review results with scrub bar, export cleaned data

## Requirements

- Python 3.8+
- OpenCV (`opencv-python`)
- NumPy
- Matplotlib
- tkinter (included with Python)

## Validation

The `validation/` folder contains `build_comparison.py` which generates an Excel comparison sheet between manual Tracker measurements and OpenCV output. Validated at <0.5% difference, R² ≈ 0.9999.
