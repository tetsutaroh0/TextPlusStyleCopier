# TextPlusStyleCopier for DaVinci Resolve

A utility script for DaVinci Resolve that copies Text+ styles to multiple clips using clip color filtering.

This tool is designed for fast subtitle and title workflow customization while keeping each clip’s original text and layout.

---

# Features

* Copy Text+ styles to multiple clips at once
* Target clips by clip color
* Optional track filtering

  * Apply to all tracks
  * Or only a specific video track (V1 / V2 / etc.)
* Preserve individual properties:

  * Text content
  * Position
  * Size
  * Rotation
  * Pivot
* Persistent UI window
* Works with Power Bin / Media Pool source clips
* Keeps original playhead position
* Automatically removes temporary reference clip
* Supports repeated execution without reopening the script

---

# Requirements

* DaVinci Resolve Studio / Free
* Python environment installed
* DaVinci Resolve scripting enabled

Tested on:

* Windows
* DaVinci Resolve 21.x

---

# Installation

## 1. Install Python

Python 3.x is required.

Download:
https://www.python.org/downloads/

Make sure:

* "Add Python to PATH" is enabled during installation

---

## 2. Install the Script

Copy the script file into the DaVinci Resolve Scripts folder.

Example (Windows):

```text

C:\\ProgramData\\Blackmagic Design\\DaVinci Resolve\\Fusion\\Scripts\\Utility

```

---

# Usage

## Basic Workflow

1. Put your reference Text+ clip into the Power Bin or Media Pool
2. Select the reference Text+ clip
3. Assign a clip color to target Text+ clips on the timeline
4. Run the script
5. Choose:

   * Target clip color
   * Target track
   * Preserve options
6. Click Apply

---

# Preserve Options

You can preserve individual properties while applying styles:

| Option   | Description                 |
| -------- | --------------------------- |
| Text     | Keeps original text content |
| Position | Keeps clip position         |
| Size     | Keeps text size             |
| Rotation | Keeps rotation              |
| Pivot    | Keeps pivot settings        |

---

# Track Filtering

You can limit style application to a specific video track.

| Value | Target     |
| ----- | ---------- |
| 0     | All Tracks |
| 1     | V1         |
| 2     | V2         |
| 3     | V3         |

etc.

---

# Notes

* Only Text+ clips are supported
* Fusion Titles other than Text+ may not work
* The script temporarily appends the source clip to the timeline internally
* Temporary clips are automatically removed afterward

---

# Included Versions

| File                      | Description         |
| ------------------------- | ------------------- |
| TextPlusStyleCopier_JP.py | Japanese UI version |
| TextPlusStyleCopier.py    | English UI version  |

---

# License

MIT License

---

# Author

Tetsu
