\# Text+ Style Copier for DaVinci Resolve



A Python script for DaVinci Resolve that copies the \*\*Text+ style\*\* from a reference clip (selected in Media Pool / Power Bin) and applies it to multiple Text+ clips on the timeline based on \*\*clip color filtering\*\*.



This tool is designed to speed up subtitle / title styling workflows.



\---



\## ✨ Features



\* 🎯 Apply Text+ style to multiple clips at once

\* 🎨 Filter target clips by \*\*clip color\*\*

\* 🔁 Persistent UI (no need to relaunch every time)

\* 🧠 Remembers last used color and settings

\* 📝 Keeps original text content (only style is copied)




\---



\## 🧰 Requirements



\* DaVinci Resolve Studio (recommended)

\* Python scripting enabled

\* Fusion UIManager available



> ⚠️ UI may not work in the free version depending on Resolve version.



\---



\## 🚀 Installation



1\. Save the script as:



```

TextPlusStyleCopier.py

```



2\. Place it in your Resolve Scripts folder:



\*\*Windows\*\*



```

C:\\ProgramData\\Blackmagic Design\\DaVinci Resolve\\Fusion\\Scripts\\Edit

```



\*\*Mac\*\*



```

/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility

```



3\. Restart DaVinci Resolve



4\. Run from:



```

Workspace → Scripts → TextPlusStyleCopier

```



\---



\## 🧑‍💻 How to Use



\### Step 1 — Select Source Style



Select a \*\*Text+ clip in Media Pool / Power Bin\*\*

→ This will be the style reference



\---



\### Step 2 — Mark Target Clips



On the timeline, assign a \*\*clip color\*\* (e.g. Orange)

to all Text+ clips you want to update



\---



\### Step 3 — Run Script



\* Choose the target clip color in UI

\* Enable / disable:



&#x20; \* \*\*Keep Clip Color After Apply\*\*

\* Click \*\*Apply\*\*



\---



\### Step 4 — Repeat Freely



The UI stays open, so you can:



\* Change source clip

\* Change target color

\* Apply again instantly



\---



\## 🧠 How It Works



1\. Selected Media Pool Text+ is temporarily added to timeline

2\. Its Fusion TextPlus node settings are extracted

3\. All timeline Text+ clips with selected color are processed

4\. Style is applied while preserving original text (`StyledText`)

5\. Temporary clip is deleted

6\. Playhead position is restored



\---



\## ⚠️ Limitations



\* Resolve scripting API does \*\*not reliably support multi-selection\*\*

&#x20; → Clip color filtering is used instead



\* Reference clip may appear very briefly in timeline

&#x20; → This cannot be fully avoided due to API limitations



\---



\## 💡 Tips



\* Use different colors for different styling groups

\* Combine with Power Bin for reusable style templates

\* Keep a "master style library" project for reuse



\---



\## 🛠 Possible Extensions



\* Style preset buttons (Title / Subtitle / Highlight)

\* Track-based filtering

\* Selection-based processing (if API improves)

\* Visual color picker UI



\---



\## 📄 License



Free to use, modify, and distribute.



\---



\## 🙌 Credits



Created with help from ChatGPT and real-world DaVinci Resolve workflow testing.



\---



Enjoy faster Text+ styling 🚀



