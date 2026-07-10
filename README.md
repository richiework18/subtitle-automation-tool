# subtitle-automation-tool
# Filmora Bilingual Subtitle Automation

An automation tool written in Python that turns a bilingual (Chinese/Indonesian) Microsoft Word script into Wondershare Filmora title clips automatically. 

This script is designed to streamline the repetitive process of manual subtitle creation by directly generating and injecting title clips into a Filmora project file (`.wfp`) without needing to manually click inside the Filmora software.

## 🚀 Features
* **Auto-Parsing:** Reads a Word document and automatically pairs each Chinese line with its Indonesian translation. It intelligently skips headings, links, or summaries at the top.
* **Dynamic Template Selection:** Automatically chooses between "z1" or "Z2" custom title templates per line, based on the Indonesian line's character count (default rule: <=58 characters uses z1, >=59 characters uses Z2).
* **Direct `.wfp` Injection:** Writes brand-new title clips directly into a copy of a Filmora `.wfp` project file using raw text editing and XML manipulation.

## ⚠️ Requirements & Limitations
Before using this tool, please take note of the following limitations:
* **Shell Project Required:** You need a "shell" `.wfp` project that already contains at least one "z1" title clip and one "Z2" title clip to act as donor templates.
* **Timing Estimation:** Timing is NOT read from the audio. Duration is estimated from assumed reading speed (default is 15 Indonesian characters/second, with a minimum of 2 seconds). You will still need to manually nudge the clip timing in Filmora to match the spoken audio perfectly.
* **Approximation:** The z1/Z2 split is a character-count approximation of true on-screen pixel width.
* **Always Use a Copy:** Always run this script on a COPY of your project to keep your original `.wfp` safe.

## 🛠️ Installation
This script requires the `python-docx` library to read Microsoft Word files. 

Install the required dependency using pip:
```bash
pip install python-docx

The script runs via the command line interface (CLI).

1. Replace all titles in a fresh project (starts at 0:00):
python filmora_bilingual_titles.py script.docx shell.wfp output.wfp

2. Append new titles after existing clips in the project:
python filmora_bilingual_titles.py script.docx shell.wfp output.wfp --mode append

3. Custom threshold or reading speed:
You can modify the threshold for the z1 template or adjust the characters-per-second (cps) reading speed.
python filmora_bilingual_titles.py script.docx shell.wfp output.wfp --threshold 60 --cps 13
