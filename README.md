# HRV BatchExtract

*[Leia em português](README_PT.md)*

HRV BatchExtract is a Windows application that automates the extraction of short-term heart rate variability (HRV) results from **Kubios HRV Scientific**. It opens each RR-interval file in Kubios, applies the same analysis settings to every file, copies the Kubios result tables and compiles them into one Excel workbook.

**HRV BatchExtract does not calculate HRV.** Every index is calculated by Kubios. The values in the output are copied as text from the Kubios result tables, so they are identical to the values Kubios displays.

HRV BatchExtract is an independent project. It is not affiliated with, endorsed by or distributed with Kubios Oy, and it does not include or modify Kubios.

## Versions

| Folder | Interface language | Launcher |
|---|---|---|
| [`en/`](en) | English | `RUN_HRV_BatchExtract.bat` |
| [`pt/`](pt) | Portuguese | `RODAR_HRV_BatchExtract.bat` |

Both versions run the same procedure. They differ only in the language of the interface, messages and output names.

## Requirements

- Windows 10 or 11 (64-bit), one monitor, display scaling unchanged during the run.
- **Kubios HRV Scientific Lite 4.3.0** (free, requires a Kubios account), or a paid Kubios HRV Scientific licence. The application was developed and tested with Lite 4.3.0 in English. A paid edition or another version must be tested on a few files before use.
- Python 3.10 or later (any distribution, including Anaconda or the one bundled with Spyder).
- Internet access on the first run, to install the Python libraries: pyautogui, pyperclip, pygetwindow, pandas, numpy, openpyxl, pillow and winsdk.

## Input files

- One plain-text file (.txt) per participant, with one RR interval per line, a single column, no header.
- Each file already cut to the **5-min** analysis segment (Task Force, 1996). HRV BatchExtract does not select or edit segments.
- RR intervals in seconds or milliseconds (detected automatically).
- All files in one folder. The file name is the participant code (the suffix `_5min` is removed). File names must not contain personal identifiers.

## Settings applied to every file

- Artefact correction: threshold-based, **medium**, reapplied and checked on screen for every file.
- Analysis sample: 00:00:00 to 00:05:00.
- All other Kubios settings: Kubios defaults.

## How to run

1. Download the folder of your language (`en` or `pt`) and keep the two files together in one local folder.
2. Connect the computer to mains power, disable sleep and screen lock, close other programs and silence notifications.
3. Double-click the launcher (`RUN_HRV_BatchExtract.bat`). The first run installs the libraries.
4. In the window, choose the folder with the .txt files. Check the path to `launch_KubiosHRV.exe` (the default installation path is pre-filled).
5. Click **RUN**.
6. **Do not use the mouse or keyboard during the run.** The application controls them.
7. To stop immediately, move the mouse to the top-left corner of the screen. **Stop** ends the run after the current file. Running again on the same folder continues with the files not yet processed.

## Output

A subfolder `hrv_batchextract_results` is created inside the data folder:

- `HRV_BatchExtract.xlsx`: sheet *HRV* (one row per participant: code, 51 Kubios indices, SD1% and SD2%), sheet *Verification* (unit, mean RR check, SD1 vs RMSSD/√2 check, correction label, number and percentage of corrected beats) and sheet *Method* (settings applied).
- `raw/`: the text copied from the Kubios result tables for each file.
- `correction_screenshots/`: a screenshot of the Kubios beat-correction field for each file.

The original RR files are never modified. Multiscale entropy is not extracted (not available in the Lite edition).

## Limitations

- The application operates the Kubios interface by screen position. It is validated only for Kubios HRV Scientific Lite 4.3.0 in English on Windows.
- The computer cannot be used during a run.
- The application does not judge signal quality. Segment selection, visual inspection and exclusion criteria remain the investigator's responsibility.

## Licence

MIT Licence (see [LICENSE](LICENSE)). You may download, use and modify your own copy, keeping the copyright notice. Changes to this repository are made only by its owner.

## Citation

A study protocol describing HRV BatchExtract and its validation is in preparation. Until it is published, please cite this repository.

The code was written with the assistance of a generative AI tool (Claude, Anthropic). The author defined the requirements, tested the application on real data, verified its output against Kubios and is responsible for the code.
