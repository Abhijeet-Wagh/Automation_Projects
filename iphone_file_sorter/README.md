# iPhone File Sorter

Sort files copied from an iPhone's **Internal Storage** (or any folder tree) into separate folders by type:

- `Images/`
- `Videos/`
- `Documents/`
- `Excel/`
- `PDF/`
- `Other/`

## Requirements

- Python 3.9+
- For the web UI: `streamlit` (see `requirements.txt`)
- Notebook / CLI sorter logic uses the standard library only

## Quick start — Web UI (recommended)

### Easiest on Windows
1. Connect and unlock your iPhone (tap **Trust**)
2. Double-click `Start_UI.bat`
3. Wait for the browser to open (usually http://localhost:8501)
The UI has **two panels (tabs)**:

### 1. Copy files
Uses **Apple AFC** (via `pymobiledevice3`) to copy real files from the iPhone **Media**
area (`DCIM`, `Downloads`, `Recordings`, …) and, when iOS allows, selected **Apps**
(WhatsApp / Telegram). Windows MTP/Explorer-style copy is avoided because it often
creates **empty folders**.

1. On Windows, prefer **Python 3.12** (3.13 often fails installing `lzfse`).
   Easiest: double-click `Start_UI_Py312.bat`  
   Or in Anaconda Prompt:

   ```bat
   conda create -n iphone_copy python=3.12 -y
   conda activate iphone_copy
   cd %USERPROFILE%\Documents\Automation_Projects\iphone_file_sorter
   python -m pip install -r requirements.txt
   python -m streamlit run app.py
   ```

2. Unlock iPhone → tap **Trust**
3. **Refresh devices** → select iPhone → **Load folder tree**
4. Expand **Media** (and **Apps** if listed) → check folders → **Browse** destination → **Copy**

Behavior:
- Copies real files over USB, **file-by-file** (progress + per-file timeout)
- Stuck/failed files are skipped; re-run resumes already-copied files
- An Excel log is saved in the destination folder
- Explorer “WhatsApp” folders are MTP-only; AFC may not see app sandboxes unless iOS grants access

### 2. Sort files
1. Open the **Sort files** tab
2. **Browse** source folder (files already on the laptop)
3. **Browse** destination folder for sorted output
4. Preview, then start sorting into Images / Videos / Documents / Excel / PDF / Other

### Or from a terminal
```bat
cd path\to\iphone_file_sorter
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Quick start — Jupyter Notebook

1. Copy your iPhone folders from File Explorer  
   (`This PC → Apple iPhone → Internal Storage`) into a local folder, for example:

   `C:\Users\YourName\Documents\iPhone_Copy`

2. Open `sort_files.ipynb` in Jupyter Notebook / JupyterLab / VS Code.

3. In the **Settings** cell, set:
   - `SOURCE` → your copied iPhone folder
   - `DESTINATION` → where sorted folders should go
   - `DRY_RUN = True` first

4. Run all cells and check the preview.

5. Set `DRY_RUN = False`, re-run Settings + the final cell to copy files.

## Quick start — Command line

1. Copy Internal Storage to a local folder (same as above).

2. Open **Command Prompt** or **PowerShell** and go to this tool folder:

   ```bat
   cd path\to\iphone_file_sorter
   ```

3. Dry run first (safe — no files changed):

   ```bat
   python sort_files.py "C:\Users\YourName\Documents\iPhone_Copy" "C:\Users\YourName\Documents\iPhone_Sorted" --dry-run
   ```

4. Copy files into sorted folders:

   ```bat
   python sort_files.py "C:\Users\YourName\Documents\iPhone_Copy" "C:\Users\YourName\Documents\iPhone_Sorted"
   ```

## Options

| Option | Meaning |
|--------|---------|
| `--dry-run` | Preview actions without copying/moving |
| `--move` | Move files instead of copying (default is copy) |

## Notes

- Default mode **copies** files so originals stay safe.
- Duplicate names are renamed: `IMG_1.jpg` → `IMG_1_1.jpg`.
- Hidden junk like `Thumbs.db` / `.DS_Store` is skipped.
- Chats (iMessage/WhatsApp) are usually **not** in Internal Storage folders; this tool only sorts files that are already on disk.
