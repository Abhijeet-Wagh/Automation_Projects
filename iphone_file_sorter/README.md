# iPhone File Sorter

Sort files copied from an iPhone's **Internal Storage** (or any folder tree) into separate folders by type:

- `Images/`
- `Videos/`
- `Documents/`
- `Excel/`
- `PDF/`
- `Other/`

## Requirements

- Python 3.9+ (standard library only — no pip packages needed)

## Quick start — Jupyter Notebook (recommended if you use notebooks)

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
