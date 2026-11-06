# PCC Shipping Station

Internal desktop app for PCC shipping operators. Generates shipping documentation (Certificate of Compliance, Packing List, Serial Numbers list) and manages the photo/serial capture workflow per Work Order.

Built with Python + tkinter, packaged as a Windows EXE.

## Where things live

- **Production EXE:** `\\PCC2\PCC-I-Shares\Production\Industrial\Shipping App\dist\PCC Shipping.exe`
- **App data (operator/admin lists):** `\\PCC2\PCC-I-Shares\Production\Industrial\Shipping App\`
- **Industrial WO folders:** `\\PCC2\PCC-I-Shares\Production\Industrial\`
- **MAG-ITAR WO folders:** `\\PCC2\PCC-I-Shares\Production\MAG - ITAR\`

## Setting up for development

You work from a **local clone on your own PC**, not directly on the network share.

### One-time setup

1. Install [Python 3.12](https://python.org/downloads) (check "Add to PATH" during install)
2. Install [Git for Windows](https://git-scm.com/download/win) and [GitHub Desktop](https://desktop.github.com)
3. Clone this repo to `C:\dev\pcc-shipping-station\`
4. Install required Python packages:
        py -3.12 -m pip install pyinstaller pillow python-docx openpyxl python-barcode
### Running from source (for testing changes)
        py -3.12 app_logic.py
### Building the EXE
        build_exe.bat

Output: `dist\PCC Shipping.exe`

### Deploying to production

1. Make sure no operator has the app running (they hold a lock on the EXE)
2. Copy `dist\PCC Shipping.exe` to `\\PCC2\PCC-I-Shares\Production\Industrial\Shipping App\dist\`
3. Test on one shipping PC before announcing to the team

## Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md) for the branch and pull request workflow.

## Project files

- `app_logic.py` — main application source (single file; the EXE is built from it)
- `PCC Shipping App.spec` — PyInstaller recipe used by `build_exe.bat`
- `PROJECT_NOTES.md` — dated change log and design notes (read this first)
- `tests/` — test suites (`python tests/test_slip_labels.py` etc.; set the `SCRIPTS` path at the top of each to your clone)
- `tools/` — read-only diagnostic scripts
- `build_exe.bat` — PyInstaller build script
- `logo.png` — company logo (used at runtime + embedded as fallback)
- `README.md` — this file
- `CONTRIBUTING.md` — how to contribute changes
- `.gitignore` — files Git should ignore (build output, caches, etc.)
