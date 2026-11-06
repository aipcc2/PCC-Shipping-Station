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
        py -3.12 PCC_Shipping.py
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

- `PCC_Shipping.py` — main application source
- `portal_upload.py` — sends CoC / Packing List PDFs to the PCC Customer Portal (see below)
- `tests/` — unit tests: `py -3.12 -m unittest discover -s tests -v`
- `build_exe.bat` — PyInstaller build script
- `logo.png` — company logo (used at runtime + embedded as fallback)
- `README.md` — this file
- `CONTRIBUTING.md` — how to contribute changes
- `.gitignore` — files Git should ignore (build output, caches, etc.)

## Customer Portal upload

After **Save and Close**, the CoC and Packing List are also sent to the PCC Customer Portal as PDFs, so the
customer sees them on their order. Details are in `portal_upload.py`; in short:

- **Industrial jobs only. MAG - ITAR jobs are never sent.** The app refuses them when the shipment is
  queued, again before the PDF is made, and again right before upload (division must be exactly
  "Industrial", files must sit under the Industrial production folder, and no folder name, file name or
  field may mention ITAR / USML / export controlled). The portal refuses them too.
- **Never blocks shipping.** Close only writes a small job file on the PC. A background thread makes the
  PDF with Excel (same look as the printed form; the PDF is saved next to the .xlsx in the WO folder) and
  uploads it. If the portal or network is down it retries (1 min, 5 min, 15 min, then hourly, up to 7 days).
- **Visible status.** The footer shows "Portal: ✓ 3 sent today · 1 waiting · ⚠ 1 failed"; click it for
  the list and **Retry now**. A failed upload pops up a warning so the office can upload it by hand.
- **Off until set up on each PC.** Admin → Portal upload: station id + station key (from the portal
  admin), Test connection, turn on, Save. Stored only on that PC in
  `%LOCALAPPDATA%\PCC Shipping\portal_upload.json`; the queue is in `%LOCALAPPDATA%\PCC Shipping\portal_queue\`.
- Needs Excel on the shipping PC (already used for the .xlsx files). No new Python packages.
