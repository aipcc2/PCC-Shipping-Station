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
        py -3.12 -m pip install pyinstaller pillow python-docx openpyxl python-barcode reportlab
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
- `pdf_docs.py` — draws the PDF copy of the CoC, Packing List and Serial list (no Excel needed)
- `portal_upload.py` — sends those PDFs to the PCC Customer Portal (Industrial only; see below)
- `PCC Shipping App.spec` — PyInstaller recipe used by `build_exe.bat` (bundles the two modules above into the one exe)
- `PROJECT_NOTES.md` — dated change log and design notes (read this first)
- `tests/` — test suites (`python tests/test_slip_labels.py` etc.; set the `SCRIPTS` path at the top of each to your clone)
- `tools/` — read-only diagnostic scripts
- `build_exe.bat` — PyInstaller build script
- `logo.png` — company logo (used at runtime + embedded as fallback)
- `README.md` — this file
- `CONTRIBUTING.md` — how to contribute changes
- `.gitignore` — files Git should ignore (build output, caches, etc.)

## PDFs and the Customer Portal upload

**PDFs.** On **Save and Close** every document is written twice: the .xlsx as before, and a .pdf with the
same name in the same folder (`4041-11-7-26-COC.xlsx` + `4041-11-7-26-COC.pdf`). `pdf_docs.py` draws the
PDFs itself with reportlab (the shipping PCs have no Excel), following the v64 .xlsx layouts. A PDF that
fails shows "+ PDF FAILED" in the close summary but never stops the close. If you change
`write_coc_xlsx` / `write_packing_list_xlsx`, change `pdf_docs.py` to match.

**Portal upload.** The CoC, Packing List and Serial list PDFs are then sent to the PCC Customer Portal so
the customer sees them on their order:

- **Industrial jobs only. MAG - ITAR jobs are never sent.** `portal_upload.itar_reason()` refuses a job
  unless the division is exactly "Industrial", every file is under the Industrial production share (not
  MAG - ITAR), and no folder name, file name or field mentions ITAR / USML / export controlled. Checked
  when the job is queued, when the uploader picks it up, and right before upload. Offline (Outbox)
  shipments are not sent. The portal refuses ITAR jobs too.
- **Never blocks shipping.** Close only writes a small job file on the PC; a background thread uploads.
  Retries 1 min, 5 min, 15 min, then hourly, up to 7 days.
- **Visible status.** Footer (bottom right): "Portal: ✓ 3 sent today · 1 waiting · ⚠ 1 failed"; click
  it for the list and **Retry now**. A failed upload pops up a warning so the office can upload by hand.
- **Off until set up on each PC.** Admin → Portal upload: station id + key (from the portal admin),
  Test connection, turn on, Save. Stored only on that PC in
  `%LOCALAPPDATA%\PCC Shipping\portal_upload.json`; queue in `%LOCALAPPDATA%\PCC Shipping\portal_queue\`.
- Tests: `py -3.12 -m unittest tests.test_pdf_docs tests.test_portal_upload -v` (run from the repo folder).
