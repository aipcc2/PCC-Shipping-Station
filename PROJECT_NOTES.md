# PCC Shipping App — Project Notes

**App:** PCC Shipping App (internal window title: "PCC Shipping Station")
**Type:** Python / tkinter desktop GUI, packaged as a Windows SINGLE-FILE exe
**Build style:** ONE self-contained `.exe`, built from ONE source file (`app_logic.py`)
**Project root:** `...\PCC Shipping App\` — this file lives here; the code + build
files live in the `Scripts\` subfolder.

---

## 1. Folder structure

```
PCC Shipping App\                          <- project root
├── PROJECT_NOTES.md                       <- this file
└── Scripts\                               <- source + build files
    ├── app_logic.py                       <- THE ENTIRE APP. One file. Edit this, then rebuild.
    ├── PCC Shipping App.spec              <- PyInstaller recipe (single-file / onefile)
    ├── build_exe.bat                      <- one-click build -> dist\PCC Shipping App.exe
    ├── logo.png                           <- Packing List logo; bundled INTO the exe at build time
    ├── admins.json, operators.json        <- local reference copies (app reads the LIVE ones from \\PCC2)
    ├── PCC Shipping App-Old Reference.py  <- pre-2025 original source; archive only, safe to delete
    └── dist\
        └── PCC Shipping App.exe           <- THE BUILT APP: one ~36 MB file, nothing beside it
```

> The Packing List logo is baked into the exe: the **same** logo lives both as
> `logo.png` (in `Scripts\`, bundled at build time) and as base64 in the code
> (`_EMBEDDED_LOGO_B64`). Either renders the correct logo, so the exe needs no external
> `logo.png` at runtime. To change the logo, replace `logo.png` **and** re-encode it
> into `_EMBEDDED_LOGO_B64`, then rebuild.

---

## 2. What this is / architecture

The app is a shipping-station tool: pick a division + operator, look up a work
order on the `\\PCC2` share, attach CCA photos and serial numbers, and generate
the shipping paperwork (Serial list, Certificate of Compliance, Packing List)
as `.xlsx` / `.docx` files into the work-order folder.

It is packaged as **one self-contained Windows exe**, built directly from the
single source file `app_logic.py` (which ends with
`if __name__ == "__main__": ShippingApp().mainloop()`). There is no launcher and no
runtime-loaded module: PyInstaller analyses `app_logic.py`, bundles the Python
runtime + all libraries (openpyxl, python-barcode, Pillow, python-docx, tkinter),
and packs everything into the single `PCC Shipping App.exe`.

> To change behavior you EDIT `app_logic.py` and REBUILD the exe. (There used to be
> a "thin-launcher / onedir" build that let you swap `app_logic.py` without
> rebuilding; it was retired 2026-08-19 in favor of one simple file — see the
> changelog.) The heavy libraries are imported lazily inside functions; the `.spec`
> lists them in `hiddenimports` + `collect_all()` so the single exe always contains them.

---

## 3. Files in `Scripts\`

| File | Purpose |
|---|---|
| `app_logic.py` | **The entire app.** All behavior + the embedded logo live here. |
| `PCC Shipping App.spec` | PyInstaller recipe — single-file, windowed, built from `app_logic.py`. |
| `build_exe.bat` | One-click build. Run on Windows to (re)build the one exe. |
| `admins.json`, `operators.json` | Local reference copies; the app reads the LIVE ones from `\\PCC2`. |
| `PCC Shipping App-Old Reference.py` | Pre-2025 original source. Archive only — not built, safe to delete. |
| `dist\PCC Shipping App.exe` | The built app you deploy — one file. |

---

## 4. How to BUILD the exe

On a **Windows** machine with **Python 3.12** installed:

1. Double-click **`build_exe.bat`** in `Scripts\` — or run
   `py -3.12 -m PyInstaller --noconfirm --clean "PCC Shipping App.spec"`.
2. It installs/updates PyInstaller + the libraries, cleans old output, and builds.
3. Output: **`Scripts\dist\PCC Shipping App.exe`** — one ~36 MB file. (`build\` is
   PyInstaller scratch; safe to ignore or delete.)

Every change is a rebuild — code OR a new library. There is no "swap the file
without rebuilding" path anymore. If you add a NEW third-party library, also add it
to `collect_all(...)` / `hiddenimports` in `PCC Shipping App.spec` before building.

Currently bundled heavy libs: **openpyxl, python-barcode, Pillow (PIL),
python-docx** (plus tkinter and the standard library).

---

## 5. How to DEPLOY

Copy the single **`PCC Shipping App.exe`** wherever it needs to run — a station's
local disk, or the network share. **Nothing else goes with it** (no folder, no
`app_logic.py`, no `logo.png`, no `_internal`). Point desktop shortcuts at that exe.

To push an update: rebuild the exe, then copy that one file over the old one.

---

## 6. Offline capability (install on each station's LOCAL disk)

Offline mode only works if the app runs from the station's **local disk** — an exe
running from `\\PCC2` cannot launch when PCC2 is down.

1. Copy `PCC Shipping App.exe` to each station locally, e.g.
   `C:\PCC Shipping App\PCC Shipping App.exe`, and point the shortcut at that copy.
2. Run it once WHILE PCC2 is up — this caches the operator/admin rosters locally
   (`%LOCALAPPDATA%\PCC Shipping App\cache`) so names work during an outage.

During an outage the app opens (no freeze), shows the orange OFFLINE banner, and can
ship urgent orders to the local Outbox (`%LOCALAPPDATA%\PCC Shipping App\OfflineOutbox`).
After PCC2 returns, copy each Outbox shipment's four subfolders into the real WO folder
(its `_OFFLINE_SHIPMENT_README.txt` explains how), then delete the Outbox entry.

---

## 7. Key paths & default values (from `app_logic.py`)

- Production roots:
  - Industrial: `\\PCC2\PCC-I-Shares\Production\Industrial`
  - MAG - ITAR: `\\PCC2\PCC-I-Shares\Production\MAG - ITAR`
- App data dir: `\\PCC2\PCC-I-Shares\Production\Industrial\Shipping App`
  - `admins.json`, `operators.json` live here (auto-created if missing).
  - `History\` lives here too — one JSON record per shipment event (see 2026-07-22 change log).
- Admin login: id `admin`. Passwords are salted **PBKDF2** (legacy SHA-256 hashes
  still work and auto-upgrade on next login). **No password is hardcoded in source.**
  A fresh install (no `admins.json`) generates a random one-time password into
  `admin_initial_password.txt`. **Rotate the admin password via the ⚙ Admin panel
  after deploy** — the previously shipped default is publicly known.
- Output subfolders created inside a WO folder: `Shipping Pictures`,
  `Shipping Serials`, `Packing List`, `Certificate of Compliance`.
- Output filenames key off the **Engineering #** and date, e.g.
  `3156LF-7-15-26 serials.xlsx`, `...-COC.xlsx`, `...-Packing_List.xlsx`.
  Same-day repeats get a timestamp suffix (never overwrites).
- Logo: prints on the Packing List. The same logo is baked in TWO ways — `logo.png`
  (in `Scripts\`, bundled) and base64 in the code (`_EMBEDDED_LOGO_B64`); both are the
  correct logo, so no external file is needed at runtime. `_find_logo()` prefers a
  bundled / next-to-exe `logo.png`, else the embedded copy.
- Window: 880x750, min 720x500. The main screen scrolls (Canvas + scrollbar) and
  the Save buttons stay pinned in a bottom footer, so nothing is cut off on small
  screens.

---

## 8. Change log

- **2026-10-28 (v64 — HISTORY: ONE ROW PER SHIPMENT + ENG# SEARCH; RECEIVING WINDOW MAXIMIZES; exe rebuilt 10/28)**
  Parts (1) and (2) of the 2026-10-24 plan, built on the shipping PC (entry written 2026-11-07 from the code;
  live exe `\\PCC2\PCC-I-Shares\Production\PCC Shipping App.exe` = `Scripts\dist`, 10/28 3:07 PM).
  - **History** lists ONE row per shipment: `group_shipment_history(records, voided_ids)` groups the draft_save /
    close / delete events that share a `draft_id`; events without a draft_id, and every receive / backfill /
    void, stay as their own rows so nothing disappears. Columns: Status (Closed / Draft / Deleted / Voided /
    Received / Imported), WO, Eng #, P/N, Customer, Started, Saved, Closed, Serials. New **"Eng# contains"**
    filter. Double-click a shipment -> new `ShipmentDetail` window with the full save timeline (double-click an
    event there for the old `HistoryDetail` card). CSV export writes every event of the rows shown.
  - **Started by / when** is now recorded: `ShippingApp._started_at/_started_by` are stamped at WO load, saved in
    the draft (`save_shipment_draft(..., started_at, started_by)`), restored when a draft is resumed, and written
    on the draft_save / close History events. Older shipments fall back to their first save.
  - **Receiving window** is no longer `transient` (Dhruvil 2026-10-28 could not maximize it): normal
    maximize/minimize buttons and its own taskbar entry.
  - `write_received_xlsx`: when `_home_unknown` is set the error now says the job's folders can't be checked from
    this PC (Rebuild on a PC that can open them) instead of "no work-order folder known".
  Still open: Receiving placement with 2+ -970 WOs + the 3 failing test_receiving.py checks; the 3694
  template-folder fix. A stale 10/5 exe still sits at `...\Industrial\Shipping App\PCC Shipping App.exe`.

- **2026-10-27 (SHIPPING REPORT AUTO-UPDATE — runs inside the close, result shown + recorded; exe rebuilt 10/27)**
  Part (3) of the 2026-10-24 plan, built on the shipping PC (entry written 2026-11-06 from the code). Before: the
  update ran on a SILENT daemon thread — a report open in Excel, or the app moving on / closing right after the close,
  made it fail with no trace (Dhruvil: "last week I have to update manually"). Now `auto_update_shipping_report`
  runs SYNCHRONOUSLY in the close flow (progress step "Updating the Shipping Report…", history-only, a few seconds)
  and RETURNS an outcome: `('updated', path)`, `('none', …)` (no report yet — opt-in: build it once in Receiving),
  `('unknown', …)` (history / folders not reachable from this PC — nothing written) or `('failed', …)` (the file is
  open on another PC, or a share error). The close summary shows "Shipping Report: updated ✓" / the reason, and
  `_record_report_update` stamps the close's History event with `report_update: {status, detail, ts}`. Jobs that
  are not test-and-return show no line. Still open from 10/24: History Eng# search + one row per shipment; the
  Receiving placement question and the 3 failing test_receiving.py checks.

- **2026-10-24 (HAND-OFF to the shipping PC — findings + an APPROVED-IN-PRINCIPLE plan, NOTHING of it built yet)**
  Dhruvil (2026-10-24): "In the history I should be able to search it with engineering number" · "it is very
  confusing as there are multiple entries for one shipment with many in drafts and many in closed … there should be
  only one entry for any shipment either in draft or closed, Once shipment close it should remove from draft. and any
  entry should show starting time with whom started and closing or saving draft by whom and when" · "in receiving
  are you sure it is auto updating everytime when we add new shipments? … last week I have to update manually" ·
  "Confirm before you build". Then: "I can start from shipping pc from here on wards to give you access to mag-itar
  file … save everything till here and I will copy to shipping pc".
  **Findings (read-only, desk PC):**
  - History has filters WO / Serial / Customer / Operator / Division — no Eng#.
  - Every Save Draft writes its OWN `draft_save` event (Sep-Oct: 210 draft saves vs 93 closes); the viewer lists
    events, so one shipment shows as up to 8 draft rows + its close (WO 16582.1: 38 draft rows + 6 closes = 7
    shipments). The shipment identity already exists: `draft_id` is on every draft save and on 88 of 93 closes. The
    draft FILE is deleted at close (0 of 4 draft files on Industrial belong to a closed shipment) — the clutter is the
    viewer only. "Started" is not recorded anywhere today (first draft save is the best proxy).
  - Receiving auto-update: `auto_update_shipping_report` IS started after every ONLINE close of a …900/…970 job
    (18 such closes 10/5-10/22, all on SAHIL-NEW, exe v63 built 10/8). But it is SILENT: (a) a report open in Excel
    on any PC -> the write raises and is swallowed; (b) it runs on a daemon thread, so closing the app right after a
    close kills it; (c) nothing records the outcome. Could not check the actual report files from the desk PC (MAG -
    ITAR not reachable here) -> verify on the shipping PC with `_handoff_to_shipping_pc\tools\check_report_autoupdate.py`.
  **Plan proposed (Dhruvil's "Go" still pending — he moved to the shipping PC instead):** (1) "Eng # contains" filter
  in History; (2) ONE row per shipment grouped by `draft_id` — status Draft / Closed / Deleted / Voided, started
  (time + operator), last saved (time + operator), closed (time + operator + station); the row's detail shows the
  full save timeline; record a real `started_at` / `started_by` at WO load from now on; (3) Shipping Report update as
  a step INSIDE the close (not a daemon), result in the close summary ("updated ✓" / "NOT updated: open in Excel …")
  and recorded in History. Open question asked: keep every save reachable only in the row's timeline (recommended) or
  also a "Show every save" checkbox.

- **2026-10-19 (RECEIVING — Dhruvil's 2026-10-06 change documented + the "can't check from this PC" safety stop restored)**
  Dhruvil confirmed (2026-10-19): "Yes I did change on 10/6 and shipping reports now goes to 970 wo instead of 900 ·
  whatever you saw in the app right now is latest". **The 2026-10-06 change (made by Dhruvil, no entry was written
  then):** `build_receiving_state` homes the Shipping Report in the -970 COATED WO folder for EVERY job with a coated
  side (`if cw or cbf:` first — reverses "-900 tree whenever the job ever shipped for testing"); the sibling-tree
  migration moves an existing -900 report into the -970 home on the next write; a job with no coated side still homes
  in the -900 tree; and `write_received_xlsx` re-derives coated / in-house from the MERGED 'Coated & shipped' sheet
  (the sheet is authoritative: a board coated only in the COC folder counted — 4096: sheet 461 vs summary 456).
  **Side effect found + fixed (approved, 2026-10-19):** putting `if cw or cbf:` first made the later branches
  unreachable, including the TRI-STATE safety stop — on a PC that cannot list the -900 job folder holding the
  existing report the app answered "no report" for the -970 WO, so a Build / Save could create a SECOND report. Now
  that branch asks `_coated_tree_holds_report` first: None -> wo_path None + `_home_unknown` ("could not check from
  this PC", nothing written), as before 10/6. **Tests:** test_receiving.py updated to the 10/6 rule (fixture creates
  the coated WO 16154.1; 8 rewritten checks; r5 / r6 unreachable-folder checks pass again): 449 ok, **3 still
  failing** (r4 / r4 / r5 "no ping-pong") + the run stops at the clean-up after them. Those three show the report
  FOLLOWS the newest -970 WO (17002.1 -> 17003.1 -> rev-twin 17010.1; always one file, rows right). A "stay where it
  is" resolver change was tried and REVERTED: it breaks five 9/23 tests where coated-only jobs deliberately move the
  report into each new -970 WO on its first close and rename it on a label change (970C -> 970D). **OPEN for
  Dhruvil:** "Follow the newest -970 WO" (as coated-only jobs since 9/23; then just update the 3 tests + the clean-up)
  vs "Stay where it is" (changes label-change / rev-folder behaviour too). His last answer (before the full picture)
  was "Stay where it is"; the corrected question was interrupted.
  **Exe:** the floor exe is still the 2026-10-08 build — it has the 10/6 change but NOT the 10/19 packing label or
  this safety stop.

- **2026-10-19 (PACKING LABEL — Apex's #lbl-pack, 3 x 2 in, in the Shipping Labels and Slip popup)**
  Dhruvil: "Will you be able to add packing label section from this to our shipping app?" Decisions: in the printer-icon
  popup; "I need to replicate whatever is in apex" (one label per print, QTY starts at 1); 3 x 2 in; printer "3 x 2
  Shipping"; "Go - build it". Apex 3.11 unchanged since the 9/24 frozen copy.
  **What changed (`Scripts\app_logic.py`):** `build_pack_html` / `_pack_sheet` = Apex's #lbl-pack markup (the ITAR/Industrial
  buttons and WO/Eng pickers, screen-only in Apex, left out; inputs as static spans) with Apex's CSS verbatim
  (`_PACK_APEX_CSS` = Apex lines 42 + 192-212) and its existing print rules; `@page{size:3in 2in;margin:0}`; fields P/N
  (Customer P/N), QTY (`pack_qty`, starts "1"), CAGE (`pack_cage`, "1WXP2", ITAR/MAG only), PO, D/C (YY+ISO week, Apex
  calcDateCode, `pack_datecode_yyww`; ITAR/MAG only), DATE (today MM/DD/YYYY; Industrial only), WO#/Lot, each with Code128
  `code128_svg(v, id, 1, 12)` (JsBarcode width 1 / height 12; `code128_svg` gained module_width/height params). The mode
  comes from the job's DIVISION (`pack_mode_for`; Apex makes you click) - the one intended difference. Form code
  `pack_form_code` = ENG-WO-PACK-YYYYMMDD-RAND4 (in the markup, hidden in print as in Apex). Third page kind in slip_data
  (`_SLIP_KINDS`: printed/outputs/pages "pack"/"pack_html"; date + date code frozen at print): `write_slip_files(...,
  "pack")` files `<WO>\Shipping Slip\<eng>-<m-d-yy>-Packing_Label.html`; `slip_print_pages(...,"pack")` opens a
  self-closing print copy (`Print_Packing_Label`); re-emit, History (Open Packing Label, rows, Output files), close-record
  outputs, offline README, status chip all carry it. Popup: "Packing Label" section (label type from the division, Packing
  QTY, CAGE on ITAR/MAG, Print Packing Label button); the slip's QTY-vs-serials prompt is skipped for it.
  **Verified:** test_slip_labels.py section 14 (16 checks) incl. element-sequence FIDELITY vs Apex's #lbl-pack and its CSS
  verbatim; slip suite 308 checks; headless-Chrome print-to-PDF = one 216 x 144 pt (3 x 2 in) page per mode, form code not
  on paper; hidden Tk smoke of both modes.
  **Found, NOT changed:** test_receiving.py now fails at "wo_path = NEWEST testing WO folder" because of the 2026-10-06
  Receiving edit made outside this session (report homes in the -970 coated WO folder; coated sheet authoritative) - that
  edit has no changelog entry here and the test still pins the old rule; it fails identically on the pre-10/19 source.
  **Deploy:** `build_exe.bat`; first run: printer icon -> Print Packing Label on a MAG-ITAR job (CAGE + D/C) and an
  Industrial job (DATE), pick the 3 x 2 Shipping printer, Margins None; scan the barcodes.

- **2026-10-05 (REV ON THE PAPERWORK = THE BOM REV — Monday's BOM Rev first, the folder name's BOM rev as the fallback; dotted / hyphenated revs read whole)**
  Dhruvil asked which rev Monday fills in. Answer (read from Monday's own column metadata): the app reads column
  `lookup_mm06bttc`, titled **"BOM Rev"** on the Production Orders board — a mirror of the Engineering board's BOM Rev
  through its "Engineering" relation column. The board also mirrors Assy Rev (`lookup_mm065m3t`), CC Rev
  (`lookup_mm68fshn`, the coating rev) and PCB Rev (`lookup_mm06bed5`); the app reads none of those. Until now the
  folder name's rev WON (the LAST rev in "Customer PN [PCC Eng - BOM Rev - Assy Rev]" = the Assy rev when filled) and
  Monday only filled a blank; and dotted revs were cut short ("Rev. 2.02" -> "2", "Rev. 1.1.10" -> "1" — 8 of the 45
  Industrial job folders). Dhruvil (verbatim): "It should always carry BOM rev. If monday has BOM rev column then it
  should always prioritize that to foldar name".
  **What changed (`Scripts\app_logic.py`):** `_merge_monday` — a real Monday BOM rev replaces the folder's
  (`monday_bom_rev`: '', '-', '--', 'N/A' = none; a typed 'Rev' word dropped, 'Rev 03' -> '03'; two DIFFERENT values
  from several linked Engineering items = none, never a guess); `parse_wo_path` now takes the folder's BOM rev = the
  FIRST rev in the name (`bom_rev_from_inside`) as the fallback; `_revs_in` reads dots / hyphens inside a rev ('2.02',
  '1.1.10', date revs like '5-4-23'; the '-' placeholders still skipped) — `rev_from_inside` (the last rev) kept for
  compatibility; the load status line says "Monday pre-filled PO …, BOM Rev …" only when that rev is the one used;
  the slip pre-fill's last-resort Monday rev goes through `monday_bom_rev`. The rev feeds only the COC's Rev default
  and the Shipping Slip / Labels Revision (and its barcode). Customer / Eng# / P/N stay folder-first.
  **Verified:** live read-only comparison of all 45 Industrial job folders against Monday: the folder's BOM rev (new
  parser) equals Monday's BOM rev on 33, differs on 0; 1 folder has no rev (Monday's '8-24-21' used); 3 have no rev
  anywhere; 8 Eng#s are not on Monday (folder rev used). Before the two format fixes 3 differed (a date rev '5-4-23'
  cut to '5'; Monday's 'Rev 03'). New section 13 in test_slip_labels.py (8 checks); all 7 suites green, 934 checks;
  hidden Tk smoke OK.
  **Also found (2026-09-29, NOT changed — offered):** searching 3694 in Industrial failed because Monday's active WO
  16655.1 had no folder: the Rev C job folder `198 Prism\PC1060C_3694 Rev C\Production Data\` still held the empty,
  un-renamed template folder "Work Order Number" (also `184 Wabtec-Trimble\AtoD-1-082218 -2429-1` and
  `379 Liber8Now\GSR10 [4033 Box - Rev 1.1.11]`), and with a Monday hit the folder search looks for Monday's WO only
  ("Found in Monday, not on the share"). Offered, not built: name the placeholder folder in that message, fall back to
  the typed number, never offer a template folder as a WO.
  **Deploy:** `build_exe.bat`. Check: a job whose folder name's last rev differs from its BOM rev (e.g. a 3486LF job,
  folder "Rev. A Rev. 1") now shows A on the COC and the slip; a dotted rev (e.g. 3414 "2.02") shows whole.

- **2026-09-25, round 2 (SHIPPING LABELS AND SLIP — Dhruvil's first-run feedback: Chrome tab closes itself, one page for Print both, barcodes under their values, pick-only dropdowns + shared box sizes, Hand Carry, printer-icon button bottom-left, Monday PO second lookup)**
  Dhruvil built the exe (9:40) and ran a 3-box test, then (verbatim, abridged): "Customer po is in automatically
  fills up from monday" (= NOT filling) · "barcode is not aligned with actual value … not looks directly under the
  value" · "when we click on prepared by or ship via it is not drops down untill we click on down arrow" · "Need hand
  carry option in ship via" · "When I click on print both somehow it only prints 3 shipping labels and not prints 1
  shipping slip" · "Can we remove that chrome pop up … When we click on print it should only print without going in
  chrome … If you can remove that chrome pops up without any compromisation" · "move that button … at bottom left
  corner. Keep only logo of Print … separate from everything else. Also title should be Shipping Labels and Slip".
  Decisions (his answers): dropdowns "Pick only but there should be option to add new option in the box size as we
  keep getting new sizes every time"; printing: "opening in chrome is fine but once we click on print it should close
  automatically that chrome tab for both" -> **Chrome stays, the tab closes itself** (a silent direct-to-Zebra path was
  prototyped and offered — headless Edge print-to-PDF + the Windows PDF engine at 203 dpi + GDI, layout identical to
  Chrome's print within 2 dots — and parked; see "Not built" below).
  **What changed (`Scripts\app_logic.py`):**
  - **PO from Monday — second lookup.** Monday matches a WO number EXACTLY (`16588.1` -> PO USPO61707; `16588` ->
    nothing) while the folder search finds `16588.1` from `16588`, so a short WO (or an Eng# whose Monday picker was
    skipped) loaded the job with no Monday data -> blank PO / Customer P/N on the slip, COC and Packing List. Now, when
    no usable hit is in hand after the folder is chosen, `_load_workorder` asks Monday once more with the folder's own
    WO number (`monday_folder_wo`) and takes the answer only when exactly ONE item carries it
    (`pick_monday_hit_for_folder`); skipped when that number is what was typed. The old validation rule moved
    unchanged into `monday_hit_matches_folder`. Runs before ConfirmDialog, which still shows the PO first.
  - **Barcodes under their values** (the ONE deliberate difference from Apex's label): `_slip_barcode_row` gives each
    `.barcode-row` an invisible copy of its row's title (same classes/text -> same width, + the row's 4 px gap) and a
    box that centres the barcode in the space left — exactly under the value, which Apex centres in that same space.
    A barcode wider than that space keeps its full size and slides left (box min-width = its own width); only one
    wider than the whole row shrinks, as Apex's max-width:100% does. Four app-side rules in `_SLIP_PAGE_CSS`; the
    barcodes themselves are untouched. (A re-emit of labels printed before this build renders the new alignment.)
  - **Chrome tab closes itself after printing** (`_PRINT_THEN_CLOSE_FN/_JS`): `window.print()` then `window.close()`
    one second later, only when print() held the script > 400 ms (Chrome/Edge hold it while the dialog is up; a
    browser that doesn't, or printing blocked, leaves the tab open as before). Chrome lets a page close a tab that
    holds only that page — how os.startfile opens it. Cancel closes it too (the page can't tell Print from Cancel).
    Used by the 4x2 board labels and by the slip/labels **print copy**: `slip_print_pages` still files the record
    page(s) in `<WO>\Shipping Slip\` (unchanged names, auto-print, stay open from History ▸ Open) and now opens ONE
    throw-away print copy from `%TEMP%\PCC Shipping App - Print` (`<eng>-<m-d-yy>-Print_Slip|Print_Labels|
    Print_Slip_and_Labels.html`); day-old print copies and `labels4x2_*` renders there are tidied. If the print copy
    can't be written the filed page(s) open as before.
  - **Print both = ONE page** (`build_slip_labels_html`: the slip, then every box label, 1 + N 4x6 pages) — Chrome
    shows one print dialog at a time, so the second tab (the slip) never printed by itself. Builders refactored:
    `_slip_sheet` / `_labels_sheets` + `build_slip_html(data, print_copy=False)` / `build_labels_html(...)`.
  - **Dropdowns pick-only like Apex's `<select>`s** (a click anywhere opens the list): Prepared By (typing only when
    the roster is empty), Ship Via (`-- Select --` + list), every box Size (`--` + shared sizes + `+ Add new size…`);
    the mouse wheel over them scrolls the form and never re-picks (`_no_wheel`). **Hand Carry** added to Ship Via.
  - **Shared box sizes**: `+ Add new size…` -> type once (normalised, e.g. `18 X 12 x 6` -> `18x12x6`) -> added to
    `\\PCC2\...\Shipping App\box_sizes.json` beside operators.json (re-read first, atomic write, cached locally for
    outages; offline -> used for this shipment only, operator told). Apex's list until the first size is added.
    **Admin ▸ Box sizes** tab to add / remove (never the last one). Functions `normalize_box_size`,
    `box_size_problem`, `load_box_sizes`, `save_box_sizes`, `add_box_size`.
  - **Button**: the footer's text button is now a printer-icon-only button (🖨) alone in the bottom-LEFT corner
    (3-column grid, Save Draft / Save and Close still centred), name on hover (`_HoverTip`: "Shipping Labels and
    Slip"); grey with no WO, blue when ready, green once the slip or labels were sent to print. Popup title + header
    + its error boxes: "Shipping Labels and Slip".
  **Verified:** all 7 suites green (926 checks; test_slip_labels.py 284 — new sections 3b print copies,
  5b barcode alignment, 11 box sizes, 12 Monday second chance; fidelity comparison documents and undoes the one
  alignment difference); a hidden Tk smoke test (every window withdrawn before it could map; share + Monday
  stubbed) built the real popup, footer and Admin tab and exercised Add-new-size; headless-Edge print-to-PDF of the
  combined page = 4 pages of 4x6 for 3 boxes; before/after renders at 203 dpi sent to Dhruvil; two independent reviewers (print pipeline / Tk + storage + Monday) — seven confirmed defects, all fixed and re-verified: (1) a barcode sliding left sat 6 px from the paper edge, under Code 128's quiet zone -> `.bc-box` keeps Apex's 8 px either side (min-width = svg + 16; slid = 14 px from the edge, too-wide = Apex's 356 px); (2) a test name with a non-cp1252 character crashed a redirected run; (3) Admin ▸ Box sizes wrote back the list it read on opening, erasing a size another station added meanwhile -> add / remove go through `add_box_size` / new `remove_box_size`, which read the share FRESH (`_read_shared_box_sizes`, never the offline cache; an unreadable file is never overwritten); (4) '+ Add new size…' was below the 10-row fold and clipped -> Size dropdown 20 wide, as many rows as entries (<= 25), header widened to match; (5) the second Monday lookup also ran after the first one FAILED (two timeouts) -> `monday_lookup(status=)` / `_monday_last_ok`: no retry after a failure, and the answer in hand (an Eng# lookup) is tried before asking again; (6) after the add-size dialog the keyboard focus sat on the popup, so Tab + Return printed the slip -> focus back on the Size row in every branch; (7) minor: the popup re-probed PCC2 (2 s) per open / add while the app already knew it was offline -> `online=` hint passed through. A stubbed run of the real `_load_workorder` (no Tk, no network) confirmed the four Monday paths; the reviewers' own repro scripts now pass.
  **Not built (parked):** printing with no browser at all. Prototype facts for later: the Zebra reports 203 dpi,
  a 4.000 x 6.000 in page, printable area = full page, offset 0 (also through Remote Desktop Easy Print); headless
  Edge `--print-to-pdf` + Windows.Data.Pdf (PowerShell WinRT) at 812 x 1218 reproduces Chrome's print layout with no
  dot > 2 px off; a screenshot render (`--screenshot`, DSF 203/96) drifts up to 3 dots on the slip (font-metric
  rounding) — don't use it. The one open question was the driver's halftoning of the grey lines / form code.
  **Deploy:** `build_exe.bat` (ships this + everything since 09-16). First run: 🖨 -> 3 boxes -> Print both -> ONE
  dialog prints slip + 3 labels, the tab closes itself; a Size -> + Add new size… -> it appears on every station;
  scan the label barcodes; load a WO typed short (e.g. 16588) -> PO fills.

- **2026-09-25 (SHIPPING SLIP + SHIPPING LABELS — Apex-identical 4×6 pages printed via the browser; Dhruvil builds the exe)**
  Dhruvil (2026-09-24): "Will you be able to copy shipping slip and shipping label section in our app? I want
  exact format. Since we have monday connect you can autofill all the data from monday." Reviewed
  `\\PCC2\PCC-I-Shares\Production\PCC_Apex_Charlie_3.11.html` (3,606 lines; a frozen byte-identical copy now
  lives in `_handoff_to_shipping_pc\tests\fixtures\` — the fidelity test compares against it). Apex's Monday
  lookup reads the SAME five columns our `monday_lookup` already reads (WO/PO/Eng/Rev/Customer + item name =
  Customer P/N), so the autofill data was already in the app. His decisions: (1) reuse Apex's HTML/CSS verbatim
  and print from the browser (pixel-identical) — not an Excel/Word approximation; (2) a button on the loaded-WO
  screen, usable any time before the close, reprints allowed; (3) **no Google Sheet** — the shipment's History
  record only; (4) the 4×6 printer is the **Zebra ZP 450 CTP**.
  - **Button** `🚚 Shipping Slip & Labels…` beside Save Draft / Save and Close (enabled while a WO is loaded;
    turns `✓` once a page printed for the loaded shipment). **`SlipLabelsPopup`** (modelled on the Packing List
    popup): Customer, Customer P/N, Customer PO #, Work Order/Lot, Engineering #, Revision, Qty Shipped (= scanned
    count), Date, Date Code (**WWYY**, as the Apex slip shows it), Prepared By (operator), Notes — pre-filled from
    the folder + the Monday hit + COC/Packing data; then **Boxes** (count spinbox → rows: Size from Apex's list
    `5x5x5 … Pallet, Other`, Qty, Weight lbs; live totals — typed ONCE, feeding both the slip's table and each
    label's weight; Apex makes you type them twice), Customer Parts No/Yes, FAIR / SI / PPAP / C of C, Ship Via
    (Apex's list), Label notes. Buttons 🖨 Print Slip · 🖨 Print Labels (N boxes) · 🖨 Print both · Save · Cancel.
    Required fields exactly as Apex (slip: Customer, WO, Qty, Prepared By; labels: Customer, WO, Qty per box); a
    COC-style qty-vs-serial-count warning (never blocks).
  - **Pages** (`build_slip_html` / `build_labels_html`, module level, Tk-free): Apex's `#lbl-ship` / `#lbl-shipl`
    markup element-for-element and its CSS rules VERBATIM (the test proves tag/id/class/inline-style sequence
    equality and rule-by-rule presence), `@page{size:4in 6in;margin:0}`, one 4×6 page per box on the labels file
    (Apex's "Print All" in one job), the footer QR (`<WO>-1`, 38 px, level M) via the embedded qrcode.js, the
    form codes in Apex's format (`ENG-WO-INITIALS-SLIP-YYYYMMDD-RAND4` / `…-XX-LABEL-…`), a non-printing hint bar
    naming the Zebra ZP 450 CTP, auto-print like the 4×2 page. **Barcodes:** python-barcode's SVG writer was NOT
    used (mm units, no viewBox, different Code128 A/B/C switching → different bars); instead the bundled package's
    Code128 pattern table drives a straight port of JsBarcode 3.11's CODE128-auto selection and SVG DOM
    (1.3 px/module, 26 px), so the bars are identical to Apex's (`code128_modules` / `code128_svg`, decoded and
    pinned by tests). Substitutions (documented in the test): the `.box-gate` / `.label-nav` blocks are not
    printed; text inputs/selects become `<span>`/`<div>` with the same classes and inline styles; checkboxes/
    radios stay real inputs frozen with `onclick="return false"`; blank fields keep Apex's control height via one
    app-side `:empty::before{content:"\00a0"}` rule (the strict audit's one major finding — blank Weight /
    Ship Via bands printed ~15 px short).
  - **Files:** `<WO folder>\Shipping Slip\<eng>-<m-d-yy>-Shipping_Slip.html` and `…-Shipping_Labels.html`
    (`SLIP_FOLDER`, created as a LEAF only inside an existing WO folder; `unique_file_path` + exclusive create —
    a same-day reprint gets `_HHMMSS`, never overwrites; new form codes per print). WO folder unreachable →
    `%TEMP%\PCC Shipping App - Print` with a "printed, but NOT saved to the share" message. **History:** the form
    travels like `coc_data` — `self.slip_data`, drafts (`slip_data`), `close_record["slip"]` (+ mirrored
    `outputs.slip_html/labels_html`), the offline close/Outbox path; `HistoryDetail` shows the slip block with
    "Open Shipping Slip / Labels" buttons (window now 820 px, the Open buttons in their own row — for all
    records); 📄 Re-emit paperwork regenerates missing slip/label pages from the record, never overwriting; the
    "switch work order?" guard also protects an unsaved slip form. Save Draft still needs photos or serials.
  - **Known, accepted differences from Apex (Dhruvil to confirm):** (a) **typeface** — Apex's CSS prints Roboto
    (Google Fonts, online); the page ships with `SLIP_FONT_LINK = False` (no network) and prints Arial unless
    Roboto is installed on the PC — everything below the big Qty row then sits ~4 px lower than an online Apex
    print. Exact match: install the free Roboto font on the print stations (offline-safe), or set
    `SLIP_FONT_LINK = True` (needs internet at print time, falls back to Arial after 2.5 s). (b) Apex prints the
    grey placeholder text of EMPTY inputs (`-- Select --`, `qty`, `lbs`) and native control chrome (select
    chevrons, calendar/datalist glyphs); the app prints blanks — filled fields are identical; sub-millimetre
    offsets on the label's PO field and the slip's Ship Via band come from that chrome. (c) Form-code date =
    local date (Apex: UTC — differs after ~7 pm Central). (d) WWYY from the slip's Date (Apex: today).
    (e) Labels require a qty per box (Apex's Print All prints without validation). (f) One box row is pre-filled
    with the qty (Apex's table starts blank).
  - **Verification:** 2 implementers → 2 adversarial review rounds (8 lenses × 3 verifiers; 39 confirmed defects
    fixed, each mutation-tested) → strict blocker/major-only audit: pass 1 found the blank-field height issue
    (fixed), pass 2 **clean** (0 findings, 48 minor notes = the items above). The fidelity lens printed the LIVE
    Apex page (filled through its own JS) and the app's pages to PDF with headless Edge and pixel-diffed them:
    identical apart from the font/chrome items in (a)/(b). New suite `test_slip_labels.py` (226 checks) — all
    seven suites green, 867 checks. Tk widgets were verified headless only (no GUI launched): a two-minute
    click-through on first launch is advisable.
  - **Deploy:** `build_exe.bat` (the exe on the floor is still the 2026-09-16 build, so this ships together with
    the 2026-09-23 opt-in Shipping Reports). First launch: load a WO → 🚚 button → 2 boxes → Print both → check
    the two pages on the Zebra next to an Apex print of the same job; History → the record → Open Shipping Slip.
    Optional: install Roboto on the print stations (see (a)). The offline Outbox README still says "four
    subfolders" (a printed slip adds `Shipping Slip`) — wording only.

- **2026-09-23 (IN-HOUSE TESTING → OPT-IN SHIPPING REPORTS + COATED-ONLY JOBS; Dhruvil builds the exe)**
  *(This desk copy of PROJECT_NOTES is missing the 2026-09-09..16 entries written on the shipping PC —
  "Shipping Reports" rename, Note column, auto-update after close, Rebuild ALL, re-emit paperwork. Copy
  that file over this one rather than re-writing them.)*
  Dhruvil: "Now we are not shipping any boards for testing. They sent test fixture in house for us to
  test here … we have to still track whatever we shipped for coated boards." Boards are serialized
  **-970 from the start** (one in-progress batch still carries -900 serials and ships as -970 with the
  same seq — no -900 shipment, no receive). His decisions: (a) the app does **NOT record in-house test
  results** — it tracks coated shipments only; (b) boards from the old flow are **still at the customer**
  (History 2026-09-23: 4065=52, 4041=23, 4094=7, 4092=3, 4096=2, 4095=2, 4102/4103/4104=10 each) so the
  Receiving flow stays exactly as it was for them; (c) a job may have a -900 WO folder, a -970 one or both
  ("depends on the job") — the report lives in whichever EXISTS, never fabricated. Live-History pre-flight
  found that **coated-only jobs already exist** — Anduril 4006 `125-1030-970C` (21 boards), 4007
  `125-1031-970D` (81), 4008 `125-1032-970C` (49), 4023 `131-3722-970A` (68): the coated P/N can carry a
  REV LETTER after 970 (PCC's label; serials embed only the digits: `000851-1251030970-C00`,
  `000945-1251032970/C-C`), and their 8/27 "Anduril test damage" closes ship -900-labelled boards under the
  coated WO (4006×8, 4007×3, 4023×3). One record (4016LF, 2 boards, 2026-07-30) has Eng `N/A` + pn = the
  folder name (pre-v58 bug) and cannot be tracked until corrected.
  - **Q1** a coated-only job's file and header use the REAL label: `125-1030-970C Shipping Report.xlsx`,
    A1 `SHIPPING REPORT — P/N 125-1030-970C` (a job with any -900 history keeps the -900 name — also when
    a Rebuild re-homes its report into the -970 WO folder).
  - **Q2** coated-only jobs ARE listed in the Receiving dropdown, tagged `(in-house tested — coated only)`;
    scanning a board into one stops (`Nothing to receive here` — Skip / Supervisor override…).
  - **Q3** -900-labelled boards shipped under a coated WO ("Case B") go on the **Coated & shipped** sheet
    with the Note `scanned with a -900 testing label` — never into shipped / still-at-customer;
    `has_testing` is LABEL-based (a -900-labelled close/import or a genuine receive), never inferred from
    serial form. Decision (a) 2026-09-23 (taken by Claude, flagged): a receive whose serials are all
    coated-form and/or all Case-B (a supervisor override on a coated-only job) does NOT flip the job to
    "testing".
  - **OPT-IN (verbatim)**: "Don't make any reports of any jobs until operator specifically clicks and
    builds it manually for first time. As we only care about current jobs; most of the jobs are old. So
    anytime we rebuild you can show all the jobs first but create reports only for those user selects
    and do it manually. Once report is created app needs to update that report only with everyday
    shipping updates." → **the report FILE's existence is the switch** (no registry file).
  What changed (`app_logic.py` 9,414 → 11,756 lines; test_receiving.py 223 → 526 checks, six suites green,
  641 checks total):
  - **Opt-in.** `_report_lookup(state) -> (has, path, found)` / `_report_exists(state)` (TRI-STATE: True /
    False / None = "can't check from this PC"; counts the current name, the legacy name, every spelling the
    job ever carried, a same-name copy in another WO of the job, the sibling tree; a copy under `Archive\`
    only as a merge base, never as opt-in). `auto_update_shipping_report` (after every close, daemon
    thread) and the Receiving Save are **UPDATE-ONLY** — they never create a file or the `Shipping Reports`
    leaf; Save says "No Shipping Report exists for this job yet … Use 🔁 Build report". Per job the button
    reads **🔁 Build report** (first time, with a confirm) or **🔁 Rebuild report**; chips show
    `📄 Report: <path>` / "No report for this job yet …" / "Can't check … from this PC". **🔁 Rebuild ALL
    reports → `ReportPickerDialog`**: every job listed (Eng#, P/N + tag, last shipment, status `has report`
    / `no report` / `report archived — rebuild re-homes it` / `can't check from this PC` / `no -900/-970
    work-order folder found` / `another job's file at its report location`), rows with a report pre-ticked,
    Select all / none / with report, **Rebuild selected (N)** builds ONLY ticked rows (`_collect_report_rows`
    → `_rebuild_report_rows`: fresh walk + folder scan per ticked job), summary counts rebuilt / created /
    folder-truth / coated-only / moved / not ticked / skipped. If a user deletes a report the app stops
    touching that job (the chips name any older copy that still keeps it opted in).
  - **Coated-only jobs.** `list_testing_jobs` registers a job seen only through -970-side closes/imports
    under its canonical -900 key (`testing: False`, `coated_label`), letter-tolerant via `_tr_pn_parts` /
    `_tr_testing_label` / `_tr_embed_keys` (core digits + side; `125-1030-970C` → key `125-1030-900`, embed
    keys `1251030900` / `1251030970`); real -900 entries keyed by (Eng#, core+side) so spelling twins
    (4065 dashed vs flat) are ONE entry whose spelling follows the newest CLOSE. `build_receiving_state`:
    home precedence **-900 close > -900 backfill > -970 close > -970 backfill** (`home_side`), real label(s)
    captured (`coated_pn_label(s)`, `testing_pn_labels`), `testing_label_seen` / `hist_has_testing`
    (snapshotted — a folder-truth Rebuild can never rename the file), Case-B routing on closes AND
    coated-labelled imports (`coat_label_mismatch`), the 4065 direction (-970 serials under a -900 label →
    coated) untouched. `receiving_summary["has_testing"]`; `_report_pn` names the file.
  - **Placement.** `_resolve_received_wo_dir`: fast path unchanged; then the home tree — the recorded WO by
    name, else the WO that already HOLDS the report (`_sibling_wo_holding_report` / `_wo_holds_report`),
    else the newest live WO (a true first placement); the sibling (-900↔-970) tree only when the home tree
    has no live WO, and there it is the holder or — ONLY on a manual Build/Rebuild (`allow_tree_switch`) —
    the newest WO (`tree_switched`; the popup says the report moved); an unlistable job folder → refuse
    ("can't check"), never a guess. `_stale_received_candidates` spans both job trees, every report-name
    spelling × both folder names, `Archive\` one level as base only; `_under_archive` — nothing under an
    archive is ever deleted; `_job_dir_is` — the Eng# must stand as a whole token in the job-folder name
    (4016 ≠ 4016LF); `_head_names_eng` — a copy whose `Engineering #` head names another job is never
    merged, deleted or used as a base (a user note after the Eng# is ignored). First-time builds refuse
    when the copy scan was partial ("could not confirm that no copy … exists elsewhere").
  - **Wording.** `sent without testing` → **`not shipped for testing`** everywhere (old cells upgrade,
    placeholder-to-placeholder, no alert); coated-only heads: A3 `In-house tested — not shipped for
    testing · Coated boards shipped (P/N <label>): N`, RCV/IH/OUT subtitles "not applicable (nothing
    shipped for testing on this job)" with a `(none — …)` row, file opens on **Coated & shipped**; mixed
    jobs gain `(k not shipped for testing)` after `Coated boards on record: N` (4065 shows 50, 4041 1 —
    the direct-coat boards, counted per board); Note column carries the Case-B tag and the operator's
    note; only one sheet is left `tabSelected` (Excel no longer opens the file with sheets grouped).
  - **Safety nets.** Three blocking stops in Receiving (ReceiveIssueDialog, Skip / Supervisor override…):
    `Nothing to receive here` (coated-only job), `Not in shipping records` (a board never shipped for
    testing — was accepted with a ⚠ before), `Coated shipment on paperwork only` (a Case-B board known only
    from folder paperwork); an override answers only the question it was asked (`_dup_checks_due`,
    `_pn_check_due`), so a supervisor may see two stops for one board. ⏪ Import older shipments warns and
    confirms when the scanned serials belong to the OTHER side of the -900/-970 pair. 🔁 Rebuild's heads-up
    adds coated COC-vs-Packing-List disagreements and `coated board shipped N× — check records`; folder-side
    Case B: a -900-form serial in a coated-labelled file (or a label-less file under the -970 tree) is a
    coated row with the Note.
  Pre-deploy checks run against the live History from the desk PC: P/N spellings consistent (0 mismatches);
  0 -970-labelled imports carrying -900 serials; no junk-labelled close on 4006/4007/4023; every job-folder
  name on record satisfies the whole-token Eng# rule; the 4 -900 imports whose `production_data` is a -970
  tree (4041, 4092×2, 4065) have no placement effect (their -900 closes outrank them).
  **How it was built/verified:** 3 implementers, then **7 adversarial review rounds** (8–9 lenses each,
  every finding verified by 3 independent agents; ~100 confirmed defects fixed, every fix mutation-tested —
  reverting it makes a test fail). The rounds stopped converging (each still 10–15 mostly-minor items, and a
  round-6 fix regressed multi-token Eng#s until round 7), so the stop criterion was a **final strict audit**
  (9 lenses, blocker/major only, unanimous 3-of-3 verification, real folder/serial shapes only, incl.
  BEFORE-vs-AFTER differential builds on 4092/4065/4094 shapes and a 39-step headless end-to-end run):
  **zero findings**; 68 minor notes, all documented trade-offs / cosmetics / hypothetical shapes.
  **Operator-visible behaviour notes:** a report a human Rebuild moved into the sibling tree is updated in
  place from then on (only a manual Rebuild moves it again); when the -900 tree of an old job is ARCHIVED the
  after-close update silently stops for that job until someone runs 🔁 Rebuild report (chips/picker say
  "archived — Rebuild re-homes it"; the re-homed file keeps its -900 name inside the -970 WO); the
  `Build report` label / question appear only when NO copy exists anywhere; a share blip during the opt-in
  check yields "can't check from this PC" (nothing written); on the desk PC (no MAG-ITAR access) every
  Anduril job shows "can't check" — build/rebuild them on the shipping PC; a -970 board scanned with no job
  picked shows "Which job?" + the coated hint (pick the job, then the stop fires); another job's same-named
  file dropped at a job's report location blocks that job's updates until moved away; un-archiving a -900
  tree that had a Rebuild-moved report brings the frozen archived copy back — the next write updates THAT
  copy and folds the -970 copy in; the manual Rebuild walks History twice (a few extra seconds).
  **Open items / decisions still Dhruvil's:** (1) a supervisor-overridden receive on a coated-only job is
  written to History with the job's canonical -900 key as `pn` (a label PCC never printed) — harmless for
  counting; (2) folder-only Case-B boards (coated paperwork History never saw) can still flip a job to
  "testing" if a supervisor overrides the `Coated shipment on paperwork only` stop — the remedy is the ⏪
  import the stop points at; (3) a Case-B board that changes rev between its 8/27 close and a later receive
  scan would not match (rev-sensitive identity for -900-form serials under the coated key) — live shapes
  ship and receive the same rev; (4) the 4016LF record (Eng `N/A`) — 2 boards untracked until corrected;
  (5) 4065 spelling twins (`134-2201-900` / `1342201900`): the two old files fold into ONE on the first
  update — confirm on the shipping PC (one file, row count = sum).
  **Deploy:** `build_exe.bat` → on the shipping PC: 📦 Receiving → 🔁 Rebuild ALL reports → the picker
  pre-ticks the old-flow jobs that already have a report; tick 4006/4007/4008/4023 if their reports are
  wanted → Rebuild selected; then a manual click-through: Build report on one coated-only job, a -970 scan
  into it (stop), a -900 scan of a never-shipped board into a mixed job (stop), a normal receive on an
  old-flow job (unchanged), Open file on both kinds. Source + tests are in `_handoff_to_shipping_pc\`.

- **2026-09-08 (COATED "COMING TWICE" — import+close dedup; Dhruvil builds the exe)**
  Dhruvil: the Coated & shipped sheet listed every serial twice (134-2201/4065, 136-3312/4092,
  142-2200/4041). Diagnosed from History: every doubled coated serial was in BOTH a
  `backfill_shipped` (his manual import, e.g. ship_date 7/27/2026) AND a real `close` (e.g.
  2026-08-03) — the SAME coated shipment entered twice — and `coated_log` appended one row per
  event. The testing side had the mirror risk but its multi-events were `close`+`close` (real
  re-test re-ships, which MUST still count twice).
  - `build_receiving_state` now COLLECTS ship occurrences per serial during the walk
    (`_test_occ`/`_coat_occ`, event-type + WO) and FINALIZES after: a real **close is the
    record of a shipment** (one per close event → a re-test re-ship still counts twice); a
    **backfill IMPORT counts only when NO close covers that serial** (a genuinely pre-history
    shipment). Coated log/bucket built the same way — one row per actual coated shipment,
    preferring the close record. Verified against live History: coated rows/ distinct now
    60/60 (4065), 54/54 (4092), 50/50 (4041) with zero duplicates; the failed-and-reshipped
    boards (4041×8, 4094×5) still correctly count as 2 shipments and sit at "still at customer".
  - The separate **4065 "coated 61 vs 60"** was the folder-scan phantom-cell issue already
    fixed by `_cell_serial` (History alone gives 60) — it just needs the new exe.
  - **4041 "received-back 81, not 80"**: the shipment history genuinely yields 81 currently-
    received (= 32 in-house + 49 coated); all 32 in-house boards shipped once and were received
    once (pass). The one concrete data gap is seq 11 (`000011-1422200970`) — coated & shipped
    with NO testing-receive record — i.e. History vs the physical/COC truth differs by a board.
    This is a records-reconciliation item (the COC folder-scan heads-up on the shipping PC
    flags exactly this), NOT a counting bug; the tracker faithfully reflects the events.

- **2026-09-08 (RECEIVED-FILE FOLDER FIX for rev-change jobs — the 4092 case; Dhruvil builds the exe)**
  Dhruvil: the tracker "creates a whole new folder for saving the received-after-testing
  file even though we have 4 different folders for 4092 (rev change in between)." Diagnosed
  from History (4092 = MAG-ITAR `377 Anduril\136-3312-…`): the recorded `wo_folder` strings
  drift across a rev change — the SAME rev-B02 testing folder appears as both
  `…Rev B01 Rev B02` (no period) and the coated sibling as `…Rev B01 Rev. B02` (with period),
  and the old rev sits under an `Archive\` folder. `write_received_xlsx` trusted the recorded
  string and did `Path(wo_path)/…; mkdir(parents=True)` — when that exact spelling doesn't
  exist on disk it FABRICATED the whole job/WO tree beside the 4 real folders.
  - New `_resolve_received_wo_dir(state)`: the file goes into `wo_path` only when that path
    actually EXISTS; otherwise it lists the customer directory and finds the newest EXISTING
    testing (…900) WO folder for this Eng#+P/N (archived/old job folders skipped, -970
    siblings excluded, the recorded WO number preferred). `write_received_xlsx` then creates
    only the `Received after testing` subfolder inside that real folder — never a tree — and
    corrects `state['wo_path']`/`pd_path` to the resolved testing folder. If no real folder
    can be found (no share access / genuinely gone) it REFUSES with a clear message instead
    of inventing a folder (the boards are still in history; rerun 🔁 Rebuild on a PC with
    access). Healthy jobs (recorded path exists) are unchanged — fast path returns it as-is.

- **2026-09-03 (COUNT-BASED CYCLES + FOLDER TRUTH — the 4094 case; Dhruvil builds the exe)**
  Dhruvil's report: 4094's tracker missed one coated shipment (5 COC/PL/Shipping-Serials
  file sets in the -970 folders, only 4 counted) and showed 5 FAILED boards as in-house even
  though they were re-shipped for re-test. Diagnosis against the live share (job =
  `MAG - ITAR\377 Anduril\161-9658-…`): (a) only 4 coated CLOSES exist in History and no
  coated import was ever done — the 5th shipment exists only as files on disk; (b) the 5
  boards' re-ship IS in History (Aug-31 close with 5 dup-overrides) but the failed receive
  was scanned late (Sep-03, no "Received on" date), so by timestamp the receive looked newer
  than the re-ship. Fixes, per Dhruvil's answers (COC = shipment record; count-based;
  folder scan on 🔁 Rebuild only):
  - **Count-based cycles**: `ship_counts` (one per close/backfill EVENT per serial; voided
    excluded) + `recv_dates` (distinct receive dates). `_currently_received`: shipped MORE
    times than received → at customer (scan order can't lie anymore); received more → in
    house; EQUAL → the old timestamp tie-break (so received-May-shipped-June still works).
    The 5 failed boards correct themselves from History alone on any save.
  - **Folder truth on 🔁 Rebuild** (`scan_job_folders` + `fold_folder_scan`): reads EVERY
    sibling job folder of the customer dir whose name carries the Eng# and embeds the -900
    or -970 P/N (rev-change twins and renamed folders included — 4094 has TWO -970 dirs),
    and every WO's **COC / Packing List / Shipping Serials** xlsx+csv files. Serials route
    by the P/N embedded in the serial. **COC is the shipment record** (a serial in 2 COC
    files = shipped twice); PL/SS only fill in serials no COC carries; disagreements between
    the three folders come back as a **"Heads-up (not an error)"** popup listing the serials
    (his wording: notify, don't block). Voided shipments are skipped even though their files
    remain. Ship counts become max(history, folders); never-imported old shipments join the
    baseline; every coated shipment found on disk joins the coated bucket + log (this is
    what finds 4094's missing 5th shipment). A scan that can't read the folders (MAG-ITAR
    has no access from the desk PC) falls back to history-only and SAYS so — a partial scan
    is never used (it would undercount cycles).
  - **Anomaly comment** (his ask): a board shipped 2+ times whose receives never recorded a
    FAIL gets "shipped N× but never recorded as FAILED — check records" in the new **Note**
    column (status sheets; Still-at-customer also gained **Times shipped**) and in the
    heads-up popup. A genuinely failed re-ship shows "re-shipped for RE-TEST after failing".
  - Repeated imports / re-scanned receives stay idempotent: ships count per EVENT (an import
    of an already-closed shipment is skipped by the importer as before), receives count per
    DISTINCT DATE, and the file merge keys by (serial, date) — importing the same shipment
    twice or re-scanning a received board cannot double anything.
  - **Round-3 adversarial review (20 agents over the scanner + cycle logic) confirmed 16
    holes — all fixed:** (1) the scanner counted the app's OWN paperwork cells as serials
    (the COC/PL "P/N" field and the COC certification sentence embed the P/N) — phantom
    boards, broken void suppression, coated-seq pollution; now only cells shaped like real
    serials count (`_cell_serial`: ≤40 chars, numeric sequence first, embedded P/N, not the
    bare P/N). (2) Every Save re-walked history into a fresh state and silently REVERTED
    the folder truth until the next Rebuild — the last successful scan is now cached per
    job and re-folded after every fresh walk. (3) `_do_save`'s local fold now updates
    `recv_dates` too (a just-received re-test board no longer flips to "at customer" when
    the post-save re-walk fails). (4) An old paper receive imported AFTER a newer live
    receive no longer shadows it (newest receive date wins, not newest event file).
    (5) COC-wins is now aggregated PER WO (a re-ship documented in another WO without a
    COC still counts; a COC anywhere no longer suppresses it). (6) Coated rows get the same
    COC-primacy per (serial, WO), a dated row beats a blank-dated one, and folder rows are
    deduped against history closes by (serial, WO) — one shipment can never appear as two
    permanent rows on the log sheet. (7) I/O errors (share blip, locked file) now refuse
    the WHOLE scan (history-only fallback with a message); merely corrupt/odd files are
    skipped but counted and listed in the heads-up ("could not read X — its shipment is
    NOT counted"). (8) A folder-found shipment's date now joins the timestamp tie-break
    (`sts`), the heads-up popup shows even when the xlsx write is blocked, the progress
    dialog names the file being read, and workbook reads are capped at 4000 rows. Known
    limits: PDFs/Word files in the shipment folders are not readable (need xlsx/csv or an
    import); a new session shows history-only counts until someone presses 🔁 Rebuild; a
    void done after a Rebuild needs another Rebuild to purge the voided shipment from
    folder counts; a board coated twice in the SAME coated WO with one shipment only on
    paper keeps the history row only.

- **2026-09-02 (SERIAL TRACKING IS NOW REV-AGNOSTIC — v63; Dhruvil builds the exe)**
  - **Board identity = Eng# + P/N + serial number; REVISION IGNORED** (`_serial_identity`).
    Real jobs 4092 (`136-3312-900`/`-970`) and 4041 changed revision MID-JOB (separate job
    folders per rev, e.g. `136-3312-900_4092 Rev A01 Rev A` → `… Rev B01 Rev B02`), so the
    SAME physical board appears as `…-A01` in old shipments and `…-B01` later. Exact matching
    split one board into two records — imports + received boards "didn't match up." Now the
    trailing Rev is stripped whenever the job's P/N is embedded in the serial; `‹WO›-‹seq›`
    serials (no embedded P/N) still match exactly, so unrelated boards can never merge.
  - Applied EVERYWHERE serials are tracked: the shipping duplicate check (index + lookup +
    approved sets), the Receiving tracker (shipped/received/failed keys, cycle logic, xlsx
    row lookups), backfill import idempotency, and the Receiving window's session checks.
    Displays keep the full serial (with rev) — only the MATCHING ignores it.
  - **Supersedes the 2026-08-19 exact-match decision** (that reversal predated the real
    mid-job rev-change data). Suites updated: A01↔B02 same board, dup-check finds either
    form, rev-changed receives clear "Still at customer", 4092-style dashed P/N covered.
  - **"Received after testing" xlsx split into THREE sheets** (Dhruvil: coated-and-shipped
    boards were cluttering the Received sheet as if still in-house): **Received** = in-house
    only (received, not yet coated & shipped — pass awaiting coating + failed awaiting
    re-ship, with the Result column); **Coated & shipped** = lifecycle-complete boards
    (serial · received date · coated WO · coated ship date); **Still at customer**
    unchanged. Summary header on every sheet still shows all five counts.
  - **Backfilled COATED shipments now count** (Dhruvil's 4065 case: `134-2201-900`/`-970`
    fully imported by hand — the 10 coated boards never showed as coated). The tracker's
    coated bucket was fed ONLY by real `-970` close events; `build_receiving_state` now also
    reads `backfill_shipped` events whose P/N is the paired coated (-970) P/N into the coated
    map (by seq). Bonus fix: a `-970` job created by such an import no longer appears as a
    selectable job in the Receiving dropdown (testing …900 P/Ns only). Nothing needs
    re-importing — the events are already in history; 🔁 Rebuild file picks them up.
  - **Report sheets renamed/clarified** (Dhruvil: "Received says 10 in summary but there is
    no serials in first sheet" — the 10 had all moved on to Coated & shipped, so the sheet
    named "Received" was legitimately empty but looked broken). Sheet 1 is now titled
    **In-House**; every sheet's subtitle carries its own row count ("… — N board(s)"); the
    summary line on all sheets and the Receiving-window chips now show the arithmetic
    explicitly: `Received back: N = in-house X (awaiting coating …, failed …) + coated &
    shipped Y`; and an empty In-House sheet prints a note saying where the boards went.
  - **"Received after testing" file rebuilt around LOG + MERGE semantics** (Dhruvil, same
    day — three requests in one): (1) some boards *never go for testing* — they coat and
    ship directly — so the **Coated & shipped sheet now lists EVERY coated shipment on
    record** ("catch everything from shipment … without thinking anything"), with a
    *Received from testing* column showing the receive date where one exists, "no receive
    record" when the board shipped for testing but never came back, and **"sent without
    testing"** when there is no testing record at all; (2) a **new sheet "Received from
    testing"** logs EVERY receive occurrence (re-receives get their own row), no matter
    whether a shipping record exists; (3) **"lets keep one file from now on"** — the file
    is now WRITABLE and MERGED instead of regenerated: on the two log sheets the app never
    modifies or deletes an existing row (manual rows, edits and even user-added sheets are
    kept), only APPENDS boards not in the sheet yet (keyed serial+date, rev-agnostic), and
    a board already present with DIFFERENT data raises an "Overlapping data" alert — the
    file's row always wins ("we will figure it out what is wrong and what is right
    afterwards"). In-House + Still-at-customer became clearly-labeled STATUS VIEWS that are
    still rebuilt from history each save. When the newest testing WO changes, the file
    MIGRATES with all content; a same-named copy that cannot be read is never deleted and
    never blocks a save. App-written placeholder cells ('sent without testing',
    'not in shipping records', …) self-upgrade when history later learns the real value.
    Old read-only-era files upgrade automatically on the next save (they were
    write-protected, so no manual edits could exist in them). New chips/summary figure:
    "Coated boards on record: N" (log-based, independent of the receive linkage).
  - **Merge engine hardened after a 33-agent adversarial review** (same day; the review
    confirmed 11 data-loss/correctness defects in the first merge implementation, all
    fixed): (1) a cross-station **lock file** (`<file>.lock`, stale locks auto-broken
    after 3 min) plus a changed-under-us mtime/size re-check wrap the whole
    load→merge→replace, so two stations saving at once — or an Excel save landing
    mid-merge — can no longer silently clobber each other's rows; (2) a file whose
    "Received from testing" tab was **renamed or deleted is REFUSED with instructions**,
    never rebuilt (only genuinely READ-ONLY old-era files rebuild fresh — the read-only
    attribute is the proof they hold no edits); (3) **drawn objects** (Excel text
    boxes/shapes), which openpyxl silently drops on save, are detected from the raw
    drawing XML — the target refuses with instructions, stray copies are kept;
    (4) stray copies are deleted **only when their ENTIRE content provably merged**
    (rows carried verbatim, zero conflicts, only known sheets/columns, no serial-less
    content rows) — enumeration errors, unreadable files, lost headers, writable
    old-format copies and everything else fail toward KEEPING the file; stray rows merge
    with file-priority (verbatim) so their edits survive; (5) existing rows key/compare
    through Excel's **cached values** (`data_only` view), so a user formula like
    =TODAY() neither duplicates its row nor gets overwritten (uncached formulas match
    by serial and are never touched); (6) blank cells in compared columns and
    blank-date manual rows are **upgraded from history** instead of raising perpetual
    conflicts; `_date_key` accepts 2-digit years; (7) merged cells no longer crash the
    save (guarded writes step past them); appended rows land after the last real DATA
    row, not below styled-empty regions; rich text is preserved (`rich_text=True`);
    mapped-drive vs UNC aliases of the target can't be deleted (`os.path.samefile`
    guard); (8) UI: the Receiving footer **Close button now goes through _on_close**
    (unsaved-scan confirmation), `_adopt_state` rejects superseded history walks (a
    void/backfill rebuild can't be overwritten by an older in-flight walk), `_do_save`
    folds new receives into `receive_log` so a failed post-save re-walk still writes
    them to the log sheet, and supervisor-overridden coated-form receives appear on the
    log sheet (kept out of the counts). Known accepted limits: overlap alerts compare
    Result / Coated WO (the semantically critical columns) — edits to other columns are
    simply kept without comparison; a user who changes the DATE of an app row creates a
    second row for that board on the next save (visible, reconcilable by hand).
  - **Round-2 verification review (20 agents over the hardened code) confirmed 16 more
    holes — all fixed:** the lock now breaks stale locks by **atomic rename** (two
    waiters can never both break-and-proceed), verifies an **owner token** before
    unlinking on release, refreshes its mtime between merge attempts (long merges don't
    look crashed; threshold 5 min), retries transient create errors instead of silently
    dropping to unlocked mode, and never strands a lock when the token write fails.
    Stray copies now travel with an at-harvest **(mtime,size) fingerprint** and are
    deleted only if still unchanged at deletion time; deletion of merged copies is
    skipped entirely unless OUR merge is provably the live file (post-replace size
    check). A stray is "fully merged" only when **every meaningful cell** of its rows —
    including never-compared columns like Received By — is represented in the target;
    cell notes/VML/hyperlinks bar deletion (can't be carried by a value harvest), and a
    note typed above the header row does too. Modern **threaded comments** on the one
    file refuse the save with instructions (openpyxl would silently drop them; legacy
    Notes are fine and preserved). Placeholder fills **write through to the cached-value
    view** so a value typed in a copy beats same-save history; `_fill_cell` writes
    '='-strings as text; blank-dated rows appended mid-save register in the blank-date
    map (no same-save duplicates); a **disputed** record (cmp columns disagree) is never
    adopted into a blank-dated manual row — both rows stand; an incoming blank/
    placeholder value never raises a false conflict. Residual accepted risk: a >3-min
    clock-skewed station can break a live lock — the sig checks, post-replace verify and
    sig-verified deletion then bound the damage to a retried save or an extra popup,
    never silent data loss.
  - **Serials are now routed by the P/N EMBEDDED IN THE SERIAL, not the event label**
    (diagnosed against the LIVE share history: 4065's coated boards were imported with the
    dialog's P/N field on `134-2201-900` while the serials were `…1342201970…`, and the
    2026-08-03 coated close recorded its pn as the WHOLE folder name — pre-v58 parser bug —
    so neither counted as coated). `build_receiving_state` and `build_serial_history` now
    classify every serial by its embedded P/N: a …970-form serial is a coated board no
    matter what the event says; a …900-form serial belongs to the testing job; serials with
    no embedded P/N follow the event label. Partner-P/N serials are excluded from each
    other's dup index (…900↔…970 sibling rule), and coated-form serials can't enter the
    'received' map. Verified against the exact real event shapes pulled from the share.

- **2026-09-02 (VOID WRONG SHIPMENTS + ONE-LOGIN BATCH RE-SHIP — v63; Dhruvil builds the exe)**
  - **🚫 Void record** (supervisor): in 🕘 History, open a Closed/Received/Imported record →
    "Void record…" → admin login + reason → a `void` event is written referencing the target's
    event_id. The record STAYS in history (audit; row grayed "(VOIDED)", detail shows who/when/why)
    but its serials STOP counting: `build_serial_history` (dup check), `build_receiving_state`
    (tracker) and `list_testing_jobs` all skip voided targets. For wrong shipments recorded by
    mistake. Voids are always newer than their target, so the newest-first walk sees them first;
    `write_void_event()` refuses a target without an event_id.
  - **🔓 Supervisor: allow re-ship (batch override):** boards legitimately ship twice sometimes,
    and the per-serial login made big re-ship batches unusable. New button in the serial popup:
    ONE admin login → every duplicate in that session is auto-approved (no dialogs), each still
    logged individually as a `dup_override` with the approver's id. Button turns red
    "RE-SHIP ALLOWED (id) — click to end"; authorization dies with the popup.
  - History viewer: "Deleted" filter checkbox now also shows 🚫 Void markers.
  - **Void now takes effect IMMEDIATELY on the voiding PC** (Dhruvil: voided batch still
    dup-flagged on re-scan). The duplicate check caches its shipped-serials index per
    loaded job; `HistoryDetail._void` now drops that cache (keeping in-progress supervisor
    approvals) and rebuilds it, and refreshes an open Receiving window's state. Other
    stations still pick the void up on their next WO load.
  - **"Received after testing" xlsx moved INSIDE the WO folder** (Dhruvil corrected the
    location): it now lives in the job's **NEWEST testing WO folder**, next to Shipping
    Serials — `…\Production Data\<newest WO>\Received after testing\<PN> - Received after
    testing.xlsx` — and moves there automatically when a newer testing WO ships. Stale
    copies (older WO folders + the old Production-Data-level home) are auto-removed on
    each regenerate (`_cleanup_stale_received_files`; state gains `wo_path`, with a
    backfilled-WO fallback for jobs that have no recorded closes).
  - 15-check void/batch suite + full regression (all six suites) green; `py_compile` clean.

- **2026-09-01 (4×2 BOARD LABELS FOR THE HPKE200 — ported from Apex; v63; Dhruvil builds the exe)**
  - **🖨 Print 4×2 labels** button in the serial popup: generates the Apex "Serialized Pack
    4×2" label for every scanned serial — LEFT column P/N (12pt title, 10pt value, QR,
    6pt mirror text), RIGHT column Serial (same structure) — as ONE self-contained HTML
    file (one 4in×2in page per serial, `@page 4in 2in margin 0`) and opens it in the
    browser; the print dialog pops automatically, operator picks the **HPKE200** (same
    flow they already use in Apex) and all labels print as a single job. Rows SELECTED in
    the list → only those print (easy single-label reprint). P/N comes from the loaded
    job; if blank/N-A it asks once.
  - **qrcode.js embedded** (`_QRCODEJS_B64`, MIT, the same library Apex loads from its
    CDN) so the generated file needs NO internet — QR codes render offline. Pure stdlib
    otherwise: no new Python packages, no spec change; normal rebuild picks it up.
  - `write_label_html_4x2()` covered by a 13-check suite (structure, per-serial pages,
    QR wiring, HTML/JS escaping incl. hostile serials, empty-list error); all other
    suites re-run green.
  - **Floor-test polish (2026-09-01/02, per Dhruvil):** titles "P/N:"/"Serial:" 12pt→**14pt**;
    values settled at **10pt semi-bold BLACK** (11pt/900 tried 09-01, reverted after print —
    long serials like `000633-1629645-900-A01` wrapped, pushing the rev to a 2nd line);
    values now **never wrap** — `white-space:nowrap` + a JS fitter that shrinks an
    overflowing value (min 6pt) so the whole serial always prints on ONE line. QR
    **rendered at 300px, displayed at 60px** (5× oversampling) so the ~203-dpi thermal
    print gets crisp modules — same physical QR size, easy scanning (floor-verified).

- **2026-08-21 (RECEIVING/BOARD TRACKER + MULTI-DRAFT — v63; Dhruvil builds the exe)**
  - **📦 Receiving window** (new header chip): for the test-and-return customer. Operator scans
    returned boards' TESTING (-900) labels; the app auto-detects the job from the embedded P/N
    (mixed boxes across WOs sort themselves — each serial is attributed to its outbound WO from
    history). Live tracker chips: Shipped for testing · Received · Still at customer · Coated &
    shipped (matched by seq against the paired -970 closes, rev-change tolerant) · In-house
    awaiting coating. Session notes. Save writes a `receive` History event (SOURCE OF TRUTH) and
    regenerates the report file. Online-only. Per Dhruvil: NO hard blocks anywhere — every stop
    (already received / doesn't match job / scanned twice) offers Skip or **Supervisor override…**
    (admin login, audited into the event + History detail).
  - **Per-P/N report file** `…\Production Data\Received after testing\<PN> - Received after
    testing.xlsx` — two sheets (**Received**, **Still at customer**) + summary header. Fully
    REGENERATED from history each save (self-healing; manual edits don't survive); kept
    **read-only on disk** so Excel viewers on other PCs can look but never edit/lock; atomic
    station+pid temp swap; FileInUseDialog retry when a viewer physically blocks it (history is
    already saved — file catches up on next save or 🔁 Rebuild). NOTE: we can't force-close Excel
    on other PCs (accepted by Dhruvil 2026-08-21).
  - **⏪ Import older shipments (backfill):** reads every WO folder's `Shipping Serials` files
    (app layout, odd layouts via P/N-containment, CSV) to seed the shipped-for-testing baseline
    for boards shipped before History existed. Idempotent — already-recorded serials are skipped
    (verified against a FRESH history walk). One `backfill_shipped` event per WO. Backfilled
    serials also now count in the SHIPPING duplicate check (`build_serial_history` reads them).
  - **Multi-draft (same day):** a WO can now hold MANY drafts — `.shipping_draft_<id>.json` per
    shipment session (legacy `.shipping_draft.json` still supported). Loading a WO with drafts
    shows a **picker** (date · saved-by · counts): Resume selected / **Start NEW shipment** /
    Cancel — a second batch no longer waits for the supervisor to close the first, and the
    supervisor chooses exactly which draft to close. Save/Close/delete all target only the
    session's own draft file. `list_shipment_drafts()`, `_draft_path()`, draft_id in records.
  - **Adversarial review (24-agent, 4 dimensions) — 19 confirmed findings, ALL FIXED**, notably:
    a race that made Save falsely report "history unavailable"; xlsx regenerated from a stale
    cache could drop OTHER stations' received rows (now re-walks history before every regenerate);
    a truncated history walk could overwrite the report with near-empty data (walk now flags
    errors and the writer refuses); the read-only attribute stayed lifted after a failed swap;
    modal-grab loss in the backfill dialog; job-switch mid-session re-attributing scans (now
    confirmed + cleared); scans bounced while the job list loads (now waited, scan kept);
    empty-admin-roster made overrides an un-overridable dead end (now reloads + honest message);
    temp-file PID collisions between stations; HistoryViewer shows the new event types
    (📦 Received / ⏪ Imported, new "Received" filter checkbox).
  - **Progress feedback on Save Draft + Save and Close** (Dhruvil: 30s frozen screen with
    "nothing in between"). New `ProgressDialog` — a modal ⏳ box that live-updates per step
    ("Moving photo 3 of 12…", "Writing the serials Excel…", "Writing the Certificate of
    Compliance…", "Recording in shipment history…") and pumps the UI so it stays painted; its
    grab swallows stray clicks/scanner bursts during the work. Both flows were split into
    wrapper + `_save_draft_work`/`_close_shipment_work` run under try/finally, so the box can
    NEVER be left stuck; it re-asserts its grab after FileInUseDialog retries and closes before
    every result/cancel popup. Also hoisted a per-photo UNC `resolve()` out of two loops (each
    was a network round-trip per photo — part of the slowness).
  - **PASS/FAIL receiving + full board CYCLES** (boards fail testing → ship back for re-test →
    return again; some failed boards coat anyway per special instruction):
    - Receiving gets a **✔ PASSED / ✘ FAILED toggle** (default PASSED — scan the passed pile,
      flip for the failed pile; fails show red in the list and as a FAIL column in the xlsx).
      The receive event records the failed subset (`"failed"`), older events read as all-pass.
    - The tracker is **cycle-aware** (`_currently_received()` compares the latest ship vs the
      latest receive per serial, via sortable `sts` timestamps; backfilled ships sort oldest):
      a failed board re-shipped for re-test moves back to "Still at customer" and can be
      received AGAIN without a false "already received" alert. Chips + xlsx add **Failed
      awaiting re-ship**; "In-house awaiting coating" = passed only.
    - **The shipping dup-check reads receive events now:** a serial whose NEWEST event is a
      FAILED receive is EXPECTED to ship again → **no "already shipped" popup** when re-shipping
      the failed pile; a PASSED board scanned into a new testing shipment still alerts (that's
      a real mistake). "Coat after small fix" needs no special step — the coated (-970) close
      simply counts it as coated.
    - **Paste List** in Receiving (button + multi-line paste into the scan box) for importing the
      old MANUAL tracking Excels: batch-processed with ONE summary (added / skipped-already-
      received / skipped-wrong-P/N / kept-with-⚠-not-in-shipping-records) — no per-serial popups,
      re-pasting overlapping lists is repeat-safe. Uses the current PASS/FAIL toggle.
    - **"Received on" date field** (optional; e.g. `2026-05-14` or `5/14/26`) so imported old
      records carry their REAL receive date — it drives both display and cycle ordering
      (`received_date` on the event, `_parse_user_date()`); blank = today.
  - **Tests:** 130+ checks across four suites (receiving state/summary/xlsx incl. lock + stress
    1500 boards, pass/fail + cycle + received_date cases, real-writer roundtrip, backfill
    parsing, multi-draft, plus full regression of all earlier features) — ALL GREEN;
    `py_compile` clean. **Exe NOT rebuilt** — Dhruvil builds via `build_exe.bat` when ready.

- **2026-08-19 (ODD-SERIAL CHECK P/N-ANCHORED + EMBEDDED LOGO FIXED + WO SEARCH SKIPS ARCHIVES — exe rebuilt)**
  - **Odd-serial (outlier) check reworked** so it actually fires on small/real batches.
    When the job P/N is known and embedded in the serials (`<seq>-<P/N>-<Rev>`), it now
    ANCHORS on that P/N instead of needing a statistical majority: a serial NOT carrying
    the job P/N is flagged (wrong board / bad scan), a serial whose trailing Rev differs
    from the rest is flagged, and a wildly out-of-range sequence # is flagged — all working
    even for a **2-scan batch**. The batch-majority check remains for the `<WO>-<seq>`
    format. (Before: `min_batch=5` + a 70%-majority rule meant a 2-serial test never
    flagged — Dhruvil's `A00` vs `B00`, `A00` vs `000`, and `A00` vs `000000` now all
    flag.) `detect_serial_outliers(serials, pn=...)`; the popup gets the job P/N via
    SerialPopup's new `job_pn`.
  - **Odd-serial FALSE-POSITIVE #1 fixed** (WO-seq format, e.g. `16452.1-1` … `16452.1-42`):
    `_serial_signature()` no longer includes each segment's LENGTH — only the character
    class per segment — so a 1-digit vs 2-digit vs 3-digit sequence number in one WO is
    the SAME shape and is no longer flagged as odd.
  - **Odd-serial FALSE-POSITIVE #2 fixed — removed the numeric range/gap check entirely.**
    It flagged a serial whose sequence # was numerically far from the rest, but a gap just
    means non-consecutive boards (you might ship #1-5 and #45 together), and number size
    can't distinguish a legit high serial from garbage. So a high or gapped sequence number
    no longer flags (Dhruvil: `16452.1-45` among `16452.1-1..5` was false-flagging; also
    `999999`/`000999` no longer flag). A bad scan is still caught by the shape / prefix /
    P/N / Rev checks — never by "this number is big." (`_lone_number_flags` removed.)
  - **Embedded logo fixed.** `_EMBEDDED_LOGO_B64` was a black-background image; re-encoded
    it from the real `Scripts\logo.png`, so the embedded copy is now the correct logo. The
    exe no longer needs an external `logo.png` to show the right logo (it is still bundled
    too when present). Verified the round-trip decode equals `logo.png`.
  - **WO search now ignores archive folders.** Many customer folders keep old versions
    under an **`Archives` / `Archive` / `Archived` / `Old`** folder; `search_wo_folder()`
    now skips any `Production Data` whose path (below the division root) contains such a
    folder name (whole-segment, case-insensitive — so "Boldera" is NOT excluded), so an
    archived WO is neither loaded nor offered as a match — only the live copy counts.
    e.g. `…\379 Liber8Now\Archives\GSR10 […]\Production Data\…` is ignored;
    `…\379 Liber8Now\GSR10 […]\Production Data\…` counts.
  - Test suite extended (P/N-anchored cases incl. Dhruvil's three; archived-vs-live WO
    search) — all green; `py_compile` clean (no warnings); single-file exe rebuilt +
    launch-verified.

- **2026-08-19 (SIMPLIFIED TO A SINGLE-FILE BUILD — thin-launcher/onedir retired; logo.png removed; exe rebuilt + verified)**
  - At Dhruvil's request, collapsed everything to **one source file + one exe**:
    - `app_logic.py` is now built **directly** into a single `dist\PCC Shipping App.exe`
      (~36 MB); it already ends with `if __name__ == "__main__": ShippingApp().mainloop()`.
    - **Deleted** the thin launcher (`PCC Shipping App.py`), the onedir spec, the whole
      `dist\PCC Shipping App\` folder (exe + app_logic.py + logo.png + `_internal\`), and
      `dist_onefile\`. The onefile spec is now the only spec, renamed `PCC Shipping App.spec`;
      `build_exe.bat` rewritten for the one-file build.
    - **Logo:** `logo.png` is kept in `Scripts\` and **bundled into the exe** at build
      time (the exe stays self-contained). The base64 copy embedded in the code
      (`_EMBEDDED_LOGO_B64`) renders with a **black background**, so the app uses the real
      `logo.png` instead (`_find_logo()` prefers the bundled file). Keep `logo.png` present
      when building. *(First cut this day removed logo.png and relied on the embedded copy;
      reverted after the black background showed up.)*
  - **Trade-off (accepted):** every change is now a rebuild — the "swap `app_logic.py` without
    rebuilding" path is gone. Dhruvil rebuilds via `build_exe.bat` when we change something.
  - Built here with PyInstaller 6.22.2 and **verified** the exe launches (GUI came up, no
    missing-module crash). Sections 1–6 of this doc were rewritten for the single-file model.

- **2026-08-19 (SERIAL DUPLICATE HISTORY + ODD-SCAN CHECK + SHIPMENT NOTES + COC ASSEMBLY OPTIONS — onedir synced, v62; SINGLE EXE NOT yet rebuilt)**
  - **Cross-WO duplicate detection (scoped to Eng# + P/N).** As each serial is scanned it is
    checked against every serial from **closed** shipments of the same **Eng# + P/N** (a "job"),
    across all past work orders. A hit shows a scanner-proof blocking popup naming where it shipped
    before (WO + date + operator). Scope is Eng#+P/N because one Eng# can carry two P/Ns whose serials
    each run 1..N — e.g. bare `192-146-900` (testing) vs. coated `192-146-970` — so the same number
    under a different P/N is NOT a duplicate. For the customer whose serial is `‹seq›-‹P/N›-‹Rev›`
    (e.g. `000001-1921682900-AOO`) the P/N is baked into the string, so exact-string matching keeps the
    two P/Ns apart on its own. Retroactive to when History started (~2026-07-22); older shipments aren't
    indexed. Match = exact string, trimmed + case-insensitive. Runs for **all customers** (safe — exact
    matching rarely false-alarms).
  - **Matching is EXACT (Rev included).** A different Rev is a **different serial** and is NOT a
    duplicate — e.g. `000001-‹P/N›-AOO` and `000001-‹P/N›-B00` are two different boards. Match key is
    `_norm_serial` (trim + casefold), the whole string. (A rev-agnostic "strip the trailing Rev"
    variant was built and then **reverted at Dhruvil's request, 2026-08-19** — the concern was that a
    rev change must not cause a *false* duplicate, which exact matching already guarantees since the
    Rev is part of the serial. Do not re-add rev-stripping.)
  - **Supervisor-only override.** The operator can't wave a duplicate through: the dialog offers
    **Skip** (default, safe) or **Supervisor override…**, which opens the existing admin login. A valid
    login adds the serial and writes a `dup_override` history event (serial, approving admin id, the
    prior shipment it collided with); the override is also folded into the shipment's close/draft record
    (visible in History detail + CSV). Override is click-only, so a scanner burst can never trigger it.
  - **Odd-serial ("outlier") check at submit — advisory.** On **Done**, the batch is analysed by
    segment (split on `-`): it learns the normal shape and the constant parts (the P/N segment, the WO
    prefix) and flags anything that deviates — a garbage/short/long scan, a **wrong P/N baked into the
    serial** (a `-970` board in a `-900` batch), a wrong WO prefix, or a sequence number far outside the
    batch. A review dialog lists only the flagged ones with **Remove/Keep** each. Conservative: needs a
    clear majority to set the norm, so a uniform batch (or a legit internal gap) produces no dialog.
  - **Shipment notes.** A free-text box in the serial popup ("where it's going / special instructions").
    Saved into the draft (survives resume) and the History record; shown in History detail + CSV.
    **Record-only — not printed** on any document.
  - **P/N confirm.** Since the dup check is P/N-scoped, if a job loads with a blank/N-A P/N the app
    prompts to confirm it before scanning (operator can skip → check simply off for that shipment).
    Offline shipments skip the whole check (share History is unreachable).
  - **COC Assembly Process options changed** (`COCPopup.AO`) from the old "Lead Free / Leaded /
    Special PCB Assembly" to exactly: **Lead free · Leaded · Special · N/A**. (The pre-existing
    "No PCB Assembly" option was **removed** at Dhruvil's request, 2026-08-19.) Combobox, defaults to
    "Lead free"; an old draft/COC saved as "No PCB Assembly" still displays but is no longer selectable.
  - **Implementation:** pure stdlib. New module fns `build_serial_history()`, `detect_serial_outliers()`,
    `_norm_serial/_norm_key`; new dialogs `HistoricalDuplicateDialog`, `OutlierReviewDialog`,
    `PNPromptDialog`; the index builds on a background thread at WO load (no UI freeze) with a full
    re-check at submit as the backstop. 35-check test suite (P/N scoping, wrong-P/N catch, the
    `16152.111` example, no-false-positives on clean/gapped batches, notes round-trip,
    rev-agnostic matching, new COC options) all green; `py_compile` clean; parser/sort
    regressions green; every dialog also rendered + screenshot-verified in the sandbox (which
    caught a `pady`-tuple crash bug in the serial popup, since fixed). Both onedir copies synced.
  - **DEPLOY:** onedir is done (swap `app_logic.py`). The single-file station exe (`dist_onefile\`) is
    **stale** — rebuild it (`PCC Shipping App-onefile.spec`) and re-copy to each station to put this on
    the floor. Still-pending read-only Monday token swap will fold into that same rebuild.

- **2026-08-05 (COC TABLE + PRINT SCALING + DUPLEX ON ALL PRINTERS — DEPLOYED, single exe rebuilt, v61)**
  - **COC fields = full-width bordered table, centered on the page.** The 8-field block is
    now a bordered label|value table; the value cell is merged B:C so the table spans the
    full width (lines up with the serial table), and `ws.print_options.horizontalCentered`
    centers the whole sheet on the paper (the content is narrower than the page, so without
    this it hugged the left). Applied horizontalCentered to the Packing List too.
  - **Scaling corrected** — reverted v60's force-everything-onto-one-page. Now COC & Packing
    just **fit all columns onto one page WIDE** (fitToWidth=1, fitToHeight=0, fit-to-page
    enabled) and flow down naturally. (v60 had over-scaled per an over-reading of the
    "orphan lines" request; the operator only wanted fit-to-width.)
  - **Two-sided now set on ALL printers**, not just the default — `_ensure_duplex_default()`
    changed from the Windows default printer to `Get-Printer | ForEach-Object { … Set-
    PrintConfiguration TwoSidedLongEdge }` (per-printer try/catch skips PDF/Fax/OneNote), so
    whichever printer the operator picks defaults to two-sided. Still once-per-session,
    hidden window, best-effort; manual printer setting remains the fallback. ⚠ still verify
    on a REAL station (PowerShell + duplex-capable printer + AV not blocking).
  - Sample COC sent to Dhruvil for eyeball → approved. test_coc_format (centered + bordered +
    full-width merge + horizontalCentered + fit flags) + additions + parser + Monday +
    online-close all green. onedir both copies + single exe rebuilt/verified/smoke-tested.
    Read-only Monday token swap still pending.

- **2026-08-04 (COC FORMAT: center-align fields + one-page scaling — DEPLOYED, single exe rebuilt, v60)**
  - **COC fields centered** — the 8-field block (Ship Date, PO#, P/N, QTY, DC, PCC Job#,
    Type, Assembly process) was label-right + value-left; both are now center-aligned.
  - **Fixed the "2nd page with 1-2 orphan lines" spill.** Root cause: the COC writer set
    fitToWidth/Height but never enabled `pageSetUpPr.fitToPage` (the packing writer did),
    so the COC printed at 100% and spilled. Now the COC enables fit-to-page AND scales the
    whole sheet onto ONE page when a shipment has ≤ `_MAX_ONE_PAGE_SERIALS` (40) serials;
    longer lists still paginate instead of shrinking to unreadable. Same rule applied to
    the Packing List. (Scaling only shows on print/preview — verify on a real print.)
  - test_coc_format (alignment + fit flags) + additions + parser + Monday + online-close
    all green. onedir both copies + single exe rebuilt & verified (bundled v60 + markers +
    ssl/urllib/subprocess in PYZ) + smoke-tested. Read-only Monday token swap still pending.

- **2026-08-03 (SHIPPING APP: SAVE-AND-PRINT + SERIAL SORT + 2-SIDED AUTO-SET — DEPLOYED, single exe rebuilt, v59)**
  - **Save and Print** on the COC and Packing List popups — the buttons are now
    Save · 🖨 Save and Print · Cancel. "Save and Print" renders the document to a temp
    folder and opens it in Excel to review + print (Ctrl+P). The filed copy is still
    created at Save and Close. The COC/packing writers already set the print area to the
    written content only, so no blank pages.
  - **Serials sorted on save** — at Save Draft and Save and Close the whole shipment's
    serials are sorted in NATURAL increasing order ("2" before "10", "SN2" before "SN10")
    via `sort_serials()`, applied to the draft, serials xlsx, COC, packing, and history.
    LIVE SCANNING is left in scan order (so the operator can see what they scanned last).
  - **Two-sided auto-set** — on the first Save-and-Print of a session the app sets THIS
    PC's default printer to two-sided (`_ensure_duplex_default()` → PowerShell
    `Set-PrintConfiguration -DuplexingMode TwoSidedLongEdge`, once per session, hidden
    window, best-effort). Persistent + affects ALL printing on that PC (intended for a
    dedicated shipping station). ⚠ MUST be verified on a REAL station: depends on the
    station's PowerShell + a duplex-capable printer + endpoint security not blocking an
    unsigned exe from spawning PowerShell. Fails silently → the manual printer setting
    (§6a) is the guaranteed fallback. Chose this over per-job pywin32 forcing (which would
    have dropped the Excel preview + bloated the exe).
  - `subprocess` added to both specs' hiddenimports. Adversarial review of Save-and-Print
    + serial-sort came back clean; 17 additions tests + parser + Monday + online-close
    regressions green; rebuilt single exe verified (bundled v59 + ssl/urllib/subprocess in
    PYZ) and smoke-tested. Read-only Monday token swap still pending (separate rebuild).

- **2026-08-03 (NEW TOOL: BOM PART VERIFIER — standalone HTML, project root)** — added
  `BOM Part Verifier.html`, a single-file browser tool for stockroom operators: upload a
  customer BOM (CSV/Excel), scan each part in the kit, and see Verified / Missing /
  Not-in-BOM live. Separate from the shipping app — no build, no deploy; just open the
  file in a browser (Excel parsing uses a CDN script, so .xlsx needs internet; CSV works
  offline).
  - **Multi P/N-type matching:** the column mapper now accepts MULTIPLE part-number
    columns (Mfr P/N, Vendor P/N, Customer P/N, Supplier P/N…) — at least one required,
    add more with "＋ Add another P/N column". Scanning ANY mapped P/N verifies the same
    BOM line (alias map), so operators can scan whichever number is on the reel/bag.
    Toast + scan log show which P/N type matched; result tables list all known P/Ns per
    line; CSV export has one column per mapped P/N type. BOM rows sharing any P/N merge
    (qtys added). Matching is case-insensitive, exact after trim.
  `split_job_name()` for the delimiter-less folder format `<P/N> <Eng> Rev. <Rev>`
  (e.g. `161-1061-970 4081 Rev. B00`), which previously returned the WHOLE name as the
  P/N. Now → pn=`161-1061-970`, eng=`4081`, rev=`B00`. Bracket `[...]` and underscore
  formats unchanged.
  - Anchored on the `Rev` keyword, but only when it is a standalone word FOLLOWED BY a
    revision value (so a bare trailing `REV`, or `REV` inside a part number like
    `REV-BOARD`, is never mistaken for it). The token before `Rev` is taken as the Eng #
    ONLY when it contains a digit — so a multi-word descriptive P/N with no Eng #
    (`15102203 Water Draw Box Rev. A`) keeps its full P/N (eng N/A) instead of losing its
    last word. `rev_from_inside()` is now case-insensitive to match.
  - Bonus: these folders are now also findable by typing the Eng # (the folder search
    used to skip them).
  - Adversarially reviewed — the review caught a multi-word-P/N regression in the first
    cut, which the digit-guard + value-required-after-Rev rule fixes. 20-case parser
    suite + Monday + online-close suites all green.
  - DEPLOYED: onedir both copies updated (edit-only) AND the **single-file exe rebuilt**
    → `dist_onefile\PCC Shipping App.exe` (v58; verified bundled code + `ssl`/`urllib` in
    the PYZ + smoke-tested to launch). Copy that one file to each station (§6a). The
    read-only Monday token swap is still pending and will require another exe rebuild.

- **2026-07-28 (MONDAY.COM WO AUTO-LOOKUP — DEPLOYED to onedir, v57)** — typing a
  Work Order # or Engineering # now looks the job up on Monday.com and pre-fills
  Customer / P/N / PO / Rev, pulling from the same board the "PCC Apex" labeling app
  uses. The standout is **PO#** — the WO folder name never contains it, so operators
  used to type it into every COC by hand. Pre-Monday build snapshot kept in the
  session scratchpad (`app_logic.offline.py`) for rollback.
  - **Augment only** — the `\\PCC2` WO folder is still found and written by the normal
    folder search; Monday only fills fields. The reliable match key is the **Work
    Order folder name** inside Production Data (always equals Monday's WO#); the fuzzy
    `Customer PN [Eng - BOM Rev - Assy Rev]` job-folder name is only a fallback.
  - **Safe by construction:** the lookup runs off the UI thread behind a "Searching
    Monday.com…" popup with a timeout (window never freezes); ANY failure — no network,
    timeout, HTTP/401, rate limit, bad data — degrades silently to the normal
    folder-only search; and it is **skipped entirely while offline**.
  - **Won't mislabel:** before the confirm step the Monday hit is verified against the
    loaded folder by WO# (or Eng#); if it doesn't match, the hit is dropped (no PO
    pre-fill). Every load still goes through the "Job Found — Please Confirm" dialog,
    which now shows the matched **Customer PO** so the operator verifies it first.
  - **Divisions:** the 3 shipping operators ship Industrial only; MAG-ITAR is handled
    by another team and they have no MAG-ITAR share access. The MAG-ITAR button stays,
    but a failed MAG-ITAR search now says "No access to MAG - ITAR — ask IT or use
    Industrial" instead of the confusing "server not reachable".
  - **Deploy:** stdlib only (`urllib`/`ssl`/`json`/`threading`), so the onedir path is
    an edit-only swap of `app_logic.py` in both onedir copies (done). Then, at Dhruvil's
    request, the **single-file exe was rebuilt** (`dist_onefile\PCC Shipping App.exe`,
    v57) so each station gets ONE file copied — networking modules were added to BOTH
    specs' `hiddenimports` and verified present in the onefile PYZ; smoke-tested to
    launch. See §6a for the single-exe deploy + its rebuild-per-change trade-off.
  - **SECURITY TODO (Dhruvil):** the Monday API token is baked into `app_logic.py`
    (`MONDAY_API_KEY`, near line 142) — it is the Apex app's WRITE-scoped token reused
    for a read-only feature. Generate a dedicated **read-only** Monday token and swap
    that one line (no rebuild needed), and rotate the exposed write token.
  - Built with a 4-agent design pass + a 4-dimension adversarial review (4 findings,
    all fixed); 30-check test suite + online-close regression green.

- **2026-07-23 (OFFLINE MODE — DEPLOYED to onedir; single-file exe NOT yet rebuilt)** —
  survive-a-PCC2-outage + ship-urgent-offline (Tiers 1+2, both divisions). Prior online
  build saved as `app_logic.py.bak-online` in both onedir copies.
  - Fast 2s reachability probe at startup (replaces the 20–90s SMB freeze); orange OFFLINE
    banner with Re-check / Open Outbox; operator & admin rosters cached locally under
    `%LOCALAPPDATA%\PCC Shipping App\cache` so the app stays usable offline.
  - Offline shipping: type the WO (can't be verified offline), photos/serials/COC/Packing as
    normal, **Save and Close generates the documents into `%LOCALAPPDATA%\PCC Shipping App\
    OfflineOutbox\<ts>_<div>_<wo>\`** (mirrors the WO folder layout) + `history_close.json`
    (source='offline') + `_OFFLINE_SHIPMENT_README.txt` with sync steps. Print → ship → sync
    the folder into the real WO folder by hand once PCC2 is back.
  - **PRECONDITION:** the app must be installed on each station's LOCAL disk — if it runs from
    `\\PCC2`, an outage means it can't even launch. (Local install = copy the app to C:\; no
    installer/build needed on the station.)
  - The single-file `dist_onefile\PCC Shipping App.exe` is STALE (predates offline) — rebuild
    it (`PCC Shipping App-onefile.spec`) if the single-file form is used for stations.
  - Adversarial review caught + fixed 2 freeze bugs (offline draft-save + delete still called
    the share history writer). Online path verified byte-equivalent against live PCC2.

- **2026-07-23 (UX packages 5–10 — DEPLOYED; single-file exe rebuilt)** — code-only for
  the onedir build; the onefile exe was rebuilt. Prior pkg4 saved as
  `app_logic.py.bak-pkg4` in both onedir locations. Production now carries the full
  UX pass. Highlights:
  - **#5 button hierarchy/flow**: sections 3–6 dim until a WO loads; division buttons show
    a ✓ on the selected one (unselected looks enabled, not disabled); status chips are
    clickable; **Save Draft** is a filled teal button with a border, distinct from the
    larger filled-green **Save and Close**; the close-confirm header is amber; History is a
    ghost chip in the header.
  - **#6 form ergonomics**: COC marks PO#/P/N required (one combined validation message),
    cross-checks QTY vs scanned serials, **auto-fills the Rev letter from the folder name**
    (e.g. `[3156LF Rev. - A]` → Rev "A", via `rev_from_inside`), Esc-cancels; the COC
    button bar is reserved at the bottom so it can't be clipped. The serial-list
    **Excel/Word choice is a separate popup with a real Cancel** (Cancel / ✕ / Esc abort
    the close with nothing written — the format is asked BEFORE any file is written/moved).
  - **#7 History**: ✓/✎/✕ glyphs + light row tints (not all-red), sort arrows, instant
    operator/division filtering, em-dash for deleted-row counts; CSV keeps plain labels.
  - **#8 Admin**: "Set password…" for an existing admin (a sole admin can finally rotate
    their own password), show-password toggles, honest "Close"/"changes saved immediately".
  - **#9 Photos**: click-to-preview pane, amber "photos will be MOVED" notice bar, orange
    header, delete disabled until a selection.
  - **#10 cleanup**: readable ShippingApp method names (old short names kept as aliases),
    hover-color fix, delete-buttons disabled until selection, dead code removed.
  - Deferred (not done): consolidating the 12 dialog classes into one BaseDialog and the
    bulk font-tuple migration — safer as their own future pass.
  - Verified: both regression suites pass on the deployed file; the rebuilt single-file
    `dist_onefile\PCC Shipping App.exe` (~35 MB) was run alone in an empty folder and
    launched fine.

- **2026-07-23 (UX package 4 + default operator — DEPLOYED)** — code-only, both runnable
  copies updated; prior pkg3 saved as `app_logic.py.bak-pkg3`.
  - **One-screen layout**: the app now opens **maximized** and section spacing was tightened
    so all six numbered steps + the status chips fit without scrolling on a 1080p station
    (content needs ~857px; a maximized 1080p window shows ~896px). Smaller/restore-down
    windows still scroll, with the Save buttons always pinned. (The footer completeness
    readout that was prototyped for this package was dropped at the owner's request.)
  - **Default operator**: `DEFAULT_OPERATOR = "Anita Patel"` (near the path constants) is
    auto-selected on a fresh start (she does ~90% of shipping). A deliberately-chosen
    operator is preserved across roster refreshes; if she isn't in the roster it falls back
    to the "select your name" placeholder. Change that one line if the primary operator changes.
  - Packages #5–#10 remain built but UNDEPLOYED in the dev working copy, pending one-by-one
    approval; production has only #1–#4 + the default-operator change.

- **2026-07-22 (UX pass, package 3: readability)** — code-only, no rebuild; both runnable
  copies synced. All changes verified to clear WCAG 4.5:1 (contrast recomputed from the
  live constants).
  - **Bigger, verifiable info card**: `_info_row` values went 8pt → 10pt (and 11pt **bold**
    for Workorder # / Customer / P/N — the fields an operator must check before shipping);
    labels 8pt → 9pt. Section headers (`_section_header`) 9pt → 11pt with a larger 26px badge.
    Dropped the stray colon in "Work order number".
  - **Contrast fixes (same hues, just darker where needed)**: `C_SUBTLE` #7A7A7A→#5F5F5F,
    `C_ORANGE` #D35400→#B34700, `C_AMBER` #B5700A→#8F5807 (also lifts white text on the
    MAG-ITAR button and dialog headers). Popup counters: serial #B5D4F4→#D6EFFF (and 13→16pt
    — the number operators watch while scanning), photo #9FE1CB→#CFF1E2, packing WO → white.
    Header subtitle 8pt → 9pt.
  - Added a small font-scale constant block (F_SECTION/F_CAPTION/F_VALUE/F_VALUE_BOLD) next
    to the colors; broader font-tuple migration deferred to the code-cleanup pass.
  - Origin: UX review package 3 of 10.

- **2026-07-22 (UX pass, package 2: crisp rendering)** — code-only, no rebuild; both
  runnable copies synced.
  - **DPI awareness**: `_enable_dpi_awareness()` (ctypes, stdlib) sets System-DPI-aware
    before the Tk window is created, so on 125%/150%-scaled monitors the app renders
    sharp instead of blurry bitmap-stretch. `_apply_dpi_scaling()` then sets `tk scaling`
    to the real DPI and grows the initial window proportionally — a no-op at 100%,
    correct physical font size on hi-DPI. (DPI lives in app_logic, not the launcher, so
    no rebuild is needed.)
  - **Widget theme**: `_init_ttk_style()` switches ttk to the `clam` base and styles
    entries, comboboxes, the History Treeview (26px rows, flat on-brand headings, green
    selection) and scrollbars to match the brand instead of default Windows gray. WO
    entry now shows a green focus border. tk.Checkbutton (history filters) is unaffected.
  - Origin: UX review package 2 of 10.

- **2026-07-22 (UX pass, package 1: scanner safety)** — hardened the UI against
  keyboard-wedge barcode scanners (which terminate every scan with Enter). Code-only,
  no rebuild; both runnable copies synced.
  - **ConfirmActionDialog** (used for Save & Close, delete shipment, discard-work
    warnings): Enter now only confirms via the confirm button and only after a 400ms
    arming delay — a scan burst or double-tapped Enter can no longer trigger an
    irreversible action the instant the dialog opens. Cancel/Escape stay instant.
  - **New DuplicateSerialDialog** replaces the native duplicate-serial messagebox:
    same blocking behavior, but the OK button stays disabled for 750ms and the
    dialog holds keyboard focus, so the trailing Enter of the *next* scan can't
    dismiss the alert unread (previously that lost the new serial silently).
    SerialPopup re-asserts its modal grab after the alert (Tk grabs don't stack).
  - **Startup focus**: the WO entry is focused as soon as the app opens, so the
    first scan of a session lands in the field instead of being dropped.
  - **Auto-open Serials on fresh load** (owner-approved flow change): after a NEW
    work order loads, the Serial popup opens automatically with the scan box
    focused — the operator can scan immediately, and a serial can no longer be
    scanned into the WO box by mistake (which used to trigger a server search and
    a "Not found" error). Draft resumes do NOT auto-open anything, so supervisors
    reviewing a draft keep full control.
  - Origin: 5-lens UX review of all 11 screens (2026-07-22); this is package 1 of 10.

- **2026-07-22** — Added a **standalone single-file (onefile) build** for network-share
  distribution. New spec `Scripts\PCC Shipping App-onefile.spec`; build output is a
  single `Scripts\dist_onefile\PCC Shipping App.exe` (~35 MB) with **no `_internal`
  folder and no external `app_logic.py`** — runs with zero dependencies beside it.
  Verified by running it alone in an empty folder. See section 5b. The onedir
  thin-launcher build is unchanged and still recommended for day-to-day edits; build
  the onefile only to finalize a release (code is baked in → changes need a rebuild).

- **2026-07-22** — Reliability, security, history & UX pass (code-only; **no exe
  rebuild** — pure Python using already-bundled libs). Both runnable copies updated
  (`dist\...\app_logic.py` and `_internal\app_logic.py`); pre-change versions saved
  as `app_logic.py.bak-pre-2026-07-22` in both spots. Verified the exe runs
  independently from the bundled `_internal` copy with the external file removed.
  - **Shipment History (new feature):** every Save & Close / draft save / delete
    writes a JSON record under `...\Shipping App\History\YYYY\YYYY-MM\`. New
    **🕘 History** button opens a searchable viewer (WO / serial # / customer /
    operator / date / division) with CSV export and per-record detail. Nothing extra
    is written inside WO folders.
  - **Close-flow data-loss fix:** photos are tracked to their moved paths and a
    safety draft is re-saved right after the move, so a cancelled/failed close is
    fully recoverable; "Keep Draft" now actually persists the draft and won't blank
    the form if the save fails.
  - **Admin auth hardening:** removed the hardcoded default password from source;
    switched to salted **PBKDF2** with backward-compatible verify + upgrade-on-login;
    roster load now **fails closed** (no default-login fallback on an unreadable
    file); fresh install writes a random one-time password; login warns while the
    known default password is still in use.
  - **Filename & date-code fixes:** document writes no longer fail when an Eng#/Job#
    contains Windows-illegal chars (`: ? * "` etc.); the YYWW date code is correct
    around New Year (was pairing calendar year with ISO week).
  - **Crash-safe file writes:** draft, `admins.json`, `operators.json` are written
    atomically (temp file + rename), so a crash or network hiccup can't corrupt them.
  - **Operator guard rails:** loading a new WO warns before discarding an
    in-progress shipment; closing a scan popup via the window **X** now keeps what
    was scanned (was silently discarded).
  - **Faster, non-blocking WO search:** single pass over the share; the "Searching…"
    window stays responsive instead of going "Not Responding". When nothing matches,
    the message now distinguishes **"server not reachable"** (share can't be opened
    from this PC — shows reachable/scanned/error) from **"searched N folders, no
    match"** (check the number or the Division toggle), so a network/access problem
    no longer masquerades as a bad WO number.
  - **UX polish:** operator name required before Save & Close; modal dialogs no
    longer hide behind the main window (`transient`); duplicate serial scan shows a
    blocking alert so the operator stops and investigates; pasting a serial list now
    **lists the duplicate serials** (with repeat counts), not just a count.

- **2026-07-16** — Folder reorganized and first thin-launcher build produced.
  - Project folder renamed to **`PCC Shipping App`**; `PROJECT_NOTES.md` moved to
    the project root; code/build files kept under `Scripts\`.
  - Built the exe successfully — `Scripts\dist\PCC Shipping App\` now exists
    (exe + editable `app_logic.py` + `logo.png` + `_internal\`).
  - Original source renamed `shipping_app.py` → `PCC Shipping App-Old Reference.py`
    (reference only).
  - Removed the earlier clutter (old one-file `PCC Shipping.exe`, `build\` cache,
    `PCC Shipping_NEW.spec`, `build_exe_OLD_onefile.bat.txt`, sample scans,
    screenshots, stray text/doc files).
  - `logo.png` is no longer in `Scripts\` (only in the built folder). Made
    `logo.png` **optional** in `PCC Shipping App.spec` so rebuilds no longer
    depend on it — the app falls back to the embedded base64 logo. Updated these
    notes to match.

- **2026-07-15** — Converted from a single-file `--onefile` build to a
  **thin-launcher onedir** build.
  - Created `app_logic.py` (exact copy of the original source + a header note).
  - Added thin launcher `PCC Shipping App.py` (pre-imports all runtime libs;
    loads `app_logic.py` next to exe with `_internal` fallback; runs as main).
  - Added `PCC Shipping App.spec` (onedir, windowed; `collect_all` for openpyxl/
    barcode/PIL/docx; bundles fallback `app_logic.py`; lists runtime-only
    submodules as `hiddenimports`).
  - Wrote `build_exe.bat` to build from the spec and auto-copy `app_logic.py`
    next to the built exe.
  - Confirmed the source already had the scroll + pinned-footer UI (kept as-is).
  - Verified: `py_compile` passes on both files; AST scan found no duplicate/
    broken defs (121 functions, 13 classes); ran all four document generators
    (serials xlsx, Word serials, COC xlsx, packing-list xlsx) with sample data —
    all produced valid files (COC embedded 7 barcodes; packing list embedded logo
    + barcodes; Word doc opened correctly).
  - App name confirmed as **PCC Shipping App** (was `PCC Shipping`).
