"""
PCC Shipping Station App — SINGLE-FILE build.
============================================================================
THIS ONE FILE IS THE WHOLE APP. There is no separate launcher and no
runtime-loaded module: PyInstaller builds THIS file directly into a single
"PCC Shipping App.exe" (see "PCC Shipping App.spec" / build_exe.bat).

  * To change behavior: edit this file, then REBUILD the exe (one ~36 MB
    file). There is no "swap the .py without rebuilding" path anymore.
  * The Packing List logo is baked in two ways: as base64 in this file
    (_EMBEDDED_LOGO_B64) and as logo.png in the Scripts folder (bundled at build
    time). Both are the correct logo, so the exe needs no external logo.png at
    runtime. To change the logo: replace that logo.png AND re-encode it into
    _EMBEDDED_LOGO_B64 (base64), then rebuild.

See PROJECT_NOTES.md for build + deploy steps.
============================================================================
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import shutil, csv, datetime, json, sys, hashlib, os, re, time, socket
import queue as _queue
from pathlib import Path
import portal_upload   # Customer Portal upload: queue + background sender + upload guard (portal_upload.py)
try:
    import pdf_docs    # PDF copies of the CoC / Packing List / Serial list (pdf_docs.py, needs reportlab)
except ImportError:
    pdf_docs = None    # the app still runs; the close summary says the PDFs failed

PRODUCTION_ROOTS = {
    "Industrial":  r"\\PCC2\PCC-I-Shares\Production\Industrial",
    "MAG - ITAR":  r"\\PCC2\PCC-I-Shares\Production\MAG - ITAR",
}
APP_DATA_DIR   = r"\\PCC2\PCC-I-Shares\Production\Industrial\Shipping App"
ADMINS_FILE    = APP_DATA_DIR + r"\admins.json"
OPERATORS_FILE = APP_DATA_DIR + r"\operators.json"
BOX_SIZES_FILE = APP_DATA_DIR + r"\box_sizes.json"   # Shipping Slip box sizes, shared (2026-09-25)
HISTORY_DIR    = APP_DATA_DIR + r"\History"
# Operator auto-selected on a fresh app start (she handles ~90% of shipping).
# If she isn't in the roster, the app falls back to the "select your name"
# placeholder — change this one line if the primary operator changes.
DEFAULT_OPERATOR = "Anita Patel"
APP_VERSION    = "v65"

# ── Offline support (PCC2 down) ──────────────────────────────────────────────
# All of the above lives on \\PCC2. These local paths let the app survive an
# outage: cached rosters keep it usable, and OfflineOutbox holds urgent
# shipments (printable documents generated locally) until they can be synced
# back to the WO folder by hand once PCC2 returns.
LOCAL_APP_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "PCC Shipping App"
CACHE_DIR     = LOCAL_APP_DIR / "cache"          # last-known-good operator/admin rosters
OUTBOX_DIR    = LOCAL_APP_DIR / "OfflineOutbox"   # urgent shipments awaiting sync

def _share_host():
    """The server name from a UNC path (\\\\PCC2\\... -> 'PCC2'), for a fast probe."""
    p = APP_DATA_DIR
    if p.startswith("\\\\"):
        return p.lstrip("\\").split("\\")[0]
    return ""

def pcc2_online(timeout=2.0):
    """Fast, non-blocking check that the file server is reachable — a direct
    socket probe to SMB port 445, so we never wait on the 20-60s filesystem
    timeout that freezes the UI. Returns True/False."""
    host = _share_host()
    if not host:
        return True   # not a UNC path (dev/sandbox) — treat as online
    try:
        socket.create_connection((host, 445), timeout=timeout).close()
        return True
    except Exception:
        return False

def _cache_roster(name, data):
    """Save a good roster locally so it's available if the share goes down."""
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        _atomic_write_json(CACHE_DIR / f"{name}.json", data)
    except Exception:
        pass

def _read_cached_roster(name):
    try:
        with open(CACHE_DIR / f"{name}.json", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def create_outbox_shipment(division, wo):
    """Create a local Outbox folder for an urgent offline shipment, mirroring the
    real WO folder layout (Shipping Pictures / Shipping Serials / Certificate of
    Compliance / Packing List) so its contents can be copied straight into the
    \\PCC2 WO folder later. Returns the folder Path, or None on failure."""
    ts = datetime.datetime.now()
    name = f"{ts:%Y%m%d_%H%M%S}_{sanitize_filename_part(division,fallback='Div')}_{sanitize_filename_part(wo,fallback='WO')}"
    root = OUTBOX_DIR / name
    try:
        for sub in (PICTURES_FOLDER, SERIALS_FOLDER, COC_FOLDER, PACKING_LIST_FOLDER):
            (root / sub).mkdir(parents=True, exist_ok=True)
        return root
    except Exception:
        return None

def write_outbox_readme(root, division, wo, operator, outputs):
    """Drop a plain-text note in an Outbox shipment explaining what it is and how
    to sync it, so anyone can complete the filing by hand once PCC2 is back."""
    ts = datetime.datetime.now()
    # The Shipping Slip / Labels pages (Dhruvil 2026-09-24) add a fifth
    # subfolder when they were printed — name it, or whoever syncs copies the
    # four standard ones and deletes the only copy of the pages with the Outbox.
    slip_here = (Path(root) / SLIP_FOLDER).is_dir()
    lines = [
        "PCC SHIPPING APP — OFFLINE SHIPMENT (PCC2 was not reachable)",
        "=" * 60, "",
        f"  Division:      {division}",
        f"  Work order:    {wo}   (typed by the operator — VERIFY before syncing)",
        f"  Operator:      {operator}",
        f"  Shipped:       {ts:%Y-%m-%d %H:%M:%S}   (station {os.environ.get('COMPUTERNAME','?')})",
        "", "TO SYNC ONCE PCC2 IS BACK:", "-" * 60,
        f"  1. On the share, open the real WO folder for {wo} under:",
        f"       {PRODUCTION_ROOTS.get(division,'<production root>')}",
        "  2. Copy the " + ("five" if slip_here else "four") + " subfolders from THIS folder into that WO folder",
        "     (Shipping Pictures, Shipping Serials, Certificate of Compliance,",
        "      Packing List" + (", Shipping Slip" if slip_here else "") + "). They match the normal layout.",
        "  3. Then delete this Outbox folder.", "",
        "  The documents in here are exactly what shipped in the box — do NOT",
        "  regenerate them (dates/serials must match the paperwork).", "",
        "Generated documents:", "-" * 60,
    ]
    for label, p in outputs:
        lines.append(f"  {label}: {Path(p).name if p else '(none)'}")
    try:
        (root / "_OFFLINE_SHIPMENT_README.txt").write_text("\n".join(lines), encoding="utf-8")
    except Exception:
        pass

# ── Monday.com WO auto-lookup ────────────────────────────────────────────────
# Typing a WO# or Eng# looks the job up on the Monday board and pre-fills
# Customer / P/N / PO / Rev — the same data the "PCC Apex" labeling app uses.
# This is an AUGMENT only: the \\PCC2 WO folder is still found and written by
# search_wo_folder(); Monday only fills fields (and is the ONLY source of the
# COC's PO#, which the folder name never contains). Any failure — no network,
# timeout, HTTP error, rate limit, bad key — degrades silently to the normal
# folder-only search, and the lookup is skipped entirely while offline.
#
# SECURITY NOTE: this token is reused from the Apex app and is WRITE-scoped, but
# this feature only READS. Replace it with a dedicated READ-ONLY Monday token
# when you can — it's the only line to change and needs no exe rebuild. To turn
# the whole feature off, set MONDAY_LOOKUP_ENABLED = False.
MONDAY_API_KEY        = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjY0NDI5NDM1MiwiYWFpIjoxMSwidWlkIjozMjA1NzU0NSwiaWFkIjoiMjAyNi0wNC0xMVQwMDo0MzowOS45MzhaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MTI3NzI1NzEsInJnbiI6InVzZTEifQ.sAXMn5AqqgnWQcDHGLMZ31wh1l5LX2SqUeqic9Rkv6Q"
MONDAY_LOOKUP_ENABLED = True          # master switch — set False to disable Monday entirely
MONDAY_API_URL        = "https://api.monday.com/v2"
MONDAY_API_VERSION    = "2023-10"
MONDAY_BOARD_ID       = "3655757841"
MONDAY_TIMEOUT        = 5.0           # seconds; socket connect/read timeout per call
# Skip results sitting in done/gone groups (matches the Apex behavior).
_MONDAY_EXCLUDE_GROUPS = ("ship", "closed", "archive", "complet", "cancel")
# Monday column ids we read (mirrors the Apex query):
#   text__1 = WO#   text6 = Eng#   text5 = PO   lookup_mm06bttc = Rev
#   board_relation0 = Customer (linked item)   item.name = Customer P/N
_MONDAY_COLS = ["text__1", "text5", "text6", "lookup_mm06bttc", "board_relation0"]

def _monday_colval(cv):
    """Extract a display string from a Monday column_value, matching Apex's
    precedence: display_value (mirror) > text > parsed JSON value. Returns ''
    for anything missing/blank so callers never see None."""
    if not cv:
        return ""
    dv = cv.get("display_value")
    if dv:
        return str(dv).strip()
    t = cv.get("text")
    if t:
        return str(t).strip()
    v = cv.get("value")
    if v:
        try:
            p = json.loads(v)
            if isinstance(p, dict):
                if p.get("value") is not None:
                    val = p["value"]
                    return (", ".join(map(str, val)) if isinstance(val, list) else str(val)).strip()
                if p.get("text"):
                    return str(p["text"]).strip()
        except Exception:
            pass
    return ""

def _monday_query(term, timeout):
    """POST the GraphQL lookup and return the parsed JSON (or raise). Stdlib only
    (urllib + ssl + json). The term is passed as a GraphQL VARIABLE so it is never
    concatenated into the query string — no injection, correct escaping."""
    import urllib.request, ssl
    frag = ('column_values(ids:%s){ id text value '
            '... on BoardRelationValue{ linked_items{ id name } } '
            '... on MirrorValue{ display_value } }') % json.dumps(_MONDAY_COLS)
    query = (
        'query ($t:[String]!) {'
        ' by_wo: items_page_by_column_values(limit:30, board_id:%s,'
        '   columns:[{column_id:"text__1", column_values:$t}]) {'
        '     items { id name group{ id title } %s } }'
        ' by_eng: items_page_by_column_values(limit:30, board_id:%s,'
        '   columns:[{column_id:"text6", column_values:$t}]) {'
        '     items { id name group{ id title } %s } }'
        ' }'
    ) % (MONDAY_BOARD_ID, frag, MONDAY_BOARD_ID, frag)
    body = json.dumps({"query": query, "variables": {"t": [term]}}).encode("utf-8")
    req = urllib.request.Request(
        MONDAY_API_URL, data=body, method="POST",
        headers={"Authorization": MONDAY_API_KEY,
                 "Content-Type": "application/json",
                 "API-Version": MONDAY_API_VERSION})
    ctx = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
        return json.loads(resp.read().decode("utf-8", "replace"))

def monday_lookup(term, timeout=None, status=None):
    """Look a WO# or Eng# up on Monday. Returns a list of dicts, each EXACTLY
    {wo,eng,po,rev,customer,cpn,group} with '' (never None) for missing values,
    or [] on ANY failure so the caller degrades to the folder-only search.
    `status`, if a dict is passed, gets status['ok'] = True only when Monday
    ANSWERED (a real 'no match' included) and False on any failure — the
    second-chance lookup is skipped after a failure (2026-09-25).
    Does ZERO Tk work — safe to run on a worker thread."""
    if status is not None:
        status["ok"] = False
    term = (term or "").strip()
    if not term or not MONDAY_LOOKUP_ENABLED or not MONDAY_API_KEY:
        return []
    if timeout is None:
        timeout = MONDAY_TIMEOUT
    try:
        data = _monday_query(term, timeout)
    except Exception:
        return []
    # GraphQL returns HTTP 200 even for errors/rate limits (top-level 'errors');
    # never trust partial data — treat as a miss and degrade to folder search.
    if not isinstance(data, dict) or data.get("errors"):
        return []
    d = data.get("data") or {}
    out, seen = [], set()
    for key in ("by_wo", "by_eng"):
        page = d.get(key) or {}
        for it in (page.get("items") or []):
            iid = str(it.get("id"))
            if iid in seen:
                continue
            seen.add(iid)
            grp = ((it.get("group") or {}).get("title") or "")
            if any(x in grp.lower() for x in _MONDAY_EXCLUDE_GROUPS):
                continue
            cm = {c.get("id"): c for c in (it.get("column_values") or [])}
            rel = cm.get("board_relation0") or {}
            linked = rel.get("linked_items") or []
            customer = (linked[0].get("name") if linked else "") or _monday_colval(rel)
            out.append({
                "wo":       _monday_colval(cm.get("text__1")),
                "eng":      _monday_colval(cm.get("text6")),
                "po":       _monday_colval(cm.get("text5")),
                "rev":      _monday_colval(cm.get("lookup_mm06bttc")),
                "customer": (customer or "").strip(),
                "cpn":      (it.get("name") or "").strip(),
                "group":    grp,
            })
    if status is not None:
        status["ok"] = True
    return out

def monday_hit_matches_folder(hit, wf):
    """Is the Monday item `hit` the job of the WO folder `wf`? The reliable key
    is the Work Order folder name inside Production Data, which always matches
    Monday's WO number; the job-folder name ("Customer PN [Eng - BOM Rev - Assy
    Rev]") is fuzzy, so its Eng# is only a secondary check. (Moved out of
    _load_workorder unchanged, 2026-09-25, so the second-chance lookup uses the
    same rule.)"""
    h = hit or {}
    wf = Path(str(wf))
    mwo = (h.get("wo") or "").strip(); meng = (h.get("eng") or "").strip().upper()
    try:
        feng = (eng_from_inside(split_job_name(wf.parent.parent.name)[1] or "") or "").strip().upper()
    except Exception:
        feng = ""
    wo_ok = bool(mwo) and (wf.name == mwo or any(wf.name.startswith(mwo + s) for s in (".", " ", "_")))
    eng_ok = bool(meng) and meng not in ("", "N/A") and feng == meng
    return wo_ok or eng_ok

def monday_folder_wo(name):
    """The WO number a WO folder's name carries: '16588.1' -> '16588.1',
    '16588.1 REWORK' -> '16588.1', '' -> ''."""
    m = re.match(r"\s*([^\s_]+)", str(name or ""))
    return m.group(1) if m else ""

def pick_monday_hit_for_folder(hits, wf):
    """Second-chance lookup (Dhruvil 2026-09-25: "Customer PO is not filling
    from Monday"): of `hits` — Monday asked with the folder's OWN WO number —
    the ONE item whose WO number is this folder's. None when there is none or
    more than one: never a guess between Monday items."""
    name = Path(str(wf)).name
    exact = []
    for h in (hits or []):
        mwo = (h.get("wo") or "").strip()
        if mwo and (name == mwo or any(name.startswith(mwo + s) for s in (".", " ", "_"))):
            exact.append(h)
    return exact[0] if len(exact) == 1 else None

if getattr(sys, 'frozen', False):
    _APP_DIR = Path(sys.executable).parent
else:
    _APP_DIR = Path(__file__).parent

# Embedded PCC logo (base64-encoded PNG, ~22 KB).
# Used as a guaranteed fallback if logo.png isn't found alongside the EXE,
# so the logo always appears on output documents.
_EMBEDDED_LOGO_B64 = (
    "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"
    "xN9dX6L9AmoUJJ4WYrH4tCxrmpfilvGWXXXV1XCQI1VoJyQWa2LZsTDlG2orqUh2kJPxe4bJZOqakZXRT/x73rFyzvBbdixMnXFk"
    "1Yt8Qkrl0IrRbGdl0tfOch4jahv8WLjrvHCuw1/uhVum/brRhSrUUJJ8Z61ZM0MKIYnjQSyIAkdksJJ+dyuRO+Li1s1PfIbOvcea"
    "WCC97FyUcri2WJMIqUZUlyHj94OMjIxORpPRPGX6lN16gz5FfH3esXqm9tZdi1I+P7z6zHIzCtSR5grtxvr5jtURmvj/gI2Fjm8J"
    "VQBi2QV1pFS7sal4hFNVqLZ82O0NZu1sSCwgFxyPVSUQC32/xN0lYhhmec7nBUl7eVR+aqeFWBamHGoo12rk5WYZvysYjca79Sa9"
    "OacghxVNLUL5zphtTBI/n3esnPP4LaTKzDyy6gUuLUjEAhf7RrN2y/ZSxzmPQSyNZSSxQBUSxGJRhUhiMWvWby1zHppyZ7nu9rpS"
    "7WyQivXx+3hOIV2xuwZXOL8JiWVXa2JJhRcvqULyXiEZvxNk5GZ0MhgNC3Pyc1h2XjYzZBtYflE+K5pWdNyUY0oW1c471s94/iaS"
    "GL44vBLEYpncKHy3cbm2rt5JpH147TaWJXFVqKnijI2F+7GUaddTcUosm0kNqyvRfAFS2SYkFnwGqZHEUsKYe+oLCKi+OOFZIrkW"
    "iQUFkhSpWMcazIkaxuRVIRlXMPR6vQeRyJrc/FxmyjUxY46RF4lYJk+ffJwkFq2oft6xbm78jTsqk2bDd0TKx4OCTISk5tTXOpFY"
    "4FzXVKHNsKzqnJE4uB+L2fVycx2du7408UuQibUqtQ/EUq4tcTenECOJpamMVCHu0n9mdzMID/FY6koSNVXpsku/jCsQEyZM+CeR"
    "Ry7IA1KKRCjWxJJXmMcmT5v8vTHPmCgOO+9YgXw8C5K+Orj8TNoMqER7F4NYtPVb52s7iaptsOqT5H80l+smgVi20bGtlpuJWFxF"
    "6d9lHv5vkmq+hOEWhILjQVCcWMy6EnfVF6hC8LwlYmwT6In+Hq8rTdTJNhYZVxxmzpz5fwaD4f6s7KzluQWtJRXrAmIhVeiH7Nzs"
    "keLQdqPBlHxNQ8nIru6GYlxXEf9PIoM5FmPtGWLBPpuGMm199dz4u0XVNthexTcwmnZVpfzW2o+Fh01Y1+wi0NOeCt0N9WZtPaQj"
    "qGGcWOh4rgqV60rcDfIEVaiuVPsM9X2PNbHw9B9luu9ri+MTZc9bGVcksrKyHiQVp44TS459YsFvRVOLfsrOyx5Oh1xtOdJ9bCtP"
    "HUKT8khtuXaq+0Gotdc2lWvnYSXGmli4oxlN+kaz5i5RtQ2w3NywIEm/e3HqSWvPWzjM0bGb6l0Yb+sXJNzRUJZUz9Or0jmtiYVU"
    "mGL0TVR1CthiGkoShxGZ7bNEkLNch9gY+X1tmSZetrHIuCKhz9b3NuQYanMLc5kxty2pWBHLLzl5Oc/TIW4TC6ksT9CE/glpSvct"
    "ST3dYNZ8sLHs6b+Ln51i14xRf6VJOB/qhyR1tBALqULOiIXvjK5IygCx8L1CIBaa1FCNGstc726un5dwB6lMdW2JZTQRC6lCM+P/"
    "Kao6BXxUYGOh6xASyxliob/fN/DdzbKNRcYViExTZiCRy3aoO/ZIBQXEkleUd9KQaxgqDnOKujLN0zsqk0+snR3PcsY9yjbN07GD"
    "y2FXSJji7tseKlMjSSxtiQXtaOucqUKQWKiuidSmEwigDWKRwiY0lGu/3uIi0NOmWcl31pWQKiSIBWQAiQdqGUlA85GPWVR1CSw3"
    "N5h1e7D0LbUFVaipXHecrq/DqqUMGRcFBoPh+glFE1y+WTMyMoINRsMuGG/tkQoKjLpUTrsiFprwg7ctSDrZWKZhb736CLv1vjAW"
    "/Ugs21qiZQeW8U14+YgMJ6o7BdQImohzLcRisXNgYsLGQpOytrrCiY0FQaJKtdl7lqS2yisEYqkv1653RSxbvtTcVQti4TYWQSxE"
    "TNiE2LQgad7q2c/cIKq6RG2pdij13RLzVkgswt6yH7mLRDUZMi5t6PX6O7Jysj4wZBumTMyZeLv42iGIgEKo7p7CKYV2SQUFxMIl"
    "mnzjc+KwVqgpSXycVI4TTRUa9t7rA1lXPxXz9I9kN98byqKIWLbM17Ejq6BGJBdiKVgc5hzsqquJEASxnHGrx2RvKNfU15WN6Cxq"
    "tgHIq96sNZIK1iZhWX25bn1dsXNigZpVX6ap5atAdF5OLMLG0lyhnbVx1tM3i6ouQQQ1rMGs3cvVMEEsWHpurtDtRooRUU2GjEsT"
    "IBRjtrEMagsc2/RG/bTMzMwbxc8OoTfp1UQaewqmFNglFRS0h3apTdhYWlBr1j61fUHyCZpsbNxYC6F4EKEoAqn0iWQ39YDEEseq"
    "52vZsTVjWH1JwhSoKeJwlwCxWCazZUKicCliAZKOOY7iVrts2HV1ZSSxkCrEpRQQCx0LFaTRrFvfWKoLEFXtAkvZVH/rflJ9pPOC"
    "mLBPaVtFMhFLotvEAge5+hbjrbgGUrHomnZVlzgO/SBDxkXFOP04j6zsrFJMfqlg2TjLlJUzceLE60Q1hyCyiMktzN3lSmIBsWQa"
    "DEPh9LWtPOHRnQtTftm2QMvGjx3cmlBE8SZigcQSOyiObZ6nZd+uIlWoNCHfXYkFy7CN5dq51svNFmIB0ei21M92rEZsXqr5V0Op"
    "Jmc3SSx846GkCkFiKdOtqzVrAkVVu4A0tGNh8iac25pYoAptX5AyY+Wc4beIqi5Ra04kiUW3vzWxcClsR12J44DgMmRcFGTkZ3TX"
    "5+iXYzXH2v8EJED//0ZkM85kMrm0Z5AqFE0Sy27sB4IznDWhSMVCVtnsgynjXt9RkXB0eyVJKG/YJxRrYrmxexiLGRjHNhGxHF1L"
    "EktFUoG7aTOwWkLqzBxrBzkUhC7YXp68xVn6j9qp/a+rL9Uadi1qSyw0yTe5SliGDY7bFiatB5HAttOKWCpTPl8987lbRVWXIDUI"
    "qtA+yQMYBdsSti9M2bZp7sgHRTUZMi4u0tPT/0yEsmzy9Ml2SQCShSnP9BupRe8NGeLaR2JS1qSH86fkOyUWY042mzItjy3/aiwn"
    "FA8ilK4OCEUqPkQs/+4WymIG9WMb52nYD1+/yGpLEie7vVRblf7HulLdXFuX/oMrxzCSOFY5U0cQ85YIxITl5hZViNqwqELaTY0V"
    "zlWharPWY/ei1HUIKkXkZjk3EQsCe++qTP5883/cl1gaSjTDdyCuCxXpGiAJ7ViY0rSlRNNLVJMh4+KhKL3ob0aT0UQSxmnud2KH"
    "BISN5WeSRNLc2dOSacgcQGrQHmfEwqWgvBwikyjWpXe4XSKxLT5BJLH0IFVocBzbOFfLvlszmm0tjs+rnjnEreVmJB0jQvgKk1tS"
    "IbiL/QLdL6RCFDaVOSaWo1Wp19eWavN2Lko5ZU0sls+IxxLvlFgQ03bv0rTVyGHUQixUjqx8AQTxX3dVIbjr1xUnfgC7kKTOoS8g"
    "lt2LUho3zUn0F1VlyLh4INXmGpJEMog8ToFA7JEAVm9y8nJ+zjJmpdIhLp3Z9Eb9IwWTC/Y6IhaoQRkGPVMPeZopAlTMK8A+kdgW"
    "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"
    "qraC3mjUZppyj0wy5rFJRiTl4xL4iY8/mDgj05TpVoaM84bq+QkJ+5aOOvbzBiIUklAgnXBCEYm7zlXBjuVb7w1nr784yEIoa8aw"
    "+nLdCx9NHudFN2gaWFbsjG5VJGIhgviJ2HuU6LZLZBgznqP2vsWxaAdkIgjlSVFFxmUCHuWtXJttqwpxiWX56GpSWTocwKi2JL4f"
    "VKHWxJIKYjmOcAynS9ROwzHUzBp5Gx3zMfxrWoiF/nLjbZl2obPd0ZNMpkh9tmn7tA+ncinDetzjhYj8XCR5P+po533pZ699uHT2"
    "q2z57JfYmvmvs8VfvNTkKtf0eUdtiSZ+/7LRR3/eMIplQOXxU7FOD50bNcc6LgrKrfeHs7GCUI6tJUIpS3pRdOMqk2nCnXTzzJj0"
    "iMJmfXNROLFMzofn4A9UL1kc5hJ6g34ktXcMHodoW5+jlwnlMoUlepwuB74mtsRCUnbN5i+d5/5xBiKP/o0VujbGWyrHm9wIIFVT"
    "mobQlJ+BTFpCU9JnvtfIrKmsrdCOrStLKq4tS5pXZ9bNbzDrindUJBbP/vit7Ya8ojbjXSpcvSGJpaCgoM1OeUhRuxeMfGdnpYbt"
    "Xqhhh5cnsx0LEjfuWTLCZYzp84aGMs3wgytGf/sTJJS3LDYUEMqdD6iIAFSsawc3CqJgdQjObXc9GMw69Qxht90Xxt54WRAKSSiN"
    "5ckv2LIviYGepOKUc2KxsyIEYgF7k6j4HZFFijjMKRC0Sezt+THLmHVFqjw3dw+7xSNANcOrT9RCrwBVlSJAtUQq9D/KIp8gNX6b"
    "7PFgoNvpOrv0jPin50NKu3r9xUJDSfI/tpUn5yE+rDWxHFiOXb6j6rY6SLzuDmqKdXFN5Ul11sSCz7sWp33fVKYd686qEJdYSP1B"
    "QRvSZ5JYfjmwYjQ2DnIv4QPLx7Dja3Rs8Zdj2USj41UcSO5c0razqIAQqnUl2nFYeUKxLJNrNze62IV93lBr1g7jhCKMslB5EBMF"
    "BlTDu0+zxZ8nML/IGB4syR5pOCrIcXzX/X1Z1OCn2Vvj3mMvvfEGezn9bfbC6yls6f/i2fG1o9j2qpRRM2cOsRuhbaJxoieRgBk3"
    "0xGxQGKZPG3yd8ZsY5I4zCl4sOqsrLPamHYpwtMvPMXLX7XcMzBysyJIzXyCY5lP35i2JTiG+dJv2Gfl3UddQ+RS0KVnn7tFM3ah"
    "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"
    "rrrq/wGkmaAzLpwtrgAAAABJRU5ErkJggg=="
)

def _find_logo():
    """
    Returns a Path to a usable logo.png. Search order:
      1. Next to the EXE (operator's own logo.png — lets you override per-site)
      2. PyInstaller bundled data dir (sys._MEIPASS) — if --add-data was used
      3. Embedded base64 logo extracted to a temp file (always available)
    """
    candidates = [_APP_DIR / "logo.png"]
    if hasattr(sys, '_MEIPASS'):
        candidates.append(Path(sys._MEIPASS) / "logo.png")
    for c in candidates:
        try:
            if c.exists():
                return c
        except Exception:
            pass
    # Fallback: extract embedded logo to a temp file
    try:
        import base64, tempfile
        if not hasattr(_find_logo, '_extracted'):
            tmp = tempfile.NamedTemporaryFile(suffix='_pcc_logo.png', delete=False)
            tmp.write(base64.b64decode(_EMBEDDED_LOGO_B64))
            tmp.close()
            _find_logo._extracted = Path(tmp.name)
        return _find_logo._extracted
    except Exception:
        return _APP_DIR / "logo.png"   # truly last resort; may not exist

LOGO_PATH = _find_logo()

PICTURES_FOLDER      = "Shipping Pictures"
SERIALS_FOLDER       = "Shipping Serials"
PACKING_LIST_FOLDER  = "Packing List"
COC_FOLDER           = "Certificate of Compliance"
# SLIP_FOLDER ("Shipping Slip") lives with the slip/label generators further
# down (Dhruvil 2026-09-24) — same leaf-only folder rule as these.
# Folder (inside a testing WO folder, next to Shipping Serials / COC / Packing
# List) holding the per-P/N board-lifecycle report. Never a WO/job candidate.
# Renamed 2026-09-15 (Dhruvil) from "Received after testing" -> "Shipping Reports";
# the legacy name is still recognised so old files MIGRATE into the new one.
RECEIVED_FOLDER        = "Shipping Reports"
RECEIVED_FOLDER_LEGACY = "Received after testing"

def _is_report_folder(name):
    """True for the Shipping Reports folder or its legacy 'Received after testing'
    name — these hold the per-job report, not a work order, so WO/job scans skip
    them (under either name, during and after the rename migration)."""
    n = str(name).strip().casefold()
    return n == RECEIVED_FOLDER.casefold() or n == RECEIVED_FOLDER_LEGACY.casefold()

def _received_fname(pn):
    """The report filename ('<PN> Shipping Report.xlsx') — the ONE place the rule
    lives, so candidate enumeration never has to build a fake path to learn it."""
    return sanitize_filename_part(pn, fallback="PN") + " Shipping Report.xlsx"

def _received_legacy_fname(pn):
    """The pre-2026-09-15 report filename, for finding + migrating old copies."""
    return sanitize_filename_part(pn, fallback="PN") + " - Received after testing.xlsx"

C_GREEN="#1B4D35"; C_GREEN_MED="#2A6B4A"; C_GREEN_LITE="#E8F3ED"
# C_SUBTLE darkened #7A7A7A->#5F5F5F, C_AMBER #B5700A->#8F5807, C_ORANGE #D35400->#B34700
# so small text / white-on-color all clear WCAG 4.5:1 (see 2026-07-22 readability pass).
C_DARK="#1A1A1A"; C_MID="#4A4A4A"; C_SUBTLE="#5F5F5F"; C_BORDER="#D0D0D0"
C_BG="#F2F4F2"; C_WHITE="#FFFFFF"; C_RED="#C0392B"
C_BLUE="#185FA5"; C_BLUE_LITE="#E6F1FB"; C_AMBER="#8F5807"
C_PURPLE="#6C3483"; C_PURPLE_LITE="#F3ECF8"
C_TEAL="#117A65"; C_TEAL_LITE="#E8F6F3"
C_ORANGE="#B34700"; C_ORANGE_LITE="#FDF2E9"

# Typography scale (a start on a shared scale; full migration is a later pass).
F_SECTION=("Segoe UI",11,"bold")   # numbered step headers 1-6
F_CAPTION=("Segoe UI",9)           # small labels / captions
F_VALUE=("Segoe UI",10)            # info-card values
F_VALUE_BOLD=("Segoe UI",11,"bold")# verification-critical values (WO / Customer / P/N)

def _hash(pw): return hashlib.sha256(pw.encode()).hexdigest()

# Salted, iterated password hashing (PBKDF2-HMAC-SHA256). New hashes are stored
# as a self-describing string:  pbkdf2_sha256$<iterations>$<salt_hex>$<hash_hex>
# Legacy accounts may still carry a bare 64-char SHA-256 hex hash; those keep
# working (see verify_password) and are upgraded to the salted format on the
# owner's next successful login.
_PBKDF2_ITERS = 200000

# SHA-256 of the password that used to be hardcoded here. Kept ONLY so the app
# can WARN when that publicly-known default is still in use — it cannot be
# reversed to the plaintext and is never used to grant access.
_KNOWN_DEFAULT_HASH = "4c5f59161c987fd4ec09b134a3a593d6a5e60a2632135204dc5a5ace5e6b5015"

def _consttime_eq(a, b):
    """Length-constant string compare. Hand-rolled so we don't import hmac /
    secrets, which aren't bundled in the frozen exe."""
    if len(a) != len(b): return False
    r = 0
    for x, y in zip(a, b): r |= ord(x) ^ ord(y)
    return r == 0

def hash_password(pw, salt=None):
    """Return a salted PBKDF2 hash string for a plaintext password."""
    if salt is None: salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, _PBKDF2_ITERS)
    return f"pbkdf2_sha256${_PBKDF2_ITERS}${salt.hex()}${dk.hex()}"

def verify_password(pw, stored):
    """Check pw against a stored hash. Returns (ok, is_legacy); is_legacy is True
    when the stored value is a pre-salting bare SHA-256 hash so callers can
    upgrade it. Supports both the salted and the legacy unsalted formats."""
    if not stored: return (False, False)
    if stored.startswith("pbkdf2_sha256$"):
        try:
            _algo, iters, salt_hex, hash_hex = stored.split("$", 3)
            dk = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"),
                                     bytes.fromhex(salt_hex), int(iters))
            return (_consttime_eq(dk.hex(), hash_hex), False)
        except Exception:
            return (False, False)
    return (_consttime_eq(_hash(pw), stored), True)  # legacy unsalted SHA-256

def _gen_password(n=12):
    """Random, human-typeable password (no ambiguous chars). os.urandom based so
    it pulls in no extra imports."""
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789"
    return "".join(alphabet[b % len(alphabet)] for b in os.urandom(n))

def _bootstrap_admins():
    """First run only (no admins.json yet): create an 'admin' account with a
    RANDOM one-time password — no password is ever hardcoded in source. The
    password is written once to a sibling file so an administrator can log in,
    set their own password, then delete the file."""
    pw = _gen_password()
    admins = [{"id": "admin", "name": "Administrator", "pw_hash": hash_password(pw)}]
    save_admins(admins)
    try:
        note = Path(APP_DATA_DIR) / "admin_initial_password.txt"
        note.write_text(
            "PCC Shipping App - initial administrator account\n\n"
            "    ID:       admin\n"
            f"    Password: {pw}\n\n"
            "Log in (gear icon), set your own admin password, then DELETE this\n"
            "file. Anyone who can read this folder can read this password.\n",
            encoding="utf-8")
    except Exception:
        pass
    return admins

def load_admins(online=None):
    if online is None: online = pcc2_online()
    if not online:
        # Offline: use the last-known-good roster cached locally (never bootstrap,
        # never touch the dead share).
        return _read_cached_roster("admins") or []
    try:
        Path(APP_DATA_DIR).mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    try:
        if Path(ADMINS_FILE).exists():
            with open(ADMINS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            _cache_roster("admins", data)   # refresh the offline cache
            return data
    except Exception:
        # File exists but couldn't be read (share hiccup / corrupt JSON). Fail
        # CLOSED: never fall back to a built-in default that would grant login,
        # and never overwrite the real file.
        return []
    # No admins file yet → first-run bootstrap with a random one-time password.
    try:
        return _bootstrap_admins()
    except Exception:
        return []

def _atomic_write_json(target, obj):
    """Write obj as JSON to `target` atomically. Dumps to a station-unique temp
    file in the SAME directory, then os.replace()'s it over the target — an
    atomic rename on NTFS/SMB. A crash, power loss, or share hiccup mid-write
    leaves the previous good file fully intact instead of an empty/half-written
    one. Raises on failure (caller decides how to surface it) after removing the
    temp file."""
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    station = os.environ.get("COMPUTERNAME", "host")
    tmp = target.with_name(f"{target.name}.{station}.{os.getpid()}.tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(obj, f, indent=2)
            f.flush()
            try: os.fsync(f.fileno())
            except Exception: pass   # best-effort durability; not all shares support fsync
        os.replace(str(tmp), str(target))
    except Exception:
        try: tmp.unlink()
        except Exception: pass
        raise

def save_admins(a):
    _atomic_write_json(ADMINS_FILE, a)

def check_admin(aid,pw,admins):
    """True if (aid,pw) matches an admin (salted or legacy hash)."""
    aid_l = aid.lower()
    for a in admins:
        if a.get("id","").lower()==aid_l and verify_password(pw, a.get("pw_hash",""))[0]:
            return True
    return False

def authenticate_admin(aid, pw, admins):
    """Verify a login and, on success, transparently upgrade a legacy unsalted
    hash to the salted format (persisted). Returns (ok, upgraded, is_default);
    is_default flags that the entered password is the publicly-known default."""
    aid_l = aid.lower()
    for a in admins:
        if a.get("id","").lower()==aid_l:
            ok, is_legacy = verify_password(pw, a.get("pw_hash",""))
            if not ok:
                return (False, False, False)
            is_default = (_hash(pw) == _KNOWN_DEFAULT_HASH)
            upgraded = False
            if is_legacy:
                try:
                    a["pw_hash"] = hash_password(pw)
                    save_admins(admins)
                    upgraded = True
                except Exception:
                    upgraded = False
            return (True, upgraded, is_default)
    return (False, False, False)

def load_operators(online=None):
    if online is None: online = pcc2_online()
    if not online:
        return _read_cached_roster("operators") or []
    try:
        if Path(OPERATORS_FILE).exists():
            with open(OPERATORS_FILE,encoding="utf-8") as f: data=json.load(f)
            _cache_roster("operators", data)   # refresh the offline cache
            return data
    except Exception: pass
    return []

def save_operators(ops):
    _atomic_write_json(OPERATORS_FILE, ops)

def search_wo_folder(division,wo,progress=None,stats=None):
    """Find a work-order folder by WO number, falling back to engineering number.

    Single filesystem traversal (the tree is walked ONCE, collecting both WO-name
    matches and engineering-number matches), versus the old two-pass scan. WO
    matches win; engineering matches are returned only if no WO match is found.

    `progress`, if given, is called periodically during the walk so the caller
    can keep its UI repainted — the walk can take several seconds on a large
    share, and without this the window freezes to "Not Responding".

    `stats`, if a dict is passed, is filled with diagnostics so the caller can
    tell "share unreachable" apart from "searched but no match":
      reachable -> did the production root open at all
      scanned   -> how many 'Production Data' folders were walked
      error     -> first scan error string, or None
    """
    root=Path(PRODUCTION_ROOTS[division]); wp=wo[:5] if len(wo)>=5 else wo
    exact,full_p,short_p=[],[],[]
    # Engineering-number form of the query (strip trailing LF, upper-case).
    eng_norm=wo.strip().upper()
    if eng_norm.endswith("LF"): eng_norm=eng_norm[:-2]
    eng_matches=[]
    scanned=0; err=None; reachable=False
    try: reachable=root.is_dir()
    except Exception as e: err=f"{type(e).__name__}: {e}"
    try:
        for i,pd in enumerate(root.rglob("Production Data")):
            if progress and (i % 25 == 0):
                try: progress()
                except Exception: pass
            if not pd.is_dir(): continue
            # Skip ARCHIVED copies: many customer folders keep old versions under an
            # "Archives" / "Archive" / "Archived" / "Old" folder — only the live,
            # non-archived WO folder should be findable. e.g. ...\379 Liber8Now\Archives\
            # GSR10 [...]\Production Data\... is ignored; ...\379 Liber8Now\GSR10 [...]\
            # Production Data\... counts. (Match is a whole folder name, case-insensitive.)
            try: rel_parts = pd.relative_to(root).parts
            except Exception: rel_parts = pd.parts
            if any(part.strip().lower() in ("archives", "archive", "archived", "old") for part in rel_parts):
                continue
            scanned+=1
            # "Received after testing" lives inside Production Data next to the WO
            # folders — it's a report folder, never a work order, so it must not be
            # offered as a WO match (the eng-# fallback would otherwise list it).
            subdirs=[e for e in pd.iterdir() if e.is_dir()
                     and not _is_report_folder(e.name)]
            # ── WO folder-name matching ──
            for e in subdirs:
                n=e.name
                if n==wo: exact.append(e)
                elif n.startswith(wo) and len(n)>len(wo) and n[len(wo)] in ("."," ","_"): full_p.append(e)
                elif n.startswith(wp) and e not in exact: short_p.append(e)
            # ── Engineering-number matching (fallback; collected in the same walk) ──
            _pn, inside = split_job_name(pd.parent.name)
            if inside is None: continue   # no [ and no _ -> no eng number to match
            folder_eng=eng_from_inside(inside).strip().upper()
            if folder_eng=="N/A": continue
            folder_eng_norm=folder_eng[:-2] if folder_eng.endswith("LF") else folder_eng
            if eng_norm==folder_eng_norm or eng_norm==folder_eng:
                for e in subdirs:
                    if e not in eng_matches: eng_matches.append(e)
    except Exception as e:
        if err is None: err=f"{type(e).__name__}: {e}"
    if stats is not None:
        stats["reachable"]=reachable; stats["scanned"]=scanned; stats["error"]=err
    if exact: return exact
    if full_p: return full_p
    if short_p: return short_p
    return eng_matches

def _clean_customer_name(name):
    """
    Strip leading numeric folder prefixes like '10 Honeywell' -> 'Honeywell'.
    Only strips when digits are followed by a separator (space, dash, dot, underscore),
    so embedded digits in real customer names are preserved.

    Examples:
      '10 Honeywell'       -> 'Honeywell'
      '0001 Acme Corp'     -> 'Acme Corp'
      '25-Boeing'          -> 'Boeing'
      '7. Lockheed Martin' -> 'Lockheed Martin'
      'Liber8now'          -> 'Liber8now'   (no separator after 8 -> kept intact)
      '3M'                 -> '3M'          (no separator -> kept intact)
      'Honeywell'          -> 'Honeywell'   (no leading digits -> unchanged)
    """
    if not name: return name
    import re
    m = re.match(r'^\d+[\s\-_.]+(.+)$', name)
    return m.group(1).strip() if m else name

def split_job_name(jn):
    """
    Split a job-folder name into (pn, inside), where `inside` is the
    engineering-number + revision portion. Handles THREE styles:

      Square bracket:  "15102203 Water Draw Box [3156LF Rev. - A Rev. - A]"
                       -> pn="15102203 Water Draw Box", inside="3156LF Rev. - A Rev. - A"
      Underscore:      "15102203 Water Draw Box_3156LF Rev. - A Rev. - A"
                       -> pn="15102203 Water Draw Box", inside="3156LF Rev. - A Rev. - A"
      Bare "Rev":      "161-1061-970 4081 Rev. B00"
                       -> pn="161-1061-970", inside="4081 Rev. B00"

    Rule: the FIRST underscore is always a divider; a P/N never contains '_'.
    If both '[' and '_' appear, whichever occurs first in the string wins. With
    neither, the "Rev" keyword divides the Eng #/rev tail from the P/N.
    Returns (pn, inside) — or (jn_stripped, None) if nothing usable is present.
    """
    if not jn:
        return ("", None)
    bp = jn.find("[")
    up = jn.find("_")

    use_bracket = use_underscore = False
    if bp >= 0 and up >= 0:
        # Both present — the earlier delimiter wins (first underscore is a divider)
        if bp < up: use_bracket = True
        else:       use_underscore = True
    elif bp >= 0:
        use_bracket = True
    elif up >= 0:
        use_underscore = True

    if use_bracket:
        pn = jn[:bp].strip()
        inside = jn[bp+1:]
        cl = inside.find("]")
        if cl >= 0:
            inside = inside[:cl]
        return (pn, inside.strip())
    if use_underscore:
        pn = jn[:up].strip()
        inside = jn[up+1:].strip()
        return (pn, inside)
    # THIRD style — no bracket, no underscore, e.g. "161-1061-970 4081 Rev. B00".
    # Treat "Rev" as a divider ONLY when it is a standalone keyword FOLLOWED BY a
    # revision value — so a trailing bare "REV", or "REV" inside a part number like
    # "REV-BOARD", is never mistaken for it. The Eng # is the token right before
    # "Rev" ONLY when it looks like one (contains a digit, e.g. "4081"/"3156LF"); a
    # plain descriptive word ("Box" in "Water Draw Box Rev. A") stays in the P/N.
    #   "161-1061-970 4081 Rev. B00"     -> pn="161-1061-970",            inside="4081 Rev. B00"
    #   "15102203 Water Draw Box Rev. A" -> pn="15102203 Water Draw Box", inside="Rev. A"  (no Eng #)
    m = re.search(r'\bRev\b\.?\s+[A-Za-z0-9]', jn, flags=re.IGNORECASE)
    if m and m.start() > 0:
        head = jn[:m.start()].strip()
        tail = jn[m.start():].strip()          # keeps "Rev..." so rev_from_inside() works
        toks = head.split()
        if len(toks) >= 2 and any(c.isdigit() for c in toks[-1]):
            return (" ".join(toks[:-1]).strip(), (toks[-1] + " " + tail).strip())
        if toks:
            return (head, tail)                # whole head is the P/N; no separate Eng #
    return (jn.strip(), None)


def eng_from_inside(inside):
    """From the inside portion of a job name, collect whitespace-separated
    tokens up to (but not including) the first 'Rev...' token -> engineering #.
    Returns 'N/A' if nothing usable."""
    if not inside:
        return "N/A"
    ep = []
    for p in inside.split():
        if p.lower().startswith("rev"):
            break
        ep.append(p)
    return " ".join(ep) if ep else "N/A"


def _revs_in(inside):
    """Every genuine rev value in the inside portion of a job name, in order —
    '3156LF Rev. - A Rev. - A' -> ['A', 'A']; '3267 Rev.  Rev. ' -> []. A rev
    may carry dots or hyphens INSIDE it ('2.02', '1.1.10', a date rev like
    '5-4-23'): until 2026-10-05 the match stopped at the first dot, so
    'Rev. 2.02' read as '2' and 'Rev. 1.1.10' as '1' (8 of 45 Industrial job
    folders). A trailing dot / hyphen is never part of it, and the '-' the
    names use for "no rev" ('Rev. - A', 'Rev. -') is still skipped."""
    if not inside:
        return []
    return [m for m in re.findall(r'\bRev\b\.?\s*-?\s*([A-Za-z0-9](?:[A-Za-z0-9.\-]*[A-Za-z0-9])?)', inside,
                                  flags=re.IGNORECASE)
            if m.lower() != "rev"]

def rev_from_inside(inside):
    """The LAST genuine rev in the inside portion of a job name (the Assy rev
    slot when both are filled), e.g. '3156LF Rev. - A' -> 'A'; '' when there's
    no real rev (e.g. '3267 Rev.  Rev. ' placeholders)."""
    revs = _revs_in(inside)
    return revs[-1] if revs else ""

def bom_rev_from_inside(inside):
    """The BOM rev in the inside portion of a job name: the FIRST genuine rev —
    job folders are named 'Customer PN [PCC Eng - BOM Rev - Assy Rev]'. The
    COC and the Shipping Slip / Labels always carry the BOM rev (Dhruvil
    2026-10-05); this is the fallback when Monday has no BOM rev for the job
    (_merge_monday). '1427765_3486LF Rev. A Rev. 1' -> 'A' (BOM A, Assy 1);
    '3414 Rev. 2.02 Rev. ' -> '2.02'; '' when there's no real rev."""
    revs = _revs_in(inside)
    return revs[0] if revs else ""

def monday_bom_rev(raw):
    """Monday's BOM rev as the paperwork should carry it. Monday's 'Rev' IS its
    'BOM Rev' column (lookup_mm06bttc, mirrored from the Engineering board's
    BOM Rev). '' when Monday has none ('', '-', 'N/A') or when the mirror shows
    several DIFFERENT revs (more than one Engineering item linked: never a
    guess); 'A, A' -> 'A'. A 'Rev' word typed into the Monday value is dropped
    ('Rev 03' -> '03') — the COC already prints 'Rev.-' in front of it."""
    vals = []
    for part in str(raw or "").split(","):
        v = re.sub(r"^rev(?:ision)?\b\.?\s*-?\s*", "", part.strip(), flags=re.IGNORECASE).strip()
        if v and v.lower() not in ("-", "--", "n/a", "na") and v not in vals:
            vals.append(v)
    return vals[0] if len(vals) == 1 else ""


def parse_wo_path(wf):
    r={"customer":"N/A","pn":"N/A","eng":"N/A","rev":""}
    try:
        pd=wf.parent; jf=pd.parent; cf=jf.parent
        if pd.name.lower()!="production data": return r
        r["customer"]=_clean_customer_name(cf.name) or "N/A"
        pn, inside = split_job_name(jf.name)
        if inside is not None:
            r["pn"]=pn if pn else "N/A"
            r["eng"]=eng_from_inside(inside)
            r["rev"]=bom_rev_from_inside(inside)     # the BOM rev (first in the name), 2026-10-05
        else:
            r["pn"]=pn if pn else "N/A"
    except Exception: pass
    return r

def sort_serials(serials):
    """Return serials in increasing NATURAL order ("2" before "10", "SN2" before
    "SN10"). Used to sort the WHOLE shipment's serials at Save Draft / Save and
    Close time (draft file, serials xlsx, COC, packing, history). Live scanning
    keeps its raw scan order elsewhere so the operator can see what they scanned
    last."""
    def key(s):
        # Tuple (type_tag, value) per segment so a digit run always sorts by its
        # numeric value, and mixed formats ("S1" vs "10") never compare str-vs-int.
        return [(1, int(t)) if t.isdigit() else (0, t.lower())
                for t in re.split(r'(\d+)', str(s)) if t != ""]
    try:
        return sorted(serials, key=key)
    except Exception:
        return list(serials)

def ensure_folder(path):
    try: path.mkdir(parents=True,exist_ok=True); return True
    except Exception as e: messagebox.showerror("Folder error",f"Could not create:\n{path}\n\n{e}"); return False

def fmt_date():
    n=datetime.datetime.now()
    return n.strftime("%#m/%#d/%Y") if sys.platform=="win32" else n.strftime("%-m/%-d/%Y")
def fmt_time(): return datetime.datetime.now().strftime("%I:%M %p")
def fmt_datecode():
    # Use the ISO year that goes WITH the ISO week, not the calendar year.
    # Pairing %y (calendar year) with %V (ISO week) produces a wrong YYWW code
    # around New Year (ISO week 1 can fall in late December, week 52/53 in early
    # January). isocalendar() keeps the year and week consistent.
    iso=datetime.date.today().isocalendar()
    return f"{iso[0]%100:02d}{iso[1]:02d}"
def fmt_file_date():
    n=datetime.datetime.now()
    if sys.platform=="win32": return n.strftime("%#m-%#d-")+n.strftime("%y")
    return n.strftime("%-m-%-d-")+n.strftime("%y")

def fmt_time_suffix():
    """Returns HHMMSS for use as a uniqueness suffix on same-day repeat shipments."""
    return datetime.datetime.now().strftime("%H%M%S")

def sanitize_filename_part(s, fallback=""):
    """Make a string safe to use as a Windows filename component. Replaces every
    character illegal on Windows ( < > : " / \\ | ? * and control chars ) with
    '_' and trims trailing dots/spaces (also illegal). Falls back to a sanitized
    `fallback`, or 'unknown', if nothing usable remains — so a document write is
    never aborted by a stray character in a free-text Eng#/Job# field."""
    bad = r'[<>:"/\\|?*\x00-\x1f]'
    def _clean(x):
        return re.sub(bad, "_", str(x)).strip().rstrip(". ").strip()
    # Require at least one alphanumeric char, so an all-punctuation field (e.g.
    # "***" -> "___") falls back to something meaningful instead of a name made
    # only of underscores.
    cleaned = _clean(s)
    if cleaned and re.search(r"[A-Za-z0-9]", cleaned):
        return cleaned
    if fallback:
        fb = _clean(fallback)
        if fb and re.search(r"[A-Za-z0-9]", fb):
            return fb
    return "unknown"

def unique_file_path(folder, stem, ext):
    """
    Return a Path inside <folder> that does not yet exist on disk.
    Tries <stem><ext> first. If it exists, returns <stem>_<HHMMSS><ext>.
    If THAT exists too (extremely unlikely), tries <stem>_<HHMMSS>_<seq><ext>.

    Used to ensure new shipments never overwrite previous shipments' files
    on the same day. The first shipment of the day gets the clean name;
    repeats add a timestamp.
    """
    folder = Path(folder)
    base = folder / f"{stem}{ext}"
    if not base.exists():
        return base
    ts = fmt_time_suffix()
    base = folder / f"{stem}_{ts}{ext}"
    if not base.exists():
        return base
    seq = 2
    while True:
        candidate = folder / f"{stem}_{ts}_{seq}{ext}"
        if not candidate.exists():
            return candidate
        seq += 1

def unique_dir_path(parent, stem):
    """
    Like unique_file_path but for directories. Tries <stem>/, then
    <stem>_<HHMMSS>/, then <stem>_<HHMMSS>_<seq>/.
    """
    parent = Path(parent)
    base = parent / stem
    if not base.exists():
        return base
    ts = fmt_time_suffix()
    base = parent / f"{stem}_{ts}"
    if not base.exists():
        return base
    seq = 2
    while True:
        candidate = parent / f"{stem}_{ts}_{seq}"
        if not candidate.exists():
            return candidate
        seq += 1

###############################################################################
# FILE WRITERS
###############################################################################

def write_serial_xlsx(serials_path,eng,wo,serials,operator):
    """
    Always creates a NEW serials xlsx — never appends. Filename uses
    Engineering # (more meaningful than WO# for cross-referencing).
    Same-day repeat shipments automatically get a timestamp suffix.
    """
    fdate = fmt_file_date()
    stem = f"{eng}-{fdate} serials"
    d,t = fmt_date(),fmt_time()
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment,Font
        ctr = Alignment(horizontal='center',vertical='center')
        bf  = Font(bold=True,size=11)
        nf  = Font(size=11)
        out = unique_file_path(serials_path, stem, ".xlsx")
        wb = Workbook(); ws = wb.active; ws.title = "Serial Numbers"
        ws.append(["Date","Time","Serial_Number","Operator","WO"])
        for c in ws[1]: c.font = bf; c.alignment = ctr
        for sn in serials:
            ws.append([d,t,sn,operator,wo])
            for c in ws[ws.max_row]: c.alignment = ctr; c.font = nf
        for col in ws.columns:
            mx = max((len(str(c.value)) for c in col if c.value), default=8)
            ws.column_dimensions[col[0].column_letter].width = max(mx+4, 12)
        wb.save(str(out)); return str(out)
    except ImportError:
        cp = unique_file_path(serials_path, stem, ".csv")
        with open(cp,"w",newline="",encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Date","Time","Serial_Number","Operator","WO"])
            for sn in serials: w.writerow([d,t,sn,operator,wo])
        return str(cp)

def write_word_serials(serials_path,eng,wo,serials):
    """
    Always creates a NEW Word doc — never appends. Filename uses Engineering #.
    """
    fdate = fmt_file_date()
    stem = f"{eng}-{fdate} serials"
    mx = max((len(s) for s in serials), default=10)
    pr = 3 if mx>25 else (5 if mx>18 else 10)
    rows = [", ".join(serials[i:i+pr]) for i in range(0,len(serials),pr)]
    try:
        from docx import Document; from docx.shared import Pt,RGBColor
        out = unique_file_path(serials_path, stem, ".docx")
        doc = Document()
        h = doc.add_heading(f"Serial Numbers — WO {wo} (Eng #{eng})", 0)
        h.runs[0].font.color.rgb = RGBColor(0x1B,0x4D,0x35)
        p = doc.add_paragraph(f"Date: {fmt_date()}     Total serials: {len(serials)}")
        p.paragraph_format.space_after = Pt(0)
        for row in rows:
            p = doc.add_paragraph(row); p.runs[0].font.size = Pt(10)
            p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
        doc.save(str(out)); return str(out)
    except ImportError:
        tp = unique_file_path(serials_path, stem, ".txt")
        with open(tp,"w",encoding="utf-8") as f:
            f.write(f"Serial Numbers — WO {wo} (Eng #{eng})\nDate: {fmt_date()}\nTotal: {len(serials)}\n\n")
            for row in rows: f.write(row+"\n")
        return str(tp)

###############################################################################
# COC XLSX
###############################################################################

def write_coc_xlsx(coc_path, eng, wo, coc_data, serials):
    """
    Always creates a NEW COC xlsx — never appends. Filename uses Engineering #.
    Same-day repeat shipments automatically get a timestamp suffix.
    """
    fdate = fmt_file_date()
    stem = f"{eng}-{fdate}-COC"
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        from openpyxl.utils.units import pixels_to_EMU
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        import tempfile
    except ImportError:
        tp = unique_file_path(coc_path, stem, ".txt")
        with open(tp, "w", encoding="utf-8") as f:
            pn = coc_data['pn']; rv = coc_data.get('rev', ''); rs = f" Rev.- {rv}" if rv else " Rev.-"
            f.write(f"CERTIFICATE OF COMPLIANCE\n\n")
            f.write(f"This is to certify...{pn}{rs} drawings.\n\n")
            for k, v in [("Ship Date", coc_data['ship_date']), ("PO #", coc_data['po']),
                         ("P/N", pn), ("QTY", f"{coc_data['qty']}ea."),
                         ("DC", f"{coc_data['datecode']} (YYWW)"), ("PCC Job#", coc_data.get('job', '')),
                         ("Type", coc_data['type']), ("Assembly process", coc_data['assembly_process'])]:
                f.write(f"{k}:\t\t{v}\n")
            f.write("\nSerial No:\n")
            for i, sn in enumerate(serials, 1): f.write(f"{i}. {sn}\n")
            f.write("\nHung Ly\nQC\nPrinted Circuits Corp.\n")
        return str(tp)

    try:
        from barcode import Code128
        from barcode.writer import ImageWriter
        from io import BytesIO
        has_barcode = True
    except ImportError:
        has_barcode = False

    out = unique_file_path(coc_path, stem, ".xlsx")
    wb = Workbook(); ws = wb.active
    ws.title = "Certificate of Compliance"
    start_row = 1

    # Styles — Aptos Display for headings, Aptos for body
    green_font = Font(name='Aptos Display', size=20, bold=True, color='2A9D5C')
    red_font = Font(name='Aptos', size=11, color='C0392B')
    normal = Font(name='Aptos', size=11)
    bold_font = Font(name='Aptos', size=11, bold=True)
    title_font = Font(name='Aptos Display', size=14, bold=True)
    center_al = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_al = Alignment(horizontal='left', vertical='center', wrap_text=True)
    center_mid = Alignment(horizontal='center', vertical='center')
    right_al = Alignment(horizontal='right', vertical='center')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'))

    # Column widths - generous sizing
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 45
    ws.column_dimensions['C'].width = 28

    r = start_row

    # Company header
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1, value="PRINTED CIRCUITS CORP.")
    c.font = green_font; c.alignment = center_al
    ws.row_dimensions[r].height = 32
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1, value="Design, Manufacturing and Assembly of Printed Circuit Boards")
    c.font = red_font; c.alignment = center_al
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1, value="5295 Webb Parkway, Lilburn, GA 30047")
    c.font = normal; c.alignment = center_al
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1, value="Phone: (770) 638 8658   Fax: (770) 638 8659")
    c.font = normal; c.alignment = center_al
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1, value="Email: contact@pcc-i.com, Website: www.pcc-i.com")
    c.font = normal; c.alignment = center_al
    r += 2

    # Title
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1, value="CERTIFICATE OF COMPLIANCE")
    c.font = title_font; c.alignment = center_al
    r += 2

    # Certification text
    pn = coc_data['pn']; rv = coc_data.get('rev', '')
    rs = f" Rev.- {rv}" if rv else " Rev.-"
    cert_text = (f"This is to certify that the following part number was manufactured "
                 f"and assembled as per drawings and specifications of the {pn}{rs} drawings.")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1, value=cert_text)
    c.font = normal; c.alignment = left_al
    ws.row_dimensions[r].height = 32
    r += 2

    # Fields
    fields = [
        ("Ship Date:", coc_data['ship_date']),
        ("PO #:", coc_data['po']),
        ("P/N:", coc_data['pn']),
        ("QTY:", f"{coc_data['qty']}ea."),
        ("DC:", f"{coc_data['datecode']} (YYWW)"),
        ("PCC Job#:", coc_data.get('job', '')),
        ("Type:", coc_data['type']),
        ("Assembly process:", coc_data['assembly_process']),
    ]
    for label, value in fields:
        # Bordered table (label | value) spanning the FULL width (A:C) so it lines up
        # with the serial table and sits centered on the page (rather than a narrow
        # block hugging the left). Value cell is merged B:C and centered; Excel draws
        # the box around the merged region from the anchor cell's (B) borders.
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = bold_font; lc.alignment = center_mid; lc.border = thin_border
        vc = ws.cell(row=r, column=2, value=value)
        vc.font = normal; vc.alignment = center_mid; vc.border = thin_border
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 20
        r += 1

    r += 1

    # Serial numbers header
    for ci, txt in enumerate(["#", "Serial Number", "Barcode"], 1):
        c = ws.cell(row=r, column=ci, value=txt)
        c.font = bold_font; c.alignment = center_mid; c.border = thin_border
    ws.row_dimensions[r].height = 22
    r += 1

    # Serial numbers with barcodes
    temp_files = []
    # Render the barcode identically to the Packing List: same module options, and
    # the bars occupy the same fraction (BC_FILL) of the barcode column, centered.
    # Pad the PNG with symmetric white space, then display it across the full width
    # of column C via a OneCellAnchor.
    from PIL import Image as PILImage
    col_c_px = int(28 * 7.0)   # column C width = 28 chars
    BC_FILL = 0.86             # must match the Packing List serial barcode fill
    barcode_w_px = col_c_px
    barcode_h_px = 30
    row_ht = 30   # was 48 — tightened to cut paper use on large (50-200 serial) shipments
    barcode_row_offset = pixels_to_EMU(max(0, (int(row_ht * 1.33) - barcode_h_px) // 2))

    serial_start_row = r
    for i, sn in enumerate(serials, 1):
        c1 = ws.cell(row=r, column=1, value=i)
        c1.font = normal; c1.alignment = center_mid; c1.border = thin_border
        c2 = ws.cell(row=r, column=2, value=sn)
        c2.font = normal; c2.alignment = center_mid; c2.border = thin_border
        c3 = ws.cell(row=r, column=3)
        c3.border = thin_border; c3.alignment = center_mid

        if has_barcode:
            try:
                from openpyxl.drawing.image import Image as XlImage
                buffer = BytesIO()
                bc_options = {
                    'module_height': 8.0,
                    'module_width': 0.25,
                    'quiet_zone': 1.0,
                    'write_text': False,
                }
                Code128(sn, writer=ImageWriter()).write(buffer, options=bc_options)
                buffer.seek(0)
                bc_im = PILImage.open(buffer).convert("RGB")
                nat_w, nat_h = bc_im.size
                padded_w = max(nat_w, int(round(nat_w / float(BC_FILL))))
                canvas = PILImage.new("RGB", (padded_w, nat_h), "white")
                canvas.paste(bc_im, ((padded_w - nat_w) // 2, 0))
                tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                canvas.save(tmp.name); tmp.close()
                temp_files.append(tmp.name)
                img = XlImage(tmp.name)
                # TwoCellAnchor spanning the barcode column (C, i.e. col index 2 to 3)
                # forces a fit-to-cell; symmetric white padding centers the bars.
                from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor as _TCA
                _from = AnchorMarker(col=2, colOff=pixels_to_EMU(0), row=r - 1, rowOff=barcode_row_offset)
                _to   = AnchorMarker(col=3, colOff=pixels_to_EMU(0), row=r,     rowOff=-barcode_row_offset)
                img.anchor = _TCA(editAs='oneCell', _from=_from, to=_to)
                ws.add_image(img)
            except Exception:
                pass

        ws.row_dimensions[r].height = row_ht
        r += 1

    r += 1

    # Signature
    ws.cell(row=r, column=1, value="Hung Ly").font = normal; r += 1
    ws.cell(row=r, column=1, value="QC").font = normal; r += 1
    ws.cell(row=r, column=1, value="Printed Circuits Corp.").font = normal

    last_row = r
    last_col = 3

    # Print area and page setup
    ws.print_area = f"A1:{get_column_letter(last_col)}{last_row}"
    from openpyxl.worksheet.properties import PageSetupProperties
    # Enable fit-to-page — WITHOUT this, fitToWidth/Height are ignored and the sheet
    # prints at 100% and spills. Scale to ONE page for normal shipments; only a
    # genuinely long serial list (> _MAX_ONE_PAGE_SERIALS) is allowed to paginate.
    ws.page_setup.fitToWidth = 1          # fit ALL columns onto one page wide (the real fix)
    ws.page_setup.fitToHeight = 0         # flow down as many pages as needed — no forced shrink
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True   # center the whole sheet on the page
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    wb.save(str(out))

    for tf in temp_files:
        try: Path(tf).unlink()
        except Exception: pass

    return str(out)


###############################################################################
# PACKING LIST XLSX
###############################################################################

def write_packing_list_xlsx(pl_path, eng, wo, pl_data, serials, logo_path=None):
    """
    Packing list — clean look matching COC styling.
    Always creates a NEW file — never appends. Filename uses Engineering #.
    Same-day repeat shipments automatically get a timestamp suffix.
    """
    fdate = fmt_file_date()
    stem = f"{eng}-{fdate}-Packing_List"
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        from openpyxl.utils.units import pixels_to_EMU
        from openpyxl.drawing.image import Image as XlImage
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        import tempfile
    except ImportError:
        tp = unique_file_path(pl_path, stem, ".txt")
        with open(tp, "w", encoding="utf-8") as f:
            f.write("PACKING LIST\n\n")
            for k in ('customer','packing_no','po','pn','wo','eng','date','datecode'):
                f.write(f"{k}: {pl_data.get(k,'')}\n")
            f.write(f"\nSerials ({len(serials)}):\n")
            for i, sn in enumerate(serials, 1): f.write(f"{i}. {sn}\n")
            f.write(f"\nTotal units: {len(serials)}\n")
        return str(tp)

    try:
        from barcode import Code128
        from barcode.writer import ImageWriter
        from io import BytesIO
        has_barcode = True
    except ImportError:
        has_barcode = False

    out = unique_file_path(pl_path, stem, ".xlsx")
    wb = Workbook(); ws = wb.active
    ws.title = "Packing List"
    start_row = 1

    # ── Hide Excel's default gridlines for clean look ─────────
    ws.sheet_view.showGridLines = False

    # ── Fonts (matching COC color palette) ────────────────────
    F_TITLE     = Font(name='Aptos Display', size=22, bold=True, color='FF2A9D5C')
    F_SUBTITLE  = Font(name='Aptos Display', size=14, bold=True, color='FFC0392B')
    F_ADDR      = Font(name='Aptos', size=10)
    F_LABEL     = Font(name='Aptos', size=8, bold=True, color='FF666666')
    F_XL_VALUE  = Font(name='Aptos', size=12, bold=True)   # renamed: shadowed the module-level F_VALUE ui font
    F_HDR       = Font(name='Aptos', size=11, bold=True)
    F_BODY      = Font(name='Aptos', size=11)
    F_BODY_BOLD = Font(name='Aptos', size=11, bold=True)
    F_TOTAL_LBL = Font(name='Aptos', size=12, bold=True)
    F_TOTAL_VAL = Font(name='Aptos', size=18, bold=True)

    AL_R   = Alignment(horizontal='right',  vertical='center')
    AL_C   = Alignment(horizontal='center', vertical='center')
    AL_LI  = Alignment(horizontal='left',   vertical='center', indent=1)
    AL_RI  = Alignment(horizontal='right',  vertical='center', indent=1)

    thin      = Side(style='thin',  color='000000')
    thick     = Side(style='thick', color='000000')
    full      = Border(left=thin, right=thin, top=thin, bottom=thin)
    btm_thick = Border(bottom=thick)

    # ── Column widths sized to fill ~7.5" letter-page printable area ─
    # All 12 columns set to equal width 9 → each info-table cell is uniformly
    # 3 columns × 9 = 27 units wide. This gives every cell the same breathing
    # room and lets values like "15102203 Water Draw Box" fit on one line.
    widths = {
        'A': 9,  'B': 9,  'C': 9,  'D': 9,  'E': 9,
        'F': 9,  'G': 9,  'H': 9,  'I': 9,
        'J': 9,  'K': 9,  'L': 9,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    r = start_row
    HDR_START = r

    # ── HEADER COMPANY INFO (right side, cols D:L) ──────────
    # Row 1: Title (green)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=12)
    c = ws.cell(row=r, column=4, value="PRINTED CIRCUITS CORP.")
    c.font = F_TITLE; c.alignment = AL_R
    ws.row_dimensions[r].height = 36; r += 1

    # Row 2: Subtitle (red)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=12)
    c = ws.cell(row=r, column=4, value="PACKING LIST")
    c.font = F_SUBTITLE; c.alignment = AL_R
    ws.row_dimensions[r].height = 24; r += 1

    # Row 3: spacer
    ws.row_dimensions[r].height = 6; r += 1

    # Row 4: address
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=12)
    c = ws.cell(row=r, column=4, value="5295 Webb Parkway, Lilburn, GA 30047")
    c.font = F_ADDR; c.alignment = AL_R
    ws.row_dimensions[r].height = 20; r += 1

    # Row 5: phone/fax
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=12)
    c = ws.cell(row=r, column=4, value="Phone: (770) 638 8658   Fax: (770) 638 8659")
    c.font = F_ADDR; c.alignment = AL_R
    ws.row_dimensions[r].height = 20; r += 1

    # Rows 6-7: small extra height so logo can be slightly taller than the
    # company info text section. Kept minimal so divider sits close to phone/fax.
    ws.row_dimensions[r].height = 12; r += 1
    ws.row_dimensions[r].height = 8
    HDR_END = r
    r += 1

    # ── LOGO inside merged single cell A:C across header rows ─
    ws.merge_cells(start_row=HDR_START, start_column=1,
                   end_row=HDR_END, end_column=3)
    if logo_path:
        try:
            lp = Path(logo_path)
            if lp.exists():
                try:
                    from PIL import Image as PILImage
                    with PILImage.open(str(lp)) as pil:
                        ow, oh = pil.size
                    ratio = oh / ow
                except Exception:
                    ratio = 1.15
                # Width-first sizing, capped at height
                target_w = 110
                target_h = int(target_w * ratio)
                max_h = 120
                if target_h > max_h:
                    target_h = max_h
                    target_w = int(target_h / ratio)
                img = XlImage(str(lp))
                img_w_emu = pixels_to_EMU(target_w)
                img_h_emu = pixels_to_EMU(target_h)
                marker = AnchorMarker(col=0, colOff=pixels_to_EMU(8),
                                       row=HDR_START - 1, rowOff=pixels_to_EMU(8))
                size = XDRPositiveSize2D(img_w_emu, img_h_emu)
                img.anchor = OneCellAnchor(_from=marker, ext=size)
                ws.add_image(img)
        except Exception:
            pass

    # ── Heavy black divider line ──────────────────────────────
    for col_i in range(1, 13):
        ws.cell(row=r, column=col_i).border = btm_thick
    ws.row_dimensions[r].height = 4; r += 1
    ws.row_dimensions[r].height = 8; r += 1

    # Shared temp_files list — collects barcode PNGs from both the info table
    # and serial table, all cleaned up after wb.save() at the end.
    temp_files = []

    # ── INFO TABLE (4 cells row 1, 3 cells row 2) ─────────────
    # Each cell now has 3 stacked rows: label, value, barcode.
    # Barcode generated from the value content; skipped if value is empty.
    # Strategy per cell based on value length:
    #   - Up to 26 chars: normal single-line display, full font size
    #   - 27-50 chars:    wrap to 2 lines, full font size (row grows automatically)
    #   - 51+ chars:      shrink-to-fit on single line (font auto-reduced by Excel)
    INFO_BC_W_PX = 110     # barcode display width in info table cells
    INFO_BC_H_PX = 22      # barcode display height
    ONE_LINE_MAX = 18      # threshold below which value fits cleanly on 1 line
    TWO_LINE_MAX = 40      # above this, even 2 lines won't fit — shrink instead

    def _value_strategy(v):
        """Return one of: 'normal', 'wrap', 'shrink' based on string length."""
        if v is None: return 'normal'
        n = len(str(v))
        if n <= ONE_LINE_MAX: return 'normal'
        if n <= TWO_LINE_MAX: return 'wrap'
        return 'shrink'

    def write_info_pair(label_row, value_row, barcode_row, col_start, col_end, label, value):
        # Merge the three stacked rows
        ws.merge_cells(start_row=label_row,   start_column=col_start, end_row=label_row,   end_column=col_end)
        ws.merge_cells(start_row=value_row,   start_column=col_start, end_row=value_row,   end_column=col_end)
        ws.merge_cells(start_row=barcode_row, start_column=col_start, end_row=barcode_row, end_column=col_end)
        # Label
        lc = ws.cell(row=label_row, column=col_start, value=label)
        lc.font = F_LABEL; lc.alignment = AL_LI
        # Value — pick alignment based on length-based strategy.
        # 'shrink' and 'wrap' are mutually exclusive in Excel so we never combine.
        vc = ws.cell(row=value_row, column=col_start, value=value)
        vc.font = F_XL_VALUE
        strat = _value_strategy(value)
        if strat == 'shrink':
            vc.alignment = Alignment(horizontal='left', vertical='center', indent=1, shrink_to_fit=True)
        elif strat == 'wrap':
            vc.alignment = Alignment(horizontal='left', vertical='center', indent=1, wrap_text=True)
        else:
            vc.alignment = AL_LI
        # Outer outline only — top on label row, bottom on barcode row,
        # left/right on edge cells, NO internal horizontal dividers.
        for cc in range(col_start, col_end + 1):
            is_first = (cc == col_start); is_last = (cc == col_end)
            kw_top = {'top': thin}
            if is_first: kw_top['left']  = thin
            if is_last:  kw_top['right'] = thin
            ws.cell(row=label_row, column=cc).border = Border(**kw_top)
            kw_mid = {}
            if is_first: kw_mid['left']  = thin
            if is_last:  kw_mid['right'] = thin
            if kw_mid:
                ws.cell(row=value_row, column=cc).border = Border(**kw_mid)
            kw_btm = {'bottom': thin}
            if is_first: kw_btm['left']  = thin
            if is_last:  kw_btm['right'] = thin
            ws.cell(row=barcode_row, column=cc).border = Border(**kw_btm)
        # Generate and embed Code128 barcode for the value, if non-empty
        if has_barcode and value is not None and str(value).strip():
            try:
                buf = BytesIO()
                Code128(str(value).strip(), writer=ImageWriter()).write(
                    buf, options={'module_height': 5.0, 'module_width': 0.20,
                                  'quiet_zone': 1.0, 'write_text': False})
                tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                tmp.write(buf.getvalue()); tmp.close()
                temp_files.append(tmp.name)
                # Approximate column-range width in pixels for centering
                col_chars = sum(widths.get(get_column_letter(c), 9) for c in range(col_start, col_end + 1))
                col_px = int(col_chars * 7.0)
                col_off = pixels_to_EMU(max(0, (col_px - INFO_BC_W_PX) // 2))
                row_off = pixels_to_EMU(2)
                img = XlImage(tmp.name)
                img_w_emu = pixels_to_EMU(INFO_BC_W_PX)
                img_h_emu = pixels_to_EMU(INFO_BC_H_PX)
                marker = AnchorMarker(col=col_start - 1, colOff=col_off,
                                       row=barcode_row - 1, rowOff=row_off)
                size = XDRPositiveSize2D(img_w_emu, img_h_emu)
                img.anchor = OneCellAnchor(_from=marker, ext=size)
                ws.add_image(img)
            except Exception:
                pass

    # Row indices: 3 stacked rows per logical row of the info table
    LBL_R1 = r;     VAL_R1 = r + 1; BC_R1 = r + 2
    LBL_R2 = r + 3; VAL_R2 = r + 4; BC_R2 = r + 5
    ws.row_dimensions[LBL_R1].height = 14
    ws.row_dimensions[BC_R1].height  = 26
    ws.row_dimensions[LBL_R2].height = 14
    ws.row_dimensions[BC_R2].height  = 26

    # Row 1 + Row 2 value-row heights adapt: if any cell in a row wraps to 2 lines,
    # grow the entire row so all 4 cells stay visually even-height. Cells that
    # only need 1 line get extra whitespace (vertically centered) — clean look.
    row1_values = [pl_data.get('customer',''), pl_data.get('packing_no',''),
                   pl_data.get('po',''), pl_data.get('pn','')]
    row2_values = [pl_data.get('wo',''), pl_data.get('eng',''),
                   pl_data.get('date',''), pl_data.get('datecode','')]
    ws.row_dimensions[VAL_R1].height = 38 if any(_value_strategy(v) == 'wrap' for v in row1_values) else 22
    ws.row_dimensions[VAL_R2].height = 38 if any(_value_strategy(v) == 'wrap' for v in row2_values) else 22

    write_info_pair(LBL_R1, VAL_R1, BC_R1, 1, 3,   "CUSTOMER",       pl_data.get('customer', ''))
    write_info_pair(LBL_R1, VAL_R1, BC_R1, 4, 6,   "PACKING LIST #", pl_data.get('packing_no', ''))
    write_info_pair(LBL_R1, VAL_R1, BC_R1, 7, 9,   "P.O. #",         pl_data.get('po', ''))
    write_info_pair(LBL_R1, VAL_R1, BC_R1, 10, 12, "P/N",            pl_data.get('pn', ''))
    write_info_pair(LBL_R2, VAL_R2, BC_R2, 1, 3,   "WO / LOT #",         pl_data.get('wo', ''))
    write_info_pair(LBL_R2, VAL_R2, BC_R2, 4, 6,   "PCC ENGINEERING #",  pl_data.get('eng', ''))
    write_info_pair(LBL_R2, VAL_R2, BC_R2, 7, 9,   "DATE",               pl_data.get('date', ''))
    write_info_pair(LBL_R2, VAL_R2, BC_R2, 10, 12, "DATE CODE",          pl_data.get('datecode', ''))

    r = BC_R2 + 1
    ws.row_dimensions[r].height = 12; r += 1

    # ── Helper: border a row of merged cells with NO internal verticals ─
    # Same pattern as the info table fix: top+bottom on all cells, left/right
    # only on edge cells. Result: each merged column appears as ONE clean cell.
    def _border_merged_row(row, start_col, end_col):
        for col in range(start_col, end_col + 1):
            kw = {'top': thin, 'bottom': thin}
            if col == start_col: kw['left']  = thin
            if col == end_col:   kw['right'] = thin
            ws.cell(row=row, column=col).border = Border(**kw)

    def _border_merged_box(start_row, end_row, start_col, end_col):
        """For multi-row blocks: outer outline only, no internal lines."""
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                kw = {}
                if row == start_row: kw['top']    = thin
                if row == end_row:   kw['bottom'] = thin
                if col == start_col: kw['left']   = thin
                if col == end_col:   kw['right']  = thin
                if kw:
                    ws.cell(row=row, column=col).border = Border(**kw)

    # ── Serial table header ───────────────────────────────────
    # # column (single cell — column A by itself)
    ws.cell(row=r, column=1, value="#").font = F_HDR
    ws.cell(row=r, column=1).alignment = AL_C
    _border_merged_row(r, 1, 1)

    # SERIAL column (merged B:E — appears as ONE cell)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    sc = ws.cell(row=r, column=2, value="Serial Number")
    sc.font = F_HDR; sc.alignment = AL_C
    _border_merged_row(r, 2, 5)

    # SERIAL BARCODE column (merged F:I — appears as ONE cell)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=9)
    bc = ws.cell(row=r, column=6, value="Barcode")
    bc.font = F_HDR; bc.alignment = AL_C
    _border_merged_row(r, 6, 9)

    # QTY column (merged J:L — appears as ONE cell)
    ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=12)
    qc = ws.cell(row=r, column=10, value="Qty")
    qc.font = F_HDR; qc.alignment = AL_C
    _border_merged_row(r, 10, 12)
    ws.row_dimensions[r].height = 22; r += 1

    # ── Serial rows with barcodes ────────────────────────────
    # Render the serial barcode so it looks identical to the COC's: same Code128
    # module options, and the barcode occupies the same fraction (BC_FILL) of its
    # column, horizontally centered. Technique: pad the barcode PNG with symmetric
    # white space so the bars are BC_FILL of the padded canvas, then display that
    # canvas across the full F:I column width via a OneCellAnchor (LibreOffice sizes
    # the image to the ext we give and renders the white margins, centering the bars).
    from openpyxl.utils import get_column_letter as _gcl
    from PIL import Image as PILImage
    col_fi_chars = sum(widths.get(_gcl(c), 9) for c in range(6, 10))  # F:I span
    col_fi_px = int(col_fi_chars * 7.0)
    BC_FILL = 0.86            # fraction of the column the bars span (matches COC)
    barcode_w_px = col_fi_px  # display the padded canvas across the whole column
    barcode_h_px = 30
    row_ht = 28
    bc_row_off = pixels_to_EMU(max(0, (int(row_ht * 1.33) - barcode_h_px) // 2))

    for i, sn in enumerate(serials, 1):
        # # cell (single)
        c1 = ws.cell(row=r, column=1, value=i)
        c1.font = F_BODY; c1.alignment = AL_C
        _border_merged_row(r, 1, 1)

        # SERIAL cell (merged B:E)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c2 = ws.cell(row=r, column=2, value=sn)
        c2.font = F_BODY_BOLD; c2.alignment = AL_C
        _border_merged_row(r, 2, 5)

        # BARCODE cell (merged F:I)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=9)
        c3 = ws.cell(row=r, column=6); c3.alignment = AL_C
        _border_merged_row(r, 6, 9)

        # QTY cell (merged J:L)
        ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=12)
        c4 = ws.cell(row=r, column=10, value=1)
        c4.font = F_BODY_BOLD; c4.alignment = AL_C
        _border_merged_row(r, 10, 12)

        if has_barcode:
            try:
                buf = BytesIO()
                Code128(sn, writer=ImageWriter()).write(
                    buf, options={'module_height': 8.0, 'module_width': 0.25,
                                  'quiet_zone': 1.0, 'write_text': False})
                buf.seek(0)
                bc_im = PILImage.open(buf).convert("RGB")
                nat_w, nat_h = bc_im.size
                padded_w = max(nat_w, int(round(nat_w / float(BC_FILL))))
                canvas = PILImage.new("RGB", (padded_w, nat_h), "white")
                canvas.paste(bc_im, ((padded_w - nat_w) // 2, 0))
                tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                canvas.save(tmp.name); tmp.close()
                temp_files.append(tmp.name)
                img = XlImage(tmp.name)
                # TwoCellAnchor spanning F:I forces LibreOffice to fit the (padded)
                # image to the full column; the symmetric white padding then yields
                # equal left/right margins -> the bars are centered at BC_FILL width.
                from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor as _TCA
                _from = AnchorMarker(col=5, colOff=pixels_to_EMU(0), row=r - 1, rowOff=bc_row_off)
                _to   = AnchorMarker(col=9, colOff=pixels_to_EMU(0), row=r,     rowOff=-bc_row_off)
                img.anchor = _TCA(editAs='oneCell', _from=_from, to=_to)
                ws.add_image(img)
            except Exception:
                pass

        ws.row_dimensions[r].height = row_ht
        r += 1

    ws.row_dimensions[r].height = 8; r += 1

    # ── Total Units row ──────────────────────────────────────
    # Empty box on the left (cols A:F across 2 rows) — one clean cell
    ws.merge_cells(start_row=r, start_column=1, end_row=r+1, end_column=6)
    _border_merged_box(r, r+1, 1, 6)

    # TOTAL UNITS box on the right (cols G:L across 2 rows) — appears as ONE cell
    # Label "TOTAL UNITS" on top half, big number on bottom half, NO inner horizontal divider
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=12)
    tlbl = ws.cell(row=r, column=7, value="TOTAL UNITS")
    tlbl.font = F_TOTAL_LBL; tlbl.alignment = AL_RI

    ws.merge_cells(start_row=r+1, start_column=7, end_row=r+1, end_column=12)
    tval = ws.cell(row=r+1, column=7, value=len(serials))
    tval.font = F_TOTAL_VAL; tval.alignment = AL_RI

    # Single outline around the entire 2-row × 6-col TOTAL UNITS area
    _border_merged_box(r, r+1, 7, 12)

    ws.row_dimensions[r].height = 18
    ws.row_dimensions[r+1].height = 26
    r += 2

    last_row = r - 1

    # ── Print area + force fit-to-page-width ─────────────────
    ws.print_area = f"A1:{get_column_letter(12)}{last_row}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1          # fit all columns onto one page wide
    ws.page_setup.fitToHeight = 0         # flow down as needed — no forced shrink
    ws.print_options.horizontalCentered = True   # center the whole sheet on the page
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4

    wb.save(str(out))
    for tf in temp_files:
        try: Path(tf).unlink()
        except Exception: pass

    return str(out)


###############################################################################
# SHIPMENT HISTORY
###############################################################################
# Every draft-save / close / delete is recorded as one small JSON file under
# the central History dir (never inside WO folders). Best-effort only: a
# history failure must NEVER block or crash the shipment action itself.

def _hist_safe(s):
    """Replace characters illegal in Windows filenames with '_' and strip."""
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(s)).strip()

def write_history_event(event_type, record):
    """
    Write one shipment event as a JSON file under
    HISTORY_DIR\\YYYY\\YYYY-MM\\<ts>_<wo>_<station>_<event_type>_<seq>.json.
    Adds standard envelope fields (schema_version, event_id, ts, station,
    win_user, app_version, ...) then merges in the caller's record dict.
    Never raises: returns the str path written, or None on ANY failure.
    """
    try:
        ts = datetime.datetime.now()
        folder = Path(HISTORY_DIR) / f"{ts:%Y}" / f"{ts:%Y-%m}"
        folder.mkdir(parents=True, exist_ok=True)
        wo_tag = _hist_safe((record or {}).get("wo") or "unknown")[:40] or "unknown"
        station = _hist_safe(os.environ.get("COMPUTERNAME", "unknown")) or "unknown"
        base = f"{ts:%Y%m%dT%H%M%S}_{ts.microsecond//1000:03d}_{wo_tag}_{station}_{event_type}"
        rec = {
            "schema_version": 1,
            "event_id": None,  # set per write attempt below
            "event_type": event_type,
            "ts": ts.strftime("%Y-%m-%d %H:%M:%S"),
            "ts_epoch": time.time(),
            "station": station,
            "win_user": os.environ.get("USERNAME", ""),
            "app_version": APP_VERSION,
            "source": "live",
        }
        rec.update(record or {})
        for seq in range(10):
            final = folder / f"{base}_{seq}.json"
            if final.exists():
                continue
            tmp = folder / (final.name + ".tmp")
            rec["event_id"] = final.stem
            try:
                with open(tmp, "x", encoding="utf-8") as f:
                    json.dump(rec, f, indent=2, default=str)
            except FileExistsError:
                continue
            except Exception:
                try: tmp.unlink()
                except Exception: pass
                return None
            if final.exists():
                try: tmp.unlink()
                except Exception: pass
                continue
            try:
                os.replace(str(tmp), str(final))
            except Exception:
                try: tmp.unlink()
                except Exception: pass
                return None
            return str(final)
        return None
    except Exception:
        return None


###############################################################################
# SERIAL DUPLICATE HISTORY + OUTLIER DETECTION
###############################################################################
# Two safety checks for the serial-scanning step (added 2026-08-19):
#   1. build_serial_history() indexes every serial from CLOSED shipments of the
#      same job (Eng# + P/N) so a serial scanned in a PAST work order is caught
#      the moment it is re-scanned (wrong board / reused-or-mislabeled barcode /
#      a returned unit shipped twice). Scope is Eng#+P/N because the same Eng#
#      can carry two P/Ns (e.g. bare vs. conformal-coated) that EACH run their own
#      1..N serials — the same serial number under a different P/N is NOT a dup.
#      For the customer whose serial is "<seq>-<P/N>-<Rev>" the P/N is baked into
#      the string, so exact-string matching separates the two P/Ns on its own.
#      Matching is EXACT (trim+casefold): a different Rev is a DIFFERENT serial and
#      is NOT treated as a duplicate (confirmed by Dhruvil 2026-08-19).
#   2. detect_serial_outliers() flags a serial that looks structurally unlike the
#      rest of the batch (a garbage scan, a wrong P/N baked into the serial, a
#      wrong WO prefix, or a sequence number far outside the batch) so the
#      operator can fix it before the paperwork is generated.
# Both are pure/stdlib and do ZERO Tk work (safe to call from a worker thread).

def _norm_serial(s):
    """Canonical form for exact duplicate matching: trimmed + case-folded."""
    return str(s if s is not None else "").strip().casefold()

def _serial_identity(s, pn=""):
    """Board identity for ALL serial tracking (dup check, receiving tracker,
    imports): **Eng# scope + P/N + serial number — REVISION IGNORED** (Dhruvil
    2026-09-02: jobs 4092/4041 changed revision mid-job, so the same physical
    board appears as ...-A01 in old shipments and ...-B01 later; exact matching
    split one board into two records and nothing lined up. This supersedes the
    2026-08-19 exact-match decision.)
    Rule: for '<seq>-<P/N>-<Rev>' serials — the serial has 3+ dash-segments AND
    the job P/N is still present after dropping the LAST segment — drop that last
    segment (the Rev) and key on the rest. Anything else (e.g. '<WO>-<seq>',
    which never embeds the P/N) keys on the exact string, so two different boards
    are never collapsed; the worst case is a missed rev-strip, never a false merge."""
    raw = str(s if s is not None else "").strip()
    pk = _alnum(pn) if pn else ""
    if pk and len(pk) >= 4:
        parts = raw.split("-")
        if len(parts) >= 3:
            candidate = "-".join(parts[:-1])
            if pk in _alnum(candidate):
                return candidate.casefold()
    return raw.casefold()

def _norm_key(s):
    """Canonical Eng#/P/N for scoping the index: trimmed + upper; '' for a
    blank/N-A value so a missing value never accidentally matches another."""
    v = str(s if s is not None else "").strip().upper()
    return "" if v in ("", "N/A", "NA", "—", "-", "NONE", "?") else v

def build_serial_history(eng, pn, progress=None):
    """Scan HISTORY_DIR for CLOSED shipments of this job (Eng# + P/N) and return a
    dict mapping normalized_serial -> {"serial","wo","ts","operator","event_id"}
    (the most recent prior occurrence). Only 'close' events count as "shipped";
    drafts/deletes/overrides are ignored. Scoping REQUIRES both Eng# and P/N to be
    known (returns {} otherwise) so the two P/Ns of one Eng# never merge and raise
    false duplicates. Never raises — returns {} on any failure so scanning is never
    blocked by a history problem."""
    ekey, pkey = _norm_key(eng), _norm_key(pn)
    if not ekey or not pkey:
        return {}
    index = {}
    skip_fail = set()        # newest event = FAILED receive -> re-ship expected, no alert
    seen_pass_recv = set()   # a NEWER pass-receive blocks an older fail from suppressing
    voided = set()           # event_ids voided by a supervisor -> their serials don't count
    pk_al = _alnum(pn)                 # this job's P/N in alnum form
    sib_al = _alnum(_pn_sibling(pn))   # the paired …900/…970 P/N — its serials never belong here
    try:
        root = Path(HISTORY_DIR)
        if not root.exists():
            return {}
        # A job's serials span the WHOLE year, so walk every YYYY/YYYY-MM folder,
        # newest-first, so the occurrence we keep per serial is the latest one.
        months = []
        for y in root.iterdir():
            if y.is_dir() and re.fullmatch(r"\d{4}", y.name):
                for m in y.iterdir():
                    if m.is_dir() and re.fullmatch(r"\d{4}-\d{2}", m.name):
                        months.append(m)
        for m in sorted(months, key=lambda p: p.name, reverse=True):
            try:
                files = sorted((f for f in m.iterdir() if f.name.lower().endswith(".json")),
                               key=lambda f: f.name, reverse=True)
            except Exception:
                continue
            for f in files:
                if callable(progress):
                    try: progress()
                    except Exception: pass
                try:
                    with open(f, encoding="utf-8") as fh:
                        rec = json.load(fh)
                except Exception:
                    continue
                # 'close' = real shipments; 'backfill_shipped' = pre-history testing
                # shipments imported via Receiving — those boards SHIPPED too, so
                # they must also trip the "already shipped before" alert. 'receive'
                # events are read ONLY to learn pass/fail: boards CYCLE for the
                # test-and-return customer, and a serial whose NEWEST event is a
                # FAILED receive is EXPECTED to ship again for re-test — no alert
                # (a PASSED board re-scanned into a testing shipment still alerts).
                if not isinstance(rec, dict):
                    continue
                et = rec.get("event_type")
                # A 'void' marker (written by a supervisor from the History viewer)
                # cancels its target record. Voids are always WRITTEN AFTER their
                # target, so the newest-first walk sees them before the target.
                if et == "void":
                    t = str(rec.get("target_event_id") or "")
                    if t: voided.add(t)
                    continue
                if str(rec.get("event_id") or "") in voided:
                    continue
                if et not in ("close", "backfill_shipped", "receive"):
                    continue
                if _norm_key(rec.get("eng")) != ekey:
                    continue
                rpn = _norm_key(rec.get("pn"))
                if et == "receive":
                    if rpn != pkey:
                        continue
                    failed = {_serial_identity(x, pn) for x in (rec.get("failed") or [])}
                    for s in (rec.get("serials") or []):
                        if _embeds(sib_al, s):     # partner P/N's board, not ours
                            continue
                        k = _serial_identity(s, pn)
                        if not k or k in index or k in skip_fail:
                            continue
                        if k in failed:
                            if k not in seen_pass_recv:   # only if no NEWER pass-receive
                                skip_fail.add(k)
                        else:
                            seen_pass_recv.add(k)
                    continue
                occ = {"wo": str(rec.get("wo") or ""),
                       "ts": str((rec.get("ship_date") if et == "backfill_shipped" else None)
                                 or rec.get("ts") or ""),
                       "operator": str(rec.get("operator") or ""),
                       "event_id": str(rec.get("event_id") or "")}
                for s in (rec.get("serials") or []):
                    # Route by the P/N EMBEDDED IN THE SERIAL — event labels can be
                    # wrong (coated boards imported under the testing P/N, or old
                    # closes whose pn was the whole folder name). A serial carrying
                    # the PARTNER P/N (…900<->…970) never belongs to this index; a
                    # serial carrying THIS P/N belongs here even off a mislabeled
                    # event. No embedded P/N -> trust the event label.
                    if _embeds(sib_al, s):
                        continue
                    if not (rpn == pkey or _embeds(pk_al, s)):
                        continue
                    # REV-AGNOSTIC identity (Eng scope + P/N + seq) — a mid-job
                    # revision change must not make the same board look new.
                    k = _serial_identity(s, pn)
                    if not k or k in index or k in skip_fail:
                        continue                  # newest-first walk: first hit wins
                    occ2 = dict(occ); occ2["serial"] = str(s)
                    index[k] = occ2
    except Exception:
        return {}
    return index

def _seg_class(seg):
    """Coarse character class of one serial segment, for shape comparison."""
    if seg == "":       return "E"
    if seg.isdigit():   return "D"
    if seg.isalpha():   return "A"
    return "X"

def _serial_segments(s):
    return str(s).split("-")

def _serial_signature(s):
    """Structural fingerprint: one class letter per dash-segment (D=digits, A=letters,
    X=mixed, E=empty); the segment COUNT is implicit in the tuple length. Segment
    LENGTH is deliberately NOT part of the signature, so a 1-digit vs 2-digit vs
    3-digit sequence number (e.g. '16452.1-1' vs '16452.1-42') is the SAME shape —
    sequence numbers naturally grow in digit-count and must not read as odd. Value-level
    anomalies (wrong prefix, out-of-range #) are caught by the later per-column checks."""
    return tuple(_seg_class(x) for x in _serial_segments(s))

def _alnum(s):
    """Alphanumerics only, upper-cased — tolerant compare so '192-146-900' and
    '192146900' match regardless of dashes/spaces."""
    return re.sub(r'[^0-9A-Za-z]', '', str(s if s is not None else "")).upper()

def _rev_after_pn(serial, pn):
    """For a '<seq>-<P/N>-<Rev>' serial, return the Rev part (segments AFTER the job
    P/N), or '' if the P/N isn't found as consecutive dash-segments."""
    segs = str(serial).split("-")
    pnsegs = [p for p in str(pn).split("-") if p != ""]
    if not pnsegs:
        return ""
    tgt = [_alnum(p) for p in pnsegs]; n = len(tgt)
    for i in range(len(segs) - n + 1):
        if [_alnum(x) for x in segs[i:i + n]] == tgt:
            return "-".join(segs[i + n:])
    return ""

def detect_serial_outliers(serials, pn="", min_batch=2, dom_frac=0.7):
    """Flag serials inconsistent with the rest of the batch so a bad scan is caught
    before the paperwork is generated. Returns [[serial, reason], ...] in first-seen
    order (one reason each); [] for a uniform batch or one smaller than min_batch.
    Advisory only — the caller lets the operator Keep or Remove each.

    If the job P/N is known AND embedded in the serials ('<seq>-<P/N>-<Rev>' style),
    it ANCHORS on that P/N — works even for a 2-scan batch:
      * any serial that does NOT carry the job P/N is flagged (wrong board / bad scan);
      * serials whose trailing Rev differs from the batch's Rev are flagged.
    Otherwise it falls back to a batch-majority shape / constant-segment check (for the
    '<WO>-<seq>' style, which needs a clear >= dom_frac majority).

    NOTE: there is deliberately NO numeric range/gap check. Sequence numbers are often
    non-consecutive (you might ship #1-5 and #45 together), and a plain numeric gap can't
    tell a legit high serial from a garbage number — so a bad scan is caught by the shape,
    prefix, P/N and Rev checks instead, never by "this number is big"."""
    from collections import Counter
    uniq, seen = [], set()
    for s in serials:
        k = _norm_serial(s)
        if k and k not in seen:
            seen.add(k); uniq.append(str(s))
    if len(uniq) < min_batch:
        return []
    flagged = {}
    order = {s: i for i, s in enumerate(uniq)}
    def flag(s, reason):
        if s not in flagged:
            flagged[s] = reason
    def result():
        return sorted(([s, r] for s, r in flagged.items()), key=lambda t: order.get(t[0], 1 << 30))

    # ── P/N-anchored mode: the job P/N is real and embedded in the serials ──
    pnd = _alnum(pn)
    if len(pnd) >= 4 and any(pnd in _alnum(s) for s in uniq):
        with_pn = []
        for s in uniq:
            if pnd in _alnum(s):
                with_pn.append(s)
            else:
                flag(s, f"does not carry this job's P/N ({pn}) — wrong board or bad scan")
        # Rev (the part after the P/N) should be the same across one shipment.
        revs = [(s, _rev_after_pn(s, pn)) for s in with_pn]
        rc = Counter(r for _, r in revs)
        if len(rc) > 1:
            top_rev = rc.most_common(1)[0][0]
            for s, r in revs:
                if r != top_rev:
                    flag(s, f"revision “{r or '(blank)'}” differs from the rest (“{top_rev or '(blank)'}”)")
        return result()

    # ── Generic batch-majority mode (e.g. '<WO>-<seq>') ──
    sigs = {s: _serial_signature(s) for s in uniq}
    dom_sig, dom_n = Counter(sigs.values()).most_common(1)[0]
    if dom_n / len(uniq) < dom_frac:
        return []   # no clear norm -> genuinely mixed; don't guess
    example = next(s for s in uniq if sigs[s] == dom_sig)
    for s in uniq:
        if sigs[s] != dom_sig:
            flag(s, f"format looks different from the rest (most read like “{example}”)")
    normal = [s for s in uniq if sigs[s] == dom_sig]
    cols = [list(c) for c in zip(*[_serial_segments(s) for s in normal])] if normal else []
    const_vals = {}
    for i, col in enumerate(cols):
        vc = Counter(col); val, cnt = vc.most_common(1)[0]
        if len(vc) > 1 and cnt / len(col) >= dom_frac:
            const_vals[i] = val
    for s in normal:
        segs = _serial_segments(s)
        for i, val in const_vals.items():
            if segs[i] != val:
                flag(s, f"“{segs[i]}” doesn’t match the batch’s “{val}”")
                break
    return result()


###############################################################################
# RECEIVING — boards returned from customer testing (added 2026-08-21)
###############################################################################
# Flow for the test-and-return customer: boards ship for testing under the -900
# P/N (serial "<seq>-<P/N>-<Rev>"), come BACK wearing that same testing label,
# get conformal-coated, and ship again under the -970 P/N with the same seq.
# The Receiving window scans returned boards, records a "receive" History event
# (the source of truth), and maintains ONE per-P/N report workbook INSIDE the
# job's NEWEST testing WO folder (next to Shipping Serials — per Dhruvil
# 2026-09-02): <pd>\<newest WO>\Received after testing\<PN> - Received after
# testing.xlsx. Since 2026-09-02 the file is WRITABLE and MERGED, not
# regenerated (Dhruvil: "If we edit and save something in that file then keep
# it"): the two log sheets keep every existing row (manual edits included), the
# app only APPENDS new boards, and overlapping-but-different data raises an
# alert; the two status sheets are rebuilt from history on every save. When the
# newest WO folder changes, the existing file is migrated (content and edits
# carried over) and old copies removed.
# "backfill_shipped" events seed the shipped-for-testing baseline for boards
# shipped before History existed, imported from the WO folders' serials files.

def testing_to_coated_pn(pn):
    """'136-9624-900' -> '136-9624-970' (confirmed customer convention: same base,
    -900 = testing, -970 = coated). '' when the P/N doesn't end in 900."""
    p = str(pn or "").strip()
    return p[:-3] + "970" if p.endswith("900") else ""

def _pn_sibling(pn):
    """The partner P/N of a test-and-return pair: …900 <-> …970. '' otherwise."""
    p = str(pn or "").strip()
    if p.endswith("900"): return p[:-3] + "970"
    if p.endswith("970"): return p[:-3] + "900"
    return ""

def _embeds(al_key, serial):
    """True if the alnum-form P/N key appears inside the serial (min 6 chars so a
    short key can't grab everything)."""
    return bool(al_key) and len(al_key) >= 6 and al_key in _alnum(serial)

def _tr_pn_parts(pn):
    """Anatomy of a test-and-return P/N LABEL for the REPORT code only (Dhruvil
    2026-09-23): '125-1030-970C' -> ('1251030', '970', 'C'), '136-9624-900' ->
    ('1369624', '900', ''), anything else (blank, 'N/A', a whole folder name) ->
    ('', '', ''). Real coated labels can carry a rev LETTER after the 970 (Anduril
    4006/4007/4008/4023: '125-1030-970C'); the serials embed only the DIGITS, so
    the coated embed key stays core+'970'. _pn_sibling / testing_to_coated_pn are
    deliberately left letter-blind — they feed the main-window duplicate check."""
    m = re.search(r'(\d{6,7})(900|970)([A-Z]{0,2})$', _alnum(pn))
    return (m.group(1), m.group(2), m.group(3)) if m else ("", "", "")

def _tr_testing_label(pn):
    """Canonical TESTING (-900) spelling of a pair member, keeping the label's own
    punctuation: '125-1030-970C' -> '125-1030-900', '1342201970' -> '1342201900',
    a -900 label -> itself (stripped); '' when it isn't a …900/…970 P/N at all.
    This is the report's ONE job key for in-house-tested (coated-only) jobs."""
    p = str(pn or "").strip()
    core, side, _ = _tr_pn_parts(p)
    if not core:
        return ""
    if side == "900":
        return p
    # _tr_pn_parts has just validated that everything after the FINAL 970 is the
    # rev-letter tail (letters + whatever punctuation the operator typed around
    # them: '125-1030-970 C', '125-1030-970-C', '125-1030-970C.', '125-1030-970 - C'
    # — review rounds 1 + 4), so replace that 970 and keep the tail-less label.
    # NEVER hand back the coated spelling itself when the substitution found
    # nothing (a coated label as the job key would walk the job as a TESTING job).
    out = re.sub(r'970[^0-9]*$', '900', p)
    return out if out != p else ""

def _coat_ident_pn(pn):
    """The P/N coated / Case-B identities are keyed by (_serial_identity): the
    digit-form coated P/N, else the job P/N itself — never the lettered real label.
    One helper so the walk, _has_testing, receiving_summary, scan_job_folders and
    write_received_xlsx cannot drift apart on a lettered label (review round 4:
    coat_label_mismatch is written in the walk and read in three other places).
    A -900 label typed WITH a rev letter ('134-2201-900A') still maps to the
    digit-form coated P/N — testing_to_coated_pn is letter-blind, and the
    lettered label itself as the identity key made every coated identity raw
    (rev-sensitive) under that key (review round 5)."""
    p = str(pn or "")
    c = testing_to_coated_pn(p)
    if not c:
        core, side, _ = _tr_pn_parts(p)
        if core and side == "900":
            c = core + "970"
    return c or p

def _tr_embed_keys(pn):
    """The two serial EMBED keys of a job key, alnum and DIGITS ONLY:
    (pk_al, ck_al) = the job's own P/N and its coated side as they sit inside a
    serial — '125-1030-900' -> ('1251030900', '1251030970'); a -970 key -> (its
    digits, ''). A -900 key typed WITH a rev letter ('134-2201-900A',
    '125-1030-900C') embeds as its digits too and its coated side comes from the
    core — testing_to_coated_pn is letter-blind and returned '' for it. ONE
    helper for build_receiving_state, _resolve_received_wo_dir,
    _stale_received_candidates, scan_job_folders, _fold_receive_locally,
    _coated_tree_holds_report (review round 7) and the
    Receiving window's P/N check (review round 6: the walk went digit-only in
    round 5 while those five sites kept _alnum(pn) / _alnum(testing_to_coated_pn(
    pn)), so under an only-lettered -900 key they saw a different job — the -970
    tree was never listed as a candidate or scanned, its live WO never resolved,
    the job's own rev-D board got 'Doesn't match this job', a coated-form
    override landed in `received`). A label with no parsable core keeps
    _alnum(pn), as before. Identical strings for every plain label."""
    p = str(pn or "")
    core, side, _ = _tr_pn_parts(p)
    pk_al = (core + side) if core else _alnum(p)
    ck_al = _alnum(testing_to_coated_pn(p)) or ((core + "970") if core and side == "900" else "")
    return pk_al, ck_al

# Shared wording (review round 1: one source of truth — the Rebuild heads-up
# FILTERS coated_log rows by the Case-B tag, so a reword must reach every writer).
_CASE_B_NOTE = "scanned with a -900 testing label"        # Q3 Note on a Case-B coated row
_COATED_ONLY_TAG = "  (in-house tested — coated only)"    # dropdown / picker suffix

def serial_seq(s):
    """Canonical board sequence # = first dash-segment as an int-string ('000123'
    -> '123') so the same board matches across testing/coated serials; '' if the
    first segment isn't numeric."""
    seg = str(s or "").split("-", 1)[0].strip()
    return str(int(seg)) if seg.isdigit() else ""

def _parse_user_date(s):
    """Parse an operator-typed date (the Receiving 'received on' field for
    importing OLD manual records). Returns 'YYYY-MM-DD' or None if unparseable."""
    s = str(s or "").strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%m-%d-%y",
                "%d %b %Y", "%b %d %Y", "%b %d, %Y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None

def _currently_received(state, k):
    """True if the board is (or was last known) in-house. Boards cycle for this
    customer (ship for testing → receive → FAIL → ship again for re-test →
    receive again …).

    COUNT-BASED first (Dhruvil 2026-09-03, the 4094 case): shipped MORE times
    than received → the board is out at the customer, no matter in what order
    the scans were typed in (his 5 failed boards were re-shipped Aug-31 but the
    failed receive was scanned Sep-03 — timestamps alone called them in-house).
    Received more times than shipped → in-house. Only when the counts are EQUAL
    do timestamps break the tie (a board received in May and shipped again in
    June is back at the customer)."""
    r = (state.get("received") or {}).get(k)
    if not r:
        return False
    n = (state.get("ship_counts") or {}).get(k, 0)
    m = len((state.get("recv_dates") or {}).get(k) or ())
    if n and m and n != m:
        return m > n
    s = (state.get("shipped") or {}).get(k)
    if not s:
        return True
    return (r.get("sts") or "") >= (s.get("sts") or "")

def _division_of_path(p):
    """Which division root a path lives under ('' if neither)."""
    try:
        rp = str(p or "").lower()
        for d, root in PRODUCTION_ROOTS.items():
            if rp.startswith(root.lower()):
                return d
    except Exception:
        pass
    return ""

def _iter_history_files(progress=None):
    """Yield every parseable event dict under HISTORY_DIR, newest file first
    (newest month folder, then newest filename). A single corrupt FILE is skipped,
    but a failed FOLDER enumeration (share blip) PROPAGATES — callers must know the
    walk was incomplete rather than silently acting on partial data (a truncated
    walk once regenerated a near-empty report over a complete one)."""
    root = Path(HISTORY_DIR)
    if not root.exists():
        return
    months = []
    for y in root.iterdir():
        if y.is_dir() and re.fullmatch(r"\d{4}", y.name):
            for m in y.iterdir():
                if m.is_dir() and re.fullmatch(r"\d{4}-\d{2}", m.name):
                    months.append(m)
    for m in sorted(months, key=lambda p: p.name, reverse=True):
        files = sorted((f for f in m.iterdir() if f.name.lower().endswith(".json")),
                       key=lambda f: f.name, reverse=True)
        for f in files:
            if callable(progress):
                try: progress()
                except Exception: pass
            try:
                with open(f, encoding="utf-8") as fh:
                    rec = json.load(fh)
            except Exception:
                continue
            if isinstance(rec, dict):
                yield rec

def list_testing_jobs(progress=None):
    """Jobs the Receiving window can work on, newest activity first: every
    (Eng#, P/N) seen in a close/backfill event whose P/N ends in 900 (the testing
    convention), plus any job that already has receive events. Never raises.
    In-house-tested jobs (Dhruvil 2026-09-23): a job that only ever shipped COATED
    (…970 / …970+letter) boards is APPENDED after the real entries, registered
    under its canonical -900 key ('125-1030-970C' -> '125-1030-900') and flagged
    "testing": False with the real label in "coated_label" — one job, one key, one
    report. Real entries keep their exact shape (no "testing" key; readers use
    j.get("testing", True)); a -900 label recorded in two spellings ('134-2201-900'
    / '1342201900') is ONE real entry under its newest spelling (review round 3).
    A job whose junk-labelled (folder-name / 'N/A', pre-v58)
    close carries -900-FORM serials is a testing job by the same inference
    build_receiving_state draws when it fills `shipped` — it is listed as a real
    entry under the canonical key, so the dropdown tag and the window's
    'Nothing to receive here' stop can never disagree (review round 1).
    A RECEIVE event alone makes a real entry only when it proves a testing
    shipment — at least one received serial that is neither coated-form (…970…)
    nor a Case-B board of that job (a -900-form serial shipped under the coated
    label): a supervisor-overridden receive of such a board on an in-house-tested
    job keeps the coated-only entry and its tag, exactly as build_receiving_state
    keeps has_testing False (Dhruvil 2026-09-23, decision (a))."""
    jobs = {}
    coat_jobs = {}
    form_test = {}         # Eng# -> the -900-form serials (alnum) seen on a junk label
    caseb = {}             # (Eng#, core) -> coated identities of Case-B serials (coated-labelled events)
    recv_serials = {}      # real-entry key -> every serial of every receive event for it
    ship_seen = set()      # real-entry keys that have a -900-labelled close/backfill
    voided = set()
    try:
        for rec in _iter_history_files(progress):
            et = rec.get("event_type")
            if et == "void":
                t = str(rec.get("target_event_id") or "")
                if t: voided.add(t)
                continue
            if str(rec.get("event_id") or "") in voided:
                continue
            if et not in ("close", "receive", "backfill_shipped"):
                continue
            ekey, pkey = _norm_key(rec.get("eng")), _norm_key(rec.get("pn"))
            if not ekey or not pkey:
                continue
            # TESTING (…900) labels are the receivable jobs. A …970 / …970+letter
            # label is an in-house-tested (coated-only) job: registered below under
            # its canonical -900 key, LISTED tagged and unreceivable (Dhruvil
            # 2026-09-23, Q2); any other label is junk (folder name / 'N/A', pre-v58).
            # The SIDE is read the way build_receiving_state reads its own label
            # (_tr_pn_parts: a rev letter after the 900 is tolerated too, so a
            # '125-1030-900C' close is a testing shipment on both sides — review
            # round 4); a label with no parsable core keeps the raw '…900' test.
            core, side, _ = _tr_pn_parts(pkey)
            test_lbl = (side == "900") or (not core and pkey.endswith("900"))
            if et in ("close", "backfill_shipped") and not test_lbl:
                # In-house-tested job (Dhruvil 2026-09-23): a job that ships ONLY
                # coated boards still gets a Shipping Report. Register it under its
                # canonical TESTING (-900) key — keyed by the core DIGITS so
                # '125-1030-970C', '125-1030-970' and '1251030970' are ONE job —
                # flagged so the Receiving dropdown can tag it and stop scans.
                if side == "970" and core:
                    label = str(rec.get("pn")).strip()
                    canon = _tr_testing_label(label)
                    ck = (ekey, core)
                    if canon and ck not in coat_jobs:   # newest-first: first wins
                        coat_jobs[ck] = {"eng": str(rec.get("eng")).strip(), "pn": canon,
                                         "ts": str(rec.get("ts") or ""), "testing": False,
                                         "coated_label": label, "_from_close": et == "close"}
                    elif canon and et == "close" and not coat_jobs[ck].get("_from_close"):
                        # the label SHOWN follows the newest CLOSE, as the walk's
                        # coated_pn_label does — an import typed '125-1031-970 D'
                        # must not relabel the dropdown/picker while the chips and
                        # the file say '125-1031-970D' (review round 3)
                        coat_jobs[ck]["coated_label"] = label; coat_jobs[ck]["_from_close"] = True
                    # Case-B boards of this job (the walk's rule: a -900-form serial
                    # in a coated-LABELLED close/backfill), by the same coated-digit
                    # identity build_receiving_state puts in coat_label_mismatch —
                    # a receive of one of them is not testing activity (decision (a))
                    for s in (rec.get("serials") or []):
                        if _embeds(core + "900", s) and not _embeds(core + "970", s):
                            caseb.setdefault(ck, set()).add(_serial_identity(s, core + "970"))
                elif not core:
                    # junk label: its -900-form serials ARE testing shipments (they
                    # land in `shipped`); a coated-side CLOSE label is Case B and
                    # deliberately not looked at here (Q3). Kept as the SERIALS, not
                    # a core guessed from them: the sequence digits sit right before
                    # the embedded P/N, so a regex over the serial swallowed one of
                    # them for a 6-digit core ('000021-192146900' -> '1192146') and
                    # the job never attached (review round 2). The embed test below
                    # is the one build_receiving_state uses to fill `shipped`.
                    for s in (rec.get("serials") or []):
                        if "900" in _alnum(s):
                            form_test.setdefault(ekey, []).append(_alnum(s))
                continue
            # one job, one entry: a -900 label recorded in two spellings
            # ('134-2201-900' and '1342201900' — punctuation is part of _norm_key)
            # is keyed by its core DIGITS, so the newest spelling is the one entry
            # (the walk accepts every spelling of the label; the report follows
            # the entry's name and folds the other spelling's file in — review
            # round 3: two rows in the picker made two reports, one going stale).
            # A label without a parsable core keeps its exact spelling as the key.
            # A spelling WITH a rev letter ('134-2201-900A', typed by hand the way
            # coated labels are written) is the entry's key only when the job has
            # no plain spelling at all: the walk derives its coated key from the
            # entry's pn, and a lettered key made every -970-form serial of the
            # 4065 import shape a testing shipment again (review round 5).
            rk = (ekey, "".join(_tr_pn_parts(pkey)[:2]) or pkey)   # core digits + side
            if et == "receive":
                recv_serials.setdefault(rk, []).extend(
                    str(s) for s in (rec.get("serials") or []) if str(s).strip())
            else:
                ship_seen.add(rk)
            spelled = str(rec.get("pn")).strip()
            if rk not in jobs:                # walk is newest-first: first wins
                jobs[rk] = {"eng": str(rec.get("eng")).strip(), "pn": spelled,
                            "ts": str(rec.get("ts") or ""), "_from_close": et == "close"}
            else:
                # The entry's spelling (= the report's file name) follows the newest
                # CLOSE — the WO folder's own spelling — as the coated label does
                # (review round 3): an ⏪ import's hand-typed spelling or a legacy
                # receive saved under the old two-row picker's other entry must not
                # rename an existing report, and ping-pong it back on the next close
                # (review round 7). A lettered spelling still yields to a plain one
                # (review round 5), and a close never trades a plain for a lettered.
                j = jobs[rk]
                cur_l = bool(_tr_pn_parts(j["pn"])[2]); new_l = bool(_tr_pn_parts(pkey)[2])
                if cur_l and not new_l:
                    j["pn"] = spelled                  # newest PLAIN spelling over a lettered one
                    if et == "close": j["_from_close"] = True
                elif et == "close" and not j.get("_from_close") and not (new_l and not cur_l):
                    j["pn"] = spelled; j["_from_close"] = True   # a CLOSE's spelling over an import's / receive's
    except Exception:
        pass
    # A real entry made by RECEIVE events alone (no -900-labelled close/backfill)
    # stands only when one of its receives proves a testing shipment — the walk's
    # rule for testing_label_seen (decision (a)): at least one serial that is not
    # coated-form and not a Case-B board of the job. Otherwise the entry falls
    # away and the job's coated-only entry (below) keeps its tag. An entry with a
    # -900-labelled shipment on record is untouched (old flow, byte-identical).
    for rk, sers in recv_serials.items():
        if rk in ship_seen or rk not in jobs:
            continue
        core = _tr_pn_parts(rk[1])[0]
        ck_al = (core + "970") if core else ""          # (alnum coated P/N — the walk's ck_al)
        cb = caseb.get((rk[0], core), set())
        if not any(not _embeds(ck_al, s) and _serial_identity(s, ck_al) not in cb for s in sers):
            del jobs[rk]
    # a coated-only entry is dropped whenever a REAL entry with the same Eng# and
    # core digits exists (the mixed job keeps its real -900 spelling, order and
    # shape); the survivors are appended so old-flow dropdown order is untouched
    real = {(_norm_key(j["eng"]), _tr_pn_parts(j["pn"])[0]) for j in jobs.values()}
    out = list(jobs.values())
    for j in out:
        j.pop("_from_close", None)            # private tie-break slot, never a reader's key
    for k, c in coat_jobs.items():
        c.pop("_from_close", None)
        if k in real:
            continue
        if any(_embeds(k[1] + "900", s) for s in form_test.get(k[0], ())):
            # testing shipments on record after all (junk-label -900-form serials):
            # a REAL entry under the canonical key — same shape as every other one
            out.append({"eng": c["eng"], "pn": c["pn"], "ts": c["ts"]})
        else:
            out.append(c)
    return out

def detect_job_for_serial(serial, jobs):
    """Pick the job whose P/N is embedded in a scanned serial (longest match wins;
    at least 6 alphanumerics so a short/generic P/N can't grab everything)."""
    sa = _alnum(serial)
    best, bl = None, 0
    for j in (jobs or []):
        pa = _alnum(j.get("pn"))
        if len(pa) >= 6 and pa in sa and len(pa) > bl:
            best, bl = j, len(pa)
    return best

def build_receiving_state(eng, pn, progress=None):
    """ONE walk over History -> everything Receiving needs for one job:
      shipped  {norm_serial: {serial, wo, ts, source 'close'|'backfill'}}
      received {norm_serial: {serial, ts, operator}}   (newest occurrence kept)
      coated   {seq: {serial, wo, ts}}   (closes AND backfills of the paired -970 P/N, by seq)
      receive_log [{serial, ts, sts, operator, result}]  EVERY receive occurrence
      coated_log  [{serial, wo, ts, note}]  EVERY coated-shipment occurrence; note = the
                  shipment note, with _CASE_B_NOTE appended for a -900-labelled board
                  shipped under the coated label (the report's Note column reads it)
      pd_path  the job's Production Data folder (from the newest event that knows it)
      home_side  'testing' | 'coated' | None — which side of the pair supplied wo_path
      coated_pn_label  the REAL coated label ('125-1030-970C') from the newest clean
                       coated-side close; '' when none (report name/header, Q1)
      coated_pn_labels EVERY clean coated-side label walked, newest first (closes,
                       backfills, voided closes) — every spelling the report file may
                       carry, so a rev-letter change ('970C' -> '970D') MIGRATES the
                       one file instead of orphaning it (review round 1)
      testing_pn_labels  every OTHER spelling of the job's own label walked
                         ('1342201900' for a '134-2201-900' walk) — a file named by
                         one of them is this job's report too (review round 3)
      testing_label_seen  True once a -900-LABELLED close/backfill was walked, or a
                          receive event that PROVES a testing shipment: one holding
                          at least one serial that is not coated-form (…970…) and
                          not a Case-B board of this job (has_testing, Q3; Dhruvil
                          2026-09-23 decision (a): a supervisor-overridden receive
                          of a coated-form or Case-B board leaves an in-house-
                          tested job in-house-tested — same label, tag, home, name)
      coat_label_mismatch  coated identities of the Case-B boards (a -900-form serial
                           in a coated-labelled close or import) — scan_job_folders
                           routes the same boards' paperwork the same way
      hist_has_testing  _has_testing() as HISTORY alone decides it — snapshotted
                        before any folder-truth fold; the report's NAME reads this
                        so a Rebuild can never rename the file (review round 1)
    The _log lists feed the report's log sheets ("catch everything from shipment
    ... without thinking anything" — Dhruvil 2026-09-02); the maps keep newest-
    wins semantics for the status views. Never raises; a failed walk just yields
    empty maps."""
    state = {"eng": str(eng or ""), "pn": str(pn or ""), "shipped": {}, "received": {},
             "coated": {}, "receive_log": [], "coated_log": [],
             "ship_counts": {}, "recv_dates": {}, "voided_ships": [], "board_notes": {},
             "voided_rcv_rows": set(), "voided_coat_rows": set(),
             "_test_occ": {}, "_coat_occ": {},
             "pd_path": None, "wo_path": None, "walk_error": False,
             # in-house-tested jobs (Dhruvil 2026-09-23) — every reader uses .get()
             "home_side": None, "coated_pn_label": "", "coated_pn_labels": [],
             "testing_pn_labels": [],
             "testing_label_seen": False, "coat_label_mismatch": set()}
    # (private walk slots — _bf_wo_path, _coat_wo_path, _coat_bf_wo_path,
    #  _coat_bf_label, _recv_serials — are created lazily and popped in finalize;
    #  _home_unknown is SET in finalize and deliberately left on the state: it is
    #  what _report_lookup / _collect_report_rows read to say "can't check" rather
    #  than "no WO folder known" — review round 5)
    ekey, pkey = _norm_key(eng), _norm_key(pn)
    if not ekey or not pkey:
        return state
    core_d, own_side, _ = _tr_pn_parts(pn)   # core digits: letter-tolerant label match
    # The serial embed keys are DIGITS ONLY: a -900 key typed with a rev letter
    # ('134-2201-900A') must still embed-match '…1342201900…' serials and still
    # know '1342201970' as its coated side — testing_to_coated_pn is letter-blind
    # and returned '' for it, so every -970-form serial on a -900-labelled event
    # (the 4065 import shape) fell through to the testing bucket (review round 5).
    # One helper derives them for every site that embeds serials (review round 6).
    ckey = _norm_key(testing_to_coated_pn(pn)) or (_norm_key(core_d + "970") if core_d and own_side == "900" else "")
    pk_al, ck_al = _tr_embed_keys(pn)
    voided = set()

    def _own_label(rpn):
        """True when an event's pn LABEL is the walked job's OWN label in ANY
        spelling: exact, or the same core digits AND the same side ('134-2201-900'
        and '1342201900' are one -900 label — punctuation is part of _norm_key, so
        the exact test alone split such a job in two, gave it two dropdown entries
        and two reports of which only one kept updating; '125-1030-970' and
        '125-1030-970C' are one coated label for a BackfillDialog -970 walk).
        A folder-name / 'N/A' label never is (review round 3)."""
        if rpn == pkey:
            return True
        rc, rs, _ = _tr_pn_parts(rpn)
        return bool(core_d) and rc == core_d and rs == own_side

    def _is_coat_label(rpn):
        """PREDICATE (not the module-level _coated_label, which returns the label
        to SHOW): True when an event's pn LABEL is this job's coated side — exact
        ('125-1030-970') or with the rev letter PCC prints after the 970
        ('125-1030-970C', real Anduril jobs). A folder-name/'N/A' label is not,
        and neither is the walked job's OWN label in any spelling: BackfillDialog
        walks a picked -970 WO folder with the -970 P/N typed in, and that label
        must stay the testing side of that walk or every closed coated serial
        would look 'new' (review rounds 1 + 3)."""
        if _own_label(rpn):
            return False
        if ckey and rpn == ckey:
            return True
        rc, rs, _ = _tr_pn_parts(rpn)
        return bool(core_d) and rc == core_d and rs == "970"

    def _note_coat_label(rec):
        """Remember a clean coated-side label spelling (newest first, deduped)."""
        lbl = str(rec.get("pn") or "").strip()
        if lbl and lbl.casefold() not in {x.casefold() for x in state["coated_pn_labels"]}:
            state["coated_pn_labels"].append(lbl)

    def _note_own_label(rec):
        """Remember every OTHER spelling of the walked job's own label (newest
        first, deduped) — a file named by one of them is this job's report too
        (_report_pn_spellings), so the two spellings' files fold into ONE."""
        lbl = str(rec.get("pn") or "").strip()
        if (lbl and _norm_key(lbl) != pkey and _tr_pn_parts(lbl)[0]
                and lbl.casefold() not in {x.casefold() for x in state["testing_pn_labels"]}):
            state["testing_pn_labels"].append(lbl)

    def _route_ship_event(rec, et):
        """Shared for 'close' and 'backfill_shipped': route EVERY serial by the
        P/N EMBEDDED IN THE SERIAL ITSELF — the event's pn label can be wrong
        (Dhruvil's 4065: coated boards imported under the testing P/N; an Aug-03
        close whose pn was recorded as the whole folder name). A …970… serial is
        a coated board no matter what the event says; a …900… serial is a testing
        shipment UNLESS the event's own label is this job's COATED side (Case B,
        Q3: coated log row + _CASE_B_NOTE, never `shipped`); serials with no
        embedded P/N follow the event label."""
        rpn = _norm_key(rec.get("pn"))
        is_test = _own_label(rpn); is_coat = _is_coat_label(rpn)
        is_close = (et == "close")
        wo = str(rec.get("wo") or "")
        ts = str(rec.get("ts") or "") if is_close else str(rec.get("ship_date") or rec.get("ts") or "")
        rnote = str(rec.get("notes") or "").strip()   # note typed on this shipment
        if is_test:
            state["testing_label_seen"] = True   # a -900-LABELLED shipment (Q3)
            _note_own_label(rec)                 # (its spelling, when it differs from ours)
        if is_test and is_close:
            if state["pd_path"] is None and rec.get("wo_folder"):
                try: state["pd_path"] = str(Path(str(rec["wo_folder"])).parent)
                except Exception: pass
            # The report xlsx lives INSIDE the newest testing WO's folder
            # (walk is newest-first, so the first close seen wins).
            if state["wo_path"] is None and rec.get("wo_folder"):
                state["wo_path"] = str(rec["wo_folder"])
        if is_test and not is_close:
            if state["pd_path"] is None and rec.get("production_data"):
                state["pd_path"] = str(rec["production_data"])
            # Fallback home for the xlsx when the job has NO recorded closes:
            # a backfilled WO's folder (a real close, seen first, always wins).
            if state.get("_bf_wo_path") is None and wo and rec.get("production_data"):
                try: state["_bf_wo_path"] = str(Path(str(rec["production_data"])) / wo)
                except Exception: pass
        # In-house-tested job (Dhruvil 2026-09-23): a coated (-970) shipment is the
        # report's LOWEST-priority home — adopted after the walk ONLY when no -900
        # close/backfill anchor exists. Detected from the SERIALS (…970… embedded)
        # too, so a close whose pn label is 'N/A' or a folder name (pre-v58 shape)
        # still anchors. Never touches pd_path/wo_path here.
        coat_evt = (not is_test) and (is_coat or any(_embeds(ck_al, s) for s in (rec.get("serials") or [])))
        if coat_evt and is_close and rec.get("wo_folder") and state.get("_coat_wo_path") is None:
            state["_coat_wo_path"] = str(rec["wo_folder"])
        if coat_evt and not is_close and wo and rec.get("production_data") and state.get("_coat_bf_wo_path") is None:
            try: state["_coat_bf_wo_path"] = str(Path(str(rec["production_data"])) / wo)
            except Exception: pass
        # the REAL coated label ('125-1030-970C') for the report name/header (Q1):
        # newest CLEAN coated-side close wins; a coated-labelled import is the fallback
        if is_coat and _tr_pn_parts(rpn)[0]:      # (_is_coat_label already excludes our own label)
            slot = "coated_pn_label" if is_close else "_coat_bf_label"
            if not state.get(slot):
                state[slot] = str(rec.get("pn") or "").strip()
            _note_coat_label(rec)          # every spelling the file may be named by
        ev_counted = set()   # a serial scanned twice in ONE event = one shipment
        for s in (rec.get("serials") or []):
            # Case B (Dhruvil 2026-09-23, Q3): a -900-form serial scanned into a
            # shipment whose label is the COATED side (the 8/27 "Anduril test damage"
            # closes on 4006/4007/4023; an ⏪ import of the in-progress -900-serialized
            # batch under its coated WO's label) IS a coated shipment — it goes to
            # the coated log with a Note, never to shipped/ship_counts. Only the
            # EVENT'S LABEL decides this (close or backfill alike — review round 2:
            # list_testing_jobs already files such a job as coated-only, and the
            # window's stop / report name must agree); the 4065 direction (-970
            # serials under a -900 label -> coated) is untouched.
            case_b = is_coat and not _embeds(ck_al, s) and _embeds(pk_al, s)
            if _embeds(ck_al, s) or (is_coat and not _embeds(pk_al, s)) or case_b:
                if str(s).strip():
                    kc = _serial_identity(s, _coat_ident_pn(pn))
                    onote = rnote
                    if case_b:
                        onote = (rnote + "  ·  " + _CASE_B_NOTE) if rnote else _CASE_B_NOTE
                        state["coat_label_mismatch"].add(kc)
                    if ("C", kc) not in ev_counted:
                        # record ONE coated occurrence per event; import(backfill)
                        # vs real close is reconciled after the walk so the SAME
                        # coated shipment recorded twice isn't listed twice
                        # (Dhruvil 2026-09-08: coated serials "coming twice").
                        ev_counted.add(("C", kc))
                        state["_coat_occ"].setdefault(kc, []).append(
                            (et, wo, ts, str(s), serial_seq(s), onote))
                    if rnote:
                        state["board_notes"].setdefault(kc, rnote)   # newest-first: first wins
                continue
            if not (is_test or _embeds(pk_al, s)):
                continue          # unrelated serial on a mislabeled event
            k = _serial_identity(s, pn)   # rev-agnostic (Dhruvil 2026-09-02)
            if not k:
                continue
            if rnote:
                state["board_notes"].setdefault(k, rnote)   # note carries to this board
            if k not in ev_counted:
                # record ONE testing occurrence per event; ship_count is finalized
                # after the walk (a real close and a redundant backfill IMPORT of
                # the same shipment count once; two closes = a genuine re-test).
                ev_counted.add(k)
                state["_test_occ"].setdefault(k, []).append((et, wo))
            cur = state["shipped"].get(k)
            if is_close:
                # a real close always beats a backfill entry for the same serial
                # (regardless of event-file walk order); among closes, the newest
                # (seen first) wins.
                if cur is None or cur.get("source") == "backfill":
                    state["shipped"][k] = {"serial": str(s), "wo": wo, "ts": ts,
                                           "sts": ts, "source": "close"}
            else:
                if cur is None:
                    # sts="" -> a backfilled (pre-history) shipment always sorts
                    # OLDER than any recorded receive, which is correct.
                    state["shipped"][k] = {"serial": str(s), "wo": wo, "ts": ts,
                                           "sts": "", "source": "backfill"}

    try:
        for rec in _iter_history_files(progress):
            et = rec.get("event_type")
            if et == "void":
                t = str(rec.get("target_event_id") or "")
                if t: voided.add(t)
                continue
            if str(rec.get("event_id") or "") in voided:
                # remember WHAT was voided: the shipment's folder files still
                # exist, and the folder scan must not resurrect them
                if et in ("close", "backfill_shipped") and _norm_key(rec.get("eng")) == ekey:
                    state["voided_ships"].append({
                        "wo": str(rec.get("wo") or ""),
                        "serials": {_norm_serial(x) for x in (rec.get("serials") or [])
                                    if str(x).strip()}})
                # Drop this voided record's rows from the LOG sheets so a re-added
                # correction merges CLEAN (Dhruvil 2026-09-16: voided the 106 receive,
                # re-added it as FAIL — the stale PASS row must not linger + conflict).
                # Keyed (identity, date) exactly like a sheet row.
                if _norm_key(rec.get("eng")) == ekey:
                    # A spelling of our OWN label that survives only on voided events
                    # is still a name the report FILE may carry: the after-close update
                    # renames the file to the newest close's spelling, and when that
                    # close is later voided (a typo fixed in 🕘 History) the renamed
                    # file must stay findable — or it is orphaned and the next Build
                    # makes a second one. The coated side has had the same mirror
                    # (_note_coat_label below) since review round 3 (review round 5).
                    if _own_label(_norm_key(rec.get("pn"))):
                        _note_own_label(rec)
                    if et == "receive":
                        vdate = str(rec.get("received_date") or "") or str(rec.get("ts") or "")
                        vdk = _date_key(vdate)
                        for s in (rec.get("serials") or []):
                            ki = _serial_identity(s, pn)
                            if ki:
                                state["voided_rcv_rows"].add((ki, vdk))
                    else:   # only a coated (…970) close/backfill has LOG-sheet rows
                        cpn_ = _coat_ident_pn(pn)
                        vdate = str(rec.get("ts") or "") if et == "close" \
                                else str(rec.get("ship_date") or rec.get("ts") or "")
                        vdk = _date_key(vdate)
                        # same Case-B rule as the live routing: a -900-form serial in a
                        # voided coated-LABELLED close/backfill also had a coated LOG row
                        vrpn = _norm_key(rec.get("pn")); v_coat = _is_coat_label(vrpn)
                        if v_coat and _tr_pn_parts(vrpn)[0]:
                            _note_coat_label(rec)   # a report built under a since-voided label
                        for s in (rec.get("serials") or []):
                            if _embeds(ck_al, s) or (v_coat and _embeds(pk_al, s)):
                                kc = _serial_identity(s, cpn_)
                                if kc:
                                    state["voided_coat_rows"].add((kc, vdk))
                continue
            if et in ("close", "backfill_shipped"):
                if _norm_key(rec.get("eng")) != ekey:
                    continue
                _route_ship_event(rec, et)
            elif et == "receive":
                if _norm_key(rec.get("eng")) != ekey or not _own_label(_norm_key(rec.get("pn"))):
                    continue        # (own label in any spelling — a receive saved under the
                                    #  job's other -900 spelling is this job's receive)
                ts = str(rec.get("ts") or ""); op = str(rec.get("operator") or "")
                # Does this receive PROVE a testing shipment? Decided in finalize, not
                # here (Dhruvil 2026-09-23, decision (a)): the walk is newest-first, so
                # a receive is seen BEFORE the (older) coated close that fills
                # coat_label_mismatch — only once the walk is complete can a received
                # -900-form serial be told apart from a Case-B board of this job.
                state.setdefault("_recv_serials", []).extend(
                    str(s) for s in (rec.get("serials") or []) if str(s).strip())
                _note_own_label(rec)
                # Historical imports carry the ORIGINAL receive date the operator
                # typed; it drives both display and cycle-ordering.
                rdate = str(rec.get("received_date") or "")
                ts_disp = rdate or ts
                sts = (rdate + " 00:00:00") if rdate else ts
                failed = {_serial_identity(x, pn) for x in (rec.get("failed") or [])}
                rnote = str(rec.get("notes") or "").strip()   # note typed on this receive
                if state["pd_path"] is None and rec.get("production_data"):
                    state["pd_path"] = str(rec["production_data"])
                for s in (rec.get("serials") or []):
                    k = _serial_identity(s, pn)
                    if not k:
                        continue
                    if rnote:
                        state["board_notes"].setdefault(k, rnote)   # newest-first: first wins
                    entry = {"serial": str(s), "ts": ts_disp, "sts": sts, "operator": op,
                             "result": "fail" if k in failed else "pass", "note": rnote}
                    state["receive_log"].append(dict(entry))
                    if _embeds(ck_al, s):
                        # a coated-form serial is not a testing receive for the
                        # COUNTS/received map — but a supervisor-overridden scan
                        # is deliberate, so the log sheet still shows it
                        continue
                    # one receive per DISTINCT DATE (a re-scan of the same board
                    # on the same day is not a second receive)
                    state["recv_dates"].setdefault(k, set()).add((ts_disp or "")[:10])
                    # keep the entry with the NEWEST receive date — the walk is
                    # newest-FILE-first, but an old paper receive imported later
                    # (received_date) must not shadow a genuinely newer receive
                    prev = state["received"].get(k)
                    if prev is None or (entry.get("sts") or "") > (prev.get("sts") or ""):
                        state["received"][k] = entry
    except Exception:
        # The walk died mid-way (share blip): the maps are INCOMPLETE. Flag it so
        # callers refuse to save/regenerate from partial data — overwriting the
        # report from a truncated walk would silently erase rows.
        state["walk_error"] = True
    # ── home of the report (Dhruvil 2026-09-23): -900 close > -900 backfill >
    #    -970 close > -970 backfill. The coated slots are private and popped; the
    #    only public trace is home_side. Strings only — nothing is created here,
    #    write_received_xlsx still refuses unless the folder EXISTS on disk. ─────
    bf = state.pop("_bf_wo_path", None)
    cw = state.pop("_coat_wo_path", None); cbf = state.pop("_coat_bf_wo_path", None)
    if not state.get("coated_pn_label"):
        state["coated_pn_label"] = state.get("_coat_bf_label") or ""
    state.pop("_coat_bf_label", None)
    # -900-form boards received back are proof of a -900 shipment — UNLESS every
    # one of them is a Case-B board of this very job (the 8/27 "Anduril test
    # damage" boards wear -900 labels but History filed their shipment as COATED,
    # coat_label_mismatch): a supervisor-overridden receive of one of those is
    # not a testing shipment either (review round 3: it stripped an in-house-
    # tested job of its home and orphaned its existing report).
    cpn_i = _coat_ident_pn(pn)
    _is_caseb = lambda s: _serial_identity(s, cpn_i) in state["coat_label_mismatch"]
    # Receive events (Dhruvil 2026-09-23, decision (a) on the open Q3 question): a
    # receive flips testing_label_seen iff at least ONE of its serials is NOT
    # coated-form (…970… embedded) AND is NOT a Case-B board of this job — i.e. a
    # genuine -900-labelled board came back, so a real testing shipment went out.
    # A supervisor-overridden receive of a coated-form board, or of one of the 8/27
    # "Anduril test damage" boards (coat_label_mismatch), proves nothing: the job
    # keeps its coated label, its dropdown tag, its coated home and its file name.
    # Decided HERE because coat_label_mismatch is only complete after the walk.
    if any(not _embeds(ck_al, s) and not _is_caseb(s) for s in state.pop("_recv_serials", [])):
        state["testing_label_seen"] = True
    all_caseb = all(_is_caseb(v.get("serial")) for v in state["received"].values())   # True when nothing was received
    if cw or cbf:
        # Dhruvil 2026-10-06: boards are no longer sent for testing — the Shipping
        # Report now homes in the -970 COATED WO folder for EVERY job that has a
        # coated side (reverses the earlier "-900 tree whenever the job ever shipped
        # for testing"). The coated anchor wins over a -900 close's wo_path; the
        # sibling-tree migration (_stale_received_candidates searches both trees)
        # moves an existing -900 report into the -970 home on the next write. A job
        # with NO coated side at all still homes in the -900 tree (below).
        # Safety stop kept (2026-10-19, approved by Dhruvil): when this PC cannot
        # check the job's folders — e.g. the -900 job folder that holds the existing
        # report can't be listed — _coated_tree_holds_report answers None. Homing in
        # the -970 WO anyway would answer 'no report' there and let a Build / Save
        # create a SECOND report beside the unreadable original (the migration would
        # then find a file already in its way). So: wo_path None (no writer places a
        # file) + _home_unknown ('could not check from this PC'), as before 10/6.
        if _coated_tree_holds_report(cw or cbf, state) is None:
            state["wo_path"] = None; state["home_side"] = None; state["_home_unknown"] = True
        else:
            state["wo_path"] = cw or cbf; state["home_side"] = "coated"
    elif state["wo_path"]:
        state["home_side"] = "testing"                  # -900 close, no coated side yet
    elif bf:
        state["wo_path"] = bf; state["home_side"] = "testing"     # -900 backfill
    elif (cw or cbf) and not state["shipped"] and all_caseb:
        # the coated anchor is for a job that never SHIPPED for testing. Testing
        # boards on record (`shipped`) or -900-form boards received back
        # (`received`) are proof of a -900 shipment even when that shipment
        # predates History: such a job keeps wo_path None exactly as before this
        # change ('No job folder … run Import older shipments') rather than
        # filing a -900-named testing report inside the coated WO folder (review
        # round 2: "any -900 shipment ever -> -900 tree"). A supervisor-overridden
        # receive of a COATED-form board is not such proof (the walk keeps those
        # out of `received`), so an in-house-tested job keeps its coated home.
        state["wo_path"] = cw or cbf; state["home_side"] = "coated"   # coated close > import
    elif cw or cbf:
        # THE FILE IS THE SWITCH (Dhruvil 2026-09-23: "Once report is created app
        # needs to update that report only"): a report that already lives in the
        # job's folders never loses its home to a History event — whatever
        # `shipped`/`received` now say, that file keeps being the one updated.
        # Looked for the way the WRITE path looks (the recorded WO string resolved
        # to the folder that exists, the previous WO right after a new WO's first
        # close, a rev-renamed job folder, the Production Data level — review
        # round 4: the raw newest-close string alone orphaned the report on the
        # next new WO or rev change; a rev-TWIN coated job folder or the -900
        # tree a human Rebuild parked it in — review round 5). The receive-only
        # job whose -900 shipment predates History (no report anywhere) still
        # refuses exactly as above. Read-only stats, only on this rare path;
        # never raises (review round 3). TRI-STATE: None = the folders could not
        # be checked from this PC — wo_path stays None (no writer can place a
        # file) but _home_unknown makes the opt-in switch answer 'could not
        # check', never 'no report' / 'no work-order folder known' (review round 5).
        held = _coated_tree_holds_report(cw or cbf, state)
        if held:
            state["wo_path"] = cw or cbf; state["home_side"] = "coated"
        elif held is None:
            state["_home_unknown"] = True
    if state["pd_path"] is None and state["home_side"] == "coated":
        try: state["pd_path"] = str(Path(state["wo_path"]).parent)
        except Exception: pass
    # ── finalize ship counts + coated log (Dhruvil 2026-09-08): reconcile a
    #    shipment recorded as BOTH a manual import and a real close. Closes are
    #    the record of an actual shipment (one per close event = one shipment,
    #    so a re-test re-ship still counts twice); a backfill IMPORT counts only
    #    when NO close covers that serial (it was a pre-history shipment). ─────
    for k, occs in state.pop("_test_occ", {}).items():
        closes = [o for o in occs if o[0] == "close"]
        if closes:
            state["ship_counts"][k] = len(closes)
        else:
            state["ship_counts"][k] = len({_alnum(o[1]) for o in occs}) or 1
    for kc, occs in state.pop("_coat_occ", {}).items():
        closes = [o for o in occs if o[0] == "close"]
        if closes:
            chosen = closes                    # the real coated shipment(s)
        else:
            seen_wo = set(); chosen = []       # imports: one per distinct WO
            for o in occs:
                w = _alnum(o[1])
                if w not in seen_wo:
                    seen_wo.add(w); chosen.append(o)
        for (et, wo, ts, serial, seq, cnote) in chosen:
            state["coated_log"].append({"serial": serial, "wo": wo, "ts": ts, "note": cnote})
            if seq and seq not in state["coated"]:
                state["coated"][seq] = {"serial": serial, "wo": wo, "ts": ts}
    # HISTORY's own answer to "did this job ever ship for testing" — taken before
    # any folder-truth fold so the report's NAME is a history fact (review round 1:
    # a -900 serial list under a coated-only job's idle -900 tree used to rename the
    # file on 🔁 Rebuild and rename it back on the next after-close update)
    state["hist_has_testing"] = _has_testing(state)
    return state

def receiving_summary(state):
    """Counts for the tracker chips. CYCLE-AWARE: 'received' counts boards whose
    LATEST movement is a receive (a failed board re-shipped for re-test moves back
    to 'outstanding'). Received boards split into coated & shipped (-970 close by
    seq), FAILED awaiting re-ship, and passed in-house awaiting coating.
    Also carries has_testing (see _has_testing) — the in-house-tested switch the
    report wording, the chips layout and the Receiving stop all read (Dhruvil
    2026-09-23); it is the one key here that is not a count."""
    shipped = state.get("shipped") or {}
    received = state.get("received") or {}
    coated_map = state.get("coated") or {}
    cur = {k: v for k, v in received.items() if _currently_received(state, k)}
    outstanding = sum(1 for k in shipped if k not in cur)
    coated = sum(1 for v in cur.values() if serial_seq(v["serial"]) in coated_map)
    failed = sum(1 for v in cur.values()
                 if v.get("result") == "fail" and serial_seq(v["serial"]) not in coated_map)
    # DISTINCT coated boards ever shipped (log-based) — independent of whether a
    # matching testing/receive record exists (boards can coat-and-ship directly).
    # Counted by board SEQUENCE where the serial has one (the same key `coated`
    # and the receive<->coated matching use): a Case-B board (-900 label under
    # the coated WO) later re-shipped with its -970 label is ONE board, not two
    # identities (review round 5); a seq-less serial falls back to its identity.
    cpn = _coat_ident_pn(state.get("pn"))
    coated_all = len({serial_seq(e.get("serial")) or _serial_identity(e.get("serial"), cpn)
                      for e in (state.get("coated_log") or []) if e.get("serial")})
    return {"shipped": len(shipped), "received": len(cur), "outstanding": outstanding,
            "coated": coated, "failed": failed, "in_house": len(cur) - coated - failed,
            "coated_all": coated_all, "has_testing": _has_testing(state)}

def _has_testing(state):
    """Did this job EVER take part in the test-and-return flow? True when History
    holds (1) a -900-LABELLED close/backfill (testing_label_seen), (2) a testing-
    side board in `shipped` (label-routed, or a -900-FORM serial on a junk-labelled
    pre-v58 close — the one serial-form inference, shared with list_testing_jobs
    so the dropdown tag, the scan stop and the report agree; the status sheets
    list those boards and the header must never contradict its own sheets), or
    (3) a receive that PROVES a testing shipment: a received board that is neither
    coated-form (…970…) nor a Case-B board of this job (Dhruvil 2026-09-23,
    decision (a) — the walk sets testing_label_seen for it; the `received` clause
    here repeats the test for hand-built states, and the walk keeps coated-form
    serials out of `received` already). Case B (-900-form boards shipped under
    the coated label, coat_label_mismatch) never flips it (Q3). False -> the
    report reads 'In-house tested — not shipped for testing'."""
    mism = state.get("coat_label_mismatch") or set()
    cpn = _coat_ident_pn(state.get("pn"))
    recv_proof = any(_serial_identity(v.get("serial"), cpn) not in mism
                     for v in (state.get("received") or {}).values())
    return (bool(state.get("testing_label_seen")) or bool(state.get("shipped"))
            or recv_proof)

def received_xlsx_path(base_dir, pn):
    """Target path of the per-P/N received-boards workbook, inside `base_dir` —
    the job's HOME work-order folder: the newest -900 testing WO for a
    test-and-return job, the -970 coated WO for an in-house-tested one (see
    _resolve_received_wo_dir; Dhruvil 2026-09-02: the report folder lives INSIDE
    the work-order folder, next to Shipping Serials / Shipping Pictures, not at
    the Production Data level). `pn` is the report's NAME P/N — see _report_pn
    (Dhruvil 2026-09-23)."""
    return Path(base_dir) / RECEIVED_FOLDER / _received_fname(pn)

def _coated_label(state):
    """The coated label to SHOW for a job (chips, A3 header, coated-only report
    name): the real recorded label ('125-1030-970C') > the digit form of the
    coated P/N > the job P/N itself — one fallback chain for every caller."""
    return (str(state.get("coated_pn_label") or "").strip()
            or testing_to_coated_pn(state.get("pn")) or str(state.get("pn") or ""))

def _job_display_pn(j):
    """The P/N a list_testing_jobs ENTRY is shown under (dropdown, picker row,
    Rebuild-ALL progress / failure text): its real coated label for a coated-only
    entry ('125-1030-970C'), else its -900 key — the job-dict twin of _coated_label,
    in one place (review round 2)."""
    return str((j or {}).get("coated_label") or (j or {}).get("pn") or "")

def _report_pn(state):
    """The P/N the Shipping Report is NAMED after — its file name and the A1
    header (write_received_xlsx) — Dhruvil 2026-09-23, Q1: a job that ever wore a
    -900 label (every existing test-and-return report, every job with boards
    still at the customer) keeps the -900 name exactly as today; an in-house-
    tested (coated-only) job is named by its REAL coated label ('125-1030-970C
    Shipping Report.xlsx'), never by an invented -900 spelling. Decided from
    HISTORY alone (state['hist_has_testing'], snapshotted by the walk before any
    folder-truth fold) so a 🔁 Rebuild can never rename the file; a hand-built
    state without the snapshot falls back to _has_testing."""
    pn = str(state.get("pn") or "")
    if _hist_has_testing(state):
        return pn
    return _coated_label(state)

def _hist_has_testing(state):
    """History's OWN has_testing — snapshotted by the walk before any folder-truth
    fold (state['hist_has_testing']); recomputed for a hand-built state without
    the snapshot. The one rule that keeps the report NAME and the folder-side
    Case-B route on the history fact rather than on a folded state (review
    round 1: a folder-truth Rebuild used to rename the file and the next history-
    only update renamed it back). Every reader goes through here (review round 4)."""
    ht = state.get("hist_has_testing")
    return _has_testing(state) if ht is None else bool(ht)

def _report_pn_spellings(state):
    """Every P/N spelling this job's report may have been named with — the -900
    key, every OTHER spelling of that key walked ('1342201900' next to
    '134-2201-900' — review round 3), EVERY clean coated label ever walked (newest
    first; the letter after the 970 is the rev, so '970C' -> '970D' happens
    mid-job) and the digit-form coated P/N — so a rename between them (a
    coated-only job gaining testing activity, a label changing letter or
    punctuation) MIGRATES the one file instead of leaving a second one behind."""
    out = []
    for p in ([str(state.get("pn") or "")]
              + [str(x or "") for x in (state.get("testing_pn_labels") or [])]
              + [str(state.get("coated_pn_label") or "")]
              + [str(x or "") for x in (state.get("coated_pn_labels") or [])]
              + [_coat_ident_pn(state.get("pn"))]):     # (digit form; letter-tolerant for a lettered -900 key — review round 6)
        p = p.strip()
        if p and p.casefold() not in {q.casefold() for q in out}:
            out.append(p)
    return out

def _wo_holds_report(wo, state, whole_tree=False):
    """Does the WO folder `wo` already hold this job's report — under ANY spelling
    the file may carry (_report_pn_spellings), current or legacy folder + file
    name? With whole_tree the WO's sibling WO folders and the Production Data
    level of that job tree count too. READ-ONLY (one stat per name + one
    read-only workbook open per hit), nothing is created. TRI-STATE (review round
    7): True / False / None = a stat or listing RAISED and nothing was found — the
    caller must not read None as 'does not hold' (pathlib re-raises an access-
    denied stat inside the 'Shipping Reports' leaf; a definite False there made
    _coated_tree_holds_report refuse a coated home whose report sits right there,
    and the chips said 'no work-order folder known' instead of 'can't check from
    this PC'). Callers: build_receiving_state (through _coated_tree_holds_report)
    keeps a coated home that already holds the report; _sibling_wo_holding_report
    asks it per WO / at the Production Data level for _resolve_received_wo_dir's
    holder-first passes (a Save / after-close update reaches a report a human
    🔁 Rebuild moved into the sibling tree; the home pass once the recorded WO is
    gone). A same-named file whose own head line names ANOTHER job (4016 / 4016LF
    share a P/N) is not a hit — the merge refuses it as base and never deletes
    it, so it must not anchor a home either (review round 5; one workbook head
    read, only on these rare paths)."""
    raised = False
    try:
        w = Path(wo)
        locs = [w]
        if whole_tree:
            pdd = w.parent
            locs.append(pdd)
            try:
                locs += [d for d in pdd.iterdir()
                         if d.is_dir() and not _is_report_folder(d.name) and d != w]
            except Exception:
                raised = True               # sibling WOs unknown: a copy may sit in one
        names = []
        for sp in _report_pn_spellings(state):
            names += [_received_fname(sp), _received_legacy_fname(sp)]
        for loc in locs:
            for sub in (RECEIVED_FOLDER, RECEIVED_FOLDER_LEGACY):
                for fn in names:
                    try:
                        if (loc / sub / fn).is_file() and not _copy_names_other_job(loc / sub / fn, state.get("eng")):
                            return True
                    except Exception:
                        raised = True       # access denied / share blip on that name: don't guess
                        continue
    except Exception:
        return None
    return None if raised else False

def _coated_tree_holds_report(cw, state):
    """build_receiving_state's 'the file is the switch' test (review round 4): does
    the COATED job tree anchored by the recorded WO string `cw` already hold this
    job's report? Resolved exactly like the write path resolves it — a probe copy
    of the state homed in that tree goes through _resolve_received_wo_dir (the
    recorded folder when it exists, else the newest live -970 WO of the job folder
    even after a rev rename; no tree switch) — then _wo_holds_report over that WO,
    its sibling WOs and the Production Data level. When that ONE job folder does
    not hold it, every other live folder of the job is asked too (review round 5):
    a rev-TWIN coated job folder ('125-1030-970B 4094 Rev. B' next to '125-1030-970
    4094 Rev. A' — the report sits in the twin the previous batch closed into) and
    the -900 tree a human 🔁 Rebuild parked it in; the write then migrates the
    file to the resolved WO exactly as it does inside the home tree. READ-ONLY,
    creates nothing, never raises. TRI-STATE (review round 5): True / False /
    None = could not check (the customer folder is not reachable from this PC, a
    matched job folder could not be listed, or a filesystem call raised) — the
    caller must not turn that into 'the folders hold no report'."""
    try:
        pr = dict(state, wo_path=str(cw), home_side="coated")
        pr.pop("allow_tree_switch", None)
        try: pr["pd_path"] = str(Path(str(cw)).parent)
        except Exception: pass
        d = _resolve_received_wo_dir(pr)
        unknown = d is None and bool(pr.get("_resolve_unknown"))
        if d is not None:
            h = _wo_holds_report(d, state, whole_tree=True)
            if h:
                return True
            if h is None:
                unknown = True              # a stat raised inside that tree: 'could not check' (review round 7)
        try:
            cust = Path(str(cw)).parent.parent.parent
            cust_ok = cust.is_dir()
        except Exception:
            cust, cust_ok = None, False
        if cust_ok:
            ek = _norm_key(state.get("eng"))
            # the walk's own embed keys (review round 7: this seventh derivation
            # site was missed by the round-6 sweep into _tr_embed_keys)
            pk_al, ck_al = _tr_embed_keys(state.get("pn"))
            for side_al in (ck_al, pk_al):           # the coated side first
                held, unreadable = _sibling_wo_holding_report(cust, ek, side_al, state)
                if held is not None:
                    return True
                if unreadable:
                    unknown = True
        return None if unknown else False
    except Exception:
        return None

_ARCHIVE_SEGS = ("archives", "archive", "archived", "old")

def _under_archive(p):
    """True when ANY path segment is an Archive/Old folder — such copies are frozen
    snapshots: they may be read as a merge BASE but are never deleted (Dhruvil
    2026-09-23). Unknown/unparseable -> True (protected)."""
    try:
        return any(str(x).strip().lower() in _ARCHIVE_SEGS for x in Path(p).parts)
    except Exception:
        return True

def _job_dir_is(nm, ek):
    """Does a customer-level job folder name belong to Eng# `ek`? The Eng# must
    stand as a WHOLE token — 'GizmoX [4041 Rev. A]', '125-1030-970C 4006 Rev. C',
    '122-9538-970 4016LF Rev. A02' — never as the head of a longer one: with a
    plain substring test the 4016LF folder was Eng 4016's too, and once the
    sibling-tree candidate scan fed the harvest path that job's report was merged
    into ours and DELETED (review round 2). Letters and digits glued to the Eng#
    break the match; punctuation and spaces around it do not. The P/N half of the
    predicate stays with the caller."""
    ek = _norm_key(ek)
    if not ek:
        return False
    return re.search(r'(?<![0-9A-Z])' + re.escape(ek) + r'(?![0-9A-Z])', _norm_key(nm)) is not None

def _pick_report_wo(cust, ek, pn_al, want_wo):
    """ONE pass of _resolve_received_wo_dir over the customer dir for ONE side of
    the job pair (pn_al = the -900 OR the -970 P/N, alnum form; a lettered folder
    name '125-1030-970C 4006 …' still contains the digits): the newest EXISTING WO
    folder of a matching job folder (recorded WO # preferred, then mtime).
    Returns (best_or_None, unreadable). `unreadable` is True when a job folder
    whose NAME matched could not be listed (partial share access from this PC):
    the caller must then REFUSE rather than guess the sibling tree — writing a
    fork there would be a second file (Dhruvil 2026-09-23)."""
    best = None; best_key = None; unreadable = False
    if len(pn_al) < 6:
        return None, False
    try:
        for jd in cust.iterdir():
            nm = jd.name
            if nm.strip().lower() in _ARCHIVE_SEGS:
                continue                # never descend the Archive/Old folder
            if not _job_dir_is(nm, ek) or pn_al not in _alnum(nm):
                continue                # this Eng# (whole token) AND this side's P/N only
            try:
                if not jd.is_dir():
                    continue
                pdd = jd / "Production Data"
                if not pdd.is_dir():
                    continue
                wos = list(pdd.iterdir())
            except Exception:
                unreadable = True       # a matching job we cannot list: don't guess
                continue
            for wo in wos:
                try:
                    if not wo.is_dir():
                        continue
                except Exception:
                    continue
                if _is_report_folder(wo.name):
                    continue
                match = bool(want_wo) and _alnum(wo.name) == want_wo
                try: mt = wo.stat().st_mtime
                except Exception: mt = 0
                key = (1 if match else 0, mt)
                if best_key is None or key > best_key:
                    best_key = key; best = wo
    except Exception:
        return None, True
    return best, unreadable

def _sibling_wo_holding_report(cust, ek, pn_al, state):
    """The WO folder of ONE side of the -900/-970 pair (`pn_al`, that side's
    serial embed key) that already HOLDS this job's report. Written for the
    SIBLING tree (review round 4) and since round 6 the HOME tree's holder search
    too, once the recorded WO folder is gone — so the home pass hands it `first`
    (the home side's key), which is not a copy-paste slip. The everyday update
    (after-close daemon, receiving Save) updates a report a human 🔁 Rebuild moved
    there IN PLACE, so the WO that holds the file
    is the answer, not the tree's newest WO by mtime: every write into 'the newest'
    migrated the file, the emptied leaf's rmdir bumped the OLD WO's mtime, and the
    file ping-ponged between two live WOs on every update. Every job folder of that
    side is looked at (a rev twin '150-3300-970B 5001 Rev. B' next to '150-3300-970
    5001 Rev. A' may be the one holding it). A copy parked at a tree's Production
    Data level resolves to that tree's newest WO (the write migrates it down, as in
    the home tree). READ-ONLY; never raises. Returns (holder_or_None, unreadable):
    `unreadable` is True when the customer folder or a job folder whose NAME
    matched could not be listed from this PC — the copy may sit there, so the
    caller must answer 'could not check', never 'nothing holds it' (review round
    5: this pass used to lean on a second, otherwise wasted _pick_report_wo
    listing of the same folders for that flag; review round 7: a stat that RAISED
    inside a WO's 'Shipping Reports' leaf counts too — _wo_holds_report's None).
    Callers: _resolve_received_wo_dir (the home pass when the recorded WO folder
    is gone, and the sibling pass — the human Rebuild prefers the holder too, see
    there) and _coated_tree_holds_report (both sides)."""
    holder = None; hkey = None
    pd_hit = None; pd_key = None
    unreadable = False
    if len(pn_al) < 6:
        return None, False
    def _mt(p):
        try: return p.stat().st_mtime
        except Exception: return 0
    try:
        for jd in cust.iterdir():
            nm = jd.name
            if nm.strip().lower() in _ARCHIVE_SEGS:
                continue                # never the Archive/Old folder
            if not _job_dir_is(nm, ek) or pn_al not in _alnum(nm):
                continue
            try:
                pdd = jd / "Production Data"
                if not jd.is_dir() or not pdd.is_dir():
                    continue
                wos = [w for w in pdd.iterdir() if w.is_dir() and not _is_report_folder(w.name)]
            except Exception:
                unreadable = True       # a matching job folder we cannot list: don't guess
                continue
            for wo in wos:
                h = _wo_holds_report(wo, state)
                if h:
                    if hkey is None or _mt(wo) > hkey:
                        holder, hkey = wo, _mt(wo)
                elif h is None:
                    unreadable = True       # a stat inside that WO raised: it may hold it (review round 7)
            if wos:
                h = _wo_holds_report(pdd, state)          # PD-level copy -> that tree's newest WO
                if h:
                    newest = max(wos, key=_mt)
                    if pd_key is None or _mt(newest) > pd_key:
                        pd_hit, pd_key = newest, _mt(newest)
                elif h is None:
                    unreadable = True
    except Exception:
        return None, True
    return (holder if holder is not None else pd_hit), unreadable

def _resolve_received_wo_dir(state):
    """The REAL on-disk WO folder to hold the Shipping Report. state['wo_path']
    is a STRING captured when a close was saved; a mid-job REV CHANGE renames/
    moves the job folder and the recorded spelling drifts (Dhruvil 2026-09-08,
    job 4092: History holds the SAME rev-B02 folder as both 'Rev B01 Rev B02' and
    'Rev B01 Rev. B02', and the old rev is under an 'Archive' folder). Trusting
    the string and mkdir(parents=True) would FABRICATE a whole new folder tree
    next to the 4 real ones.

    So: use wo_path only when it actually EXISTS; otherwise find the newest
    EXISTING WO folder for this Eng#+P/N by listing the customer directory
    (archived/old job folders skipped). Returns a Path that exists, or None when
    nothing real can be found (the caller must then refuse rather than invent a
    folder). Within the home tree the recorded WO NAME wins when it still exists
    under a renamed job folder (the fast path's twin); when the recorded WO folder
    itself is gone, the live WO that already HOLDS the report wins over the newest
    by mtime (review round 6 — see the comment at that pass), and only a first
    placement takes the newest.

    In-house-tested jobs (Dhruvil 2026-09-23): a job has a -900 job tree, a -970
    one, or both. The search runs in TWO ordered passes driven by
    state['home_side'] — the tree History anchored (testing = -900 first, coated
    = -970 first). The SIBLING tree is tried only when the home tree has no live
    WO folder. There the WO that already HOLDS the job's report wins for EVERY
    caller (a human moved it there earlier — updating it in place is the everyday
    update Dhruvil asked for, review round 3; and a human Rebuild must not shift
    it to a newer sibling WO either, review round 5: the emptied leaf's rmdir made
    the previous holder newest by mtime, so every Rebuild relocated the file
    between two live WOs). Only when NO sibling folder holds it may a caller with
    state['allow_tree_switch'] (a human 🔁 Rebuild / Rebuild ALL) PLACE the file
    in the sibling tree's newest WO — state['tree_switched'] tells the caller the
    file is about to be placed / moved INTO that tree; it stays False when the
    file already lives there. An auto-update or Save never places a report in
    the sibling tree (it does bring one HOME: when the home tree has a live WO
    again the first pass wins and the write migrates the sibling copy there, as
    it migrates any older-WO copy). A matched job folder that cannot be listed
    (partial share access) -> None, never a guess.
    state['_resolve_unknown'] (review round 1) says WHY a None came back: True
    when this PC could not even list the customer folder / the matched job folder
    (nothing can be concluded from here), False when the folders were listable
    and simply hold no live WO for the job."""
    state["tree_switched"] = False
    state["_resolve_unknown"] = False
    wp = state.get("wo_path")
    try:
        if wp and Path(wp).is_dir():
            return Path(wp)             # healthy fast path — unchanged behavior
    except Exception:
        pass                            # no access from this PC: try to resolve
    ek = _norm_key(state.get("eng")); pk_al, ck_al = _tr_embed_keys(state.get("pn"))
    if len(pk_al) < 6 or not ek:
        return None
    cust = None                         # <cust>\<job>\Production Data\<WO>
    for base, up in ((wp, 3), (state.get("pd_path"), 2)):
        if not base:
            continue
        try:
            p = Path(base)
            for _ in range(up):
                p = p.parent
            if p.is_dir():
                cust = p; break
        except Exception:
            continue
    if cust is None:
        state["_resolve_unknown"] = True    # customer folder not reachable from here
        return None
    want_wo = _alnum(Path(wp).name) if wp else ""
    first, second = ((ck_al, pk_al) if state.get("home_side") == "coated" else (pk_al, ck_al))
    best, unreadable = _pick_report_wo(cust, ek, first, want_wo)
    if best is not None:
        if want_wo and _alnum(best.name) == want_wo:
            return best                 # the recorded WO itself, under a renamed job folder
        # The home tree has live WOs but the recorded WO folder is GONE (renamed
        # or archived on its own — not a job-folder rev rename): the WO that
        # already HOLDS the report wins here too (review round 6). Ranking the
        # live WOs by mtime moved the file into the newest one, and the merge's
        # rmdir of the emptied 'Shipping Reports' leaf bumped the OLD WO's mtime
        # above it — so every after-close update and every receiving Save deleted
        # and re-created the one file in the OTHER WO, forever (round 5 fixed the
        # same ping-pong for the sibling tree only). Only a FIRST placement
        # (nothing holds it anywhere in the home tree) takes the newest by mtime.
        # One more listing of the home side, on this rare path only.
        held, h_unreadable = _sibling_wo_holding_report(cust, ek, first, state)
        if held is not None:
            return held
        if h_unreadable:
            state["_resolve_unknown"] = True
            return None                 # a home job folder we cannot list may hold it: don't guess
        return best
    if unreadable:
        state["_resolve_unknown"] = True
        return None                     # can't see the home tree: never guess the sibling
    # The sibling tree: the WO that already HOLDS the report, for every caller.
    # A Save / after-close update never PLACES a report in the sibling tree — but
    # once a human 🔁 Rebuild has re-homed it there, that file is the job's report
    # and "Once report is created app needs to update that report only" (Dhruvil
    # 2026-09-23) applies to it: updating a file that already exists there is
    # neither a move nor a creation (review round 3: the daemon used to refuse
    # silently and a receiving Save showed 'File error' for a report the chips
    # just named; review round 4: 'the newest WO of a tree that holds it
    # somewhere' moved the file back and forth between two live sibling WOs on
    # every update, and missed a rev-twin job folder; review round 5: the human
    # allow_tree_switch path still took the newest WO by mtime, so every manual
    # Rebuild relocated the file between two live sibling WOs — the emptied
    # leaf's rmdir bumps the WO the file just left). tree_switched stays False —
    # nothing moves. One listing of the sibling side; the helper reports an
    # unlistable matched job folder itself.
    held, sib_unreadable = _sibling_wo_holding_report(cust, ek, second, state)
    if held is not None:
        return held
    if not state.get("allow_tree_switch"):
        if sib_unreadable:
            state["_resolve_unknown"] = True    # a sibling job folder we cannot list may hold it
        return None
    # a human Rebuild may PLACE the file in the sibling tree's newest WO when
    # nothing there holds it (a first build, or the only copy sits in the home
    # tree / under Archive): tree_switched tells the caller so
    best, unreadable = _pick_report_wo(cust, ek, second, want_wo)
    if best is None:
        if unreadable or sib_unreadable:
            state["_resolve_unknown"] = True    # the sibling job folder exists but can't be listed from here
        return None
    state["tree_switched"] = True
    return best

def _same_file(a, b):
    """True when two paths are provably the SAME physical file (mapped drive vs
    UNC spelling of one share, etc.). False on any doubt or error — callers use
    this to avoid deleting the one true file through an alias."""
    try:
        return os.path.samefile(str(a), str(b))
    except Exception:
        return False

def _stale_received_candidates(state, target, include_archive=False):
    """Same-named report files in every OTHER possible home (the legacy
    Production-Data level and every other WO folder) — the places a copy could
    linger after the newest-WO location moved. A share blip mid-enumeration
    returns the PARTIAL list (never an error): missing a candidate only delays
    its merge; it must never cause one to be deleted unmerged.
    In-house-tested jobs (Dhruvil 2026-09-23): the report may live in EITHER job
    tree (-900 / -970 siblings under the customer dir), so every sibling job
    folder's Production Data and WO folders are candidates too, under every
    spelling in _report_pn_spellings (the -900 key and its other spellings, every
    clean coated label, the digit form).
    Archive/Old folders are entered ONLY with include_archive (the merge BASE
    when the target is missing after a tree switch) and are never deleted.
    state['_cand_partial'] (review round 4) records whether a LIVE-tree enumeration
    was swallowed: the opt-in switch (_report_lookup) must answer 'could not
    check', never 'no report', from a partial list, and a first-time build must
    refuse rather than fork. A failed listing of the Archive/Old descent goes to
    state['_arch_partial'] instead (review round 5): an archived copy can never
    make the switch True (R6) nor be deleted, so not seeing it proves nothing
    about opt-in — a standing no-list ACL on a customer's Archive folder must not
    turn every unbuilt job into 'can't check from this PC'. Both private, left on
    the state like _resolve_unknown."""
    state["_cand_partial"] = False
    state["_arch_partial"] = False
    pd = state.get("pd_path")
    if not pd:
        return []
    # BOTH the current name ('Shipping Reports' / '<PN> Shipping Report.xlsx') and
    # the legacy one ('Received after testing' / '<PN> - Received after testing.xlsx')
    # so the 2026-09-15 rename MIGRATES: the merge folds an old file's rows into the
    # one true file, then it's removed like any other fully-merged stray copy.
    fnames = [Path(target).name]
    for sp in _report_pn_spellings(state):
        for fn in (_received_fname(sp), _received_legacy_fname(sp)):
            if fn.casefold() not in {f.casefold() for f in fnames}:
                fnames.append(fn)
    combos = tuple((sub, fn) for fn in fnames for sub in (RECEIVED_FOLDER, RECEIVED_FOLDER_LEGACY))
    out = []
    try:
        pd = Path(pd)
        locs = [pd]
        # a recorded Production Data folder that no longer EXISTS (rev rename,
        # archived job tree) is not a blip — the sibling pass below still covers
        # the customer folder; only a folder that is there but cannot be listed
        # makes the list partial
        try: gone = not pd.exists()
        except Exception: gone = False
        if not gone:
            locs += [d for d in pd.iterdir() if d.is_dir() and not _is_report_folder(d.name)]
        for loc in locs:
            for sub, fn in combos:
                out.append(loc / sub / fn)
    except Exception:
        state["_cand_partial"] = True   # keep whatever was enumerated before the blip
    # the SIBLING job trees (same predicate as scan_job_folders: Eng# as a whole
    # token AND the -900 or -970 P/N in the folder name — these paths reach the
    # harvest/DELETE pass, so a prefix-related Eng# must never match). Per-entry
    # errors `continue` — a blip here never loses the own-tree candidates above.
    try:
        ek = _norm_key(state.get("eng")); pk_al, ck_al = _tr_embed_keys(state.get("pn"))
        cust = Path(pd).parent.parent
        def _job_match(nm):
            return (_job_dir_is(nm, ek)
                    and (pk_al in _alnum(nm) or (ck_al and ck_al in _alnum(nm))))
        def _add_job(jd, flag="_cand_partial"):
            try:
                pdd = jd / "Production Data"
                if not pdd.is_dir():
                    return
                subs = [d for d in pdd.iterdir()
                        if d.is_dir() and not _is_report_folder(d.name)]
            except Exception:
                state[flag] = True
                return
            for loc in [pdd] + subs:
                for sub, fn in combos:
                    out.append(loc / sub / fn)
        if ek and len(pk_al) >= 6:
            for jd in cust.iterdir():
                try:
                    if not jd.is_dir():
                        continue
                    if jd.name.strip().lower() in _ARCHIVE_SEGS:
                        if include_archive:        # one level down, same predicate
                            try:                   # (its own flag: an archived copy never decides opt-in)
                                for ad in jd.iterdir():
                                    try:
                                        if ad.is_dir() and _job_match(ad.name):
                                            _add_job(ad, "_arch_partial")
                                    except Exception:
                                        state["_arch_partial"] = True
                                        continue
                            except Exception:
                                state["_arch_partial"] = True
                        continue
                    if _job_match(jd.name):
                        _add_job(jd)
                except Exception:
                    state["_cand_partial"] = True
                    continue
    except Exception:
        state["_cand_partial"] = True
    seen = set(); res = []
    for c in out:
        k = str(c).casefold()
        if k in seen:
            continue
        seen.add(k)
        if c != Path(target) and not _same_file(c, target):
            res.append(c)
    return res

def _delete_merged_copies(items, target):
    """Remove EXACTLY the given report copies — each one either had its FULL
    content carried into the one true file this save, or provably contains
    nothing a user could have typed (read-only-era app output). Anything the
    merge was not 100% sure about never reaches this list. Every item carries
    the (mtime, size) fingerprint taken WHEN IT WAS HARVESTED: a copy that
    changed since (someone saved into it mid-merge) is left alone — its new
    content merges on the next save instead of dying here. Best-effort."""
    import stat as _stat
    for c, sig in items:
        try:
            c = Path(c)
            if c == Path(target) or _same_file(c, target) or not c.exists():
                continue
            if sig is None or _file_sig(c) != sig:
                continue                    # changed since harvest: KEEP it
            try: os.chmod(str(c), _stat.S_IWRITE)   # read-only-era copies
            except Exception: pass
            c.unlink()
            try: c.parent.rmdir()                   # only succeeds when empty
            except Exception: pass
        except Exception:
            pass

def _is_read_only(path):
    """True when the file exists and carries the read-only attribute — the mark
    every pre-merge-era report file was saved with (those were app-regenerated
    and write-protected, so they cannot contain manual edits)."""
    try:
        return Path(path).exists() and not os.access(str(path), os.W_OK)
    except Exception:
        return False

def _xlsx_has_shapes(path):
    """True when the workbook carries drawn objects (text boxes, shapes,
    connectors) — content openpyxl silently DROPS on a load->save round trip
    (it only re-writes charts and images). Detected from the raw drawing XML so
    the merge can refuse / keep the file instead of destroying a user's drawn
    'DO NOT SHIP' note. Errors count as 'has shapes' (can't prove it's safe)."""
    import zipfile
    try:
        with zipfile.ZipFile(str(path)) as z:
            for nm in z.namelist():
                if nm.startswith("xl/drawings/") and nm.endswith(".xml"):
                    data = z.read(nm)
                    if b"<xdr:sp " in data or b"<xdr:sp>" in data or b"<xdr:cxnSp" in data:
                        return True
    except Exception:
        return True
    return False

def _file_sig(p):
    """(mtime_ns, size) fingerprint, or None when the file doesn't exist —
    compared around the merge to catch the file changing under us."""
    try:
        s = Path(p).stat()
        return (s.st_mtime_ns, s.st_size)
    except Exception:
        return None

def _xlsx_blockers(path):
    """The reason this workbook must NOT be re-saved by the merge (openpyxl
    would silently destroy that content), or None when it is safe: drawn
    objects (text boxes/shapes/connectors) and modern THREADED comments are
    dropped/degraded by a load->save round trip. Errors count as blocked."""
    import zipfile
    try:
        with zipfile.ZipFile(str(path)) as z:
            names = z.namelist()
            for nm in names:
                if nm.startswith("xl/drawings/") and nm.endswith(".xml"):
                    data = z.read(nm)
                    if b"<xdr:sp " in data or b"<xdr:sp>" in data or b"<xdr:cxnSp" in data:
                        return "a drawn object (a text box or shape)"
            for nm in names:
                if "threadedcomment" in nm.lower():
                    return ("a modern threaded comment — right-click that cell and "
                            "'Delete Comment' (an old-style Note is fine to keep)")
    except Exception:
        return "content that could not be scanned"
    return None

def _xlsx_has_annotations(path):
    """True when the workbook carries cell comments/notes, VML parts or
    hyperlink relationships — typed content the row-harvest cannot carry, so a
    STRAY copy holding any of it must never be deleted. Errors count as True."""
    import zipfile
    try:
        with zipfile.ZipFile(str(path)) as z:
            for nm in z.namelist():
                low = nm.lower()
                if (low.endswith(".vml") or "/comments" in low
                        or "threadedcomment" in low or low.startswith("xl/comments")):
                    return True
    except Exception:
        return True
    return False

def _acquire_received_lock(folder, target, timeout=8.0):
    """Cross-station lock (sidecar file, exclusive-create) held around the whole
    load->merge->replace of the report. Without it two stations saving at once
    each build from their own snapshot and the second os.replace silently throws
    away the first station's merge. Returns (lock_path, owner_token) or None
    when the share genuinely can't do exclusive-create (then we proceed unlocked
    — the sig checks and sig-verified stray deletion are the backstop).

    Correctness details (2026-09-02 review round 2):
      • a stale lock (untouched > 5 min; the holder refreshes its mtime between
        merge attempts) is broken by an ATOMIC RENAME — only one contender wins
        the rename, so two waiters can never both break-and-proceed;
      • release verifies the OWNER TOKEN, so a station whose lock was broken
        never deletes the next holder's lock;
      • transient create errors (delete-pending, share blip) are retried as
        contention — a single blip must not silently drop to unlocked mode;
      • the token write is best-effort: a failed write must not strand the
        just-created lock while its owner proceeds unlocked."""
    lock = Path(folder) / (Path(target).name + ".lock")
    token = (f"{os.environ.get('COMPUTERNAME', 'pc')}|{os.getpid()}|"
             f"{os.urandom(4).hex()}")
    deadline = time.time() + timeout
    misc_errors = 0
    while True:
        try:
            fd = os.open(str(lock), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                if time.time() - lock.stat().st_mtime > 300:
                    grave = lock.with_name(lock.name + f".stale-{os.getpid()}-{os.urandom(2).hex()}")
                    lock.rename(grave)     # atomic: exactly ONE contender wins
                    try: grave.unlink()
                    except Exception: pass
                    continue
            except Exception:
                pass
            if time.time() >= deadline:
                raise RuntimeError(
                    "another station is updating this file right now — try again in a "
                    "few seconds (the scanned boards ARE saved to history)")
            time.sleep(0.4)
            continue
        except OSError:
            misc_errors += 1
            if misc_errors >= 5:
                return None                # share can't do exclusive-create at all
            if time.time() >= deadline:
                raise RuntimeError(
                    "the file's lock could not be checked (share hiccup) — try again in "
                    "a few seconds (the scanned boards ARE saved to history)")
            time.sleep(0.4)
            continue
        try:
            os.write(fd, token.encode("utf-8", "replace"))
        except Exception:
            pass                           # best-effort: we still own the lock
        finally:
            try: os.close(fd)
            except Exception: pass
        return (lock, token)

def _refresh_received_lock(lk):
    """Touch the lock's mtime so long merges aren't mistaken for a crash."""
    if lk:
        try: os.utime(str(lk[0]), None)
        except Exception: pass

def _release_received_lock(lk):
    """Unlink the lock ONLY if it is still ours (or its token write failed and
    it is empty) — a broken-and-replaced lock belongs to another station now."""
    if not lk:
        return
    lock, token = lk
    try:
        try:
            content = lock.read_text(encoding="utf-8", errors="replace")
        except FileNotFoundError:
            return
        except Exception:
            return                         # can't verify ownership: leave it
        if content == token or content == "":
            lock.unlink()
    except Exception:
        pass

def _load_wb(path):
    """load_workbook keeping in-cell rich text when this openpyxl supports it
    (otherwise the merge's re-save would flatten a user's partly-bold cell)."""
    from openpyxl import load_workbook
    try:
        return load_workbook(str(path), rich_text=True)
    except TypeError:
        return load_workbook(str(path))

def _load_wb_values(path):
    """Cached-VALUES view of the same workbook (data_only): a user formula like
    =TODAY() must key/compare by the value Excel cached, not the formula string.
    None when the view can't be built — callers fall back to raw values."""
    from openpyxl import load_workbook
    try:
        return load_workbook(str(path), data_only=True)
    except Exception:
        return None

def _date_key(v):
    """Normalize a date-ish value (an Excel cell or a history string) to a
    comparable key: 'YYYY-MM-DD' when recognizable, else the trimmed casefolded
    string. Lets a datetime cell, '2026-08-21 10:00:00', '8/21/2026' and
    '8/21/26' all identify the same board row across merges; doubles as the
    general value normalizer for overlap comparison."""
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if not s:
        return ""
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        return f"{int(m.group(3)):04d}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2})(?!\d)", s)
    if m:
        return f"{2000 + int(m.group(3)):04d}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
    return s.casefold()

RCV_SHEET, COAT_SHEET = "Received from testing", "Coated & shipped"
IH_SHEET, OUT_SHEET = "In-House", "Still at customer"
RCV_COLS = ["#", "Serial Number", "Result", "Received", "Received By", "Testing WO", "Shipped", "Note"]
COAT_COLS = ["#", "Serial Number", "Coated WO", "Shipped", "Received from testing", "Note"]
# Texts the app itself writes into informational cells — the ONLY existing cell
# values the merge is allowed to improve on a later save (anything else in an
# existing row is treated as the user's and never touched).
_XLSX_PLACEHOLDERS = {"", "—", "sent without testing", "not shipped for testing",
                      "no receive record", "not in shipping records"}
# ('sent without testing' is the pre-2026-09-23 spelling — kept recognised so
#  existing cells upgrade to 'not shipped for testing' on their next save)

def write_received_xlsx(state):
    """MERGE the per-P/N workbook with history (Dhruvil 2026-09-02: "lets keep
    one file from now on... If we edit and save something in that file then keep
    it"). The file is WRITABLE. Two LOG sheets — 'Received from testing' (EVERY
    receive, whether or not a shipping record exists) and 'Coated & shipped'
    (EVERY coated shipment on record; a board with no testing record is included
    and annotated 'not shipped for testing') — keep every existing row untouched,
    manual rows and edits included; the app only APPENDS boards not in the sheet
    yet, and a board already present with DIFFERENT data is reported as an
    overlap in state['xlsx_conflicts'] (callers alert the user; the file's row
    always wins). Two STATUS sheets — 'In-House' and 'Still at customer' — stay
    derived views rebuilt from history on every save.

    Data-safety hardening (2026-09-02 adversarial review; "nothing the user
    typed may ever be silently destroyed"):
      • a cross-station LOCK plus a changed-under-us retry wrap the whole
        load->merge->replace, so two stations (or an Excel save landing
        mid-merge) can't silently clobber each other's rows;
      • a file whose 'Received from testing' sheet is missing/renamed is
        REFUSED with instructions, never rebuilt — only a genuinely READ-ONLY
        old-era file (which cannot hold manual edits) rebuilds fresh;
      • drawn objects (text boxes/shapes) that openpyxl cannot round-trip make
        the target refuse (with instructions); stray copies with them are kept;
      • a stray copy is deleted ONLY when its ENTIRE content provably merged
        (rows carried verbatim, zero conflicts, nothing outside the known
        footprint) — every doubt fails toward keeping the file on disk;
      • existing rows key/compare through Excel's CACHED values, so a user
        formula neither duplicates its row nor gets overwritten.
    Raises PermissionError when Excel physically blocks the swap (caller shows
    the retry dialog) — the History event is already saved by then."""
    import stat as _stat
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.cell.cell import MergedCell
    state["xlsx_conflicts"] = []
    if state.get("walk_error"):
        # Incomplete history walk — merging from partial data could mis-flag
        # overlaps and rebuild the status sheets half-empty. Caller retries.
        raise RuntimeError("history could not be read completely — file left unchanged")
    wo_path = state.get("wo_path")
    if not wo_path:
        # Distinguish 'could not CHECK the home from this PC' (the -900/-970 job
        # folders are unreadable here) from a job that genuinely has no work-order
        # folder yet. Telling the operator "no work-order folder known" when the
        # folder merely can't be reached is misleading — point them at a PC that can
        # open it instead (review: the unlistable-sibling refusal).
        if state.get("_home_unknown"):
            raise RuntimeError(
                "this job's work-order folders can't be checked from this PC. "
                "Nothing was written — the boards ARE saved to history. Run 🔁 Rebuild "
                "on a PC that can open this job's folders.")
        raise RuntimeError("no work-order folder known for this job yet")
    # Resolve to a REAL testing WO folder — a rev change may have renamed/moved
    # the recorded one, and mkdir(parents=True) on a stale string would build a
    # whole phantom folder tree (Dhruvil 2026-09-08, job 4092).
    real_dir = _resolve_received_wo_dir(state)
    if real_dir is None:
        head = ("the job's work-order folder could not be found on the share — the "
                "recorded path\n  " + str(wo_path) + "\ndoes not exist (a rev change may "
                "have renamed it) and ")
        if not state.get("allow_tree_switch") and not state.get("_resolve_unknown"):
            # a Save / after-close update never PLACES the report in the sibling
            # tree (only a human 🔁 Rebuild may; both update one already there,
            # and both bring a sibling copy HOME once the home tree has a live WO
            # again) — so this PC may well be able to open the job's folders, and
            # the sibling WO may even hold the report; say THAT, not "go to
            # another PC" (review round 2)
            raise RuntimeError(head +
                "the job's home tree holds no live work-order folder "
                "(a save never switches between the -900 or -970 trees). Nothing was written "
                "— the boards ARE saved to history. Use 🔁 Rebuild report on this PC: it may "
                "re-home the report to the sibling (-900 or -970) work-order folder.")
        raise RuntimeError(head +
            "no matching -900 or -970 work-order folder was found. "
            "Nothing was written — the boards ARE saved to history. Run 🔁 Rebuild "
            "on a PC that can open this job's folders.")
    wo_path = str(real_dir)
    state["wo_path"] = wo_path                    # keep caller/cleanup consistent
    try:
        state["pd_path"] = str(real_dir.parent)   # the home job's Production Data
    except Exception:
        pass
    folder = real_dir / RECEIVED_FOLDER           # real_dir exists → only the
    try: created_leaf = not folder.exists()       # (remembered: a REFUSED first build must not
    except Exception: created_leaf = False        #  leave an empty leaf behind — review round 5)
    folder.mkdir(parents=True, exist_ok=True)     # subfolder is created, never a tree
    # named by the -900 key for every test-and-return job, by the REAL coated
    # label for an in-house-tested one (Dhruvil 2026-09-23, Q1) — the same P/N
    # goes into the A1 header below
    rpn = _report_pn(state)
    target = received_xlsx_path(wo_path, rpn)
    summ = receiving_summary(state)
    shipped = state.get("shipped") or {}; received = state.get("received") or {}
    coated_map = state.get("coated") or {}
    pn = state["pn"]; cpn = _coat_ident_pn(pn)
    cur = {k: v for k, v in received.items() if _currently_received(state, k)}
    done_keys = {k for k, v in cur.items() if serial_seq(v["serial"]) in coated_map}
    inhouse = {k: v for k, v in cur.items() if k not in done_keys}
    out_count = sum(1 for k in shipped if k not in cur)

    ctr = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    bf = Font(bold=True, size=11); nf = Font(size=11)
    tf = Font(bold=True, size=14, color="1B4D35")
    note_f = Font(size=9, italic=True, color="5F5F5F")
    hdr_fill = PatternFill("solid", fgColor="E8F3ED")
    fail_font = Font(size=11, bold=True, color="C0392B")
    ph = {p.casefold() for p in _XLSX_PLACEHOLDERS}

    # ── computed LOG rows ("catch everything ... without thinking anything") ──
    rcv_hist = []
    # a Case-B board (Q3: a -900-labelled board shipped under the coated WO) is by
    # design never in `shipped`; a supervisor-overridden receive of one used to
    # read 'not in shipping records' on this sheet while the very next sheet lists
    # its coated shipment — point at that row instead (review round 6; cols 6-7
    # are FILL-only, so an existing placeholder cell upgrades and nothing alerts)
    caseb_rows = {}
    for e in (state.get("coated_log") or []):
        if _CASE_B_NOTE in str(e.get("note") or ""):
            kc = _serial_identity(e.get("serial"), cpn)
            if kc not in caseb_rows or str(e.get("ts") or "") > str(caseb_rows[kc].get("ts") or ""):
                caseb_rows[kc] = e
    for e in (state.get("receive_log") or []):
        sh = shipped.get(_serial_identity(e.get("serial"), pn)) or {}
        cb = {} if sh else (caseb_rows.get(_serial_identity(e.get("serial"), cpn)) or {})
        rcv_hist.append(["", str(e.get("serial") or ""),
                         "FAIL" if e.get("result") == "fail" else "PASS",
                         e.get("ts", ""), e.get("operator", ""),
                         sh.get("wo", "") or (f"{cb.get('wo', '')} (coated, -900 label)" if cb
                                              else "not in shipping records"),
                         sh.get("ts", "") or cb.get("ts", ""), str(e.get("note") or "")])
    recv_by_seq = {}
    for v in received.values():
        q = serial_seq(v.get("serial"))
        if q and q not in recv_by_seq:
            recv_by_seq[q] = v
    ship_seqs = {serial_seq(v.get("serial")) for v in shipped.values()}
    coat_hist = []
    for e in (state.get("coated_log") or []):
        q = serial_seq(e.get("serial"))
        rv = recv_by_seq.get(q)
        link = (str(rv.get("ts") or "") if rv
                else ("no receive record" if q and q in ship_seqs else "not shipped for testing"))
        coat_hist.append(["", str(e.get("serial") or ""), e.get("wo", ""), e.get("ts", ""), link,
                          str(e.get("note") or "")])
    # In-house-tested job (Dhruvil 2026-09-23): a job that never wore a -900 label
    # gets neutral wording — no "Shipped for testing: 0" row of zeros, the status
    # sheets say "not applicable", and the file opens on 'Coated & shipped'. The
    # header carries the report's NAME P/N (the REAL coated label for such a job,
    # Q1); a mixed job says how many coated boards never went out for testing.
    has_testing = summ["has_testing"]
    # (k_nt counts BOARDS the way 'Coated boards on record' does — by sequence,
    #  identity for a seq-less serial, review round 5 — and a board is 'not shipped
    #  for testing' only when EVERY coated row of it says so: a Case-B board
    #  (-900 label, coated WO) re-shipped with its -970 label is one board, not
    #  two identities, and the two numbers in A3 share one basis — review round 7)
    by_board = {}
    for r in coat_hist:
        by_board.setdefault(serial_seq(r[1]) or _serial_identity(r[1], cpn), []).append(r[4])
    k_nt = sum(1 for links in by_board.values() if all(l == "not shipped for testing" for l in links))

    def _write_head(ws, subtitle, log_sheet):
        note = ("log sheet — the app only APPENDS new boards here and never changes or "
                "removes existing rows; your edits are kept. Overlapping data raises an alert."
                if log_sheet else
                "STATUS VIEW — rebuilt from shipping history on every save; edits here are "
                "NOT kept (edit the log sheets instead).")
        if has_testing:
            totals = (f"Shipped for testing: {summ['shipped']}    ·    Still at customer: {summ['outstanding']}    ·    "
                      f"Received back: {summ['received']}  =  in-house {summ['in_house'] + summ['failed']} "
                      f"(awaiting coating {summ['in_house']}, failed {summ['failed']})"
                      f"  +  coated & shipped {summ['coated']}    ·    "
                      f"Coated boards on record: {summ['coated_all']}"
                      + (f" ({k_nt} not shipped for testing)" if k_nt else ""))
        else:
            totals = (f"In-house tested — not shipped for testing    ·    "
                      f"Coated boards shipped (P/N {_coated_label(state)}): {summ['coated_all']}")
        lines = [f"SHIPPING REPORT — P/N {rpn}",
                 f"Engineering # {state['eng']}    ·    {subtitle}",
                 totals,
                 f"Updated {datetime.datetime.now():%Y-%m-%d %H:%M} — {note}"]
        for i, txt in enumerate(lines, 1):
            try:
                cell = ws.cell(row=i, column=1)
                if isinstance(cell, MergedCell):
                    continue                 # user merged the title block: skip, don't crash
                cell.value = txt
            except Exception:
                continue
        try:
            ws["A1"].font = tf; ws["A2"].font = nf; ws["A3"].font = bf; ws["A4"].font = note_f
        except Exception:
            pass

    def _fit(ws, widen_only):
        try:
            for col in ws.columns:
                try:
                    mx = max((len(str(c.value)) for c in col if c.value), default=8)
                    want = min(max(mx + 3, 12), 60)
                    letter = col[0].column_letter
                    if not widen_only or want > (ws.column_dimensions[letter].width or 0):
                        ws.column_dimensions[letter].width = want
                except Exception:
                    continue
        except Exception:
            pass

    def _find_hdr(ws, name):
        for r in range(1, min(ws.max_row, 60) + 1):
            if str(ws.cell(row=r, column=1).value or "").strip() == "#":
                return r
        raise RuntimeError(
            f"the '{name}' sheet has lost its header row (the one starting with '#'). "
            f"Fix that sheet — or delete the file — then use 🔁 Rebuild report. "
            f"Nothing was changed.")

    def _build():
        """One complete load->merge pass; returns (wb, deletable, migrate_from,
        mig_sig). Runs inside the lock; may rerun when the file changed under us."""
        nonlocal summ, done_keys, inhouse   # re-derived from the merged coated sheet
        state["xlsx_conflicts"] = []
        conflicts = state["xlsx_conflicts"]
        deletable = []                         # [(path, sig-at-harvest)]
        migrate_from = None; mig_sig = None
        wb = wbv = None
        if target.exists():
            try:
                wb = _load_wb(target)          # the ONE true file: unreadable = hard stop
            except Exception as e:
                raise RuntimeError(
                    f"the existing file could not be read ({type(e).__name__}: {e}):\n"
                    f"{target}\nIt was NOT overwritten — nothing typed into it is lost. "
                    f"Close/repair (or move away) that file, then use 🔁 Rebuild report.")
            # the file at the target must be THIS job's: two jobs sharing a P/N
            # (4016 / 4016LF) name their reports identically, and another job's
            # copy dropped here by hand would have been adopted whole — its rows
            # merged among ours and its own 'Engineering #' head line rewritten
            # to ours, destroying the one mark that tells the two apart (review
            # round 5; bases and strays have refused such a copy since rounds 2-3)
            a2_t = _report_head_line(wb)
            if _head_names_eng(a2_t, state.get("eng")) is False:
                raise RuntimeError(
                    f"the file at\n{target}\nnames another job (Engineering # {_eng_from_head_line(a2_t)}) — "
                    f"nothing was changed; move that file away, then use 🔁 Rebuild report.")
            blocked = _xlsx_blockers(target)
            if blocked:
                raise RuntimeError(
                    f"the file contains {blocked}, and saving would silently destroy it. "
                    "Move that content into a normal cell, then try again — the file was "
                    f"NOT changed:\n{target}")
            wbv = _load_wb_values(target)
            if RCV_SHEET not in wb.sheetnames:
                if _is_read_only(target):
                    wb = wbv = None            # genuine read-only-era file: rebuild fresh
                else:
                    raise RuntimeError(
                        "the file's 'Received from testing' sheet is missing (renamed or "
                        "deleted?). To protect your edits nothing was changed — rename that "
                        "sheet back to 'Received from testing' (or move the file away), "
                        f"then try again:\n{target}")
        else:
            legacy = _find_legacy_received_file(state, target)
            if legacy is None and state.get("_cand_partial"):
                # ONE FILE: a FIRST-TIME build may only go ahead when every live
                # place a copy could sit in was listed. The candidate enumeration
                # behind _find_legacy_received_file was PARTIAL (a job folder under
                # the customer directory could not be listed from this PC) and
                # found no base — the report may well exist where we could not
                # look, and a fresh workbook here would be a second live file
                # (review round 5: the opt-in switch already answered 'could not
                # check' for this shape since round 4, but the human Rebuild's
                # write did not). The UPDATE path above (target exists) is
                # untouched; an unlistable Archive folder does not refuse
                # (_arch_partial — an archived copy is only a base, never the file).
                raise RuntimeError(
                    "could not confirm that no copy of this report exists elsewhere — a job "
                    "folder under the customer directory could not be listed from this PC. "
                    "Nothing was created (an existing report would have been updated). Try "
                    "again, or run 🔁 Rebuild on a PC that can open the job's folders.")
            # (a legacy base is loaded WHOLE, so legacy notes/hyperlinks survive —
            # only true blockers, which openpyxl can't round-trip, disqualify it)
            if legacy is not None and not _xlsx_blockers(legacy):
                lsig = _file_sig(legacy)
                try:
                    lwb = _load_wb(legacy)
                except Exception:
                    lwb = None                 # unreadable stray: kept on disk, fresh start
                if lwb is not None and RCV_SHEET in lwb.sheetnames:
                    wb = lwb; wbv = _load_wb_values(legacy)
                    migrate_from = legacy; mig_sig = lsig
                # any other legacy copy flows through the stray pass below, which
                # harvests what it safely can and keeps the file when in doubt
        if wb is None:
            wb = Workbook()
            for nm in list(wb.sheetnames):
                wb.remove(wb[nm])              # drop the default empty 'Sheet'
            wbv = None

        def _ctx(name, cols, ident_pn, date_ci):
            fresh = name not in wb.sheetnames
            if fresh:
                ws = wb.create_sheet(name)
                for _ in range(5): ws.append([])
                ws.append(cols); hdr = 6
                for c in ws[hdr]: c.font = bf; c.alignment = ctr; c.fill = hdr_fill
            else:
                ws = wb[name]; hdr = _find_hdr(ws, name)
                # Upgrade a header that predates a newly-added column (the 'Note'
                # column, 2026-09-16): fill only BLANK trailing header cells so an
                # existing file's col-8 note values get their label; never overwrite
                # a heading the user set.
                if hdr:
                    for ci in range(1, len(cols) + 1):
                        try:
                            hc = ws.cell(row=hdr, column=ci)
                            if not isinstance(hc, MergedCell) and str(hc.value or "").strip() == "":
                                hc.value = cols[ci - 1]
                                hc.font = bf; hc.alignment = ctr; hc.fill = hdr_fill
                        except Exception:
                            pass
            wsv = wbv[name] if (wbv is not None and name in wbv.sheetnames) else None

            def val(r, ci):
                # Excel's CACHED value first: a user formula must key/compare by
                # what it shows, not by its '=...' source text
                if wsv is not None:
                    try:
                        v = wsv.cell(row=r, column=ci).value
                        if v is not None:
                            return v
                    except Exception:
                        pass
                return ws.cell(row=r, column=ci).value

            ident = lambda s: _serial_identity(str(s or "").strip(), ident_pn)
            c_ = {"name": name, "cols": cols, "ws": ws, "wsv": wsv, "hdr": hdr,
                  "fresh": fresh, "val": val, "ident": ident, "date_ci": date_ci,
                  "existing": {}, "blank": {}, "nofill": set(),
                  "max_no": 0, "nrows": 0, "last": hdr}
            width = max(ws.max_column, len(cols))
            for r in range(hdr + 1, ws.max_row + 1):
                if any(ws.cell(row=r, column=k).value not in (None, "")
                       for k in range(1, width + 1)):
                    c_["last"] = r             # last row with real content: appends go
                                               # below DATA, not below styled-empty rows
                sn = val(r, 2)
                if sn is None or not str(sn).strip():
                    continue
                c_["nrows"] += 1
                dv = val(r, date_ci)
                if isinstance(dv, str) and dv.startswith("="):
                    dk = ""                    # uncached user formula: unknown date —
                    c_["nofill"].add(r)        # match by serial, never overwrite it
                else:
                    dk = _date_key(dv)
                c_["existing"].setdefault((ident(sn), dk), r)
                if not dk:
                    c_["blank"].setdefault(ident(sn), r)
                try: c_["max_no"] = max(c_["max_no"], int(val(r, 1)))
                except Exception: pass
            return c_

        def _writable_row(ws, r, n):
            try:
                return not any(isinstance(ws.cell(row=r, column=k), MergedCell)
                               for k in range(1, n + 1))
            except Exception:
                return False

        def _put_row(c_, vals, fail=False, note5=False):
            ws = c_["ws"]; n = len(vals)
            r = c_["last"] + 1
            guard = 0
            while not _writable_row(ws, r, n) and guard < 200:
                r += 1; guard += 1             # step past a user's merged-cell region
            for k, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=k, value=v)
                if isinstance(v, str) and v.startswith("="):
                    cell.data_type = "s"       # a carried-over '=...' STRING stays text
                cell.font = nf; cell.alignment = ctr
            if fail:
                ws.cell(row=r, column=3).font = fail_font
            if note5:
                ws.cell(row=r, column=5).font = note_f
            c_["last"] = max(c_["last"], r); c_["nrows"] += 1
            return r

        def _fill_cell(c_, r, ci, v):
            try:
                cell = c_["ws"].cell(row=r, column=ci)
                if isinstance(cell, MergedCell):
                    return
                raw = cell.value
                if isinstance(raw, str) and raw.startswith("="):
                    return                     # never overwrite a user formula
                cell.value = v
                if isinstance(v, str) and v.startswith("="):
                    cell.data_type = "s"       # a carried-over '=...' STRING stays text
                if c_["wsv"] is not None:
                    # write-through to the cached-values view, so later batches
                    # see the fill instead of the stale placeholder
                    try: c_["wsv"].cell(row=r, column=ci).value = v
                    except Exception: pass
            except Exception:
                pass

        def _reconcile(c_, key, row, cmp_cis, fill_cis, src):
            """Row index when the sheet already has this board's row (compared
            and placeholder-upgraded in place); None -> the caller appends it."""
            r = c_["existing"].get(key)
            if r is None and key[1]:
                # a row with a BLANK/unknown date is the same board waiting for
                # its date — but adopt it ONLY when the compared columns agree:
                # a DISPUTED record must stand as its own row, not overwrite the
                # user's blank-dated one (review round 2).
                r2 = c_["blank"].get(key[0])
                if r2 is not None:
                    compat = True
                    for ci in cmp_cis:
                        have = c_["val"](r2, ci)
                        hv = ("" if have is None else str(have).strip()).casefold()
                        iv = str(row[ci - 1] or "").strip().casefold()
                        if hv and hv not in ph and iv and iv not in ph \
                                and _date_key(have) != _date_key(row[ci - 1]):
                            compat = False; break
                    if compat:
                        r = r2
                        if r not in c_["nofill"]:
                            _fill_cell(c_, r, c_["date_ci"], row[c_["date_ci"] - 1])
                        c_["existing"][key] = r
                        c_["blank"].pop(key[0], None)
            if r is None:
                return None
            for ci, label in cmp_cis.items():
                iv = str(row[ci - 1] or "").strip()
                if not iv or iv.casefold() in ph:
                    continue                   # missing data is not disagreement
                have = c_["val"](r, ci)
                hv = ("" if have is None else str(have).strip()).casefold()
                if hv in ph:
                    _fill_cell(c_, r, ci, row[ci - 1])   # blank -> upgrade, no conflict
                    continue
                if _date_key(have) != _date_key(row[ci - 1]):
                    conflicts.append({"sheet": c_["name"], "serial": str(row[1]),
                                      "field": label, "file": str(have),
                                      "history": str(row[ci - 1]), "note": src})
            for ci in fill_cis:
                have = c_["val"](r, ci)
                hv = ("" if have is None else str(have).strip()).casefold()
                if str(row[ci - 1] or "").strip() and hv in ph \
                        and _date_key(have) != _date_key(row[ci - 1]):
                    _fill_cell(c_, r, ci, row[ci - 1])
            # NOTE column (the last one): backfill this occurrence's note onto a row
            # that has none yet — so a pre-notes row, or one whose note was added
            # later, shows it — but NEVER overwrite a note already there (a manual
            # edit carries forward). Empty-only fill, so it can't clobber user text.
            nci = len(c_["cols"])
            inc_note = str(row[nci - 1] or "").strip() if len(row) >= nci else ""
            if inc_note and not str(c_["val"](r, nci) or "").strip():
                _fill_cell(c_, r, nci, inc_note)
            return r

        def _merge_rows(c_, rows, src, cmp_cis, fill_cis):
            """Fold rows (file copies first, then history) into the sheet. A NEW
            (serial, date) key appends verbatim; an existing key only compares /
            upgrades placeholders. Returns (conflicts_raised, unmerged) — where
            `unmerged` counts key-matched rows whose NON-blank cells could not
            all be carried into the sheet (the source copy must then be kept)."""
            n0 = len(conflicts); unmerged = 0
            comp = {}
            for row in rows:
                row = list(row[:len(c_["cols"])])
                if len(row) < len(c_["cols"]):
                    row += [""] * (len(c_["cols"]) - len(row))
                if not str(row[1] or "").strip():
                    continue
                key = (c_["ident"](row[1]), _date_key(row[c_["date_ci"] - 1]))
                prev = comp.get(key)
                if prev is None:
                    comp[key] = row
                    continue
                for ci, label in cmp_cis.items():
                    pv = str(prev[ci - 1] or "").strip(); iv = str(row[ci - 1] or "").strip()
                    if not pv or pv.casefold() in ph or not iv or iv.casefold() in ph:
                        continue               # missing data is not disagreement
                    if _date_key(prev[ci - 1]) != _date_key(row[ci - 1]):
                        conflicts.append({"sheet": c_["name"], "serial": str(row[1]),
                                          "field": label, "file": str(prev[ci - 1]),
                                          "history": str(row[ci - 1]),
                                          "note": src + " — two records disagree"})
            order = {s: i for i, s in enumerate(sort_serials([str(x[1]) for x in comp.values()]))}
            for key, row in sorted(comp.items(),
                                   key=lambda kv: (order.get(str(kv[1][1]), 0),
                                                   _date_key(kv[1][c_["date_ci"] - 1]))):
                r = _reconcile(c_, key, row, cmp_cis, fill_cis, src)
                if r is not None:
                    # deletability audit: EVERY meaningful cell of this incoming
                    # row must now be represented in the sheet's row (edits in
                    # never-compared columns, e.g. 'Received By', count too)
                    for ci in range(2, len(c_["cols"]) + 1):
                        iv = str(row[ci - 1] or "").strip()
                        if not iv or iv.casefold() in ph:
                            continue           # blank/placeholder: nothing to lose
                        if _date_key(c_["val"](r, ci)) != _date_key(row[ci - 1]):
                            unmerged += 1
                            break
                    continue
                c_["max_no"] += 1
                vals = [c_["max_no"]] + list(row[1:])
                fail = (c_["name"] == RCV_SHEET
                        and str(row[2] or "").strip().upper() == "FAIL")
                note5 = (c_["name"] == COAT_SHEET
                         and str(row[4] or "").strip().casefold() in ph)
                r = _put_row(c_, vals, fail=fail, note5=note5)
                c_["existing"][key] = r
                if not key[1]:
                    # a just-appended blank-dated row must be visible to the
                    # blank-date fallback for the SAME save's later batches
                    c_["blank"].setdefault(key[0], r)
            return len(conflicts) - n0, unmerged

        def _purge_voided(name, vset, live_set, ident_pn, date_ci):
            """Drop LOG rows a supervisor VOIDED (Dhruvil 2026-09-16: voided the 106
            receive, re-added it as FAIL) so the corrected re-add merges CLEAN — but
            KEEP a row that matches a still-LIVE occurrence (the same-day corrected
            re-add, told apart by its compared value), so the live row is never
            churned on later saves. Runs BEFORE _ctx reads the sheet: indices valid."""
            if not vset or name not in wb.sheetnames:
                return
            ws_ = wb[name]
            hdr = _find_hdr(ws_, name)
            if not hdr:
                return
            drop = []
            for r in range(hdr + 1, ws_.max_row + 1):
                sn = ws_.cell(row=r, column=2).value
                if sn is None or not str(sn).strip():
                    continue
                ident = _serial_identity(str(sn), ident_pn)
                dk = _date_key(ws_.cell(row=r, column=date_ci).value)
                if (ident, dk) not in vset:
                    continue
                discr = str(ws_.cell(row=r, column=3).value or "").strip().casefold()
                if (ident, dk, discr) in live_set:
                    continue                  # the corrected LIVE row — keep it
                drop.append(r)
            # delete from BOTH the workbook AND its cached-values view (wbv): _ctx
            # reads existing rows through wbv, so a row left there would still be
            # seen by the merge even after it's gone from the saved file.
            wsv_ = wbv[name] if (wbv is not None and name in wbv.sheetnames) else None
            for r in reversed(drop):          # bottom-up: earlier indices stay valid
                try: ws_.delete_rows(r, 1)
                except Exception: pass
                if wsv_ is not None:
                    try: wsv_.delete_rows(r, 1)
                    except Exception: pass
        _live_rcv = {(_serial_identity(e.get("serial"), pn), _date_key(e.get("ts")),
                      ("fail" if e.get("result") == "fail" else "pass"))
                     for e in (state.get("receive_log") or [])}
        _live_coat = {(_serial_identity(e.get("serial"), cpn), _date_key(e.get("ts")),
                       str(e.get("wo") or "").strip().casefold())
                      for e in (state.get("coated_log") or [])}
        _purge_voided(RCV_SHEET, state.get("voided_rcv_rows") or set(), _live_rcv, pn, 4)
        _purge_voided(COAT_SHEET, state.get("voided_coat_rows") or set(), _live_coat, cpn, 4)

        rc = _ctx(RCV_SHEET, RCV_COLS, pn, 4)
        cc = _ctx(COAT_SHEET, COAT_COLS, cpn, 4)
        # stray copies FIRST — their rows are FILE data and carry over verbatim;
        # a copy may be deleted only when clean, fully carried, conflict-free AND
        # unchanged since harvest (the sig travels with it to deletion time)
        for cpath in _stale_received_candidates(state, target):
            try:
                if not cpath.exists() or (migrate_from is not None
                                          and (cpath == migrate_from
                                               or _same_file(cpath, migrate_from))):
                    continue
            except Exception:
                continue                       # can't even stat it: leave it alone
            ssig = _file_sig(cpath)
            got = _extract_received_rows(cpath, state.get("eng"))
            if got is None:
                continue                       # unreadable/unknown/another job's: keep it on disk
            d1, u1 = _merge_rows(rc, got["rcv"], "another copy of the file", {3: "Result"}, (6, 7))
            d2, u2 = _merge_rows(cc, got["coat"], "another copy of the file", {3: "Coated WO"}, (5,))
            if got["clean"] and d1 + d2 == 0 and u1 + u2 == 0 and ssig is not None \
                    and not _under_archive(cpath):
                deletable.append((cpath, ssig))
        _merge_rows(rc, rcv_hist, "shipping history", {3: "Result"}, (6, 7))
        _merge_rows(cc, coat_hist, "shipping history", {3: "Coated WO"}, (5,))
        # The append-only 'Coated & shipped' sheet is the AUTHORITATIVE record of
        # every coated shipment (Dhruvil 2026-10-06, 4096): a board coated only in
        # the COC folder (no in-app close) sits on this sheet but NOT in the history
        # 'coated' map, so a history-only save (auto-update / receive Save) wrongly
        # dropped it into In-House and under-counted "coated on record" (the sheet
        # kept 461 while the summary read 456). Re-derive the status from the MERGED
        # coated sheet so every save agrees with it — regardless of whether this run
        # folded a folder scan.
        _sheet_coat_seq = set()
        try:
            _ch = cc.get("hdr") or 6
            for _r in range(_ch + 1, cc["ws"].max_row + 1):
                _sv = cc["ws"].cell(row=_r, column=2).value
                _q = serial_seq(_sv) if _sv else ""
                if _q:
                    _sheet_coat_seq.add(_q)
        except Exception:
            pass
        if _sheet_coat_seq:
            _coat_seqs = set(coated_map.keys()) | _sheet_coat_seq
            done_keys = {k for k, v in cur.items() if serial_seq(v["serial"]) in _coat_seqs}
            inhouse = {k: v for k, v in cur.items() if k not in done_keys}
            summ = dict(summ)
            summ["coated_all"] = max(int(summ.get("coated_all", 0) or 0), len(_sheet_coat_seq))
            summ["coated"] = sum(1 for v in cur.values() if serial_seq(v["serial"]) in _coat_seqs)
            summ["in_house"] = summ["received"] - summ["coated"] - summ["failed"]
        if has_testing:
            rc_sub = f"EVERY board received back from testing — {rc['nrows']} row(s)"
            cc_sub = f"EVERY coated shipment on record (testing or not) — {cc['nrows']} row(s)"
        else:
            # an in-house-tested job's RCV sheet is normally empty; a supervisor-
            # overridden scan of a coated-form / Case-B board IS logged there without
            # making the job a testing job (Dhruvil 2026-09-23, decision (a)), so the
            # subtitle must not say "none" above such a row
            rc_sub = ("received back from customer testing — none (this job's boards are tested in-house)"
                      if not rc["nrows"] else
                      f"received back from customer testing — {rc['nrows']} row(s), supervisor-overridden scans "
                      f"(this job's boards are tested in-house)")
            cc_sub = f"EVERY coated shipment on record — {cc['nrows']} row(s) (boards tested in-house)"
        for c_, sub in ((rc, rc_sub), (cc, cc_sub)):
            if c_["fresh"] or str(c_["ws"].cell(row=1, column=1).value or "").upper().startswith(
                    ("SHIPPING REPORT", "RECEIVED AFTER TESTING")):
                _write_head(c_["ws"], sub, True)
            _fit(c_["ws"], widen_only=not c_["fresh"])

        # ── STATUS sheets: derived views, rebuilt every save ──────────────
        notes = receiving_anomaly_notes(state)
        ship_n = state.get("ship_counts") or {}
        board_notes = state.get("board_notes") or {}
        # NOTES that CARRY FORWARD (Dhruvil 2026-09-16): a note already in the file
        # — written by the app from a receive/ship event OR typed by hand in any
        # sheet — is preserved across rebuilds. Harvest serial->note from the merged
        # LOG sheets and the OLD status views BEFORE rebuilding; a board with no file
        # note falls back to its latest note from history (board_notes).
        carried = {}
        def _harvest_notes(sh):
            if sh is None: return
            hdr = None
            for r in range(1, min(sh.max_row, 60) + 1):
                if str(sh.cell(row=r, column=1).value or "").strip() == "#":
                    hdr = r; break
            if hdr is None: return
            note_ci = None
            for ci in range(1, sh.max_column + 1):
                if str(sh.cell(row=hdr, column=ci).value or "").strip().casefold() == "note":
                    note_ci = ci; break
            if note_ci is None: return
            for r in range(hdr + 1, sh.max_row + 1):
                sn = sh.cell(row=r, column=2).value
                nv = sh.cell(row=r, column=note_ci).value
                if sn and nv is not None and str(nv).strip():
                    carried[_serial_identity(str(sn), pn)] = str(nv).strip()   # last non-empty wins
        for cx in (rc, cc):
            try: _harvest_notes(cx["ws"])
            except Exception: pass
        for _nm in (IH_SHEET, OUT_SHEET):
            try:
                if _nm in wb.sheetnames: _harvest_notes(wb[_nm])
            except Exception: pass
        def _disp_note(k):
            base = carried.get(k) or board_notes.get(k) or ""
            an = notes.get(k, "")
            if an and an not in base:
                return (base + "  ·  " + an) if base else an
            return base
        for nm in (IH_SHEET, OUT_SHEET):
            if nm in wb.sheetnames:
                try: wb.remove(wb[nm])
                except Exception: pass
        ws = wb.create_sheet(IH_SHEET)
        for _ in range(5): ws.append([])
        ws.append(RCV_COLS)
        for c in ws[6]: c.font = bf; c.alignment = ctr; c.fill = hdr_fill
        for i, sn in enumerate(sort_serials([v["serial"] for v in inhouse.values()]), 1):
            k = _serial_identity(sn, pn)
            v = inhouse[k]
            sh = shipped.get(k) or {}
            res = "FAIL" if v.get("result") == "fail" else "PASS"
            dn = _disp_note(k)
            ws.append([i, v["serial"], res, v.get("ts", ""), v.get("operator", ""),
                       sh.get("wo", "") or ("—" if not sh else ""), sh.get("ts", ""),
                       dn])
            for c in ws[ws.max_row]: c.font = nf; c.alignment = ctr
            if res == "FAIL":
                ws.cell(row=ws.max_row, column=3).font = fail_font
            if dn:
                ws.cell(row=ws.max_row, column=8).font = note_f
            if not sh:
                ws.cell(row=ws.max_row, column=6, value="not in shipping records")
        if not inhouse:
            ws.append([])
            why = ("all received boards have been coated & shipped — see the 'Coated & shipped' sheet"
                   if done_keys else
                   ("no boards are in-house right now" if has_testing else
                    "this job's boards are tested in-house — nothing was shipped for testing"))
            ws.append(["", f"(none — {why})"])
            ws.cell(row=ws.max_row, column=2).font = Font(size=11, italic=True, color="5F5F5F")
            ws.cell(row=ws.max_row, column=2).alignment = left
        _write_head(ws, (f"IN-HOUSE — received back, not yet coated & shipped — {len(inhouse)} board(s)"
                         if has_testing else
                         f"IN-HOUSE — not applicable (nothing shipped for testing on this job) — {len(inhouse)} board(s)"),
                    False)
        _fit(ws, widen_only=False)

        ws2 = wb.create_sheet(OUT_SHEET)
        for _ in range(5): ws2.append([])
        ws2.append(["#", "Serial Number", "Testing WO", "Shipped", "Times shipped", "Note"])
        for c in ws2[6]: c.font = bf; c.alignment = ctr; c.fill = hdr_fill
        out_serials = [v["serial"] for k, v in shipped.items() if k not in cur]
        for i, sn in enumerate(sort_serials(out_serials), 1):
            k = _serial_identity(sn, pn)
            sh = shipped[k]
            n_sh = ship_n.get(k, 1)
            note = _disp_note(k)
            if not note and n_sh >= 2:
                note = "re-shipped for RE-TEST after failing"
            ws2.append([i, sh["serial"], sh.get("wo", ""), sh.get("ts", ""),
                        n_sh, note])
            for c in ws2[ws2.max_row]: c.font = nf; c.alignment = ctr
            if note:
                ws2.cell(row=ws2.max_row, column=6).font = note_f
        if not out_serials and not has_testing:
            ws2.append([])
            ws2.append(["", "(none — nothing was shipped for testing on this job; boards are tested in-house)"])
            ws2.cell(row=ws2.max_row, column=2).font = Font(size=11, italic=True, color="5F5F5F")
            ws2.cell(row=ws2.max_row, column=2).alignment = left
        _write_head(ws2, (f"shipped for testing, NOT yet received back — {out_count} board(s)"
                          if has_testing else
                          f"Still at customer — not applicable (nothing shipped for testing on this job) — {out_count} board(s)"),
                    False)
        _fit(ws2, widen_only=False)

        order2 = {RCV_SHEET: 0, COAT_SHEET: 1, IH_SHEET: 2, OUT_SHEET: 3}
        try:
            wb._sheets.sort(key=lambda w: order2.get(w.title, len(order2)))
            # openpyxl keeps the tabSelected flag of whichever sheet the user last
            # had active in Excel and only moves activeTab: the saved file then
            # opened with TWO sheets selected ('[Group]' — a note typed on one
            # sheet lands on the other's LOG row too). Exactly one selected tab,
            # the active one (review round 7).
            idx = 0 if has_testing else 1        # in-house-tested: open on 'Coated & shipped'
            for w in wb.worksheets:
                try: w.sheet_view.tabSelected = False
                except Exception: pass
            wb.active = idx
            wb.worksheets[idx].sheet_view.tabSelected = True
        except Exception:
            pass
        return wb, deletable, migrate_from, mig_sig

    # ── save under the cross-station lock, redoing the merge if the file
    #    changes under us (an Excel save landing mid-merge must never be
    #    clobbered by a snapshot taken before it) ────────────────────────────
    lock = _acquire_received_lock(folder, target)
    try:
        station = re.sub(r'[^0-9A-Za-z-]', '', os.environ.get("COMPUTERNAME", "pc")) or "pc"
        tmp = folder / f"~{station}_{os.getpid()}_{target.name}.tmp"
        for _attempt in range(3):
            _refresh_received_lock(lock)   # long merges must not look crashed
            sig0 = _file_sig(target)
            wb, deletable, migrate_from, mig_sig = _build()
            try:
                wb.save(str(tmp))          # a failed save must not strand a partial tmp
                try: tmp_size = tmp.stat().st_size
                except Exception: tmp_size = None
                if _file_sig(target) != sig0:
                    try: tmp.unlink()
                    except Exception: pass
                    continue               # redo the merge from the fresh file
                if target.exists():
                    try: os.chmod(str(target), _stat.S_IWRITE)   # legacy read-only attr
                    except Exception: pass
                os.replace(str(tmp), str(target))
            except Exception:
                try: tmp.unlink()
                except Exception: pass
                raise
            if migrate_from is not None and not _under_archive(migrate_from):
                deletable.append((migrate_from, mig_sig))   # base's FULL content carried
                # (an ARCHIVED base is a frozen snapshot: read once, never deleted)
            # Delete merged copies only when OUR merge is provably the live file:
            # if another writer landed right behind us, it never saw these copies'
            # rows — they must stay on disk and merge into ITS next save instead.
            now_sig = _file_sig(target)
            if tmp_size is not None and now_sig is not None and now_sig[1] == tmp_size:
                _delete_merged_copies(deletable, target)
            return str(target)
        raise RuntimeError("the file kept changing while saving (someone is editing it "
                           "in Excel right now) — try again in a moment")
    finally:
        _release_received_lock(lock)
        if created_leaf:
            # the leaf was created by THIS call: after a refusal it is empty and goes
            # again (rmdir only ever succeeds on an empty folder — a file written here
            # by us or by another station keeps it)
            try: folder.rmdir()
            except Exception: pass

def auto_update_shipping_report(eng, pn):
    """Refresh a job's Shipping Report right after a shipment closes and REPORT the
    outcome (Dhruvil 2026-09-15 it should auto-update; 2026-10-24 the silent daemon
    thread didn't reliably run, so this now runs SYNCHRONOUSLY in the close flow and
    RETURNS a result the close summary shows + History records). HISTORY-ONLY (fast;
    the folder-truth scan stays on 🔁 Rebuild). Never raises.

    Returns None when there is nothing to say (the close isn't a test-and-return
    job, so there is no report for it), else a (status, detail) pair:
      ('updated', path)    the report was rewritten;
      ('none',    msg)     a test-and-return job with NO report yet — OPT-IN
                           (Dhruvil 2026-09-23: a report is created only by a manual
                           first Build; auto-update only UPDATES one);
      ('unknown', msg)     couldn't read this job's history / reach its folders from
                           this PC — nothing was written;
      ('failed',  msg)     the write was attempted and failed (file open elsewhere,
                           share error).
    The report is keyed by the canonical TESTING (…900) P/N resolved from
    list_testing_jobs so the SAME file is updated; a coated-only job is reached via
    the entry that list appends for it (one entry per Eng#+core since review round 3)."""
    try:
        ekey = _norm_key(eng)
        m = re.search(r'(\d{6,7})(?:900|970)', _alnum(pn))
        if not ekey or not m:
            return None
        core = m.group(1)
        jobs = [j for j in list_testing_jobs()
                if _norm_key(j.get("eng")) == ekey and core in _alnum(j.get("pn"))]
        if not jobs:
            return None
        last = None
        for job in jobs:
            st = build_receiving_state(job["eng"], job["pn"])
            if not st or st.get("walk_error"):
                last = ("unknown", "couldn't read this job's history from this PC"); continue
            if not st.get("wo_path"):
                last = ("unknown", "the job's work-order folder isn't reachable from this PC"); continue
            exists, _p = _report_exists(st)
            if exists is None:
                last = ("unknown", "couldn't reach this job's folders from this PC"); continue
            if not exists:
                last = ("none", "no Shipping Report yet — build it once in Receiving to start auto-updates"); continue
            try:
                out = write_received_xlsx(st)
                return ("updated", str(out))
            except PermissionError:
                return ("failed", "the report file is open on another PC — it'll update on the next close or 🔁 Rebuild")
            except Exception as e:
                return ("failed", f"{type(e).__name__}: {e}")
        return last
    except Exception as e:
        return ("failed", f"{type(e).__name__}: {e}")

def _record_report_update(hist_path, result):
    """Stamp the just-written close event with the Shipping Report auto-update
    outcome (Dhruvil 2026-10-24: record it in History). Best-effort; never raises."""
    try:
        if not hist_path or not result:
            return
        p = Path(hist_path)
        rec = json.loads(p.read_text(encoding="utf-8"))
        rec["report_update"] = {"status": result[0], "detail": result[1],
                                "ts": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
        _atomic_write_json(p, rec)
    except Exception:
        pass

def _report_exists(state):
    """OPT-IN switch (Dhruvil 2026-09-23): does this job's Shipping Report already
    exist? -> (answer, path). answer is TRI-STATE (review round 1):
      True   the file is at its current home (resolved READ-ONLY: no tree switch,
             nothing created) — path = that target — OR a legacy/stale copy the
             merge would fold in exists in the own or SIBLING tree (path = the
             current target when the home WO is live, else the copy that was
             actually found, so nothing ever points at a folder that does not exist);
      False  the folders were listable from this PC and hold no report — the job
             is NOT opted in;
      None   could not check: this PC cannot list the job's folders (no share
             access, unreachable path) or a filesystem call raised. Callers treat
             it like False for WRITING (nothing is created or touched) but must
             never tell the operator "no report exists".
    Archive/Old copies do NOT count: a report the operator deleted to opt out must
    stay deleted (R6) even when an older rev's frozen copy sits under Archive —
    that copy is a merge BASE for a manual Build only (_find_legacy_received_file);
    _report_lookup still NAMES it as `found` so the UI can say 'archived' rather
    than 'no report' (review round 4). A state with no wo_path is (False, None),
    or (None, None) when build_receiving_state could not CHECK the coated home
    (state['_home_unknown'], review round 5) — 'could not check', not 'no
    work-order folder known'.
    The only code paths that may CREATE a report are the manual 🔁 Build report /
    Rebuild ALL picker. Callers that also need the COPY the answer rests on use
    _report_lookup."""
    has, path, _found = _report_lookup(state)
    return has, path

def _report_lookup(state):
    """_report_exists plus WHERE the answer came from: (answer, path, found) —
    `found` is the existing file that made the answer True (the target itself, or
    the legacy/stray copy in another folder that the merge would fold in), None
    otherwise. The chips name it when it is not the target, so an operator opting
    a job OUT (R6: delete the report) can see which older copy still keeps the job
    opted in — a copy the merge keeps on disk forever because it is not provably
    deletable (a user-added sheet, a drawn text box) re-opts the job in on the
    next save until it is deleted too (review round 2; R1 as written counts any
    candidate the merge would fold in — flagged for Dhruvil).
    Review round 4: (1) a state with NO wo_path answers (False, None, None) — every
    writer and 📄 Open file gate on wo_path, so the switch must not say 'report
    exists' about a folder resolved from pd_path alone (a receive event's
    production_data after the job's only -900 record was voided); (2) a PARTIAL
    candidate enumeration (share blip, an unlistable job folder) with nothing found
    answers None, never False — a first-time Build offered on that False would
    have forked the report; (3) an ARCHIVE-only copy keeps the answer False (R6:
    not opted in) but is returned as `found`, so the chips / picker / Save popup
    can say 'exists only under Archive — 🔁 Rebuild report re-homes it' instead
    of 'no report yet' (a manual Build does use that copy as its merge base).
    Review round 5: (4) a state with no wo_path whose coated home could not be
    CHECKED (state['_home_unknown'], set by build_receiving_state when the
    'file is the switch' probe could not see the folders) answers (None, None,
    None) — 'could not check', not 'no work-order folder known'; (5) a file at
    the target whose own head line names ANOTHER job (4016 / 4016LF) is not this
    job's report: it is passed over like any other such copy and the rest of the
    enumeration decides (the write refuses it too); (6) an unlistable Archive/Old
    folder (state['_arch_partial']) does not make the answer None — an archived
    copy never decides opt-in.
    Review round 6: (7) when that foreign file sits at the target and the
    enumeration then finds one of OUR copies elsewhere, `path` is OUR copy, never
    the target: the chips printed '📄 Report: <foreign file>' and 📄 Open file
    opened the other job's workbook as this job's report, while every writer
    (rightly) refused. The blocker is left on the given state as
    state['_report_blocked_by'] = (target, its Eng#) — a cache key like the
    others, popped when nothing blocks — so the chips can say what to move away.
    Review round 7: (8) with that foreign file at the target and NO copy of our
    own anywhere the answer is (False, None, archived) — never (False, <the
    foreign file>): 📄 Open file trusted the path and opened the other job's
    workbook as this job's report, while the chips / Save tail / 'Build report?'
    promised a first build the writer was bound to refuse. The blocker stays on
    the state for every consumer to name. (9) 'names another job' is decided by
    _head_names_eng (whole tokens, an Eng# may have several — '4044 LF')."""
    state.pop("_report_blocked_by", None)
    if not state.get("wo_path"):
        if state.get("_home_unknown"):
            return None, None, None
        return False, None, None
    try:
        st = dict(state)
        st.pop("allow_tree_switch", None)           # read-only look: never switch trees
        d = _resolve_received_wo_dir(st)
        unknown = d is None and bool(st.get("_resolve_unknown"))
        base = d or state.get("wo_path")
        target = received_xlsx_path(base, _report_pn(state))
    except Exception:
        return None, None, None
    blocked = False
    try:
        if target.exists():
            a2_t = _report_head_line_at(target)
            # ours unless the head line names ANOTHER job (whole-token, multi-
            # token compare — review round 7); no head line / a user-typed first
            # token -> undecidable -> the file at the target is this job's report
            blocked = _head_names_eng(a2_t, state.get("eng")) is False
            if not blocked:
                return True, target, target
            # another job's same-named file sits AT our target (4016 / 4016LF):
            # not this job's report — the enumeration decides, whatever it finds
            # is the path (never that file), and the blocker is recorded for the
            # chips (review rounds 5 + 6)
            state["_report_blocked_by"] = (target, _eng_from_head_line(a2_t))
        if d is not None:
            st["wo_path"] = str(d)
            try: st["pd_path"] = str(d.parent)
            except Exception: pass
        # ONE enumeration, archive included; live copies decide, an archived
        # copy (the customer folder's own Archive/Old subfolder — the only place
        # the enumeration descends into) is only reported, never counts
        try: cust = Path(st.get("pd_path")).parent.parent
        except Exception: cust = None
        def _arch(c):
            try:
                rel = c.relative_to(cust).parts
                return bool(rel) and rel[0].strip().lower() in _ARCHIVE_SEGS
            except Exception:
                return _under_archive(c)
        archived = None
        cands = _stale_received_candidates(st, target, include_archive=True)
        partial = bool(st.get("_cand_partial"))
        for c in cands:
            if not c.exists():
                continue
            if _copy_names_other_job(c, state.get("eng")):
                # another job's same-named copy sitting in our tree (4016 /
                # 4016LF): the merge refuses it as base and never deletes it,
                # so it must not opt THIS job in either — or an R6 opt-out
                # delete is undone by the next after-close update (review
                # round 3). One workbook load, only when the target is gone.
                continue
            if _arch(c):
                if archived is None:
                    archived = c
                continue
            return True, (target if d is not None and not blocked else c), c
    except Exception:
        return None, target, None
    if unknown or partial:
        return None, target, None      # could not check — a live copy may sit where we could not look
    # a False answer never hands out the OTHER job's file as this job's path:
    # 📄 Open file opened it as this job's report and the chips promised a Build
    # the writer was bound to refuse (review round 7) — the blocker on the state
    # says what sits there, path stays None until it is moved away
    return False, (None if blocked else target), archived

_REPORT_CACHE_KEYS = ("_report_exists", "_report_path", "_report_at_home", "_report_found_at",
                      "_report_blocked_by")

def _annotate_report_exists(state):
    """Cache the opt-in answer ON the state (Dhruvil 2026-09-23, R4): worker
    threads call this right after build_receiving_state, so the Receiving window
    never has to touch the share from the Tk thread to know whether a job has a
    report. Keys: state['_report_exists'] (True / False / None = could not check),
    state['_report_path'] (Path or None — None too when another job's file sits
    at the target and no copy of ours exists, review round 7), state['_report_at_home'] (True when the
    answer rests on the file AT that path; False = the report is only a legacy /
    stray copy elsewhere that the next Save / after-shipment update (or 🔁 Rebuild
    report) brings home), state['_report_found_at'] (the copy the answer rests
    on — see _report_lookup; with a False answer it is the ARCHIVE-only copy a
    manual Rebuild would re-home, else None) and state['_report_blocked_by']
    (set by _report_lookup: (target, Eng#) of another job's file sitting at this
    job's target; absent otherwise). Recomputed on every fresh walk; never
    raises. _mark_report_written / _forget_report_cache are the only other
    writers of these keys."""
    try:
        has, path, found = _report_lookup(state)
    except Exception:
        has, path, found = None, None, None
        state.pop("_report_blocked_by", None)
    state["_report_exists"] = None if has is None else bool(has)
    state["_report_path"] = path
    at_home = None
    if has and path is not None:
        # from the lookup's own evidence, never a fresh stat of the path: with
        # another job's same-named file sitting at the target a stat said 'at
        # home' about a file that is not this job's report (review round 6)
        try: at_home = found is not None and Path(found) == Path(path)
        except Exception: at_home = None
    state["_report_at_home"] = at_home
    state["_report_found_at"] = found
    return state

def _mark_report_written(state, out):
    """The cache after a SUCCESSFUL write_received_xlsx: the report exists, at
    `out`, at home (it was just verified-written there — a legacy/stray copy the
    walk had found was merged in). One writer for all the keys, so the chips can
    never keep saying 'exists under an older name or folder' about the file the
    action just brought home (review round 2)."""
    state["_report_exists"] = True
    state["_report_path"] = Path(out)
    state["_report_at_home"] = True
    state["_report_found_at"] = Path(out)
    state.pop("_report_blocked_by", None)   # a write succeeded: nothing sits in the way
    return state

def _forget_report_cache(state):
    """Drop the cached opt-in answer so the next reader recomputes it from disk
    (_receive_save_gate does when the key is absent). Used when a Save's pre-write
    re-walk FAILED and the window kept its older state: that state's answer dates
    from the job pick, and a report deleted since (opt-out, R6) must not be
    re-created from it — nor an existing one skipped (review round 2)."""
    for k in _REPORT_CACHE_KEYS:
        state.pop(k, None)
    return state

def _receive_save_gate(state):
    """The ONE decision a receiving save makes about Excel (Dhruvil 2026-09-23,
    R3 — UPDATE-ONLY): True = an existing report is refreshed, False = the job has
    no report and nothing is written, None = could not check from this PC (nothing
    written either, but the operator is told THAT, not "no report"). Uses the
    answer cached by the walk; computes it when a hand-built state has none.
    Tk-free so the tests exercise the real gate."""
    if "_report_exists" not in state:
        _annotate_report_exists(state)
    return state.get("_report_exists")

def _last_shipment_date(state):
    """'YYYY-MM-DD' of the newest SHIPMENT a walked state holds — closes and
    imports, testing or coated side — for the picker's 'Last shipment' column
    (review round 4: a list_testing_jobs entry's own ts is its newest event of
    ANY kind, a receiving save included, and read as a shipment date it made an
    old-flow job look current). '' when no shipment is on record. Never raises."""
    best = ""
    try:
        for ts in ([v.get("ts") for v in (state.get("shipped") or {}).values()]
                   + [e.get("ts") for e in (state.get("coated_log") or [])]):
            d = _date_key(ts)
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", d) and d > best:
                best = d
    except Exception:
        pass
    return best

def _fold_receive_locally(state, kept, operator, rdate, ts):
    """_do_save's display fallback after the History event is written: fold this
    session's boards into the window's state so the tracker (and, if the re-walk
    fails, the xlsx written from this state) already shows them. Mirrors the
    walk's receive branch (review round 4): EVERY kept board lands in receive_log
    (the RCV sheet lists supervisor-overridden scans too), but a COATED-form
    serial (…970… embedded) stays out of `received` / recv_dates exactly as
    build_receiving_state keeps it out — otherwise a failed re-walk after an
    overridden coated scan on an in-house-tested job flipped has_testing and the
    file's header/subtitles/active sheet to testing wording until the next
    successful walk (decision (a)). Tk-free; `kept` = [{'serial', 'result'}]."""
    ck_al = _tr_embed_keys(state.get("pn"))[1]      # the walk's own coated key (review round 6)
    pn = state.get("pn", "")
    for e in kept:
        entry = {"serial": e["serial"], "ts": rdate or ts,
                 "sts": (rdate + " 00:00:00") if rdate else ts,
                 "operator": operator, "result": e.get("result", "pass")}
        # …into the LOG: if the re-walk fails, the xlsx is written from this
        # state and the log sheet must still carry the new boards.
        state.setdefault("receive_log", []).append(dict(entry))
        if _embeds(ck_al, e["serial"]):
            continue                    # coated-form: logged, never a testing receive
        k = _serial_identity(e["serial"], pn)
        state.setdefault("received", {})[k] = entry
        # …and the RECEIVE COUNT: the count-based cycle check would otherwise
        # call a just-received re-test board 'still at customer' until the
        # next successful walk (review round 3).
        state.setdefault("recv_dates", {}).setdefault(k, set()).add((rdate or ts)[:10])
    return state

def _collect_report_rows(box, jobs=None):
    """🔁 Rebuild ALL, step 1 (Dhruvil 2026-09-23, OPT-IN): one history walk per
    job and the read-only opt-in check, filled into `box` for the picker. Tk-free
    — runs on a worker thread; the UI thread only reads the box. Row shape:
      {job, st, has, path, found, last_ts, coated_only, status, tickable}
    status is one of 'has report' / 'no report' / 'report archived — rebuild
    re-homes it' (the only copy sits under Archive: not opted in, so `has` is
    False and the row is not pre-ticked, but a Rebuild uses it as the base —
    review round 4) / 'no WO folder known' / 'history unreadable' / "can't check
    from this PC" (the opt-in check could not see the job's folders) / 'no
    -900/-970 work-order folder found' (folders listable but no live WO in either
    tree — a write would be refused, so the row is not offered; 'report found —
    …' when a copy exists somewhere all the same) / "another job's file at its
    report location" (a foreign same-named file at the target and no copy of ours
    — the write would refuse it, review round 7); only the first three can be
    ticked. A job without wo_path whose coated home could not be CHECKED from
    this PC (_home_unknown) reads "can't check from this PC" too, not 'no WO
    folder known' (review round 5). last_ts is the newest SHIPMENT on record (a
    receive is not one — review round 4; blank when nothing shipped, review round
    5), the entry's newest event only when the walk failed. One job's failure
    never stops the rest (its row shows 'history unreadable', disabled)."""
    box.setdefault("rows", []); box.setdefault("n", 0); box.setdefault("i", 0)
    box.setdefault("cur", "")
    try:
        if jobs is None:
            jobs = list_testing_jobs()
        box["n"] = len(jobs)
        for idx, j in enumerate(jobs, 1):
            box["i"] = idx
            box["cur"] = f"{j.get('eng')} · {_job_display_pn(j)}"
            row = {"job": j, "st": None, "has": False, "path": None, "found": None,
                   "last_ts": str(j.get("ts") or "")[:10],
                   "coated_only": not j.get("testing", True),
                   "status": "history unreadable", "tickable": False}
            try:
                st = build_receiving_state(j["eng"], j["pn"])
                if st and not st.get("walk_error"):
                    row["st"] = st
                    row["last_ts"] = _last_shipment_date(st)     # blank = nothing shipped on record
                    if st.get("wo_path"):
                        has, path, found = _report_lookup(st)
                        if has is None:
                            row.update(path=path, status="can't check from this PC")
                        else:
                            # would the write even find a home? Same read-only resolve
                            # the rebuild will do (sibling tree allowed) — a row with
                            # no live WO in either tree would only fail, every time
                            live = _resolve_received_wo_dir(dict(st, allow_tree_switch=True))
                            if live is None:
                                row.update(has=bool(has), path=path, found=found,
                                           status=("report found — no live -900/-970 work-order folder"
                                                   if (has or found) else "no -900/-970 work-order folder found"))
                            elif not has and st.get("_report_blocked_by"):
                                # another job's file at the target, no copy of ours: the
                                # write would refuse it — not offered (review round 7)
                                row.update(has=False, path=path, found=found,
                                           status="another job's file at its report location")
                            else:
                                row.update(has=bool(has), path=path, found=found, tickable=True,
                                           status=("has report" if has else
                                                   ("report archived — rebuild re-homes it" if found
                                                    else "no report")))
                    elif st.get("_home_unknown"):
                        row["status"] = "can't check from this PC"   # the coated home could not be probed
                    else:
                        row["status"] = "no WO folder known"
            except Exception:
                pass
            box["rows"].append(row)
    except Exception as e:
        box["error"] = f"{type(e).__name__}: {e}"
    finally:
        box["done"] = True
    return box["rows"]

def _rebuild_report_rows(rows, box):
    """🔁 Rebuild ALL, step 3: the full rebuild (folder truth + write) for the
    TICKED picker rows only. Each ticked job is walked AGAIN right before its
    write (review round 1): a close on another station while the picker was open
    was skipped by its own after-close update (no file yet) and would otherwise
    be missing from the report the picker creates; the step-1 state is the
    fallback when that walk fails. The (slow, share-bound) folder scan runs
    BEFORE that walk and is folded into the fresh state, so a close landing
    during the scan is in the file too (review round 3). A manual action, so
    the cross-tree move is allowed (allow_tree_switch). Counters in `box`: ok /
    created / folders / coated_only / moved / failed (each failure keeps the
    refusal's first line — the picker is how a coated-only job gets its FIRST
    report, and 'RuntimeError' alone told the operator nothing). Nothing happens
    to a row that isn't passed in — no file, no folder. Tk-free; runs on a
    worker thread."""
    for k in ("ok", "created", "folders", "coated_only", "moved", "i"):
        box.setdefault(k, 0)
    box.setdefault("failed", []); box["n"] = len(rows); box.setdefault("cur", "")
    try:
        for idx, row in enumerate(rows, 1):
            j = row.get("job") or {}
            box["i"] = idx
            box["cur"] = f"{j.get('eng')} · {_job_display_pn(j)}"
            st = row.get("st")
            try:
                if not st or st.get("walk_error") or not st.get("wo_path"):
                    box["failed"].append(f"{_job_display_pn(j)}: no state / no work-order folder"); continue
                try:      # folder truth first — it only needs the job's identity/folders
                    scan = scan_job_folders(st)
                    if not (scan and scan.get("ok")):
                        scan = None
                except Exception:
                    scan = None
                try:      # fresh walk: shipments closed while the picker was open / the scan ran
                    fresh = build_receiving_state(j.get("eng"), j.get("pn"))
                    if fresh and not fresh.get("walk_error") and fresh.get("wo_path"):
                        st = fresh; row["st"] = st
                except Exception:
                    pass
                if scan is not None:
                    try:
                        fold_folder_scan(st, scan); box["folders"] += 1
                    except Exception:
                        pass
                st["allow_tree_switch"] = True
                write_received_xlsx(st); box["ok"] += 1
                # a FIRST-TIME build is one with no copy anywhere (not even under
                # Archive); a copy that existed and changed trees was MOVED, a
                # first build into the sibling tree was merely placed (review round 4)
                had = bool(row.get("has") or row.get("found"))
                if not had: box["created"] += 1
                if row.get("coated_only"): box["coated_only"] += 1
                if st.get("tree_switched") and had: box["moved"] += 1
            except Exception as e:
                why = (str(e).strip().splitlines() or [""])[0][:140]
                box["failed"].append(f"{_job_display_pn(j)}: {type(e).__name__}" + (f": {why}" if why else ""))
            finally:
                if isinstance(st, dict): st.pop("allow_tree_switch", None)
    finally:
        box["done"] = True
    return box

def _backfill_serials_look_coated(pn, serials):
    """BackfillDialog guard (Dhruvil 2026-09-23, plan change 10): True when EVERY
    scanned serial embeds the SIBLING P/N of the typed one and none embeds the
    typed one — a -900 typed with …970 serials is the operator about to import
    coated shipments as 'shipped for testing' (the 4065-style mislabel that also
    plants a -970 folder as a testing anchor); a coated label typed (auto-filled
    from a -970 WO folder, rev letter included) with …900-form serials is the
    in-progress batch, which History then files as COATED shipments (Case B).
    Core-DIGIT based via _tr_pn_parts, so the letter after the 970 no longer
    blinds it (review round 2). Pure; never raises."""
    try:
        core, side, _ = _tr_pn_parts(pn)
        sers = [s for s in (serials or []) if str(s).strip()]
        if not sers or not core or not side:
            return False
        own = core + side; sib = core + ("900" if side == "970" else "970")
        return (all(_embeds(sib, s) for s in sers)
                and not any(_embeds(own, s) for s in sers))
    except Exception:
        return False

def _find_legacy_received_file(state, target):
    """When the newest testing WO changes, the report MOVES with it — but 'keep
    one file': the previous location's copy is the merge BASE (its content and
    manual edits carry over into the new location) and is deleted only after the
    new file saves successfully. Picks the most recently modified LIVE copy; an
    archived copy is the base only when no live copy exists at all (review round
    3: a frozen old-rev snapshot someone had opened in Excel outranked the live
    previous-WO copy by mtime — its stale rows came back, the operator's sheet was
    stranded and two live files remained). Archived copies are candidates HERE
    only (never deleted)."""
    best = None; best_key = None
    for c in _stale_received_candidates(state, target, include_archive=True):
        try:
            if not c.exists():
                continue
            if _copy_names_other_job(c, state.get("eng")):
                continue               # another job's report: never our base
            key = (0 if _under_archive(c) else 1, c.stat().st_mtime)   # live beats archived
            if best_key is None or key > best_key:
                best = c; best_key = key
        except Exception:
            continue
    return best

_ENG_TOKEN_RE = r'[0-9A-Z][0-9A-Z./-]*'     # what one Eng# token looks like: '4041', '4016LF', 'LF'

def _head_eng_tokens(a2):
    """The whitespace-separated TOKENS of an A2 head line's Eng# segment
    ('Engineering # 4006    ·    …', see _write_head): everything after
    'Engineering #' up to the '·' separator (or the end of the cell), normalised
    and split. [] when the cell is not a head line at all. The ONE parse behind
    _eng_from_head_line and _head_names_eng (review round 6: one implementation
    for the delete / target gate and the opt-in / base gate)."""
    m = re.match(r'\s*Engineering\s*#\s*([^·]*)', str(a2 or ""))
    return _norm_key(m.group(1)).split() if m else []

def _eng_from_head_line(a2):
    """The FIRST Eng# token of an A2 head line, for messages and diagnostics
    ('names another job (Engineering # 4016LF)') — '' when the cell is no head
    line or its first token does not look like an Eng# (Eng#s are letters /
    digits with '-', '.', '/' at most). NOT the decider: an Eng# may have several
    tokens (see _head_names_eng), so a one-token read can never say 'this file is
    ours' or 'another job's' on its own."""
    toks = _head_eng_tokens(a2)
    tok = toks[0] if toks else ""
    return tok if re.fullmatch(_ENG_TOKEN_RE, tok) else ""

def _head_names_eng(a2, eng):
    """Does a report copy's A2 head line name THIS job (`eng`)? TRI-STATE:
    True = ours, False = another job's file, None = undecidable — no head line
    (a legacy-era file), a blank `eng`, or a first token that is no Eng# shape
    (a user typed over the app-owned cell: 'Engineering # — 4044 checked by
    JD'); the caller then decides on other evidence, as it always did for a
    legacy file, and _write_head rewrites the line on the next save.
    The compare is by WHOLE TOKENS: the job's Eng# tokens must be the LEADING
    tokens of the segment, text after them is a user edit and ignored
    ('Engineering # 4044 — checked by JD' names 4044; review round 6). An Eng# may
    have SEVERAL tokens — the folder-name parser (eng_from_inside) joins every
    token before 'Rev' with a space, so '136-9630-900 GizmoX [4044 LF Rev. A]'
    is job '4044 LF' and its own head line reads 'Engineering # 4044 LF    ·'.
    Round 6 compared the first token only, so such a job's OWN report was
    'another job's file' for good: the switch read False, every update was
    refused, and the next WO forked a second live file (review round 7). 4016 vs
    4016LF still differ (one token each, whole-token compare); a job whose Eng#
    is a token-prefix of another job's ('4044' vs '4044 LF', sharing a P/N)
    cannot be told from a user note after '4044' and reads as ours — the
    round-6 trade-off, unchanged."""
    ek = _norm_key(eng).split()
    toks = _head_eng_tokens(a2)
    if not ek or not toks or not re.fullmatch(_ENG_TOKEN_RE, toks[0]):
        return None
    return toks[:len(ek)] == ek

def _report_head_line(wb):
    """The app-owned A2 head line of a report copy ('Engineering # 4006    ·    …'
    — see _write_head), read from the 'Received from testing' sheet (else the
    first sheet). None when unreadable. Works on a full and on a read_only
    workbook (ws['A2'] in both modes). Never raises."""
    try:
        ws = wb[RCV_SHEET] if RCV_SHEET in wb.sheetnames else wb.worksheets[0]
        return ws["A2"].value
    except Exception:
        return None

def _report_copy_eng(wb):
    """The (first) Eng# token a report copy names in its own head line — for
    messages ('names another job (Engineering # 4016LF)'). '' when the copy
    carries no such line (a legacy-era or foreign file). Never raises. The
    DECISION 'ours / another job's' is _head_names_eng(_report_head_line(wb),
    eng), never a compare of this token (review round 7)."""
    return _eng_from_head_line(_report_head_line(wb))

def _report_head_line_at(path):
    """_report_head_line for a file on disk without a full parse: openpyxl's
    read_only mode streams the sheet, so reading A2 costs a zip open, not the
    whole workbook (review round 5: the full rich-text parse ran per candidate
    in _report_lookup AND again in _find_legacy_received_file before the merge
    loaded the chosen base twice more — seconds each over SMB on a 900-row
    report). Always closed again (read_only keeps the file handle open — the
    copy may be unlinked or replaced right after). None on any failure."""
    from openpyxl import load_workbook
    wb = None
    try:
        wb = load_workbook(str(path), read_only=True)
        return _report_head_line(wb)
    except Exception:
        return None
    finally:
        try:
            if wb is not None: wb.close()
        except Exception:
            pass

def _report_copy_eng_at(path):
    """_report_copy_eng for a file on disk (read-only zip open, see
    _report_head_line_at). '' on any failure."""
    return _eng_from_head_line(_report_head_line_at(path))

def _copy_names_other_job(path, eng):
    """True when a report copy's head line names a DIFFERENT job than `eng` —
    it is another job's file and must never be merged into ours, used as our
    base or deleted (review round 2: two jobs sharing a P/N whose Eng#s are
    prefix-related, 4016 / 4016LF, name their reports identically). A copy that
    cannot be read, has no head line or an unreadable Eng# -> False (decided on
    other evidence). The compare is _head_names_eng's whole-token, multi-token
    one (review round 7)."""
    try:
        return _head_names_eng(_report_head_line_at(path), eng) is False
    except Exception:
        return False

def _extract_received_rows(path, eng=None):
    """Harvest ANOTHER copy of the report so its content merges into the one
    true file. Returns None when NOTHING can be assumed about the file (the
    caller must then KEEP it on disk), or {'rcv': rows, 'coat': rows, 'clean':
    bool} where clean=True asserts the copy's ENTIRE user-visible content is in
    those rows — only then may the caller delete it. Old READ-ONLY-era copies
    are clean and empty (app-generated, write-protected: nothing to preserve);
    a copy with drawn objects, user-added sheets, extra columns, serial-less
    content rows or a lost header merges what it safely can but stays on disk.
    With `eng`, a copy whose own head line names a DIFFERENT Eng# is another
    job's report -> None (kept, never merged, never deleted — review round 2)."""
    try:
        if not Path(path).exists():
            return None
        owb = _load_wb(path)
    except Exception:
        return None
    if eng is not None and _norm_key(eng):
        if _head_names_eng(_report_head_line(owb), eng) is False:
            return None                        # another job's file: not ours to touch
    # drawn objects, threaded/legacy comments, VML: the row-harvest can't carry
    # them, so their presence bars deletion (the rows still merge)
    out = {"rcv": [], "coat": [],
           "clean": _xlsx_blockers(path) is None and not _xlsx_has_annotations(path)}
    try:
        if RCV_SHEET not in owb.sheetnames:
            if _is_read_only(path):
                return {"rcv": [], "coat": [], "clean": True}
            return None       # unknown provenance: keep the file, merge nothing
        owbv = _load_wb_values(path)
        known = {RCV_SHEET, COAT_SHEET, IH_SHEET, OUT_SHEET}
        if any(nm not in known for nm in owb.sheetnames):
            out["clean"] = False                       # user-added sheets
        for slot, nm, ncols in (("rcv", RCV_SHEET, len(RCV_COLS)),
                                ("coat", COAT_SHEET, len(COAT_COLS))):
            if nm not in owb.sheetnames:
                out["clean"] = False
                continue
            ows = owb[nm]
            owsv = owbv[nm] if (owbv is not None and nm in owbv.sheetnames) else None

            def val(r, ci):
                if owsv is not None:
                    v = owsv.cell(row=r, column=ci).value
                    if v is not None:
                        return v
                return ows.cell(row=r, column=ci).value

            hdr = None
            for r in range(1, min(ows.max_row, 60) + 1):
                if str(ows.cell(row=r, column=1).value or "").strip() == "#":
                    hdr = r; break
            if hdr is None:
                out["clean"] = False                   # lost header: can't harvest
                continue
            try:
                if getattr(ows, "_hyperlinks", None):
                    out["clean"] = False           # links aren't carried by values
            except Exception:
                out["clean"] = False
            width = max(ows.max_column, ncols)
            # above the header row only the app's OWN A1-A4 head lines belong;
            # a user note anywhere else there (incl. col A row 5+) bars deletion
            for r in range(1, hdr):
                if any(ows.cell(row=r, column=ci).value not in (None, "")
                       for ci in range(1, width + 1)
                       if not (ci == 1 and r <= 4)):
                    out["clean"] = False
                    break
            for r in range(hdr + 1, ows.max_row + 1):
                sn = val(r, 2)
                rowvals = [val(r, ci) for ci in range(1, ncols + 1)]
                if any(ows.cell(row=r, column=ci).value not in (None, "")
                       for ci in range(ncols + 1, width + 1)):
                    out["clean"] = False               # values beyond the known columns
                if sn is None or not str(sn).strip():
                    if any(v not in (None, "") for v in rowvals):
                        out["clean"] = False           # a serial-less content row
                    continue
                out[slot].append(rowvals)
    except Exception:
        return None
    return out

def write_void_event(target_rec, reason, admin_id):
    """Supervisor VOID of a wrong shipment/receive record: the target stays in
    history (audit) but its serials stop counting — the duplicate check and the
    Receiving tracker skip voided records. Returns the written path or None."""
    tid = str((target_rec or {}).get("event_id") or "")
    if not tid:
        return None
    return write_history_event("void", {
        "wo": str(target_rec.get("wo") or ""),
        "division": str(target_rec.get("division") or ""),
        "eng": str(target_rec.get("eng") or ""),
        "pn": str(target_rec.get("pn") or ""),
        "target_event_id": tid,
        "target_type": str(target_rec.get("event_type") or ""),
        "target_ts": str(target_rec.get("ts") or ""),
        "serial_count": len(target_rec.get("serials") or []),
        "reason": str(reason or ""),
        "approved_by": str(admin_id or "")})

def _reemit_file_date(recorded_name, ship_date):
    """The date token the app puts in a paperwork filename (m-d-yy). Prefer the
    ORIGINAL filename's token (a verbatim restore); else derive it from the ship
    date. Returns '' if neither is usable (caller falls back to today)."""
    m = re.search(r'-(\d{1,2}-\d{1,2}-\d{2})-', str(recorded_name or ""))
    if m:
        return m.group(1)
    d = _parse_user_date(ship_date)          # 'YYYY-MM-DD' or None
    if d:
        try:
            dt = datetime.datetime.strptime(d, "%Y-%m-%d")
            if sys.platform == "win32":
                return dt.strftime("%#m-%#d-") + dt.strftime("%y")
            return dt.strftime("%-m-%-d-") + dt.strftime("%y")
        except Exception:
            pass
    return ""

def reemit_paperwork(rec):
    """Re-generate the COC + Packing List for a past 'close' shipment from its
    History record, using the app's OWN writers so the output (Code128 barcodes,
    logo, layout) is identical to the app's format, and restoring the ORIGINAL
    ship-date filename. For paperwork that was DELETED or LOST from a WO folder;
    History is the source of truth (Dhruvil 2026-09-09, the 4092 case: 26 coated
    boards' COC/Packing files had been deleted from the share).

    Also rebuilds the Shipping Slip / Shipping Labels pages from rec['slip']
    (Dhruvil 2026-09-24) — only the pages that were printed, under their
    original names, with the recorded form codes (no new random digits).

    Safety, consistent with the rest of the app:
      • Only a 'close' record re-emits (its coc/packing/serials ARE the shipment).
      • NEVER overwrites a file still on disk — that document is left untouched.
      • NEVER fabricates a folder tree — writes only into a WO folder that already
        exists (the recorded output folder, else the recorded WO folder with the
        standard subfolder created as a LEAF); refuses otherwise, so no share
        access or a renamed/archived folder can't spawn a phantom tree.
    Returns {ok, wrote:[paths], kept:[msgs], errors:[msgs]}; never raises."""
    res = {"ok": False, "wrote": [], "kept": [], "errors": []}
    if str((rec or {}).get("event_type")) != "close":
        res["errors"].append("Only a closed-shipment record can re-emit paperwork.")
        return res
    eng = sanitize_filename_part(rec.get("eng_safe") or rec.get("eng") or "", "unknown")
    wo = str(rec.get("wo") or "")
    serials = [str(s) for s in (rec.get("serials") or []) if str(s).strip()]
    # The Shipping Slip / Labels (Dhruvil 2026-09-24) are rebuilt from the
    # record's slip_data and need no serials; the COC and Packing List do.
    slip = rec.get("slip") or {}
    no_serials = "This record has no serials to put on the paperwork."
    if not serials and not slip:
        res["errors"].append(no_serials)
        return res
    out = rec.get("outputs") or {}
    wof = rec.get("wo_folder")

    def _target_folder(recorded_path, subfolder):
        # 1) the folder the file was originally in, if it still exists on disk
        if recorded_path:
            par = Path(str(recorded_path)).parent
            if par.is_dir():
                return par
        # 2) else the recorded WO folder + the standard subfolder — created as a
        #    LEAF only, and only inside a WO folder that already exists (never a
        #    whole tree; see the 4092 rev-change lesson).
        if wof and Path(str(wof)).is_dir():
            d = Path(str(wof)) / subfolder
            try:
                d.mkdir(exist_ok=True)
                return d
            except Exception:
                return None
        return None

    # plan rows: (label, recorded path, subfolder, filename kind, ext, ship date, writer,
    #             unfiled code) — the last one is the form code of a slip page that was
    #             printed but never filed on the share (see below); None otherwise.
    plan = []
    coc = rec.get("coc") or {}
    pk = rec.get("packing") or {}
    ship_date = str((coc.get("ship_date") if coc else "") or (pk.get("date") if pk else ""))
    if (coc or pk) and not serials:
        res["errors"].append(no_serials)
    if coc and serials:
        plan.append(("COC", out.get("coc_xlsx"), COC_FOLDER, "COC", ".xlsx", ship_date,
                     lambda folder: write_coc_xlsx(str(folder), eng, wo, coc, serials), None))
    if pk and serials:
        plan.append(("Packing List", out.get("packing_xlsx"), PACKING_LIST_FOLDER, "Packing_List",
                     ".xlsx", ship_date,
                     lambda folder: write_packing_list_xlsx(str(folder), eng, wo, pk, serials,
                                                            logo_path=str(LOGO_PATH)), None))
    # Shipping Slip / Labels pages: only the ones that were actually printed —
    # a recorded share path, or a printed stamp with NO path (printed while the
    # share was down: the copy went to the temp print folder and was never
    # filed, so re-emit is what files it). A page that never existed is not
    # "missing". Rebuilt from what the page was rendered from, with its recorded
    # form code.
    if slip:
        sout = slip.get("outputs") or {}
        spr = slip.get("printed") or {}
        for key, (pkey, kind, label, _ck) in _SLIP_KINDS.items():
            recorded = sout.get(key) or out.get(key)
            if recorded and _is_temp_print_path(recorded):
                recorded = None                 # never "restore" into %TEMP%
            if recorded or spr.get(pkey):
                plan.append((label, recorded, SLIP_FOLDER, kind, ".html", str(slip.get("date") or ""),
                             lambda folder, k=key: _reemit_write_slip_html(folder, slip, k),
                             None if recorded else _slip_page_code(slip, key)))
    if not plan:
        if not res["errors"]:
            res["errors"].append("This record has no COC/Packing List"
                                 + ("/Shipping Slip" if slip else "") + " data to rebuild from.")
        return res

    for label, recorded, subfolder, kind, ext, doc_date, writer, unfiled_code in plan:
        try:
            recname = Path(str(recorded)).name if recorded else ""
            fd = _reemit_file_date(recname, doc_date)
            orig_name = recname or f"{eng}-{fd or fmt_file_date()}-{kind}{ext}"
            folder = _target_folder(recorded, subfolder)
            if folder is None:
                res["errors"].append(
                    f"{label}: the shipment's folder isn't reachable from here — nothing "
                    f"written. Rerun on a PC with share access, or restore the folder first.")
                continue
            if unfiled_code is not None:
                # Printed while the share was down: there is no original name to
                # restore, and the clean <eng>-<date>-<kind> name may already belong
                # to ANOTHER print of the same day (the 9:00 online print the 9:30
                # offline reprint corrected, or another shipment of the job). So
                # "already on disk" is decided by this print's form code (drawn once
                # per print — it identifies the paper), and a new copy takes the next
                # free name, the _HHMMSS rule a same-day reprint gets. Sending it
                # through the name check below left the paper that actually shipped
                # unfiled for good (review round 2).
                have = _reemit_filed_page(folder, kind, unfiled_code)
                if have is not None:
                    res["kept"].append(f"{label}: {have.name} already on disk — left untouched")
                    continue
                target = unique_file_path(folder, orig_name[:-len(ext)], ext)
                written = Path(writer(folder))
                if written != target and not target.exists():
                    os.replace(str(written), str(target))
                    written = target
                res["wrote"].append(str(written))
                continue
            target = folder / orig_name
            if target.exists():
                res["kept"].append(f"{label}: {orig_name} already on disk — left untouched")
                continue
            written = Path(writer(folder))          # app writer -> a today-dated temp name
            if written.name != orig_name:
                if target.exists():                 # a copy appeared meanwhile: don't clobber
                    try: written.unlink()
                    except Exception: pass
                    res["kept"].append(f"{label}: {orig_name} already on disk — left untouched")
                    continue
                os.replace(str(written), str(target))
                written = target
            res["wrote"].append(str(written))
        except Exception as e:
            res["errors"].append(f"{label}: {type(e).__name__}: {e}")
    res["ok"] = bool(res["wrote"]) and not res["errors"]
    return res

def parse_serials_from_shipping_files(wo_folder, pn):
    """Backfill reader for ONE WO folder: collect serials from every .xlsx/.csv in
    its 'Shipping Serials' subfolder. Primary: the app's own layout (a header row
    with a 'Serial…' column, plus a 'Date' column for the ship date). Fallback for
    odd/manual layouts: any text cell whose alphanumerics embed the job P/N.
    Returns (serials_in_order, ship_date, files_read); never raises."""
    out, seen = [], set()
    ship_date = ""; files = 0
    pa = _alnum(pn)
    def add(v):
        v = str(v).strip(); k = _norm_serial(v)
        if v and k and k not in seen:
            seen.add(k); out.append(v)
    try:
        ser_dir = Path(wo_folder) / SERIALS_FOLDER
        if not ser_dir.is_dir():
            return [], "", 0
        for f in sorted(ser_dir.iterdir()):
            n = f.name.lower()
            rows = None
            try:
                if n.endswith(".xlsx") and not n.startswith("~"):
                    from openpyxl import load_workbook
                    wb = load_workbook(str(f), read_only=True, data_only=True)
                    rows = [list(r) for r in wb.active.iter_rows(values_only=True)]
                    wb.close()
                elif n.endswith(".csv"):
                    with open(f, encoding="utf-8-sig", newline="") as fh:
                        rows = [r for r in csv.reader(fh)]
                else:
                    continue
            except Exception:
                continue
            files += 1
            hi = scol = dcol = None
            for i, r in enumerate(rows[:10]):
                for j, c in enumerate(r):
                    if isinstance(c, str) and "serial" in c.lower():
                        hi, scol = i, j; break
                if hi is not None: break
            if hi is not None:
                hdr = rows[hi]
                dcol = next((j for j, c in enumerate(hdr)
                             if isinstance(c, str) and "date" in c.lower()), None)
                for r in rows[hi + 1:]:
                    if scol < len(r) and r[scol] not in (None, ""):
                        add(r[scol])
                        if dcol is not None and not ship_date and dcol < len(r) and r[dcol]:
                            ship_date = str(r[dcol])
            elif len(pa) >= 6:
                for r in rows:
                    for c in r:
                        if isinstance(c, str) and pa in _alnum(c):
                            add(c)
    except Exception:
        pass
    return out, ship_date, files

def _extract_file_date(rows):
    """Best-effort ship date from ONE shipment file's cells: a 'Date' header
    column first (value below or beside it), else the first date-looking cell.
    Returns 'YYYY-MM-DD' or ''."""
    hi = dcol = None
    for i, r in enumerate(rows[:12]):
        for j, c in enumerate(r):
            if isinstance(c, str) and "date" in c.lower() and len(c.strip()) < 30:
                hi, dcol = i, j
                break
        if hi is not None:
            break
    def _dk(v):
        dk = _date_key(v)
        return dk if re.match(r"^\d{4}-\d{2}-\d{2}$", dk) else ""
    if hi is not None:
        hr = rows[hi]
        if dcol + 1 < len(hr) and hr[dcol + 1] not in (None, ""):
            got = _dk(hr[dcol + 1])              # 'Date:' label with value beside
            if got:
                return got
        for r in rows[hi + 1:]:
            if dcol < len(r) and r[dcol] not in (None, ""):
                got = _dk(r[dcol])
                if got:
                    return got
                break
    for r in rows:
        for c in r:
            if isinstance(c, (datetime.datetime, datetime.date)):
                return _date_key(c)
    for r in rows:
        for c in r:
            if isinstance(c, str):
                got = _dk(c)
                if got:
                    return got
    return ""

def _cell_serial(c, pk_al, ck_al):
    """The serial in ONE spreadsheet cell, or None. Guards against the app's own
    paperwork (review round 3): the COC/Packing-List 'P/N' field and the COC
    certification sentence also EMBED the P/N but are not serials — they became
    phantom boards and broke the voided-shipment skip. A real serial for this
    customer is '<seq>-<P/N>-<rev>': short, numeric sequence first, embedded P/N."""
    if not isinstance(c, str):
        return None
    v = c.strip()
    if not v or len(v) > 40:
        return None                       # the cert sentence is a whole paragraph
    al = _alnum(v)
    pa = ck_al if (ck_al and ck_al in al) else (pk_al if pk_al and pk_al in al else "")
    if not pa or al == pa:
        return None                       # no P/N embedded / the bare P/N field
    if not serial_seq(v):
        return None                       # no numeric sequence in front
    # The P/N must sit CONTIGUOUSLY right after the leading sequence — a real
    # serial is '<seq>-<P/N>-<rev>' (e.g. '000021-1342201970-A00'). A COC 'P/N'
    # cell that holds the whole FOLDER NAME ('134-2201-970 4065 Rev. A00 Rev.
    # A00', from the pre-v58 parser bug) also embeds the P/N and starts with a
    # number ('134' reads as a seq), but there the P/N is the field's own
    # hyphenated value, not a serial. _alnum() hides that by stripping hyphens,
    # so require the chars after the first separator to BEGIN with the P/N.
    tail = re.split(r'[-_\s]', v, maxsplit=1)
    if len(tail) < 2 or not _alnum(tail[1]).startswith(pa):
        return None
    return v

def _file_pn_label_side(rows, core):
    """Which side of the pair a shipment file's own P/N cell names: '970' when a
    cell is the job's coated label ('125-1030-970C', 'P/N: 125-1030-970C'), '900'
    for the testing label, '' when the file carries no such cell (or `core` is
    unknown). Only a WHOLE cell counts — never a serial (its leading sequence
    digits fail the full match) and never the cert sentence."""
    if not core:
        return ""
    try:
        for r in rows:
            for c in r:
                if not isinstance(c, str) or len(c) > 40:
                    continue
                m = re.fullmatch(r'(?:PN)?(\d{6,7})(900|970)([A-Z]{0,2})', _alnum(c))
                if m and m.group(1) == core:
                    return m.group(2)
    except Exception:
        pass
    return ""

def scan_job_folders(state, progress=None):
    """SHIPMENT TRUTH FROM THE JOB FOLDERS (Dhruvil 2026-09-03, the 4094 case:
    "you have to look for 900 and 970 folder in each job and scan either packing
    list or coc or shipping serials folder to be accurate"). Reads every WO's
    COC / Packing List / Shipping Serials xlsx+csv files in EVERY job folder of
    this Eng# whose name embeds the testing or coated P/N — rev-change and
    renamed sibling folders included (4094 has TWO -970 folders). Only cells
    shaped like real serials count (_cell_serial), routed by the P/N embedded
    in the serial itself, exactly like the history walk.

    COC is the shipment record (his call), aggregated PER WO: within one WO a
    serial's ship count = its COC-file count when any COC there carries it, else
    max(Packing List, Shipping Serials); WO counts then add up (a re-ship
    documented in another WO without a COC still counts). Disagreements between
    the three folders — and files that exist but can't be parsed — come back in
    result['discrepancies'] as a heads-up, never an error. Shipments the app
    VOIDED are skipped even though their files still exist.

    Never raises. result['ok'] False = the folders/files could not be read (no
    access from this PC, or an I/O error mid-scan) — the caller then rebuilds
    history-only and says so; a PARTIAL scan is never used (it would undercount
    ship cycles). Unparseable-but-readable files only skip that one file,
    counted in result['skipped'] and surfaced in the heads-up.

    Coated side (Dhruvil 2026-09-23, in-house testing): the same three-folder
    disagreement check runs for coated serials — lines tagged '(WO …, coated)' in
    result['discrepancies'] — and a coated board found shipped 2+ times across
    the files is listed in result['coated_headsup'] ("coated board shipped N×").
    Both are POPUP text only; nothing is written into the report. Folder twin of
    the engine's Case B (the 8/27 "Anduril test damage" shipments): a -900-form
    serial in a file whose own P/N cell is the COATED label — or, when the file
    carries no P/N cell, in a file under the COATED job folder — is a coated row
    with the Note 'scanned with a -900 testing label', never a testing shipment;
    otherwise 🔁 Rebuild would turn an in-house-tested job's report into 'Shipped
    for testing: 8 · Still at customer: 8'. Review round 1: this re-route applies
    ONLY to a job whose HISTORY never shipped for testing (hist_has_testing False)
    or to boards History already classified as Case B (coat_label_mismatch) — a
    test-and-return job's folders route exactly as before (embedded P/N decides)."""
    res = {"ok": False, "error": "", "files": 0, "jobs": [], "batches": [],
           "ship_counts": {}, "ship_info": {}, "coated_rows": [],
           "discrepancies": [], "coated_headsup": [], "skipped": []}
    pn = str(state.get("pn") or ""); eng = str(state.get("eng") or "")
    pk_al, ck_al = _tr_embed_keys(pn)         # the walk's own keys (review round 6: a lettered
    ek = _norm_key(eng)                        # -900 key gave ck_al '' -> the -970 tree never scanned)
    pd = state.get("pd_path")
    if not pd or len(pk_al) < 6 or not ek:
        res["error"] = "the job's folder is not known yet"
        return res
    # Folder-side Case B is for in-house-tested jobs (History never saw a -900
    # label) and for boards History already filed as Case B — decided from the
    # history snapshot, never from an earlier fold of this same scan
    hist_testing = _hist_has_testing(state)
    caseb_hist = state.get("coat_label_mismatch") or set()
    core_d = _tr_pn_parts(pn)[0]
    try:
        cust = Path(pd).parent.parent          # <customer>\<job>\Production Data
        job_dirs = [d for d in cust.iterdir() if d.is_dir()
                    and _job_dir_is(d.name, ek)
                    and (pk_al in _alnum(d.name) or (ck_al and ck_al in _alnum(d.name)))]
    except Exception as e:
        res["error"] = f"{type(e).__name__}: {e}"
        return res
    if not job_dirs:
        res["error"] = "no matching job folders found"
        return res

    voided = state.get("voided_ships") or []
    def _is_voided(wo_name, serset):
        for v in voided:
            try:
                vs = v.get("serials") or set()
                if serset and vs and serset <= vs:
                    vw, bw = _alnum(v.get("wo")), _alnum(wo_name)
                    if not vw or not bw or vw == bw or vw in bw:
                        return True
            except Exception:
                pass
        return False

    import csv as _csv
    from openpyxl import load_workbook
    SUBS = ((COC_FOLDER, "coc"), (PACKING_LIST_FOLDER, "pl"), (SERIALS_FOLDER, "ss"))
    try:
        for jd in sorted(job_dirs, key=lambda d: d.name):
            res["jobs"].append(str(jd))
            pdir = jd / "Production Data"
            if not pdir.is_dir():
                continue
            # the COATED job tree ('125-1030-970C 4006 Rev. C'): paperwork here
            # belongs to coated shipments, whatever label a serial wears (Case B)
            coat_tree = bool(ck_al) and ck_al in _alnum(jd.name) and pk_al not in _alnum(jd.name)
            for wo in sorted(x for x in pdir.iterdir() if x.is_dir()):
                if _is_report_folder(wo.name):
                    continue
                for sub, slot in SUBS:
                    d = wo / sub
                    if not d.is_dir():
                        continue
                    for f in sorted(d.iterdir()):
                        n = f.name.lower()
                        if f.is_dir() or n.startswith("~"):
                            continue
                        if callable(progress):
                            try: progress(f"{wo.name} · {f.name[:40]}")
                            except Exception: pass
                        rows = None
                        try:
                            if n.endswith(".xlsx"):
                                wb = load_workbook(str(f), read_only=True, data_only=True)
                                rows = []
                                for ws_ in wb.worksheets:
                                    for r in ws_.iter_rows(values_only=True):
                                        rows.append(list(r))
                                        if len(rows) > 4000:
                                            break
                                wb.close()
                            elif n.endswith(".csv"):
                                with open(f, encoding="utf-8-sig", newline="") as fh:
                                    rows = [r for r in _csv.reader(fh)]
                            else:
                                continue       # pdf/docx/etc: not readable
                        except OSError as e:
                            # a share/permission blip mid-scan would silently
                            # UNDERCOUNT ship cycles — refuse the whole scan
                            res["error"] = f"{type(e).__name__} on {f.name}: {e}"
                            return res
                        except Exception:
                            res["skipped"].append(f"{wo.name}\\{sub}\\{f.name}")
                            continue           # corrupt/odd file: skip, but SAY so
                        res["files"] += 1
                        # the file's OWN P/N cell ('P/N: 125-1030-970C' on the COC of
                        # the in-progress -900-serialized batch that ships as -970):
                        # '970' = coated paperwork, '900' = testing, '' = no cell
                        flabel = _file_pn_label_side(rows, core_d)
                        test_s, coat_s, caseb_s, seen_f = [], [], set(), set()
                        for r in rows:
                            for c in r:
                                v = _cell_serial(c, pk_al, ck_al)
                                if v is None or _norm_serial(v) in seen_f:
                                    continue
                                seen_f.add(_norm_serial(v))
                                if ck_al and ck_al in _alnum(v):
                                    coat_s.append(v)
                                elif ((flabel == "970" or (not flabel and coat_tree))
                                      and (not hist_testing
                                           or _serial_identity(v, _coat_ident_pn(pn)) in caseb_hist)):
                                    # Case B on disk: a -900-labelled board in coated
                                    # paperwork = a coated shipment (Q3, label-based)
                                    coat_s.append(v); caseb_s.add(_norm_serial(v))
                                else:
                                    test_s.append(v)
                        if not test_s and not coat_s:
                            continue
                        serset = {_norm_serial(x) for x in test_s + coat_s}
                        if _is_voided(wo.name, serset):
                            continue
                        res["batches"].append({"job": jd.name, "wo": wo.name,
                                               "slot": slot, "file": f.name,
                                               "date": _extract_file_date(rows),
                                               "test": test_s, "coat": coat_s,
                                               "caseb": caseb_s})
    except Exception as e:
        # a scan that died mid-way would UNDERCOUNT ship cycles — refuse it whole
        res["error"] = f"{type(e).__name__}: {e}"
        return res

    # ── Case B is decided per WO FOLDER, not per file (review round 4): the app's
    #    own Shipping-Serials file carries no P/N cell, so under the -900 job tree
    #    it routed the in-progress batch's -900-form boards as TESTING shipments
    #    while the COC / Packing List of the SAME shipment (P/N cell = the coated
    #    label) filed them as coated — one board on 'Coated & shipped' AND on
    #    'Still at customer'. A serial another file of the same WO already files
    #    as Case B is Case B in every file of that WO. A label-less list ALONE
    #    (no coated-labelled file names its boards) still routes as before. ──
    cb_wo = {}
    for b in res["batches"]:
        for s in (b.get("caseb") or ()):
            cb_wo.setdefault((b["job"], b["wo"]), set()).add(s)
    for b in res["batches"]:
        cb = cb_wo.get((b["job"], b["wo"]))
        if not cb or not b["test"]:
            continue
        keep = []
        for s in b["test"]:
            if _norm_serial(s) in cb:
                b["coat"].append(s); b["caseb"].add(_norm_serial(s))
            else:
                keep.append(s)
        b["test"] = keep

    # ── testing side: PER-WO counts, COC wins within a WO (a re-ship that was
    #    documented in another WO without a COC still counts) ────────────────
    per = {"coc": {}, "pl": {}, "ss": {}}      # {(identity, wo): n files}
    info = {}
    for b in res["batches"]:
        for s in b["test"]:
            k = _serial_identity(s, pn)
            if not k:
                continue
            kw = (k, b["wo"])
            per[b["slot"]][kw] = per[b["slot"]].get(kw, 0) + 1
            cur = info.get(k)
            if cur is None or (b["date"] or "") >= (cur.get("date") or ""):
                info[k] = {"serial": s, "wo": b["wo"], "date": b["date"]}
    ships = {}
    for kw in set(per["coc"]) | set(per["pl"]) | set(per["ss"]):
        n = per["coc"].get(kw) or max(per["pl"].get(kw, 0), per["ss"].get(kw, 0))
        ships[kw[0]] = ships.get(kw[0], 0) + n
    res["ship_counts"] = ships
    res["ship_info"] = info

    # ── coated side: same COC primacy per (serial, WO); a DATED row beats a
    #    blank-dated one (one COC + one dateless serials file = ONE row) ─────
    coat_by = {}                               # {(identity, wo): {slot: [entry]}}
    per_c = {"coc": {}, "pl": {}, "ss": {}}    # {(identity, wo): n files} — coated
    disp_c = {}
    for b in res["batches"]:
        for s in b["coat"]:
            k = _serial_identity(s, _coat_ident_pn(pn))
            entry = {"serial": s, "wo": b["wo"], "ts": b["date"]}
            if _norm_serial(s) in (b.get("caseb") or ()):
                entry["note"] = _CASE_B_NOTE            # same tag as the history walk
            coat_by.setdefault((k, b["wo"]), {}).setdefault(b["slot"], []).append(entry)
            if k:
                kw = (k, b["wo"])
                per_c[b["slot"]][kw] = per_c[b["slot"]].get(kw, 0) + 1
                disp_c.setdefault(k, s)
    coat_rows = {}
    for (k, wo), slots in coat_by.items():
        rows = slots.get("coc") or slots.get("pl") or slots.get("ss") or []
        dated = [r for r in rows if r.get("ts")]
        for r in (dated or rows):
            kk = (k, wo, _date_key(r.get("ts")))
            if kk not in coat_rows:
                coat_rows[kk] = r
    res["coated_rows"] = list(coat_rows.values())

    have = {slot: any(b["slot"] == slot and b["test"] for b in res["batches"])
            for slot in ("coc", "pl", "ss")}
    names = {"coc": "COC", "pl": "Packing List", "ss": "Shipping Serials"}
    flagged = set()
    for kw in sorted(set(per["coc"]) | set(per["pl"]) | set(per["ss"])):
        got = {slot: per[slot].get(kw, 0) for slot in ("coc", "pl", "ss") if have[slot]}
        if len(set(got.values())) > 1 and kw[0] not in flagged:
            flagged.add(kw[0])
            disp = (info.get(kw[0]) or {}).get("serial") or kw[0]
            res["discrepancies"].append(
                f"{disp} (WO {kw[1]}):  "
                + " · ".join(f"{names[s]} {v}×" for s, v in got.items()))
    # ── coated side heads-up (Dhruvil 2026-09-23): the same disagreement line,
    #    tagged 'coated'; and a coated board documented as shipped 2+ times
    #    (COC primacy per WO, WOs added up — a same-WO same-date duplicate COC
    #    still collapses to ONE row above, so this line is how it gets seen) ──
    have_c = {slot: any(b["slot"] == slot and b["coat"] for b in res["batches"])
              for slot in ("coc", "pl", "ss")}
    flagged_c = set()
    for kw in sorted(set(per_c["coc"]) | set(per_c["pl"]) | set(per_c["ss"])):
        got = {slot: per_c[slot].get(kw, 0) for slot in ("coc", "pl", "ss") if have_c[slot]}
        if len(set(got.values())) > 1 and kw[0] not in flagged_c:
            flagged_c.add(kw[0])
            res["discrepancies"].append(
                f"{disp_c.get(kw[0]) or kw[0]} (WO {kw[1]}, coated):  "
                + " · ".join(f"{names[s]} {v}×" for s, v in got.items()))
    cships = {}
    for kw in set(per_c["coc"]) | set(per_c["pl"]) | set(per_c["ss"]):
        n = per_c["coc"].get(kw) or max(per_c["pl"].get(kw, 0), per_c["ss"].get(kw, 0))
        cships[kw[0]] = cships.get(kw[0], 0) + n
    for k in sorted(cships):
        if cships[k] >= 2:
            res["coated_headsup"].append(
                f"{disp_c.get(k) or k}:  coated board shipped {cships[k]}× — check records")
    for name in res["skipped"][:8]:
        res["discrepancies"].append(f"could not read {name} — its shipment is NOT counted")
    res["ok"] = True
    return res

def fold_folder_scan(state, scan):
    """Fold the folder truth into a freshly walked state (🔁 Rebuild only —
    Dhruvil's chosen timing): ship counts become max(history, folders), old
    never-imported shipments join the shipped baseline, and every coated
    shipment found on disk joins the coated bucket + log. History receive
    events remain the only receive truth."""
    if not (scan and scan.get("ok")):
        return state
    sc = state.setdefault("ship_counts", {})
    for k, n in (scan.get("ship_counts") or {}).items():
        if n > sc.get(k, 0):
            sc[k] = n
    shipped = state.setdefault("shipped", {})
    for k, inf in (scan.get("ship_info") or {}).items():
        if k not in shipped:
            # a KNOWN file date must join the timestamp tie-break; '' (=older
            # than any receive) is only right when the date is truly unknown
            d = str(inf.get("date") or "")
            shipped[k] = {"serial": inf.get("serial", ""), "wo": inf.get("wo", ""),
                          "ts": d, "sts": (d + " 00:00:00") if d else "",
                          "source": "folder"}
    coated = state.setdefault("coated", {})
    clog = state.setdefault("coated_log", [])
    seen = {(_norm_serial(e.get("serial")), _date_key(e.get("ts"))) for e in clog}
    # a history close IS the record of that shipment — its folder files must not
    # add a second row just because the file's date differs from the close ts
    seen_wo = {(_norm_serial(e.get("serial")), _alnum(e.get("wo"))) for e in clog}
    # A folder-side Case-B row (a -900-labelled board in coated paperwork, no
    # History close for it) is a Case-B IDENTITY of this job too: the folded
    # state's _has_testing / _coated_only, the RCV cross-reference and the next
    # scan's caseb_hist must agree with the sheet the fold just wrote (review
    # round 7 — an overridden receive of such a board read as a genuine -900
    # return and the next Rebuild re-routed the same boards to 'Still at
    # customer' while their Case-B rows stayed on the log). Folder-only ones are
    # remembered apart (state['_caseb_folder']) because the daemon's history-
    # only walk cannot see them — the Save sweep warns before such a receive.
    mism = state.setdefault("coat_label_mismatch", set())
    cb_folder = state.setdefault("_caseb_folder", set())
    cpn_f = _coat_ident_pn(state.get("pn"))
    for row in (scan.get("coated_rows") or []):
        key = (_norm_serial(row.get("serial")), _date_key(row.get("ts")))
        if key in seen or (key[0], _alnum(row.get("wo"))) in seen_wo:
            continue
        seen.add(key)
        ent = {"serial": row.get("serial", ""), "wo": row.get("wo", ""),
               "ts": row.get("ts", "")}
        if row.get("note"):
            ent["note"] = str(row["note"])   # folder-side Case B keeps its tag
            if _CASE_B_NOTE in ent["note"]:
                kc = _serial_identity(row.get("serial"), cpn_f)
                if kc and kc not in mism:
                    cb_folder.add(kc)
                    mism.add(kc)
        clog.append(ent)
        q = serial_seq(row.get("serial"))
        if q and q not in coated:
            coated[q] = {"serial": row.get("serial", ""), "wo": row.get("wo", ""),
                         "ts": row.get("ts", "")}
    state["folder_scan"] = {"files": scan.get("files", 0), "jobs": scan.get("jobs") or [],
                            "discrepancies": scan.get("discrepancies") or [],
                            "coated_headsup": scan.get("coated_headsup") or []}
    return state

def receiving_anomaly_notes(state):
    """{identity: note} for boards whose records look inconsistent: shipped 2+
    times but NEVER recorded as FAILED (Dhruvil 2026-09-03: 'write one comment
    that user understand that board shipped twice even though it received as
    passed'). Shown in the status sheets' Note column and the heads-up popup.
    TESTING side only: ship_counts holds -900-labelled shipments, so nothing here
    fires for coated boards or for the Case-B boards of an in-house-tested job
    (Dhruvil 2026-09-23, Q3 — those are coated LOG rows with their own Note)."""
    pn = state.get("pn")
    fails = {_serial_identity(e.get("serial"), pn)
             for e in (state.get("receive_log") or []) if e.get("result") == "fail"}
    notes = {}
    for k, n in (state.get("ship_counts") or {}).items():
        if n >= 2 and k not in fails:
            notes[k] = f"shipped {n}× but never recorded as FAILED — check records"
    return notes


###############################################################################
# 4x2 BOARD LABELS (HPKE200) — added 2026-08-21
###############################################################################
# Replicates the PCC Apex "Serialized Pack 4x2" label inside this app: a 4in x 2in
# label, LEFT column = P/N (title 12pt, value 10pt, QR code, 6pt mirror text) and
# RIGHT column = Serial (same structure). The serial popup's "Print 4x2 labels"
# button writes ONE self-contained HTML file (qrcode.js embedded below, so no
# internet needed) with one 4x2 page per serial, then opens it in the browser —
# the operator picks the HPKE200 in the print dialog exactly like they do in Apex,
# and all labels print as a single job.
#
# qrcode.js (c) Shim Sangmin, MIT license — https://github.com/davidshimjs/qrcodejs
# (the same library the Apex app loads from its CDN), embedded as base64 so the
# generated label file also works with no network.
_QRCODEJS_B64 = (
    "dmFyIFFSQ29kZTshZnVuY3Rpb24oKXtmdW5jdGlvbiBhKGEpe3RoaXMubW9kZT1jLk1PREVfOEJJVF9CWVRFLHRoaXMuZGF0YT1h"
    "LHRoaXMucGFyc2VkRGF0YT1bXTtmb3IodmFyIGI9W10sZD0wLGU9dGhpcy5kYXRhLmxlbmd0aDtlPmQ7ZCsrKXt2YXIgZj10aGlz"
    "LmRhdGEuY2hhckNvZGVBdChkKTtmPjY1NTM2PyhiWzBdPTI0MHwoMTgzNTAwOCZmKT4+PjE4LGJbMV09MTI4fCgyNTgwNDgmZik+"
    "Pj4xMixiWzJdPTEyOHwoNDAzMiZmKT4+PjYsYlszXT0xMjh8NjMmZik6Zj4yMDQ4PyhiWzBdPTIyNHwoNjE0NDAmZik+Pj4xMixi"
    "WzFdPTEyOHwoNDAzMiZmKT4+PjYsYlsyXT0xMjh8NjMmZik6Zj4xMjg/KGJbMF09MTkyfCgxOTg0JmYpPj4+NixiWzFdPTEyOHw2"
    "MyZmKTpiWzBdPWYsdGhpcy5wYXJzZWREYXRhPXRoaXMucGFyc2VkRGF0YS5jb25jYXQoYil9dGhpcy5wYXJzZWREYXRhLmxlbmd0"
    "aCE9dGhpcy5kYXRhLmxlbmd0aCYmKHRoaXMucGFyc2VkRGF0YS51bnNoaWZ0KDE5MSksdGhpcy5wYXJzZWREYXRhLnVuc2hpZnQo"
    "MTg3KSx0aGlzLnBhcnNlZERhdGEudW5zaGlmdCgyMzkpKX1mdW5jdGlvbiBiKGEsYil7dGhpcy50eXBlTnVtYmVyPWEsdGhpcy5l"
    "cnJvckNvcnJlY3RMZXZlbD1iLHRoaXMubW9kdWxlcz1udWxsLHRoaXMubW9kdWxlQ291bnQ9MCx0aGlzLmRhdGFDYWNoZT1udWxs"
    "LHRoaXMuZGF0YUxpc3Q9W119ZnVuY3Rpb24gaShhLGIpe2lmKHZvaWQgMD09YS5sZW5ndGgpdGhyb3cgbmV3IEVycm9yKGEubGVu"
    "Z3RoKyIvIitiKTtmb3IodmFyIGM9MDtjPGEubGVuZ3RoJiYwPT1hW2NdOyljKys7dGhpcy5udW09bmV3IEFycmF5KGEubGVuZ3Ro"
    "LWMrYik7Zm9yKHZhciBkPTA7ZDxhLmxlbmd0aC1jO2QrKyl0aGlzLm51bVtkXT1hW2QrY119ZnVuY3Rpb24gaihhLGIpe3RoaXMu"
    "dG90YWxDb3VudD1hLHRoaXMuZGF0YUNvdW50PWJ9ZnVuY3Rpb24gaygpe3RoaXMuYnVmZmVyPVtdLHRoaXMubGVuZ3RoPTB9ZnVu"
    "Y3Rpb24gbSgpe3JldHVybiJ1bmRlZmluZWQiIT10eXBlb2YgQ2FudmFzUmVuZGVyaW5nQ29udGV4dDJEfWZ1bmN0aW9uIG4oKXt2"
    "YXIgYT0hMSxiPW5hdmlnYXRvci51c2VyQWdlbnQ7cmV0dXJuL2FuZHJvaWQvaS50ZXN0KGIpJiYoYT0hMCxhTWF0PWIudG9TdHJp"
    "bmcoKS5tYXRjaCgvYW5kcm9pZCAoWzAtOV1cLlswLTldKS9pKSxhTWF0JiZhTWF0WzFdJiYoYT1wYXJzZUZsb2F0KGFNYXRbMV0p"
    "KSksYX1mdW5jdGlvbiByKGEsYil7Zm9yKHZhciBjPTEsZT1zKGEpLGY9MCxnPWwubGVuZ3RoO2c+PWY7ZisrKXt2YXIgaD0wO3N3"
    "aXRjaChiKXtjYXNlIGQuTDpoPWxbZl1bMF07YnJlYWs7Y2FzZSBkLk06aD1sW2ZdWzFdO2JyZWFrO2Nhc2UgZC5ROmg9bFtmXVsy"
    "XTticmVhaztjYXNlIGQuSDpoPWxbZl1bM119aWYoaD49ZSlicmVhaztjKyt9aWYoYz5sLmxlbmd0aCl0aHJvdyBuZXcgRXJyb3Io"
    "IlRvbyBsb25nIGRhdGEiKTtyZXR1cm4gY31mdW5jdGlvbiBzKGEpe3ZhciBiPWVuY29kZVVSSShhKS50b1N0cmluZygpLnJlcGxh"
    "Y2UoL1wlWzAtOWEtZkEtRl17Mn0vZywiYSIpO3JldHVybiBiLmxlbmd0aCsoYi5sZW5ndGghPWE/MzowKX1hLnByb3RvdHlwZT17"
    "Z2V0TGVuZ3RoOmZ1bmN0aW9uKCl7cmV0dXJuIHRoaXMucGFyc2VkRGF0YS5sZW5ndGh9LHdyaXRlOmZ1bmN0aW9uKGEpe2Zvcih2"
    "YXIgYj0wLGM9dGhpcy5wYXJzZWREYXRhLmxlbmd0aDtjPmI7YisrKWEucHV0KHRoaXMucGFyc2VkRGF0YVtiXSw4KX19LGIucHJv"
    "dG90eXBlPXthZGREYXRhOmZ1bmN0aW9uKGIpe3ZhciBjPW5ldyBhKGIpO3RoaXMuZGF0YUxpc3QucHVzaChjKSx0aGlzLmRhdGFD"
    "YWNoZT1udWxsfSxpc0Rhcms6ZnVuY3Rpb24oYSxiKXtpZigwPmF8fHRoaXMubW9kdWxlQ291bnQ8PWF8fDA+Ynx8dGhpcy5tb2R1"
    "bGVDb3VudDw9Yil0aHJvdyBuZXcgRXJyb3IoYSsiLCIrYik7cmV0dXJuIHRoaXMubW9kdWxlc1thXVtiXX0sZ2V0TW9kdWxlQ291"
    "bnQ6ZnVuY3Rpb24oKXtyZXR1cm4gdGhpcy5tb2R1bGVDb3VudH0sbWFrZTpmdW5jdGlvbigpe3RoaXMubWFrZUltcGwoITEsdGhp"
    "cy5nZXRCZXN0TWFza1BhdHRlcm4oKSl9LG1ha2VJbXBsOmZ1bmN0aW9uKGEsYyl7dGhpcy5tb2R1bGVDb3VudD00KnRoaXMudHlw"
    "ZU51bWJlcisxNyx0aGlzLm1vZHVsZXM9bmV3IEFycmF5KHRoaXMubW9kdWxlQ291bnQpO2Zvcih2YXIgZD0wO2Q8dGhpcy5tb2R1"
    "bGVDb3VudDtkKyspe3RoaXMubW9kdWxlc1tkXT1uZXcgQXJyYXkodGhpcy5tb2R1bGVDb3VudCk7Zm9yKHZhciBlPTA7ZTx0aGlz"
    "Lm1vZHVsZUNvdW50O2UrKyl0aGlzLm1vZHVsZXNbZF1bZV09bnVsbH10aGlzLnNldHVwUG9zaXRpb25Qcm9iZVBhdHRlcm4oMCww"
    "KSx0aGlzLnNldHVwUG9zaXRpb25Qcm9iZVBhdHRlcm4odGhpcy5tb2R1bGVDb3VudC03LDApLHRoaXMuc2V0dXBQb3NpdGlvblBy"
    "b2JlUGF0dGVybigwLHRoaXMubW9kdWxlQ291bnQtNyksdGhpcy5zZXR1cFBvc2l0aW9uQWRqdXN0UGF0dGVybigpLHRoaXMuc2V0"
    "dXBUaW1pbmdQYXR0ZXJuKCksdGhpcy5zZXR1cFR5cGVJbmZvKGEsYyksdGhpcy50eXBlTnVtYmVyPj03JiZ0aGlzLnNldHVwVHlw"
    "ZU51bWJlcihhKSxudWxsPT10aGlzLmRhdGFDYWNoZSYmKHRoaXMuZGF0YUNhY2hlPWIuY3JlYXRlRGF0YSh0aGlzLnR5cGVOdW1i"
    "ZXIsdGhpcy5lcnJvckNvcnJlY3RMZXZlbCx0aGlzLmRhdGFMaXN0KSksdGhpcy5tYXBEYXRhKHRoaXMuZGF0YUNhY2hlLGMpfSxz"
    "ZXR1cFBvc2l0aW9uUHJvYmVQYXR0ZXJuOmZ1bmN0aW9uKGEsYil7Zm9yKHZhciBjPS0xOzc+PWM7YysrKWlmKCEoLTE+PWErY3x8"
    "dGhpcy5tb2R1bGVDb3VudDw9YStjKSlmb3IodmFyIGQ9LTE7Nz49ZDtkKyspLTE+PWIrZHx8dGhpcy5tb2R1bGVDb3VudDw9Yitk"
    "fHwodGhpcy5tb2R1bGVzW2ErY11bYitkXT1jPj0wJiY2Pj1jJiYoMD09ZHx8Nj09ZCl8fGQ+PTAmJjY+PWQmJigwPT1jfHw2PT1j"
    "KXx8Yz49MiYmND49YyYmZD49MiYmND49ZD8hMDohMSl9LGdldEJlc3RNYXNrUGF0dGVybjpmdW5jdGlvbigpe2Zvcih2YXIgYT0w"
    "LGI9MCxjPTA7OD5jO2MrKyl7dGhpcy5tYWtlSW1wbCghMCxjKTt2YXIgZD1mLmdldExvc3RQb2ludCh0aGlzKTsoMD09Y3x8YT5k"
    "KSYmKGE9ZCxiPWMpfXJldHVybiBifSxjcmVhdGVNb3ZpZUNsaXA6ZnVuY3Rpb24oYSxiLGMpe3ZhciBkPWEuY3JlYXRlRW1wdHlN"
    "b3ZpZUNsaXAoYixjKSxlPTE7dGhpcy5tYWtlKCk7Zm9yKHZhciBmPTA7Zjx0aGlzLm1vZHVsZXMubGVuZ3RoO2YrKylmb3IodmFy"
    "IGc9ZiplLGg9MDtoPHRoaXMubW9kdWxlc1tmXS5sZW5ndGg7aCsrKXt2YXIgaT1oKmUsaj10aGlzLm1vZHVsZXNbZl1baF07aiYm"
    "KGQuYmVnaW5GaWxsKDAsMTAwKSxkLm1vdmVUbyhpLGcpLGQubGluZVRvKGkrZSxnKSxkLmxpbmVUbyhpK2UsZytlKSxkLmxpbmVU"
    "byhpLGcrZSksZC5lbmRGaWxsKCkpfXJldHVybiBkfSxzZXR1cFRpbWluZ1BhdHRlcm46ZnVuY3Rpb24oKXtmb3IodmFyIGE9ODth"
    "PHRoaXMubW9kdWxlQ291bnQtODthKyspbnVsbD09dGhpcy5tb2R1bGVzW2FdWzZdJiYodGhpcy5tb2R1bGVzW2FdWzZdPTA9PWEl"
    "Mik7Zm9yKHZhciBiPTg7Yjx0aGlzLm1vZHVsZUNvdW50LTg7YisrKW51bGw9PXRoaXMubW9kdWxlc1s2XVtiXSYmKHRoaXMubW9k"
    "dWxlc1s2XVtiXT0wPT1iJTIpfSxzZXR1cFBvc2l0aW9uQWRqdXN0UGF0dGVybjpmdW5jdGlvbigpe2Zvcih2YXIgYT1mLmdldFBh"
    "dHRlcm5Qb3NpdGlvbih0aGlzLnR5cGVOdW1iZXIpLGI9MDtiPGEubGVuZ3RoO2IrKylmb3IodmFyIGM9MDtjPGEubGVuZ3RoO2Mr"
    "Kyl7dmFyIGQ9YVtiXSxlPWFbY107aWYobnVsbD09dGhpcy5tb2R1bGVzW2RdW2VdKWZvcih2YXIgZz0tMjsyPj1nO2crKylmb3Io"
    "dmFyIGg9LTI7Mj49aDtoKyspdGhpcy5tb2R1bGVzW2QrZ11bZStoXT0tMj09Z3x8Mj09Z3x8LTI9PWh8fDI9PWh8fDA9PWcmJjA9"
    "PWg/ITA6ITF9fSxzZXR1cFR5cGVOdW1iZXI6ZnVuY3Rpb24oYSl7Zm9yKHZhciBiPWYuZ2V0QkNIVHlwZU51bWJlcih0aGlzLnR5"
    "cGVOdW1iZXIpLGM9MDsxOD5jO2MrKyl7dmFyIGQ9IWEmJjE9PSgxJmI+PmMpO3RoaXMubW9kdWxlc1tNYXRoLmZsb29yKGMvMyld"
    "W2MlMyt0aGlzLm1vZHVsZUNvdW50LTgtM109ZH1mb3IodmFyIGM9MDsxOD5jO2MrKyl7dmFyIGQ9IWEmJjE9PSgxJmI+PmMpO3Ro"
    "aXMubW9kdWxlc1tjJTMrdGhpcy5tb2R1bGVDb3VudC04LTNdW01hdGguZmxvb3IoYy8zKV09ZH19LHNldHVwVHlwZUluZm86ZnVu"
    "Y3Rpb24oYSxiKXtmb3IodmFyIGM9dGhpcy5lcnJvckNvcnJlY3RMZXZlbDw8M3xiLGQ9Zi5nZXRCQ0hUeXBlSW5mbyhjKSxlPTA7"
    "MTU+ZTtlKyspe3ZhciBnPSFhJiYxPT0oMSZkPj5lKTs2PmU/dGhpcy5tb2R1bGVzW2VdWzhdPWc6OD5lP3RoaXMubW9kdWxlc1tl"
    "KzFdWzhdPWc6dGhpcy5tb2R1bGVzW3RoaXMubW9kdWxlQ291bnQtMTUrZV1bOF09Z31mb3IodmFyIGU9MDsxNT5lO2UrKyl7dmFy"
    "IGc9IWEmJjE9PSgxJmQ+PmUpOzg+ZT90aGlzLm1vZHVsZXNbOF1bdGhpcy5tb2R1bGVDb3VudC1lLTFdPWc6OT5lP3RoaXMubW9k"
    "dWxlc1s4XVsxNS1lLTErMV09Zzp0aGlzLm1vZHVsZXNbOF1bMTUtZS0xXT1nfXRoaXMubW9kdWxlc1t0aGlzLm1vZHVsZUNvdW50"
    "LThdWzhdPSFhfSxtYXBEYXRhOmZ1bmN0aW9uKGEsYil7Zm9yKHZhciBjPS0xLGQ9dGhpcy5tb2R1bGVDb3VudC0xLGU9NyxnPTAs"
    "aD10aGlzLm1vZHVsZUNvdW50LTE7aD4wO2gtPTIpZm9yKDY9PWgmJmgtLTs7KXtmb3IodmFyIGk9MDsyPmk7aSsrKWlmKG51bGw9"
    "PXRoaXMubW9kdWxlc1tkXVtoLWldKXt2YXIgaj0hMTtnPGEubGVuZ3RoJiYoaj0xPT0oMSZhW2ddPj4+ZSkpO3ZhciBrPWYuZ2V0"
    "TWFzayhiLGQsaC1pKTtrJiYoaj0haiksdGhpcy5tb2R1bGVzW2RdW2gtaV09aixlLS0sLTE9PWUmJihnKyssZT03KX1pZihkKz1j"
    "LDA+ZHx8dGhpcy5tb2R1bGVDb3VudDw9ZCl7ZC09YyxjPS1jO2JyZWFrfX19fSxiLlBBRDA9MjM2LGIuUEFEMT0xNyxiLmNyZWF0"
    "ZURhdGE9ZnVuY3Rpb24oYSxjLGQpe2Zvcih2YXIgZT1qLmdldFJTQmxvY2tzKGEsYyksZz1uZXcgayxoPTA7aDxkLmxlbmd0aDto"
    "Kyspe3ZhciBpPWRbaF07Zy5wdXQoaS5tb2RlLDQpLGcucHV0KGkuZ2V0TGVuZ3RoKCksZi5nZXRMZW5ndGhJbkJpdHMoaS5tb2Rl"
    "LGEpKSxpLndyaXRlKGcpfWZvcih2YXIgbD0wLGg9MDtoPGUubGVuZ3RoO2grKylsKz1lW2hdLmRhdGFDb3VudDtpZihnLmdldExl"
    "bmd0aEluQml0cygpPjgqbCl0aHJvdyBuZXcgRXJyb3IoImNvZGUgbGVuZ3RoIG92ZXJmbG93LiAoIitnLmdldExlbmd0aEluQml0"
    "cygpKyI+Iis4KmwrIikiKTtmb3IoZy5nZXRMZW5ndGhJbkJpdHMoKSs0PD04KmwmJmcucHV0KDAsNCk7MCE9Zy5nZXRMZW5ndGhJ"
    "bkJpdHMoKSU4OylnLnB1dEJpdCghMSk7Zm9yKDs7KXtpZihnLmdldExlbmd0aEluQml0cygpPj04KmwpYnJlYWs7aWYoZy5wdXQo"
    "Yi5QQUQwLDgpLGcuZ2V0TGVuZ3RoSW5CaXRzKCk+PTgqbClicmVhaztnLnB1dChiLlBBRDEsOCl9cmV0dXJuIGIuY3JlYXRlQnl0"
    "ZXMoZyxlKX0sYi5jcmVhdGVCeXRlcz1mdW5jdGlvbihhLGIpe2Zvcih2YXIgYz0wLGQ9MCxlPTAsZz1uZXcgQXJyYXkoYi5sZW5n"
    "dGgpLGg9bmV3IEFycmF5KGIubGVuZ3RoKSxqPTA7ajxiLmxlbmd0aDtqKyspe3ZhciBrPWJbal0uZGF0YUNvdW50LGw9YltqXS50"
    "b3RhbENvdW50LWs7ZD1NYXRoLm1heChkLGspLGU9TWF0aC5tYXgoZSxsKSxnW2pdPW5ldyBBcnJheShrKTtmb3IodmFyIG09MDtt"
    "PGdbal0ubGVuZ3RoO20rKylnW2pdW21dPTI1NSZhLmJ1ZmZlclttK2NdO2MrPWs7dmFyIG49Zi5nZXRFcnJvckNvcnJlY3RQb2x5"
    "bm9taWFsKGwpLG89bmV3IGkoZ1tqXSxuLmdldExlbmd0aCgpLTEpLHA9by5tb2Qobik7aFtqXT1uZXcgQXJyYXkobi5nZXRMZW5n"
    "dGgoKS0xKTtmb3IodmFyIG09MDttPGhbal0ubGVuZ3RoO20rKyl7dmFyIHE9bStwLmdldExlbmd0aCgpLWhbal0ubGVuZ3RoO2hb"
    "al1bbV09cT49MD9wLmdldChxKTowfX1mb3IodmFyIHI9MCxtPTA7bTxiLmxlbmd0aDttKyspcis9YlttXS50b3RhbENvdW50O2Zv"
    "cih2YXIgcz1uZXcgQXJyYXkociksdD0wLG09MDtkPm07bSsrKWZvcih2YXIgaj0wO2o8Yi5sZW5ndGg7aisrKW08Z1tqXS5sZW5n"
    "dGgmJihzW3QrK109Z1tqXVttXSk7Zm9yKHZhciBtPTA7ZT5tO20rKylmb3IodmFyIGo9MDtqPGIubGVuZ3RoO2orKyltPGhbal0u"
    "bGVuZ3RoJiYoc1t0KytdPWhbal1bbV0pO3JldHVybiBzfTtmb3IodmFyIGM9e01PREVfTlVNQkVSOjEsTU9ERV9BTFBIQV9OVU06"
    "MixNT0RFXzhCSVRfQllURTo0LE1PREVfS0FOSkk6OH0sZD17TDoxLE06MCxROjMsSDoyfSxlPXtQQVRURVJOMDAwOjAsUEFUVEVS"
    "TjAwMToxLFBBVFRFUk4wMTA6MixQQVRURVJOMDExOjMsUEFUVEVSTjEwMDo0LFBBVFRFUk4xMDE6NSxQQVRURVJOMTEwOjYsUEFU"
    "VEVSTjExMTo3fSxmPXtQQVRURVJOX1BPU0lUSU9OX1RBQkxFOltbXSxbNiwxOF0sWzYsMjJdLFs2LDI2XSxbNiwzMF0sWzYsMzRd"
    "LFs2LDIyLDM4XSxbNiwyNCw0Ml0sWzYsMjYsNDZdLFs2LDI4LDUwXSxbNiwzMCw1NF0sWzYsMzIsNThdLFs2LDM0LDYyXSxbNiwy"
    "Niw0Niw2Nl0sWzYsMjYsNDgsNzBdLFs2LDI2LDUwLDc0XSxbNiwzMCw1NCw3OF0sWzYsMzAsNTYsODJdLFs2LDMwLDU4LDg2XSxb"
    "NiwzNCw2Miw5MF0sWzYsMjgsNTAsNzIsOTRdLFs2LDI2LDUwLDc0LDk4XSxbNiwzMCw1NCw3OCwxMDJdLFs2LDI4LDU0LDgwLDEw"
    "Nl0sWzYsMzIsNTgsODQsMTEwXSxbNiwzMCw1OCw4NiwxMTRdLFs2LDM0LDYyLDkwLDExOF0sWzYsMjYsNTAsNzQsOTgsMTIyXSxb"
    "NiwzMCw1NCw3OCwxMDIsMTI2XSxbNiwyNiw1Miw3OCwxMDQsMTMwXSxbNiwzMCw1Niw4MiwxMDgsMTM0XSxbNiwzNCw2MCw4Niwx"
    "MTIsMTM4XSxbNiwzMCw1OCw4NiwxMTQsMTQyXSxbNiwzNCw2Miw5MCwxMTgsMTQ2XSxbNiwzMCw1NCw3OCwxMDIsMTI2LDE1MF0s"
    "WzYsMjQsNTAsNzYsMTAyLDEyOCwxNTRdLFs2LDI4LDU0LDgwLDEwNiwxMzIsMTU4XSxbNiwzMiw1OCw4NCwxMTAsMTM2LDE2Ml0s"
    "WzYsMjYsNTQsODIsMTEwLDEzOCwxNjZdLFs2LDMwLDU4LDg2LDExNCwxNDIsMTcwXV0sRzE1OjEzMzUsRzE4Ojc5NzMsRzE1X01B"
    "U0s6MjE1MjIsZ2V0QkNIVHlwZUluZm86ZnVuY3Rpb24oYSl7Zm9yKHZhciBiPWE8PDEwO2YuZ2V0QkNIRGlnaXQoYiktZi5nZXRC"
    "Q0hEaWdpdChmLkcxNSk+PTA7KWJePWYuRzE1PDxmLmdldEJDSERpZ2l0KGIpLWYuZ2V0QkNIRGlnaXQoZi5HMTUpO3JldHVybihh"
    "PDwxMHxiKV5mLkcxNV9NQVNLfSxnZXRCQ0hUeXBlTnVtYmVyOmZ1bmN0aW9uKGEpe2Zvcih2YXIgYj1hPDwxMjtmLmdldEJDSERp"
    "Z2l0KGIpLWYuZ2V0QkNIRGlnaXQoZi5HMTgpPj0wOyliXj1mLkcxODw8Zi5nZXRCQ0hEaWdpdChiKS1mLmdldEJDSERpZ2l0KGYu"
    "RzE4KTtyZXR1cm4gYTw8MTJ8Yn0sZ2V0QkNIRGlnaXQ6ZnVuY3Rpb24oYSl7Zm9yKHZhciBiPTA7MCE9YTspYisrLGE+Pj49MTty"
    "ZXR1cm4gYn0sZ2V0UGF0dGVyblBvc2l0aW9uOmZ1bmN0aW9uKGEpe3JldHVybiBmLlBBVFRFUk5fUE9TSVRJT05fVEFCTEVbYS0x"
    "XX0sZ2V0TWFzazpmdW5jdGlvbihhLGIsYyl7c3dpdGNoKGEpe2Nhc2UgZS5QQVRURVJOMDAwOnJldHVybiAwPT0oYitjKSUyO2Nh"
    "c2UgZS5QQVRURVJOMDAxOnJldHVybiAwPT1iJTI7Y2FzZSBlLlBBVFRFUk4wMTA6cmV0dXJuIDA9PWMlMztjYXNlIGUuUEFUVEVS"
    "TjAxMTpyZXR1cm4gMD09KGIrYyklMztjYXNlIGUuUEFUVEVSTjEwMDpyZXR1cm4gMD09KE1hdGguZmxvb3IoYi8yKStNYXRoLmZs"
    "b29yKGMvMykpJTI7Y2FzZSBlLlBBVFRFUk4xMDE6cmV0dXJuIDA9PWIqYyUyK2IqYyUzO2Nhc2UgZS5QQVRURVJOMTEwOnJldHVy"
    "biAwPT0oYipjJTIrYipjJTMpJTI7Y2FzZSBlLlBBVFRFUk4xMTE6cmV0dXJuIDA9PShiKmMlMysoYitjKSUyKSUyO2RlZmF1bHQ6"
    "dGhyb3cgbmV3IEVycm9yKCJiYWQgbWFza1BhdHRlcm46IithKX19LGdldEVycm9yQ29ycmVjdFBvbHlub21pYWw6ZnVuY3Rpb24o"
    "YSl7Zm9yKHZhciBiPW5ldyBpKFsxXSwwKSxjPTA7YT5jO2MrKyliPWIubXVsdGlwbHkobmV3IGkoWzEsZy5nZXhwKGMpXSwwKSk7"
    "cmV0dXJuIGJ9LGdldExlbmd0aEluQml0czpmdW5jdGlvbihhLGIpe2lmKGI+PTEmJjEwPmIpc3dpdGNoKGEpe2Nhc2UgYy5NT0RF"
    "X05VTUJFUjpyZXR1cm4gMTA7Y2FzZSBjLk1PREVfQUxQSEFfTlVNOnJldHVybiA5O2Nhc2UgYy5NT0RFXzhCSVRfQllURTpyZXR1"
    "cm4gODtjYXNlIGMuTU9ERV9LQU5KSTpyZXR1cm4gODtkZWZhdWx0OnRocm93IG5ldyBFcnJvcigibW9kZToiK2EpfWVsc2UgaWYo"
    "Mjc+Yilzd2l0Y2goYSl7Y2FzZSBjLk1PREVfTlVNQkVSOnJldHVybiAxMjtjYXNlIGMuTU9ERV9BTFBIQV9OVU06cmV0dXJuIDEx"
    "O2Nhc2UgYy5NT0RFXzhCSVRfQllURTpyZXR1cm4gMTY7Y2FzZSBjLk1PREVfS0FOSkk6cmV0dXJuIDEwO2RlZmF1bHQ6dGhyb3cg"
    "bmV3IEVycm9yKCJtb2RlOiIrYSl9ZWxzZXtpZighKDQxPmIpKXRocm93IG5ldyBFcnJvcigidHlwZToiK2IpO3N3aXRjaChhKXtj"
    "YXNlIGMuTU9ERV9OVU1CRVI6cmV0dXJuIDE0O2Nhc2UgYy5NT0RFX0FMUEhBX05VTTpyZXR1cm4gMTM7Y2FzZSBjLk1PREVfOEJJ"
    "VF9CWVRFOnJldHVybiAxNjtjYXNlIGMuTU9ERV9LQU5KSTpyZXR1cm4gMTI7ZGVmYXVsdDp0aHJvdyBuZXcgRXJyb3IoIm1vZGU6"
    "IithKX19fSxnZXRMb3N0UG9pbnQ6ZnVuY3Rpb24oYSl7Zm9yKHZhciBiPWEuZ2V0TW9kdWxlQ291bnQoKSxjPTAsZD0wO2I+ZDtk"
    "KyspZm9yKHZhciBlPTA7Yj5lO2UrKyl7Zm9yKHZhciBmPTAsZz1hLmlzRGFyayhkLGUpLGg9LTE7MT49aDtoKyspaWYoISgwPmQr"
    "aHx8ZCtoPj1iKSlmb3IodmFyIGk9LTE7MT49aTtpKyspMD5lK2l8fGUraT49Ynx8KDAhPWh8fDAhPWkpJiZnPT1hLmlzRGFyayhk"
    "K2gsZStpKSYmZisrO2Y+NSYmKGMrPTMrZi01KX1mb3IodmFyIGQ9MDtiLTE+ZDtkKyspZm9yKHZhciBlPTA7Yi0xPmU7ZSsrKXt2"
    "YXIgaj0wO2EuaXNEYXJrKGQsZSkmJmorKyxhLmlzRGFyayhkKzEsZSkmJmorKyxhLmlzRGFyayhkLGUrMSkmJmorKyxhLmlzRGFy"
    "ayhkKzEsZSsxKSYmaisrLCgwPT1qfHw0PT1qKSYmKGMrPTMpfWZvcih2YXIgZD0wO2I+ZDtkKyspZm9yKHZhciBlPTA7Yi02PmU7"
    "ZSsrKWEuaXNEYXJrKGQsZSkmJiFhLmlzRGFyayhkLGUrMSkmJmEuaXNEYXJrKGQsZSsyKSYmYS5pc0RhcmsoZCxlKzMpJiZhLmlz"
    "RGFyayhkLGUrNCkmJiFhLmlzRGFyayhkLGUrNSkmJmEuaXNEYXJrKGQsZSs2KSYmKGMrPTQwKTtmb3IodmFyIGU9MDtiPmU7ZSsr"
    "KWZvcih2YXIgZD0wO2ItNj5kO2QrKylhLmlzRGFyayhkLGUpJiYhYS5pc0RhcmsoZCsxLGUpJiZhLmlzRGFyayhkKzIsZSkmJmEu"
    "aXNEYXJrKGQrMyxlKSYmYS5pc0RhcmsoZCs0LGUpJiYhYS5pc0RhcmsoZCs1LGUpJiZhLmlzRGFyayhkKzYsZSkmJihjKz00MCk7"
    "Zm9yKHZhciBrPTAsZT0wO2I+ZTtlKyspZm9yKHZhciBkPTA7Yj5kO2QrKylhLmlzRGFyayhkLGUpJiZrKys7dmFyIGw9TWF0aC5h"
    "YnMoMTAwKmsvYi9iLTUwKS81O3JldHVybiBjKz0xMCpsfX0sZz17Z2xvZzpmdW5jdGlvbihhKXtpZigxPmEpdGhyb3cgbmV3IEVy"
    "cm9yKCJnbG9nKCIrYSsiKSIpO3JldHVybiBnLkxPR19UQUJMRVthXX0sZ2V4cDpmdW5jdGlvbihhKXtmb3IoOzA+YTspYSs9MjU1"
    "O2Zvcig7YT49MjU2OylhLT0yNTU7cmV0dXJuIGcuRVhQX1RBQkxFW2FdfSxFWFBfVEFCTEU6bmV3IEFycmF5KDI1NiksTE9HX1RB"
    "QkxFOm5ldyBBcnJheSgyNTYpfSxoPTA7OD5oO2grKylnLkVYUF9UQUJMRVtoXT0xPDxoO2Zvcih2YXIgaD04OzI1Nj5oO2grKyln"
    "LkVYUF9UQUJMRVtoXT1nLkVYUF9UQUJMRVtoLTRdXmcuRVhQX1RBQkxFW2gtNV1eZy5FWFBfVEFCTEVbaC02XV5nLkVYUF9UQUJM"
    "RVtoLThdO2Zvcih2YXIgaD0wOzI1NT5oO2grKylnLkxPR19UQUJMRVtnLkVYUF9UQUJMRVtoXV09aDtpLnByb3RvdHlwZT17Z2V0"
    "OmZ1bmN0aW9uKGEpe3JldHVybiB0aGlzLm51bVthXX0sZ2V0TGVuZ3RoOmZ1bmN0aW9uKCl7cmV0dXJuIHRoaXMubnVtLmxlbmd0"
    "aH0sbXVsdGlwbHk6ZnVuY3Rpb24oYSl7Zm9yKHZhciBiPW5ldyBBcnJheSh0aGlzLmdldExlbmd0aCgpK2EuZ2V0TGVuZ3RoKCkt"
    "MSksYz0wO2M8dGhpcy5nZXRMZW5ndGgoKTtjKyspZm9yKHZhciBkPTA7ZDxhLmdldExlbmd0aCgpO2QrKyliW2MrZF1ePWcuZ2V4"
    "cChnLmdsb2codGhpcy5nZXQoYykpK2cuZ2xvZyhhLmdldChkKSkpO3JldHVybiBuZXcgaShiLDApfSxtb2Q6ZnVuY3Rpb24oYSl7"
    "aWYodGhpcy5nZXRMZW5ndGgoKS1hLmdldExlbmd0aCgpPDApcmV0dXJuIHRoaXM7Zm9yKHZhciBiPWcuZ2xvZyh0aGlzLmdldCgw"
    "KSktZy5nbG9nKGEuZ2V0KDApKSxjPW5ldyBBcnJheSh0aGlzLmdldExlbmd0aCgpKSxkPTA7ZDx0aGlzLmdldExlbmd0aCgpO2Qr"
    "KyljW2RdPXRoaXMuZ2V0KGQpO2Zvcih2YXIgZD0wO2Q8YS5nZXRMZW5ndGgoKTtkKyspY1tkXV49Zy5nZXhwKGcuZ2xvZyhhLmdl"
    "dChkKSkrYik7cmV0dXJuIG5ldyBpKGMsMCkubW9kKGEpfX0sai5SU19CTE9DS19UQUJMRT1bWzEsMjYsMTldLFsxLDI2LDE2XSxb"
    "MSwyNiwxM10sWzEsMjYsOV0sWzEsNDQsMzRdLFsxLDQ0LDI4XSxbMSw0NCwyMl0sWzEsNDQsMTZdLFsxLDcwLDU1XSxbMSw3MCw0"
    "NF0sWzIsMzUsMTddLFsyLDM1LDEzXSxbMSwxMDAsODBdLFsyLDUwLDMyXSxbMiw1MCwyNF0sWzQsMjUsOV0sWzEsMTM0LDEwOF0s"
    "WzIsNjcsNDNdLFsyLDMzLDE1LDIsMzQsMTZdLFsyLDMzLDExLDIsMzQsMTJdLFsyLDg2LDY4XSxbNCw0MywyN10sWzQsNDMsMTld"
    "LFs0LDQzLDE1XSxbMiw5OCw3OF0sWzQsNDksMzFdLFsyLDMyLDE0LDQsMzMsMTVdLFs0LDM5LDEzLDEsNDAsMTRdLFsyLDEyMSw5"
    "N10sWzIsNjAsMzgsMiw2MSwzOV0sWzQsNDAsMTgsMiw0MSwxOV0sWzQsNDAsMTQsMiw0MSwxNV0sWzIsMTQ2LDExNl0sWzMsNTgs"
    "MzYsMiw1OSwzN10sWzQsMzYsMTYsNCwzNywxN10sWzQsMzYsMTIsNCwzNywxM10sWzIsODYsNjgsMiw4Nyw2OV0sWzQsNjksNDMs"
    "MSw3MCw0NF0sWzYsNDMsMTksMiw0NCwyMF0sWzYsNDMsMTUsMiw0NCwxNl0sWzQsMTAxLDgxXSxbMSw4MCw1MCw0LDgxLDUxXSxb"
    "NCw1MCwyMiw0LDUxLDIzXSxbMywzNiwxMiw4LDM3LDEzXSxbMiwxMTYsOTIsMiwxMTcsOTNdLFs2LDU4LDM2LDIsNTksMzddLFs0"
    "LDQ2LDIwLDYsNDcsMjFdLFs3LDQyLDE0LDQsNDMsMTVdLFs0LDEzMywxMDddLFs4LDU5LDM3LDEsNjAsMzhdLFs4LDQ0LDIwLDQs"
    "NDUsMjFdLFsxMiwzMywxMSw0LDM0LDEyXSxbMywxNDUsMTE1LDEsMTQ2LDExNl0sWzQsNjQsNDAsNSw2NSw0MV0sWzExLDM2LDE2"
    "LDUsMzcsMTddLFsxMSwzNiwxMiw1LDM3LDEzXSxbNSwxMDksODcsMSwxMTAsODhdLFs1LDY1LDQxLDUsNjYsNDJdLFs1LDU0LDI0"
    "LDcsNTUsMjVdLFsxMSwzNiwxMl0sWzUsMTIyLDk4LDEsMTIzLDk5XSxbNyw3Myw0NSwzLDc0LDQ2XSxbMTUsNDMsMTksMiw0NCwy"
    "MF0sWzMsNDUsMTUsMTMsNDYsMTZdLFsxLDEzNSwxMDcsNSwxMzYsMTA4XSxbMTAsNzQsNDYsMSw3NSw0N10sWzEsNTAsMjIsMTUs"
    "NTEsMjNdLFsyLDQyLDE0LDE3LDQzLDE1XSxbNSwxNTAsMTIwLDEsMTUxLDEyMV0sWzksNjksNDMsNCw3MCw0NF0sWzE3LDUwLDIy"
    "LDEsNTEsMjNdLFsyLDQyLDE0LDE5LDQzLDE1XSxbMywxNDEsMTEzLDQsMTQyLDExNF0sWzMsNzAsNDQsMTEsNzEsNDVdLFsxNyw0"
    "NywyMSw0LDQ4LDIyXSxbOSwzOSwxMywxNiw0MCwxNF0sWzMsMTM1LDEwNyw1LDEzNiwxMDhdLFszLDY3LDQxLDEzLDY4LDQyXSxb"
    "MTUsNTQsMjQsNSw1NSwyNV0sWzE1LDQzLDE1LDEwLDQ0LDE2XSxbNCwxNDQsMTE2LDQsMTQ1LDExN10sWzE3LDY4LDQyXSxbMTcs"
    "NTAsMjIsNiw1MSwyM10sWzE5LDQ2LDE2LDYsNDcsMTddLFsyLDEzOSwxMTEsNywxNDAsMTEyXSxbMTcsNzQsNDZdLFs3LDU0LDI0"
    "LDE2LDU1LDI1XSxbMzQsMzcsMTNdLFs0LDE1MSwxMjEsNSwxNTIsMTIyXSxbNCw3NSw0NywxNCw3Niw0OF0sWzExLDU0LDI0LDE0"
    "LDU1LDI1XSxbMTYsNDUsMTUsMTQsNDYsMTZdLFs2LDE0NywxMTcsNCwxNDgsMTE4XSxbNiw3Myw0NSwxNCw3NCw0Nl0sWzExLDU0"
    "LDI0LDE2LDU1LDI1XSxbMzAsNDYsMTYsMiw0NywxN10sWzgsMTMyLDEwNiw0LDEzMywxMDddLFs4LDc1LDQ3LDEzLDc2LDQ4XSxb"
    "Nyw1NCwyNCwyMiw1NSwyNV0sWzIyLDQ1LDE1LDEzLDQ2LDE2XSxbMTAsMTQyLDExNCwyLDE0MywxMTVdLFsxOSw3NCw0Niw0LDc1"
    "LDQ3XSxbMjgsNTAsMjIsNiw1MSwyM10sWzMzLDQ2LDE2LDQsNDcsMTddLFs4LDE1MiwxMjIsNCwxNTMsMTIzXSxbMjIsNzMsNDUs"
    "Myw3NCw0Nl0sWzgsNTMsMjMsMjYsNTQsMjRdLFsxMiw0NSwxNSwyOCw0NiwxNl0sWzMsMTQ3LDExNywxMCwxNDgsMTE4XSxbMyw3"
    "Myw0NSwyMyw3NCw0Nl0sWzQsNTQsMjQsMzEsNTUsMjVdLFsxMSw0NSwxNSwzMSw0NiwxNl0sWzcsMTQ2LDExNiw3LDE0NywxMTdd"
    "LFsyMSw3Myw0NSw3LDc0LDQ2XSxbMSw1MywyMywzNyw1NCwyNF0sWzE5LDQ1LDE1LDI2LDQ2LDE2XSxbNSwxNDUsMTE1LDEwLDE0"
    "NiwxMTZdLFsxOSw3NSw0NywxMCw3Niw0OF0sWzE1LDU0LDI0LDI1LDU1LDI1XSxbMjMsNDUsMTUsMjUsNDYsMTZdLFsxMywxNDUs"
    "MTE1LDMsMTQ2LDExNl0sWzIsNzQsNDYsMjksNzUsNDddLFs0Miw1NCwyNCwxLDU1LDI1XSxbMjMsNDUsMTUsMjgsNDYsMTZdLFsx"
    "NywxNDUsMTE1XSxbMTAsNzQsNDYsMjMsNzUsNDddLFsxMCw1NCwyNCwzNSw1NSwyNV0sWzE5LDQ1LDE1LDM1LDQ2LDE2XSxbMTcs"
    "MTQ1LDExNSwxLDE0NiwxMTZdLFsxNCw3NCw0NiwyMSw3NSw0N10sWzI5LDU0LDI0LDE5LDU1LDI1XSxbMTEsNDUsMTUsNDYsNDYs"
    "MTZdLFsxMywxNDUsMTE1LDYsMTQ2LDExNl0sWzE0LDc0LDQ2LDIzLDc1LDQ3XSxbNDQsNTQsMjQsNyw1NSwyNV0sWzU5LDQ2LDE2"
    "LDEsNDcsMTddLFsxMiwxNTEsMTIxLDcsMTUyLDEyMl0sWzEyLDc1LDQ3LDI2LDc2LDQ4XSxbMzksNTQsMjQsMTQsNTUsMjVdLFsy"
    "Miw0NSwxNSw0MSw0NiwxNl0sWzYsMTUxLDEyMSwxNCwxNTIsMTIyXSxbNiw3NSw0NywzNCw3Niw0OF0sWzQ2LDU0LDI0LDEwLDU1"
    "LDI1XSxbMiw0NSwxNSw2NCw0NiwxNl0sWzE3LDE1MiwxMjIsNCwxNTMsMTIzXSxbMjksNzQsNDYsMTQsNzUsNDddLFs0OSw1NCwy"
    "NCwxMCw1NSwyNV0sWzI0LDQ1LDE1LDQ2LDQ2LDE2XSxbNCwxNTIsMTIyLDE4LDE1MywxMjNdLFsxMyw3NCw0NiwzMiw3NSw0N10s"
    "WzQ4LDU0LDI0LDE0LDU1LDI1XSxbNDIsNDUsMTUsMzIsNDYsMTZdLFsyMCwxNDcsMTE3LDQsMTQ4LDExOF0sWzQwLDc1LDQ3LDcs"
    "NzYsNDhdLFs0Myw1NCwyNCwyMiw1NSwyNV0sWzEwLDQ1LDE1LDY3LDQ2LDE2XSxbMTksMTQ4LDExOCw2LDE0OSwxMTldLFsxOCw3"
    "NSw0NywzMSw3Niw0OF0sWzM0LDU0LDI0LDM0LDU1LDI1XSxbMjAsNDUsMTUsNjEsNDYsMTZdXSxqLmdldFJTQmxvY2tzPWZ1bmN0"
    "aW9uKGEsYil7dmFyIGM9ai5nZXRSc0Jsb2NrVGFibGUoYSxiKTtpZih2b2lkIDA9PWMpdGhyb3cgbmV3IEVycm9yKCJiYWQgcnMg"
    "YmxvY2sgQCB0eXBlTnVtYmVyOiIrYSsiL2Vycm9yQ29ycmVjdExldmVsOiIrYik7Zm9yKHZhciBkPWMubGVuZ3RoLzMsZT1bXSxm"
    "PTA7ZD5mO2YrKylmb3IodmFyIGc9Y1szKmYrMF0saD1jWzMqZisxXSxpPWNbMypmKzJdLGs9MDtnPms7aysrKWUucHVzaChuZXcg"
    "aihoLGkpKTtyZXR1cm4gZX0sai5nZXRSc0Jsb2NrVGFibGU9ZnVuY3Rpb24oYSxiKXtzd2l0Y2goYil7Y2FzZSBkLkw6cmV0dXJu"
    "IGouUlNfQkxPQ0tfVEFCTEVbNCooYS0xKSswXTtjYXNlIGQuTTpyZXR1cm4gai5SU19CTE9DS19UQUJMRVs0KihhLTEpKzFdO2Nh"
    "c2UgZC5ROnJldHVybiBqLlJTX0JMT0NLX1RBQkxFWzQqKGEtMSkrMl07Y2FzZSBkLkg6cmV0dXJuIGouUlNfQkxPQ0tfVEFCTEVb"
    "NCooYS0xKSszXTtkZWZhdWx0OnJldHVybiB2b2lkIDB9fSxrLnByb3RvdHlwZT17Z2V0OmZ1bmN0aW9uKGEpe3ZhciBiPU1hdGgu"
    "Zmxvb3IoYS84KTtyZXR1cm4gMT09KDEmdGhpcy5idWZmZXJbYl0+Pj43LWElOCl9LHB1dDpmdW5jdGlvbihhLGIpe2Zvcih2YXIg"
    "Yz0wO2I+YztjKyspdGhpcy5wdXRCaXQoMT09KDEmYT4+PmItYy0xKSl9LGdldExlbmd0aEluQml0czpmdW5jdGlvbigpe3JldHVy"
    "biB0aGlzLmxlbmd0aH0scHV0Qml0OmZ1bmN0aW9uKGEpe3ZhciBiPU1hdGguZmxvb3IodGhpcy5sZW5ndGgvOCk7dGhpcy5idWZm"
    "ZXIubGVuZ3RoPD1iJiZ0aGlzLmJ1ZmZlci5wdXNoKDApLGEmJih0aGlzLmJ1ZmZlcltiXXw9MTI4Pj4+dGhpcy5sZW5ndGglOCks"
    "dGhpcy5sZW5ndGgrK319O3ZhciBsPVtbMTcsMTQsMTEsN10sWzMyLDI2LDIwLDE0XSxbNTMsNDIsMzIsMjRdLFs3OCw2Miw0Niwz"
    "NF0sWzEwNiw4NCw2MCw0NF0sWzEzNCwxMDYsNzQsNThdLFsxNTQsMTIyLDg2LDY0XSxbMTkyLDE1MiwxMDgsODRdLFsyMzAsMTgw"
    "LDEzMCw5OF0sWzI3MSwyMTMsMTUxLDExOV0sWzMyMSwyNTEsMTc3LDEzN10sWzM2NywyODcsMjAzLDE1NV0sWzQyNSwzMzEsMjQx"
    "LDE3N10sWzQ1OCwzNjIsMjU4LDE5NF0sWzUyMCw0MTIsMjkyLDIyMF0sWzU4Niw0NTAsMzIyLDI1MF0sWzY0NCw1MDQsMzY0LDI4"
    "MF0sWzcxOCw1NjAsMzk0LDMxMF0sWzc5Miw2MjQsNDQyLDMzOF0sWzg1OCw2NjYsNDgyLDM4Ml0sWzkyOSw3MTEsNTA5LDQwM10s"
    "WzEwMDMsNzc5LDU2NSw0MzldLFsxMDkxLDg1Nyw2MTEsNDYxXSxbMTE3MSw5MTEsNjYxLDUxMV0sWzEyNzMsOTk3LDcxNSw1MzVd"
    "LFsxMzY3LDEwNTksNzUxLDU5M10sWzE0NjUsMTEyNSw4MDUsNjI1XSxbMTUyOCwxMTkwLDg2OCw2NThdLFsxNjI4LDEyNjQsOTA4"
    "LDY5OF0sWzE3MzIsMTM3MCw5ODIsNzQyXSxbMTg0MCwxNDUyLDEwMzAsNzkwXSxbMTk1MiwxNTM4LDExMTIsODQyXSxbMjA2OCwx"
    "NjI4LDExNjgsODk4XSxbMjE4OCwxNzIyLDEyMjgsOTU4XSxbMjMwMywxODA5LDEyODMsOTgzXSxbMjQzMSwxOTExLDEzNTEsMTA1"
    "MV0sWzI1NjMsMTk4OSwxNDIzLDEwOTNdLFsyNjk5LDIwOTksMTQ5OSwxMTM5XSxbMjgwOSwyMjEzLDE1NzksMTIxOV0sWzI5NTMs"
    "MjMzMSwxNjYzLDEyNzNdXSxvPWZ1bmN0aW9uKCl7dmFyIGE9ZnVuY3Rpb24oYSxiKXt0aGlzLl9lbD1hLHRoaXMuX2h0T3B0aW9u"
    "PWJ9O3JldHVybiBhLnByb3RvdHlwZS5kcmF3PWZ1bmN0aW9uKGEpe2Z1bmN0aW9uIGcoYSxiKXt2YXIgYz1kb2N1bWVudC5jcmVh"
    "dGVFbGVtZW50TlMoImh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIixhKTtmb3IodmFyIGQgaW4gYiliLmhhc093blByb3BlcnR5"
    "KGQpJiZjLnNldEF0dHJpYnV0ZShkLGJbZF0pO3JldHVybiBjfXZhciBiPXRoaXMuX2h0T3B0aW9uLGM9dGhpcy5fZWwsZD1hLmdl"
    "dE1vZHVsZUNvdW50KCk7TWF0aC5mbG9vcihiLndpZHRoL2QpLE1hdGguZmxvb3IoYi5oZWlnaHQvZCksdGhpcy5jbGVhcigpO3Zh"
    "ciBoPWcoInN2ZyIse3ZpZXdCb3g6IjAgMCAiK1N0cmluZyhkKSsiICIrU3RyaW5nKGQpLHdpZHRoOiIxMDAlIixoZWlnaHQ6IjEw"
    "MCUiLGZpbGw6Yi5jb2xvckxpZ2h0fSk7aC5zZXRBdHRyaWJ1dGVOUygiaHR0cDovL3d3dy53My5vcmcvMjAwMC94bWxucy8iLCJ4"
    "bWxuczp4bGluayIsImh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiKSxjLmFwcGVuZENoaWxkKGgpLGguYXBwZW5kQ2hpbGQo"
    "ZygicmVjdCIse2ZpbGw6Yi5jb2xvckRhcmssd2lkdGg6IjEiLGhlaWdodDoiMSIsaWQ6InRlbXBsYXRlIn0pKTtmb3IodmFyIGk9"
    "MDtkPmk7aSsrKWZvcih2YXIgaj0wO2Q+ajtqKyspaWYoYS5pc0RhcmsoaSxqKSl7dmFyIGs9ZygidXNlIix7eDpTdHJpbmcoaSks"
    "eTpTdHJpbmcoail9KTtrLnNldEF0dHJpYnV0ZU5TKCJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiwiaHJlZiIsIiN0ZW1w"
    "bGF0ZSIpLGguYXBwZW5kQ2hpbGQoayl9fSxhLnByb3RvdHlwZS5jbGVhcj1mdW5jdGlvbigpe2Zvcig7dGhpcy5fZWwuaGFzQ2hp"
    "bGROb2RlcygpOyl0aGlzLl9lbC5yZW1vdmVDaGlsZCh0aGlzLl9lbC5sYXN0Q2hpbGQpfSxhfSgpLHA9InN2ZyI9PT1kb2N1bWVu"
    "dC5kb2N1bWVudEVsZW1lbnQudGFnTmFtZS50b0xvd2VyQ2FzZSgpLHE9cD9vOm0oKT9mdW5jdGlvbigpe2Z1bmN0aW9uIGEoKXt0"
    "aGlzLl9lbEltYWdlLnNyYz10aGlzLl9lbENhbnZhcy50b0RhdGFVUkwoImltYWdlL3BuZyIpLHRoaXMuX2VsSW1hZ2Uuc3R5bGUu"
    "ZGlzcGxheT0iYmxvY2siLHRoaXMuX2VsQ2FudmFzLnN0eWxlLmRpc3BsYXk9Im5vbmUifWZ1bmN0aW9uIGQoYSxiKXt2YXIgYz10"
    "aGlzO2lmKGMuX2ZGYWlsPWIsYy5fZlN1Y2Nlc3M9YSxudWxsPT09Yy5fYlN1cHBvcnREYXRhVVJJKXt2YXIgZD1kb2N1bWVudC5j"
    "cmVhdGVFbGVtZW50KCJpbWciKSxlPWZ1bmN0aW9uKCl7Yy5fYlN1cHBvcnREYXRhVVJJPSExLGMuX2ZGYWlsJiZfZkZhaWwuY2Fs"
    "bChjKX0sZj1mdW5jdGlvbigpe2MuX2JTdXBwb3J0RGF0YVVSST0hMCxjLl9mU3VjY2VzcyYmYy5fZlN1Y2Nlc3MuY2FsbChjKX07"
    "cmV0dXJuIGQub25hYm9ydD1lLGQub25lcnJvcj1lLGQub25sb2FkPWYsZC5zcmM9ImRhdGE6aW1hZ2UvZ2lmO2Jhc2U2NCxpVkJP"
    "UncwS0dnb0FBQUFOU1VoRVVnQUFBQVVBQUFBRkNBWUFBQUNOYnlibEFBQUFIRWxFUVZRSTEyUDQvLzgvdzM4R0lBWERJQktFMERI"
    "eGdsak5CQUFPOVRYTDBZNE9Id0FBQUFCSlJVNUVya0pnZ2c9PSIsdm9pZCAwfWMuX2JTdXBwb3J0RGF0YVVSST09PSEwJiZjLl9m"
    "U3VjY2Vzcz9jLl9mU3VjY2Vzcy5jYWxsKGMpOmMuX2JTdXBwb3J0RGF0YVVSST09PSExJiZjLl9mRmFpbCYmYy5fZkZhaWwuY2Fs"
    "bChjKX1pZih0aGlzLl9hbmRyb2lkJiZ0aGlzLl9hbmRyb2lkPD0yLjEpe3ZhciBiPTEvd2luZG93LmRldmljZVBpeGVsUmF0aW8s"
    "Yz1DYW52YXNSZW5kZXJpbmdDb250ZXh0MkQucHJvdG90eXBlLmRyYXdJbWFnZTtDYW52YXNSZW5kZXJpbmdDb250ZXh0MkQucHJv"
    "dG90eXBlLmRyYXdJbWFnZT1mdW5jdGlvbihhLGQsZSxmLGcsaCxpLGope2lmKCJub2RlTmFtZSJpbiBhJiYvaW1nL2kudGVzdChh"
    "Lm5vZGVOYW1lKSlmb3IodmFyIGw9YXJndW1lbnRzLmxlbmd0aC0xO2w+PTE7bC0tKWFyZ3VtZW50c1tsXT1hcmd1bWVudHNbbF0q"
    "YjtlbHNlInVuZGVmaW5lZCI9PXR5cGVvZiBqJiYoYXJndW1lbnRzWzFdKj1iLGFyZ3VtZW50c1syXSo9Yixhcmd1bWVudHNbM10q"
    "PWIsYXJndW1lbnRzWzRdKj1iKTtjLmFwcGx5KHRoaXMsYXJndW1lbnRzKX19dmFyIGU9ZnVuY3Rpb24oYSxiKXt0aGlzLl9iSXNQ"
    "YWludGVkPSExLHRoaXMuX2FuZHJvaWQ9bigpLHRoaXMuX2h0T3B0aW9uPWIsdGhpcy5fZWxDYW52YXM9ZG9jdW1lbnQuY3JlYXRl"
    "RWxlbWVudCgiY2FudmFzIiksdGhpcy5fZWxDYW52YXMud2lkdGg9Yi53aWR0aCx0aGlzLl9lbENhbnZhcy5oZWlnaHQ9Yi5oZWln"
    "aHQsYS5hcHBlbmRDaGlsZCh0aGlzLl9lbENhbnZhcyksdGhpcy5fZWw9YSx0aGlzLl9vQ29udGV4dD10aGlzLl9lbENhbnZhcy5n"
    "ZXRDb250ZXh0KCIyZCIpLHRoaXMuX2JJc1BhaW50ZWQ9ITEsdGhpcy5fZWxJbWFnZT1kb2N1bWVudC5jcmVhdGVFbGVtZW50KCJp"
    "bWciKSx0aGlzLl9lbEltYWdlLnN0eWxlLmRpc3BsYXk9Im5vbmUiLHRoaXMuX2VsLmFwcGVuZENoaWxkKHRoaXMuX2VsSW1hZ2Up"
    "LHRoaXMuX2JTdXBwb3J0RGF0YVVSST1udWxsfTtyZXR1cm4gZS5wcm90b3R5cGUuZHJhdz1mdW5jdGlvbihhKXt2YXIgYj10aGlz"
    "Ll9lbEltYWdlLGM9dGhpcy5fb0NvbnRleHQsZD10aGlzLl9odE9wdGlvbixlPWEuZ2V0TW9kdWxlQ291bnQoKSxmPWQud2lkdGgv"
    "ZSxnPWQuaGVpZ2h0L2UsaD1NYXRoLnJvdW5kKGYpLGk9TWF0aC5yb3VuZChnKTtiLnN0eWxlLmRpc3BsYXk9Im5vbmUiLHRoaXMu"
    "Y2xlYXIoKTtmb3IodmFyIGo9MDtlPmo7aisrKWZvcih2YXIgaz0wO2U+aztrKyspe3ZhciBsPWEuaXNEYXJrKGosayksbT1rKmYs"
    "bj1qKmc7Yy5zdHJva2VTdHlsZT1sP2QuY29sb3JEYXJrOmQuY29sb3JMaWdodCxjLmxpbmVXaWR0aD0xLGMuZmlsbFN0eWxlPWw/"
    "ZC5jb2xvckRhcms6ZC5jb2xvckxpZ2h0LGMuZmlsbFJlY3QobSxuLGYsZyksYy5zdHJva2VSZWN0KE1hdGguZmxvb3IobSkrLjUs"
    "TWF0aC5mbG9vcihuKSsuNSxoLGkpLGMuc3Ryb2tlUmVjdChNYXRoLmNlaWwobSktLjUsTWF0aC5jZWlsKG4pLS41LGgsaSl9dGhp"
    "cy5fYklzUGFpbnRlZD0hMH0sZS5wcm90b3R5cGUubWFrZUltYWdlPWZ1bmN0aW9uKCl7dGhpcy5fYklzUGFpbnRlZCYmZC5jYWxs"
    "KHRoaXMsYSl9LGUucHJvdG90eXBlLmlzUGFpbnRlZD1mdW5jdGlvbigpe3JldHVybiB0aGlzLl9iSXNQYWludGVkfSxlLnByb3Rv"
    "dHlwZS5jbGVhcj1mdW5jdGlvbigpe3RoaXMuX29Db250ZXh0LmNsZWFyUmVjdCgwLDAsdGhpcy5fZWxDYW52YXMud2lkdGgsdGhp"
    "cy5fZWxDYW52YXMuaGVpZ2h0KSx0aGlzLl9iSXNQYWludGVkPSExfSxlLnByb3RvdHlwZS5yb3VuZD1mdW5jdGlvbihhKXtyZXR1"
    "cm4gYT9NYXRoLmZsb29yKDFlMyphKS8xZTM6YX0sZX0oKTpmdW5jdGlvbigpe3ZhciBhPWZ1bmN0aW9uKGEsYil7dGhpcy5fZWw9"
    "YSx0aGlzLl9odE9wdGlvbj1ifTtyZXR1cm4gYS5wcm90b3R5cGUuZHJhdz1mdW5jdGlvbihhKXtmb3IodmFyIGI9dGhpcy5faHRP"
    "cHRpb24sYz10aGlzLl9lbCxkPWEuZ2V0TW9kdWxlQ291bnQoKSxlPU1hdGguZmxvb3IoYi53aWR0aC9kKSxmPU1hdGguZmxvb3Io"
    "Yi5oZWlnaHQvZCksZz1bJzx0YWJsZSBzdHlsZT0iYm9yZGVyOjA7Ym9yZGVyLWNvbGxhcHNlOmNvbGxhcHNlOyI+J10saD0wO2Q+"
    "aDtoKyspe2cucHVzaCgiPHRyPiIpO2Zvcih2YXIgaT0wO2Q+aTtpKyspZy5wdXNoKCc8dGQgc3R5bGU9ImJvcmRlcjowO2JvcmRl"
    "ci1jb2xsYXBzZTpjb2xsYXBzZTtwYWRkaW5nOjA7bWFyZ2luOjA7d2lkdGg6JytlKyJweDtoZWlnaHQ6IitmKyJweDtiYWNrZ3Jv"
    "dW5kLWNvbG9yOiIrKGEuaXNEYXJrKGgsaSk/Yi5jb2xvckRhcms6Yi5jb2xvckxpZ2h0KSsnOyI+PC90ZD4nKTtnLnB1c2goIjwv"
    "dHI+Iil9Zy5wdXNoKCI8L3RhYmxlPiIpLGMuaW5uZXJIVE1MPWcuam9pbigiIik7dmFyIGo9Yy5jaGlsZE5vZGVzWzBdLGs9KGIu"
    "d2lkdGgtai5vZmZzZXRXaWR0aCkvMixsPShiLmhlaWdodC1qLm9mZnNldEhlaWdodCkvMjtrPjAmJmw+MCYmKGouc3R5bGUubWFy"
    "Z2luPWwrInB4ICIraysicHgiKX0sYS5wcm90b3R5cGUuY2xlYXI9ZnVuY3Rpb24oKXt0aGlzLl9lbC5pbm5lckhUTUw9IiJ9LGF9"
    "KCk7UVJDb2RlPWZ1bmN0aW9uKGEsYil7aWYodGhpcy5faHRPcHRpb249e3dpZHRoOjI1NixoZWlnaHQ6MjU2LHR5cGVOdW1iZXI6"
    "NCxjb2xvckRhcms6IiMwMDAwMDAiLGNvbG9yTGlnaHQ6IiNmZmZmZmYiLGNvcnJlY3RMZXZlbDpkLkh9LCJzdHJpbmciPT10eXBl"
    "b2YgYiYmKGI9e3RleHQ6Yn0pLGIpZm9yKHZhciBjIGluIGIpdGhpcy5faHRPcHRpb25bY109YltjXTsic3RyaW5nIj09dHlwZW9m"
    "IGEmJihhPWRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGEpKSx0aGlzLl9hbmRyb2lkPW4oKSx0aGlzLl9lbD1hLHRoaXMuX29RUkNv"
    "ZGU9bnVsbCx0aGlzLl9vRHJhd2luZz1uZXcgcSh0aGlzLl9lbCx0aGlzLl9odE9wdGlvbiksdGhpcy5faHRPcHRpb24udGV4dCYm"
    "dGhpcy5tYWtlQ29kZSh0aGlzLl9odE9wdGlvbi50ZXh0KX0sUVJDb2RlLnByb3RvdHlwZS5tYWtlQ29kZT1mdW5jdGlvbihhKXt0"
    "aGlzLl9vUVJDb2RlPW5ldyBiKHIoYSx0aGlzLl9odE9wdGlvbi5jb3JyZWN0TGV2ZWwpLHRoaXMuX2h0T3B0aW9uLmNvcnJlY3RM"
    "ZXZlbCksdGhpcy5fb1FSQ29kZS5hZGREYXRhKGEpLHRoaXMuX29RUkNvZGUubWFrZSgpLHRoaXMuX2VsLnRpdGxlPWEsdGhpcy5f"
    "b0RyYXdpbmcuZHJhdyh0aGlzLl9vUVJDb2RlKSx0aGlzLm1ha2VJbWFnZSgpfSxRUkNvZGUucHJvdG90eXBlLm1ha2VJbWFnZT1m"
    "dW5jdGlvbigpeyJmdW5jdGlvbiI9PXR5cGVvZiB0aGlzLl9vRHJhd2luZy5tYWtlSW1hZ2UmJighdGhpcy5fYW5kcm9pZHx8dGhp"
    "cy5fYW5kcm9pZD49MykmJnRoaXMuX29EcmF3aW5nLm1ha2VJbWFnZSgpfSxRUkNvZGUucHJvdG90eXBlLmNsZWFyPWZ1bmN0aW9u"
    "KCl7dGhpcy5fb0RyYXdpbmcuY2xlYXIoKX0sUVJDb2RlLkNvcnJlY3RMZXZlbD1kfSgpOw=="
)

_LABEL42_CSS = """
*{margin:0;padding:0;box-sizing:border-box;font-family:Arial,Helvetica,sans-serif}
body{background:#4a5b50}
.bar{position:sticky;top:0;background:#093824;color:#fff;padding:10px 16px;
     display:flex;gap:14px;align-items:center;flex-wrap:wrap;z-index:5}
.bar b{font-size:15px}
.bar span{font-size:12px;color:#c6ccb2}
.bar button{padding:8px 22px;border:none;border-radius:6px;background:#c6ccb2;
            color:#093824;font-weight:700;font-size:14px;cursor:pointer}
.lab{width:4in;height:2in;background:#fff;margin:10px auto;display:flex;
     overflow:hidden;border:1px solid #999}
.col{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;
     padding:0.08in 0.06in;text-align:center;min-width:0}
.col.left{border-right:1px solid #ccc}
.ttl{font-size:14pt;font-weight:900;color:#000;letter-spacing:.3px}
/* Values: 10pt semi-bold black (per floor test 2026-09-02) and NEVER wrapped —
   a too-long serial shrinks its own font (JS fitter below) instead of pushing
   the trailing rev (e.g. -A01) onto a second line. */
.val{font-size:10pt;font-weight:700;color:#000;margin-top:2px;white-space:nowrap;
     max-width:100%;line-height:1.1;min-height:13pt}
.qr{margin-top:4px;display:flex;align-items:center;justify-content:center}
/* QR is RENDERED at 300px (see the script) but shown/printed at 60px — the
   5x-oversampled bitmap keeps the modules crisp at the printer's ~203 dpi. */
.qr img,.qr canvas{width:60px!important;height:60px!important}
.sub{font-size:6pt;font-weight:700;color:#000;margin-top:1px;word-break:break-all;line-height:1.1}
@media print{
  @page{size:4in 2in;margin:0}
  body{background:#fff}
  .bar{display:none}
  .lab{margin:0;border:none;page-break-after:always;break-after:page}
  .lab:last-child{page-break-after:auto;break-after:auto}
}
"""

# The page the app opens FOR PRINTING closes its own Chrome tab once the print
# dialog is done (Dhruvil 2026-09-25: "once we click on print it should close
# automatically that chrome tab"). window.print() holds the script while
# Chrome's / Edge's print dialog is open, so the close runs after Print — or
# Cancel: the browser doesn't tell the page which was pressed. A print() that
# returns at once (a browser that doesn't hold the script, printing blocked by
# policy) leaves the tab open, as before. Chrome lets a page close a tab that
# holds only that page — how os.startfile opens it; anywhere else close() is a
# harmless no-op. The one-second pause lets Chrome hand the job to Windows
# before the tab goes. Used by the 4x2 board labels and by the Shipping Slip /
# Labels PRINT COPIES only: a filed slip / labels page in the WO folder keeps
# plain auto-print, so opening one from History still leaves it on screen.
_PRINT_THEN_CLOSE_FN = ("function _pccPrintThenClose(){var t0=Date.now();window.print();"
                        "if(Date.now()-t0>400){setTimeout(function(){window.close();},1000);}}")
_PRINT_THEN_CLOSE_JS = _PRINT_THEN_CLOSE_FN + "\nsetTimeout(_pccPrintThenClose,600);"

def write_label_html_4x2(out_dir, pn, serials, wo=""):
    """Write the self-contained 4x2 label sheet (one page per serial) and return
    its path. Layout matches the Apex spack-4x2 label exactly: P/N left, Serial
    right, QR under each with the encoded value mirrored in 6pt text."""
    import html as _html
    import base64
    pn = str(pn or "").strip()
    serials = [str(s).strip() for s in (serials or []) if str(s).strip()]
    if not serials:
        raise ValueError("no serials to print")
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    ts = datetime.datetime.now()
    fname = sanitize_filename_part(f"labels4x2_{wo or pn}_{ts:%Y%m%d_%H%M%S}", fallback="labels") + ".html"
    js = base64.b64decode(_QRCODEJS_B64).decode("utf-8", "replace")
    pn_h = _html.escape(pn)
    labs = []
    for i, sn in enumerate(serials):
        sn_h = _html.escape(sn)
        labs.append(
            '<div class="lab">'
            '<div class="col left"><div class="ttl">P/N:</div>'
            f'<div class="val">{pn_h or "&mdash;"}</div>'
            f'<div class="qr" id="qp{i}"></div>'
            f'<div class="sub">{pn_h or "&mdash;"}</div></div>'
            '<div class="col"><div class="ttl">Serial:</div>'
            f'<div class="val">{sn_h}</div>'
            f'<div class="qr" id="qs{i}"></div>'
            f'<div class="sub">{sn_h}</div></div>'
            '</div>')
    n = len(serials)
    title = _html.escape(f"4x2 labels - {wo or pn} - {n} label(s)")
    bar = ('<div class="bar"><b>4 × 2 board labels &mdash; ' + str(n) + ' label(s)</b>'
           '<span>' + (('WO ' + _html.escape(wo) + ' &middot; ') if wo else '')
           + 'P/N ' + (pn_h or '&mdash;')
           + ' &nbsp;&mdash;&nbsp; printer: HPKE200 &middot; label 4 × 2 in &middot; set Margins to None</span>'
           '<button onclick="window.print()">🖨 Print all</button></div>')
    tail = """<script>
var PN=__PN__;var SERIALS=__SN__;
/* Render at 300px, display at 60px (CSS): 5x oversampling so the thermal
   printer gets crisp modules instead of a stretched blurry 60px bitmap. */
SERIALS.forEach(function(sn,i){
 if(PN){try{new QRCode(document.getElementById("qp"+i),{text:PN,width:300,height:300,colorDark:"#000",colorLight:"#fff",correctLevel:QRCode.CorrectLevel.M});}catch(e){}}
 try{new QRCode(document.getElementById("qs"+i),{text:sn,width:300,height:300,colorDark:"#000",colorLight:"#fff",correctLevel:QRCode.CorrectLevel.M});}catch(e){}
});
/* One-line fitter: shrink an overflowing P/N or serial (never below 6pt) so the
   whole value always prints on a single line. */
document.querySelectorAll(".val").forEach(function(el){
 var fs=10;el.style.fontSize=fs+"pt";
 while(fs>6&&el.scrollWidth>el.clientWidth){fs-=0.5;el.style.fontSize=fs+"pt";}
});
""" + _PRINT_THEN_CLOSE_JS + """
</script>"""
    tail = tail.replace("__PN__", json.dumps(pn)).replace("__SN__", json.dumps(serials))
    doc = ('<!doctype html><html><head><meta charset="utf-8"><title>' + title + '</title>'
           '<script>' + js + '</script>'
           '<style>' + _LABEL42_CSS + '</style></head><body>'
           + bar + "".join(labs) + tail + '</body></html>')
    target = out / fname
    with open(target, "w", encoding="utf-8") as f:
        f.write(doc)
    return str(target)


###############################################################################
# SHIPPING SLIP + SHIPPING LABELS (Apex-identical, printed by the browser)
###############################################################################
# Dhruvil 2026-09-24: "copy the shipping slip and shipping label section of the
# Apex app — I want the exact format". So this is NOT a re-drawing: the page the
# app writes carries Apex 3.11's own markup (#lbl-ship / #lbl-shipl) and its CSS
# rules VERBATIM (frozen copy: PCC_Apex_Charlie_3.11.html), and the browser
# prints it exactly as it prints Apex — 4 x 6 in, margins 0, on the Zebra
# ZP 450 CTP. The only differences from the live Apex page:
#   • the box-count "gate" and the Prev/Next/Print-All bar are not part of the
#     printed label and are left out;
#   • text inputs / selects become static <span>/<div> elements with the SAME
#     class names and inline styles (a <div> where the original was a block or
#     table-cell control, a <span> inside the flex rows) so the printed text is
#     identical, while checkboxes / radios stay real inputs (checked as
#     recorded, frozen with onclick="return false" — NOT disabled, that greys
#     them out);
#   • the labels page repeats one label per box, each in its own 4 x 6 page
#     wrapper, so ONE print job prints every box (Apex "Print All" prints one
#     job per box).
# Everything below is Tk-free and testable (tests/test_slip_labels.py compares
# the generated element sequence and CSS against the frozen Apex file).
#
# What actually prints for the header code (.label-code): Apex's base rule is
# color:#888, but its theme layer overrides it with
# `.label-code{color:var(--ap-text-3)!important;}` (= #8a958c) and the print
# block never resets it (it forces print-color-adjust:exact instead), so the
# SLIP code prints in #8a958c — the :root variables and that override are
# copied verbatim so the cascade resolves the same way here. On the LABEL the
# code is not printed at all: Apex's print block has
# `#pack-code-display,#shipl-code-display{display:none!important;}` (kept).

SLIP_FOLDER = "Shipping Slip"          # leaf subfolder inside an existing WO folder

# Apex DEFAULTS.boxsize / DEFAULTS.shipvia: the dropdown choices, PICK-ONLY like
# Apex's <select>s (Dhruvil 2026-09-25: a click anywhere in the field opens the
# list). Box sizes keep growing ("we keep getting new sizes"), so the form's Size
# list ends with '+ Add new size…', which files the new size in a list SHARED by
# every station (box_sizes.json beside operators.json — load_box_sizes /
# add_box_size below; admins can remove a mistyped one in Admin ▸ Box sizes).
# SLIP_BOXSIZE_OPTIONS is that list until the first size is added. 'Hand Carry'
# is ours, not Apex's (Dhruvil 2026-09-25).
SLIP_BOXSIZE_OPTIONS = ['5x5x5', '5x5x10', '10x10x5', '10x10x10', '12x10x8',
                        '14x10x10', '16x12x10', '20x16x12', 'Pallet', 'Other']
SLIP_SHIPVIA_OPTIONS = ['Courier', 'Customer Pickup', 'Deliver', 'FedEx GND',
                        'FedEx Priority', 'Hand Carry', 'UPS 2nd Day', 'UPS GND']
# The dropdowns' "nothing picked" entries (Apex's own blank-option labels) and
# the Size list's last entry — none of them is ever stored as a value.
_SLIP_SIZE_NONE = "--"
_SLIP_VIA_NONE = "-- Select --"
_SLIP_ADD_SIZE = "+ Add new size…"

def normalize_box_size(s):
    """A typed box size in the list's own style: spaces collapsed and the
    dimension separator made a lower-case x ('18 X 12 x 6', '18×12×6' ->
    '18x12x6'); anything else is kept as typed ('Pallet 48x40')."""
    s = " ".join(str(s or "").split())
    return re.sub(r"(?<=\d)\s*[xX\u00d7*]\s*(?=\d)", "x", s)

def box_size_problem(s):
    """Why the (normalised) text `s` can't be a box size, or None."""
    s = str(s or "")
    if not s.strip():
        return "Type a box size."
    if len(s) > 24:
        return "Keep a box size to 24 characters or fewer."
    if s in (_SLIP_SIZE_NONE, _SLIP_VIA_NONE, _SLIP_ADD_SIZE) or s.lower().startswith("+ add"):
        return f"'{s}' can't be used as a box size."
    return None

def _clean_box_sizes(data):
    """A stored / typed list as the dropdown shows it: normalised, no blanks, no
    reserved entries, no case-insensitive duplicates, order kept."""
    out, seen = [], set()
    for x in (data if isinstance(data, list) else []):
        s = normalize_box_size(x) if isinstance(x, str) else ""
        if s and not box_size_problem(s) and s.lower() not in seen:
            seen.add(s.lower()); out.append(s)
    return out

def _box_sizes_with(sizes, s):
    """`sizes` plus the new size `s`, placed ahead of 'Pallet' / 'Other' (Apex
    ends its list with those two)."""
    head = [x for x in sizes if x.lower() not in ("pallet", "other")]
    tail = [x for x in sizes if x.lower() in ("pallet", "other")]
    return head + [s] + tail

def _read_shared_box_sizes():
    """The shared list as it is on the share NOW: (sizes, None) — Apex's list
    when nobody has added a size yet — or (None, error text) when the file is
    there but can't be read. The writers (add / remove) start from this, never
    from the offline cache or a list read earlier: writing a stale copy back
    would erase sizes other stations added meanwhile (review round 2)."""
    try:
        p = Path(BOX_SIZES_FILE)
        if not p.exists():
            return list(SLIP_BOXSIZE_OPTIONS), None
        with open(p, encoding="utf-8") as f:
            data = _clean_box_sizes(json.load(f))
        return (data or list(SLIP_BOXSIZE_OPTIONS)), None
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"

def load_box_sizes(online=None):
    """The Size dropdown's choices: the shared list (BOX_SIZES_FILE) once anyone
    has added a size, Apex's list before that. Offline, or the file unreadable:
    the last good copy cached on this PC, else Apex's list. Never raises, never
    empty. online=False (the app already knows PCC2 is down) skips the probe."""
    if online is None:
        online = pcc2_online()
    if online:
        data, _err = _read_shared_box_sizes()
        if data is not None:
            if Path(BOX_SIZES_FILE).exists():
                _cache_roster("box_sizes", data)     # refresh the offline copy
            return data
    data = _clean_box_sizes(_read_cached_roster("box_sizes"))
    return data or list(SLIP_BOXSIZE_OPTIONS)

def save_box_sizes(sizes):
    """Write the shared box-size list (atomic, like the operator roster) and
    cache it locally. Raises when it can't be written — callers report it."""
    clean = _clean_box_sizes(sizes)
    if not clean:
        raise ValueError("the box-size list can't be empty")
    _atomic_write_json(BOX_SIZES_FILE, clean)
    _cache_roster("box_sizes", clean)
    return clean

def add_box_size(size, online=None):
    """'+ Add new size…': add `size` to the SHARED list. The list is read fresh
    from the share first, so a size another station added a moment ago is not
    lost. Returns (sizes, saved, error): the list the form should show now,
    whether the size reached the share, and why not (offline / no access — the
    form still uses the size for the shipment in hand). A size already on the
    list (any case) is not added twice."""
    s = normalize_box_size(size)
    bad = box_size_problem(s)
    if bad:
        raise ValueError(bad)
    if online is None:
        online = pcc2_online()
    cur, err = _read_shared_box_sizes() if online else (None, None)
    if cur is None:                                  # offline, or the shared file unreadable
        seen = load_box_sizes(online=False)          # what this PC knows, for the form in hand
        new = seen if any(x.lower() == s.lower() for x in seen) else _box_sizes_with(seen, s)
        why = (f"The shared list could not be read ({err})." if online
               else "PCC2 (the share) is not reachable from this PC.")
        return new, False, why
    if any(x.lower() == s.lower() for x in cur):
        return cur, True, None
    new = _box_sizes_with(cur, s)
    try:
        return save_box_sizes(new), True, None
    except Exception as e:
        return new, False, f"{type(e).__name__}: {e}"

def remove_box_size(size, online=None):
    """Admin ▸ Box sizes 'Remove selected': take `size` (any case) off the
    SHARED list, read fresh first so a size another station added meanwhile is
    kept. Returns (sizes, saved, error) like add_box_size; never writes while
    the share is unreachable or unreadable, never leaves the list empty."""
    if online is None:
        online = pcc2_online()
    if not online:
        return load_box_sizes(online=False), False, "PCC2 (the share) is not reachable from this PC."
    cur, err = _read_shared_box_sizes()
    if cur is None:
        return load_box_sizes(online=False), False, f"The shared list could not be read ({err})."
    key = normalize_box_size(size).lower()
    new = [x for x in cur if x.lower() != key]
    if new == cur:
        return cur, True, None                       # already gone
    if not new:
        return cur, False, "Keep at least one box size."
    try:
        return save_box_sizes(new), True, None
    except Exception as e:
        return cur, False, f"{type(e).__name__}: {e}"

# Apex 3.11 <head> lines 10-12 load Roboto from Google Fonts; its `*{font-
# family:'Roboto',...!important}` rule (copied verbatim below) prints in Roboto
# when that font is available and in Arial when it is not. The generated page
# does NOT fetch it (standing rule: no network — and a render-blocking
# stylesheet on a shipping PC with no internet route holds the page AND the
# print dialog for the browser's connect timeout): it prints in Roboto when
# Roboto is installed on the PC, else Arial, exactly as Apex does with the web
# font unavailable. True re-enables Apex's <link> (kept as a one-line switch,
# review round 1) — the link then goes AFTER the embedded script so a pending
# stylesheet can never block the QR, and the auto-print waits for the font.
SLIP_FONT_LINK = False
_SLIP_FONT_LINK_HTML = (
    '<link rel="preconnect" href="https://fonts.googleapis.com">\n'
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n'
    '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700;900'
    '&family=Roboto+Condensed:wght@400;700&family=Roboto+Mono:wght@400;700&display=swap" '
    'rel="stylesheet">\n')

# ── Apex 3.11 CSS, VERBATIM (line ranges noted inline; extracted mechanically
# from the frozen file — do not "tidy" these, the fidelity test checks them
# character for character). Includes the two `*` rules (box-sizing + font
# stack), the label/slip/label-page rules, the print block, the :root theme
# variables, the .label-code theme override and the 3.9 print-size block.
_SLIP_APEX_CSS = r"""/* Apex 3.11 lines 14-14 */
*{margin:0;padding:0;box-sizing:border-box;font-family:'Roboto','Helvetica Neue',Arial,sans-serif!important;}
/* Apex 3.11 lines 40-41 */
.label{width:4in;height:6in;background:#fff;border:1.5px solid #555;display:none;flex-direction:column;overflow:hidden;flex-shrink:0;}
.label.active{display:flex;}
/* Apex 3.11 lines 67-71 */
.lbl-hdr-split{padding:1px 6px;border-bottom:2px solid #c9a84c;flex-shrink:0;display:flex;align-items:center;justify-content:space-between;}
.hdr-co{font-size:14px;font-weight:900;font-family:Arial,sans-serif;}
.hdr-right{text-align:right;}
.hdr-title{font-size:11px;font-weight:700;font-family:Arial,sans-serif;color:#1a1a1a;}
.label-code{font-size:9px;font-weight:700;font-family:Arial,sans-serif;color:#888;margin-top:1px;}
/* Apex 3.11 lines 73-86 */
.ncr-hdr .hdr-co{color:#7a3e00;}.rej-hdr .hdr-co{color:#6b0000;}.ship-hdr .hdr-co{color:#1a4a6e;}.shipl-hdr .hdr-co{color:#0a2a4e;}
.job-info{padding:1px 6px;border-bottom:2px solid #333;flex-shrink:0;}
.fi-row{display:flex;align-items:center;border-bottom:1px solid #ddd;padding:2px 0;gap:4px;min-height:22px;}
.fi-row:last-child{border-bottom:none;}
.fi-lbl{font-size:9px;font-weight:700;font-family:Arial,sans-serif;text-transform:uppercase;white-space:nowrap;min-width:72px;}
.op-lbl{color:#2c4a6e;}.qc-lbl{color:#4a2c2c;}.test-lbl{color:#1a5c3a;}.ncr-lbl{color:#7a3e00;}
.rej-lbl{color:#6b0000;}.ship-lbl{color:#1a4a6e;}.shipl-lbl{color:#0a2a4e;}
.fi-input{flex:1;border:none;outline:none;font-size:13px;font-weight:700;font-family:Arial,sans-serif;background:transparent;color:#000;padding:0 2px;}
.fi-input.invalid{border-bottom:2px solid #c00 !important;}
.fi-row.qty .fi-input{font-size:22px;font-weight:900;}
.fi-row.serial .fi-input{font-size:20px;font-weight:900;color:#6b0000;}
.fi-select{flex:1;border:none;outline:none;font-size:12px;font-weight:700;font-family:Arial,sans-serif;background:transparent;color:#000;padding:0 2px;cursor:pointer;}
.fi-select.invalid{border-bottom:2px solid #c00 !important;}
.fi-date{flex:1;border:none;outline:none;font-size:12px;font-weight:700;font-family:Arial,sans-serif;background:transparent;color:#000;padding:0 2px;cursor:pointer;}
/* Apex 3.11 lines 93-94 */
.footer-qr{width:38px;height:38px;flex-shrink:0;display:flex;align-items:center;justify-content:center;}
.footer-qr canvas,.footer-qr img{width:38px!important;height:38px!important;}
/* Apex 3.11 lines 100-102 */
.opt{display:flex;align-items:center;gap:4px;font-size:11px;font-weight:600;font-family:Arial,sans-serif;color:#111;cursor:pointer;}
.opt input[type=checkbox]{width:12px;height:12px;cursor:pointer;}
.opt-qc input{accent-color:#4a2c2c;}.opt-test input{accent-color:#1a5c3a;}.opt-rej input{accent-color:#6b0000;}.opt-ship input{accent-color:#1a4a6e;}
/* Apex 3.11 lines 126-135 */
.box-table{width:100%;border-collapse:collapse;font-size:10px;}
.box-table th{background:#e8ecf2;color:#2c4a6e;font-size:10px;font-weight:700;font-family:Arial,sans-serif;text-transform:uppercase;padding:3px 4px;border:1px solid #ddd;text-align:left;}
.box-table td{padding:0 3px;border:1px solid #ddd;height:8px;vertical-align:middle;}
.box-table input,.box-table select{border:none;outline:none;font-size:10px;font-weight:600;font-family:Arial,sans-serif;background:transparent;width:100%;}
.col-box{width:14%;}.col-size{width:36%;}.col-qty{width:16%;}.col-wt{width:30%;}
.box-table tfoot td{font-weight:800;font-size:10px;background:#f5f5e8;}
.ship-section-lbl{font-size:9px;font-weight:700;font-family:Arial,sans-serif;text-transform:uppercase;color:#1a4a6e;border-bottom:1px solid #c9a84c;padding-bottom:1px;margin-bottom:3px;letter-spacing:.2px;}
.ship-approved-footer{border-top:2px solid #c9a84c;padding:3px 6px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;gap:5px;}
.barcode-row{padding:1px 8px 2px;border-bottom:1px solid #eee;flex-shrink:0;display:flex;align-items:center;justify-content:center;}
.barcode-row svg{max-width:100%;height:26px;}
/* Apex 3.11 lines 545-589 */
@media print{
  /* Apex 3.9 (2026-06-10): fixed 4×6 edge-clip / next-page issue.
     Force margin:0 universally so the active label fills the page exactly.
     The previous 0.15in top margin combined with padding-top:2px pushed
     content 2-4px past the bottom of the page → blank page two with
     bottom edge. setPrintSize() also writes @page{size:...; margin:0}
     into #page-size-style so the page itself is the right size. */
  @page{margin:0;}
  html,body{margin:0 !important;padding:0 !important;background:#fff !important;}
  body{background:#fff!important;padding:0;gap:0;margin:0;display:block!important;}
  .top-bar,.action-bar,.toast,.form-code-badge,.admin-overlay,.reset-btn,.admin-fixed,.pin-overlay,.login-overlay,.logged-in-bar,.label-nav,.wo-lookup-overlay,.tag-sel-overlay,.bin-size-bar,.no-print,.parts-selects,.serial-scan-overlay,.wip-overlay,.stats-overlay,.prod-overlay,.andu-overlay,.ptrk-overlay,.dash-btns,.dsched-overlay,.app-header{display:none !important;}
  .label.active{
    display:flex !important;
    border:none !important;
    background:#fff!important;
    margin:0 !important;
    padding-top:0 !important;
    box-sizing:border-box !important;
    overflow:hidden !important;
    position:absolute !important;
    top:0 !important; left:0 !important;
    page-break-after:avoid !important;
    page-break-inside:avoid !important;
  }
  .label.active .label-nav{background:transparent!important;}
  #lbl-op.active,#lbl-qc.active,#lbl-test.active,#lbl-rej.active,#lbl-ncr.active,#lbl-ship.active,#lbl-shipl.active,#lbl-rwk.active,#lbl-sreq.active{width:4in;height:6in;}
  #lbl-kit.active{width:3in;height:2in;}
  #lbl-pack.active,#lbl-spack.active{width:3in;height:2in;}
  #lbl-pack.active.pack-3x1_5,#lbl-spack.active.spack-3x1_5{width:3in;height:1.5in;}
  #lbl-pack.active.pack-3x2_5,#lbl-spack.active.spack-3x2_5{width:3in;height:2.5in;}
  #lbl-spack4x2.active{display:flex !important;width:4in;height:2in;border:none;background:#fff!important;margin:0;}
  #lbl-bin.active{width:2in;height:1in;}
  #lbl-bin.active.bin-3x2{width:3in;height:2in;}
  #lbl-parts.active{width:2in;height:1in;}
  #lbl-gen.active{width:6in;height:4in;}
  #lbl-preq.active{width:8.5in;height:auto;min-height:5in;}
  #lbl-preq.active.preq-4x6{width:4in;height:6in;min-height:unset;}
  .label:not(.active){display:none !important;}
  .box-gate{display:none !important;}
  .kit-selects,.pack-fields,.pack-mode{display:none !important;}
  #pack-code-display,#shipl-code-display{display:none!important;}
  .footer-qr{margin-right:6px;}
  .lbl-footer{padding-right:8px;}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact;}
}
/* Apex 3.11 lines 601-621 */
:root{
  --ap-forest:#093824;
  --ap-forest-2:#0d4a30;
  --ap-forest-3:#072a1a;
  --ap-cream:#c6ccb2;
  --ap-cream-soft:#dde1ce;
  --ap-cream-50:#f2f3eb;
  --ap-taupe:#7a5c58;
  --ap-taupe-soft:#a48783;
  --ap-bg:#f2f3eb;
  --ap-surface:#ffffff;
  --ap-border:#d8dcc9;
  --ap-border-soft:#e8eadc;
  --ap-text:#142b1f;
  --ap-text-2:#4a5b50;
  --ap-text-3:#8a958c;
  --ap-shadow-sm:0 1px 2px rgba(9,56,36,.06);
  --ap-shadow:0 2px 10px rgba(9,56,36,.08),0 1px 3px rgba(9,56,36,.04);
  --ap-shadow-lg:0 12px 32px rgba(9,56,36,.12),0 4px 10px rgba(9,56,36,.06);
}

/* Apex 3.11 lines 623-623 */
*{font-family:'Roboto','Helvetica Neue',Arial,sans-serif!important;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility;}
/* Apex 3.11 lines 1127-1127 */
.label-code{color:var(--ap-text-3)!important;}
/* Apex 3.11 lines 3085-3085 */
.label.active{zoom:1.6;}
/* Apex 3.11 lines 3111-3116 */
@media print{
  .setup-overlay,.xscan-overlay,.opack-scan-overlay,.rma-scan-overlay,.oring-scan-overlay,.oring-setup-overlay{display:none!important;}
  #lbl-opack.active{width:4in!important;height:6in!important;background:#fff!important;}
  #lbl-opack-letter.active,#lbl-rma.active,#lbl-oring.active,#lbl-preq.active{width:8.5in!important;height:auto;min-height:5in;background:#fff!important;}
  .label.active{zoom:1!important;}
}
/* Apex 3.11 lines 3126-3169 */
@media print{
  /* Universal label-print invariants — apply to EVERY label size. */
  .label.active{
    box-sizing:border-box !important;
    overflow:hidden !important;
    border:none !important;
    margin:0 !important;
    padding-top:0 !important;
    background:#fff !important;
    page-break-after:avoid !important;
    page-break-inside:avoid !important;
    break-after:avoid-page !important;
    break-inside:avoid-page !important;
    position:absolute !important;
    top:0 !important; left:0 !important;
    transform:none !important;
  }
  /* 4×6 standard labels (Quality, Testing, Rejection, NCR, Assembly,
     Shipping Slip, Shipping Label, Rework, Supply Request, RMA Tag). */
  #lbl-op.active,#lbl-qc.active,#lbl-test.active,#lbl-rej.active,
  #lbl-ncr.active,#lbl-ship.active,#lbl-shipl.active,#lbl-rwk.active,
  #lbl-sreq.active{width:4in !important;height:6in !important;}
  /* 3-inch labels (kit, packing, serialized pack) — three height options. */
  #lbl-kit.active{width:3in !important;height:2in !important;}
  #lbl-pack.active,#lbl-spack.active{width:3in !important;height:2in !important;}
  #lbl-pack.active.pack-3x1_5,#lbl-spack.active.spack-3x1_5{width:3in !important;height:1.5in !important;}
  #lbl-pack.active.pack-3x2_5,#lbl-spack.active.spack-3x2_5{width:3in !important;height:2.5in !important;}
  /* Spack 4×2 (the simplified layout). */
  #lbl-spack4x2.active{width:4in !important;height:2in !important;}
  /* Small labels (bin & parts). */
  #lbl-bin.active{width:2in !important;height:1in !important;}
  #lbl-bin.active.bin-3x2{width:3in !important;height:2in !important;}
  #lbl-parts.active{width:2in !important;height:1in !important;}
  /* General label (6×4 landscape). */
  #lbl-gen.active{width:6in !important;height:4in !important;}
  /* Parts Request — selectable 4×6 OR letter. */
  #lbl-preq.active{width:8.5in !important;height:auto !important;min-height:5in;max-height:11in !important;}
  #lbl-preq.active.preq-4x6{width:4in !important;height:6in !important;min-height:unset !important;max-height:6in !important;}
  /* Outsourced Pack — 4×6 (slip) and letter (full-page version). */
  #lbl-opack.active{width:4in !important;height:6in !important;}
  #lbl-opack-letter.active{width:8.5in !important;height:auto !important;min-height:5in;max-height:11in !important;}
  /* RMA Receiving + Outsourced Receiving — letter. */
  #lbl-rma.active,#lbl-oring.active{width:8.5in !important;height:auto !important;min-height:5in;max-height:11in !important;}
}"""

# App-side additions only (never touch the Apex rules above): the non-printing
# hint bar, a 4 x 6 page wrapper per label so several labels print as one job
# (Apex positions .label.active absolutely at the page origin; the wrapper
# gives each label its own page to be the origin of), and the three rules that
# make the static stand-ins for Apex's <input>/<select> controls print like
# the controls did (the fidelity test pins these three and allows nothing else
# here to touch an Apex class or id). body colour = Apex's --ap-text, which
# the i/N cells and the totals inherit on Apex's page.
_SLIP_PAGE_CSS = """
body{background:#4a5b50;color:#142b1f;padding:0 0 24px;}
.bar{position:sticky;top:0;background:#093824;color:#fff;padding:10px 16px;
     display:flex;gap:14px;align-items:center;flex-wrap:wrap;z-index:5;}
.bar b{font-size:15px;}
.bar span{font-size:12px;color:#c6ccb2;}
.bar button{padding:8px 22px;border:none;border-radius:6px;background:#c6ccb2;
            color:#093824;font-weight:700;font-size:14px;cursor:pointer;}
.sheet{position:relative;width:max-content;margin:16px auto;}
/* A native control clips a long value on ONE line; a span/div would wrap, grow
   the row and push everything below it (the box table lost rows). Clip exactly
   as the control did. The two controls Apex leaves at the browser's field
   colour print black. */
span.fi-input,span.fi-select,span.fi-date,#ship-via,.box-table td>div,#shipl-weight{white-space:nowrap;overflow:hidden;}
.box-table td>div,#shipl-weight{color:#000;}
/* A blank <input>/<select> keeps its full control height; a blank span/div has
   no line box at all, so the label's "Weight (lbs): ____" band and the slip's
   Ship Via section printed shorter than Apex's and everything below them
   moved (~4 mm). A non-breaking space gives the empty stand-in the line box the
   control had and prints nothing. (Dhruvil 2026-09-25) */
#shipl-weight:empty::before,#ship-via:empty::before{content:"\\00a0";}
/* Barcodes under their values (Dhruvil 2026-09-25) — see _slip_barcode_row:
   an invisible copy of the row's title + its 4 px gap, then the barcode
   centred in the space left, exactly as the value above it is. The box keeps
   Apex's 8 px either side of the barcode: Code 128's quiet zone. */
.barcode-row.bc-under{padding-left:0;padding-right:0;justify-content:flex-start;}
.bc-sp{flex:0 1 auto;min-width:0;overflow:hidden;visibility:hidden;display:flex;}
.bc-sp>.fi-lbl{box-sizing:content-box;padding-right:4px;}
.bc-box{flex:1 1 auto;display:flex;justify-content:center;padding:0 8px;}
@media print{
  .sheet{width:4in;height:6in;margin:0;overflow:hidden;page-break-after:always;break-after:page;}
  .sheet:last-of-type{page-break-after:auto;break-after:auto;}
}
"""

# Apex's setPrintSize('ship'|'shipl') writes this into #page-size-style.
_SLIP_PAGE_SIZE_CSS = "@media print{@page{size:4in 6in;margin:0}}"

def slip_data_template():
    """A fresh slip_data dict with every key the generators read (the popup,
    the draft and the close record all carry this exact shape)."""
    return {
        "customer": "", "cpn": "", "po": "", "wo": "", "eng": "", "rev": "",
        "qty": "", "date": datetime.date.today().strftime("%Y-%m-%d"),
        "datecode_wwyy": "", "prepared_by": "", "notes": "",
        "boxes": [],                       # [{size, qty, weight}] as typed
        "customer_parts": "No",            # 'Yes' | 'No'
        "fair": False, "si": False, "ppap": False, "cofc": False,
        "ship_via": "", "label_notes": "",
        "slip_code": "", "labels_code": "",
        # packing label (2026-10-19): QTY starts at 1 and CAGE at 1WXP2 as in
        # Apex; the mode follows the division; date + date code frozen at print
        "pack_qty": "1", "pack_cage": PACK_CAGE_DEFAULT, "pack_mode": "itar",
        "pack_code": "", "pack_date": "", "pack_datecode": "",
        "printed": {"slip": None, "labels": None, "pack": None},
        "outputs": {"slip_html": None, "labels_html": None, "pack_html": None},
        # what each page was rendered from, frozen at print time (a copy of the
        # fields above) — Re-emit rebuilds THIS, not the form as edited later
        "pages": {"slip": None, "labels": None, "pack": None},
    }

# ── JS-number helpers: the page must show exactly what Apex's JS shows ────────
def _slip_js_num(x):
    """JS Number -> String: 15 -> '15', 3.5 -> '3.5', and the same shortest
    round-trip digits as the browser for the barcode SVG coordinates."""
    if isinstance(x, float):
        return str(int(x)) if x.is_integer() and abs(x) < 1e21 else repr(x)
    return str(x)

_SLIP_JS_FLOAT_RE = re.compile(r'[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?')
def _slip_js_parse_float(s):
    """JS parseFloat: the leading numeric prefix, or None for NaN."""
    m = _SLIP_JS_FLOAT_RE.match(str(s if s is not None else "").lstrip())
    return float(m.group(0)) if m else None

def _slip_js_to_fixed(x, digits=2):
    """JS Number.toFixed: rounds the EXACT binary value half-up (Python's
    format rounds half-even, e.g. 0.125 -> '0.12' where JS gives '0.13')."""
    from decimal import Decimal, ROUND_HALF_UP
    x = float(x)
    if x == 0:
        x = 0.0                            # JS prints -0 as '0.00'
    return str(Decimal(x).quantize(Decimal(1).scaleb(-digits), rounding=ROUND_HALF_UP))

def calc_box_totals(boxes):
    """Apex calcBoxTotals: (qty total, weight total) as the slip shows them —
    the qty sum as a JS number string, the weight sum as toFixed(2)+' lbs',
    '—' when no box has a numeric value."""
    tq = 0.0; hq = False; tw = 0.0; hw = False
    for b in (boxes or []):
        v = _slip_js_parse_float((b or {}).get("weight"))
        if v is not None:
            tw += v; hw = True
        v = _slip_js_parse_float((b or {}).get("qty"))
        if v is not None:
            tq += v; hq = True
    return (_slip_js_num(tq) if hq else "—",
            (_slip_js_to_fixed(tw, 2) + " lbs") if hw else "—")

# ── Date code (slip only) and the form codes ─────────────────────────────────
def slip_datecode_wwyy(date_str=None):
    """Apex calcDateCode()/updateShipDateCode(): ISO week number of the slip
    date, shown WWYY (week 39 of 2026 -> '3926'). Reproduces Apex exactly —
    the YY is the date's CALENDAR year, unlike the app's fmt_datecode() (YYWW,
    ISO year), which is deliberately left alone. Unparseable/blank -> today."""
    iso = _parse_user_date(date_str) if date_str else None
    d = (datetime.datetime.strptime(iso, "%Y-%m-%d").date() if iso
         else datetime.date.today())
    thu = d + datetime.timedelta(days=4 - d.isoweekday())      # this ISO week's Thursday
    wk = -(-((thu - datetime.date(thu.year, 1, 1)).days + 1) // 7)   # Math.ceil
    return f"{wk:02d}{d.year % 100:02d}"

def slip_code_rand():
    """Apex: Math.floor(1000+Math.random()*9000) -> 1000..9999, drawn ONCE per
    print and recorded with the shipment."""
    import random
    return random.randint(1000, 9999)

def _slip_code_date(date):
    if date is None:
        return datetime.date.today().strftime("%Y%m%d")
    if isinstance(date, (datetime.date, datetime.datetime)):
        return date.strftime("%Y%m%d")
    iso = _parse_user_date(date)
    return (iso or str(date)).replace("-", "")

def _slip_code_token(v, fallback):
    # Apex: (value||'ENG').replace(/\s+/g,'-')
    v = str(v or "")
    return re.sub(r"\s+", "-", v if v else fallback)

def slip_form_code(eng, wo, prepared_by, date=None, rand=None):
    """Apex generateCode('ship'): ENG-WO-INITIALS-SLIP-YYYYMMDD-RAND4. INITIALS =
    first letter of each space-separated part of the Prepared-By name, upper-
    cased, or 'XX'. `date` = today unless given; `rand` = a fresh draw unless
    given (re-emit passes the recorded code instead of calling this)."""
    person = "".join(p[:1] for p in str(prepared_by or "").split(" ")).upper() or "XX"
    r = slip_code_rand() if rand is None else int(rand)
    return (f"{_slip_code_token(eng, 'ENG')}-{_slip_code_token(wo, 'WO')}-{person}"
            f"-SLIP-{_slip_code_date(date)}-{r}")

def labels_form_code(eng, wo, date=None, rand=None):
    """Apex generateCode('shipl'): ENG-WO-XX-LABEL-YYYYMMDD-RAND4 — the label
    has no person field in Apex's map, so its initials are always 'XX'."""
    r = slip_code_rand() if rand is None else int(rand)
    return (f"{_slip_code_token(eng, 'ENG')}-{_slip_code_token(wo, 'WO')}-XX"
            f"-LABEL-{_slip_code_date(date)}-{r}")

# ── Code128 barcodes: JsBarcode's CODE128 (auto) DOM, bar for bar ────────────
# Apex renders each label barcode with JsBarcode(svg, value, {format:'CODE128',
# width:1.3, height:26, displayValue:false, margin:0, background:'transparent'}).
# The symbol patterns come from the bundled python-barcode package
# (barcode.charsets.code128.CODES / STOP); the code-set selection below is a
# straight port of JsBarcode's auto mode (its A/B/C choices decide the bar
# widths, and python-barcode's own optimiser chooses differently for mixed
# text like '16452.1'), so the printed bars match Apex exactly, not just the
# scanned value.
_JSB_A = r'[\x00-\x5F\xC8-\xCF]'
_JSB_B = r'[\x20-\x7F\xC8-\xCF]'
_JSB_C = r'(\xCF*[0-9]{2}\xCF*)'

def _jsb_auto(s):
    def a_len(t): return len(re.match('^' + _JSB_A + '*', t).group(0))
    def b_len(t): return len(re.match('^' + _JSB_B + '*', t).group(0))
    def c_match(t): return re.match('^' + _JSB_C + '*', t).group(0)
    def from_ab(t, is_a):
        ranges = _JSB_A if is_a else _JSB_B
        until_c = re.match('^(' + ranges + '+?)(([0-9]{2}){2,})([^0-9]|$)', t)
        if until_c:
            return until_c.group(1) + chr(204) + from_c(t[len(until_c.group(1)):])
        m = re.match('^' + ranges + '+', t)
        if not m:
            raise ValueError("Code128: character outside set")
        chars = m.group(0)
        if len(chars) == len(t):
            return t
        return chars + chr(205 if is_a else 206) + from_ab(t[len(chars):], not is_a)
    def from_c(t):
        cm = c_match(t)
        if len(cm) == len(t):
            return t
        t = t[len(cm):]
        is_a = a_len(t) >= b_len(t)
        return cm + chr(206 if is_a else 205) + from_ab(t, is_a)
    if len(c_match(s)) >= 2:
        out = chr(210) + from_c(s)
    else:
        is_a = a_len(s) > b_len(s)
        out = chr(208 if is_a else 209) + from_ab(s, is_a)
    # a lone character between two code switches becomes a SHIFT (Apex/JsBarcode)
    return re.sub('[\xCD\xCE]([\\s\\S])[\xCD\xCE]', lambda m: chr(203) + m.group(1), out, count=1)

def code128_modules(value):
    """The barcode as a string of '1' (bar) / '0' (space) modules, exactly as
    JsBarcode encodes `value` (start code, data, checksum, stop). Raises on
    text it cannot encode (callers show no barcode then, like Apex)."""
    from barcode.charsets import code128 as _c128
    bars = list(_c128.CODES)                 # 0..105
    bars.append(_c128.STOP + "11")           # 106 = stop + termination bar
    b = [ord(c) for c in _jsb_auto(str(value))]
    start = b.pop(0) - 105                   # 208/209/210 -> 103/104/105 (A/B/C)
    if start not in (103, 104, 105):
        raise ValueError("Code128: bad start code")
    cur = start; pos = 1; checksum = 0; body = ""
    while b:
        if b[0] >= 200:
            idx = b.pop(0) - 105
            nxt = {101: 103, 100: 104, 99: 105}.get(idx)
            if nxt is None and cur in (103, 104) and idx == 98:     # SHIFT
                if cur == 103:
                    b[0] = b[0] - 96 if b[0] > 95 else b[0]
                else:
                    b[0] = b[0] + 96 if b[0] < 32 else b[0]
        else:
            if cur == 103:
                c = b.pop(0); idx = c + 64 if c < 32 else c - 32
            elif cur == 104:
                idx = b.pop(0) - 32
            else:
                idx = (b.pop(0) - 48) * 10 + b.pop(0) - 48
            nxt = None
        body += bars[idx]; checksum += idx * pos; pos += 1
        if nxt is not None:
            cur = nxt
    return bars[start] + body + bars[(checksum + start) % 103] + bars[106]

def code128_svg(value, svg_id, module_width=1.3, height=26):
    """The <svg> JsBarcode leaves in Apex's `.barcode-row` for `value` (width
    1.3 px/module, 26 px high, no text, no margin, transparent background,
    viewBox kept so `.barcode-row svg{max-width:100%;height:26px}` scales it
    the same way). Non-printable characters are stripped first, as Apex's
    renderBarcode does; nothing printable -> an empty <svg> (no barcode)."""
    import math
    safe = re.sub(r'[^\x20-\x7E]', '', str(value if value is not None else "").strip())
    if not safe:
        return f'<svg id="{svg_id}"></svg>'
    try:
        bits = code128_modules(safe)
    except Exception:
        return f'<svg id="{svg_id}"></svg>'
    W = module_width; H = height       # Apex: shipping label 1.3 / 26, packing label 1 / 12
    width = math.ceil(len(bits) * W)
    rects = []; bar = 0; x = 0.0
    for i, ch in enumerate(bits):
        x = i * W
        if ch == "1":
            bar += 1
        elif bar > 0:
            rects.append((x - W * bar, W * bar)); bar = 0
    if bar > 0:
        rects.append((x - W * (bar - 1), W * bar))
    body = "".join(f'<rect x="{_slip_js_num(rx)}" y="0" width="{_slip_js_num(rw)}" height="{H}"></rect>'
                   for rx, rw in rects)
    return (f'<svg id="{svg_id}" width="{width}px" height="{H}px" x="0px" y="0px" '
            f'viewBox="0 0 {width} {H}" xmlns="http://www.w3.org/2000/svg" version="1.1" '
            f'style="transform: translate(0,0)">'
            f'<rect x="0" y="0" width="{width}" height="{H}" style="fill:transparent;"></rect>'
            f'<g transform="translate(0, 0)" style="fill:#000000;">{body}</g></svg>')

# ── The pages ────────────────────────────────────────────────────────────────
def _slip_date_iso(d):
    """'YYYY-MM-DD' (what a <input type=date> holds and what its barcode
    encodes) from whatever the popup stored; unparseable text passes through."""
    return _parse_user_date(d) or str(d or "").strip()

def _slip_date_display(d):
    """What Chrome shows in a <input type=date>: MM/DD/YYYY."""
    iso = _parse_user_date(d)
    if iso:
        return datetime.datetime.strptime(iso, "%Y-%m-%d").strftime("%m/%d/%Y")
    return str(d or "").strip()

def _slip_esc(v):
    """A recorded value as HTML text (None -> ''): the ONE escape both page
    builders and the page title use."""
    import html as _html
    return _html.escape(str(v if v is not None else ""))

def _slip_page(title, body_html, tail_js, apex_extra="", page_size_css=None, page_css_extra=""):
    """Wrap Apex markup in a self-contained page: embedded qrcode.js, the
    verbatim Apex CSS, Apex's @page rule, the app's hint bar, auto-print.
    apex_extra = more verbatim Apex rules (the packing label's);
    page_size_css / page_css_extra = another label size (the 3 x 2 packing
    label) — the 4 x 6 slip / labels pages pass neither."""
    import base64
    js = base64.b64decode(_QRCODEJS_B64).decode("utf-8", "replace")
    return ('<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8">\n'
            '<title>' + _slip_esc(title) + '</title>\n'
            '<script>' + js + '</script>\n'
            + (_SLIP_FONT_LINK_HTML if SLIP_FONT_LINK else '')
            + '<style>\n' + _SLIP_APEX_CSS + ('\n' + apex_extra if apex_extra else '') + '\n</style>\n'
            '<style id="page-size-style">' + (page_size_css or _SLIP_PAGE_SIZE_CSS) + '</style>\n'
            '<style>' + _SLIP_PAGE_CSS + page_css_extra + '</style>\n'
            '</head>\n<body>\n' + body_html + '\n<script>\n' + tail_js + '\n</script>\n</body>\n</html>\n')

# Auto-open the print dialog (as Apex does after applyCode): the 4x2 labels'
# plain timer — nothing on the page waits on the network. With SLIP_FONT_LINK
# on, wait (bounded) for the web font first so the page doesn't print in Arial
# while Roboto is still arriving.
_SLIP_AUTOPRINT_JS = "setTimeout(function(){window.print();},600);"
_SLIP_AUTOPRINT_FONT_JS = """var _printed=false;function _goPrint(){if(_printed)return;_printed=true;window.print();}
if(document.fonts&&document.fonts.ready){document.fonts.ready.then(function(){setTimeout(_goPrint,250);});}
setTimeout(_goPrint,2500);"""
def _slip_autoprint_js():
    return _SLIP_AUTOPRINT_FONT_JS if SLIP_FONT_LINK else _SLIP_AUTOPRINT_JS
_SLIP_CLOSE_AFTER_PRINT_FONT_JS = _PRINT_THEN_CLOSE_FN + """
var _printed=false;function _goPrint(){if(_printed)return;_printed=true;_pccPrintThenClose();}
if(document.fonts&&document.fonts.ready){document.fonts.ready.then(function(){setTimeout(_goPrint,250);});}
setTimeout(_goPrint,2500);"""
def _slip_print_js(print_copy=False):
    """The page's print trigger: plain auto-print on a FILED page (the copy in
    the WO folder, opened again from History), print-then-close-the-tab on the
    PRINT COPY the app opens for printing (_PRINT_THEN_CLOSE_JS). Both wait
    (bounded) for the web font when SLIP_FONT_LINK is on."""
    if not print_copy:
        return _slip_autoprint_js()
    return _SLIP_CLOSE_AFTER_PRINT_FONT_JS if SLIP_FONT_LINK else _PRINT_THEN_CLOSE_JS

_SLIP_PRINTER_HINT = 'Printer: Zebra ZP 450 CTP &middot; 4 &times; 6 in &middot; set Margins to None'
_PACK_PRINTER_HINT = 'Printer: 3 &times; 2 Shipping &middot; 3 &times; 2 in &middot; set Margins to None'
def _slip_bar(kind, n_pages, hint=None):
    return ('<div class="no-print bar"><b>' + kind + ' &mdash; ' + str(n_pages)
            + (' page' if n_pages == 1 else ' pages') + '</b>'
            '<span>' + (hint or _SLIP_PRINTER_HINT) + '</span>'
            '<button onclick="window.print()">\U0001F5A8 Print</button></div>\n')

def _slip_sheet(data):
    """The Shipping Slip's 4 x 6 sheet for one shipment (the slip page's body
    and the first page of the "Print both" print copy): Apex's #lbl-ship markup with
    the recorded values as static text, its boxes table rows (Apex
    buildBoxTable's row template, 'i/N'), totals per calcBoxTotals, the WWYY
    date code, the footer QR (WO + '-1', 38 px, level M) and the form code
    (data['slip_code'] — stamped by write_slip_files, or the recorded one).
    Returns (sheet html, the script that draws the footer QR)."""
    d = data or {}
    E = _slip_esc
    wo = str(d.get("wo") or "")
    boxes = [b or {} for b in (d.get("boxes") or [])]
    n = len(boxes)
    wwyy = str(d.get("datecode_wwyy") or "") or slip_datecode_wwyy(d.get("date"))
    tq, tw = calc_box_totals(boxes)
    rows = []
    for i, b in enumerate(boxes, 1):
        rows.append(
            '<tr><td class="col-box" style="font-weight:800;text-align:center;font-size:10px;font-family:Arial,sans-serif;">'
            + f'{i}/{n}' + '</td>'
            '<td class="col-size"><div style="border:none;outline:none;font-size:10px;font-weight:600;font-family:Arial,sans-serif;width:100%;background:transparent;">'
            + (E(b.get("size")) or "--") + '</div></td>'
            '<td class="col-qty"><div class="qty-input" style="border:none;outline:none;font-size:10px;font-weight:700;font-family:Arial,sans-serif;width:100%;">'
            + E(b.get("qty")) + '</div></td>'
            '<td class="col-wt"><div class="wt-input" style="border:none;outline:none;font-size:10px;font-weight:700;font-family:Arial,sans-serif;width:100%;">'
            + E(b.get("weight")) + '</div></td></tr>')
    cp_yes = str(d.get("customer_parts") or "").strip().lower() in ("yes", "y", "true", "1")
    def chk(key):
        return ' checked' if d.get(key) else ''
    lbl = ('<div class="label active" id="lbl-ship">\n'
        '  <div class="lbl-hdr-split ship-hdr"><div><div class="hdr-co" style="color:#1a4a6e;">Printed Circuits Corp.</div></div><div class="hdr-right"><div class="hdr-title">SHIPPING SLIP</div><div class="label-code" id="ship-code-display">' + E(d.get("slip_code")) + '</div></div></div>\n'
        '  <div id="ship-form" style="display:flex;flex:1;flex-direction:column;overflow:hidden;">\n'
        '    <div class="job-info" style="border-bottom:1px solid #ccc;">\n'
        '      <div class="fi-row"><span class="fi-lbl ship-lbl">Customer</span><span class="fi-select fi-input" id="ship-cust">' + E(d.get("customer")) + '</span></div>\n'
        '      <div class="fi-row"><span class="fi-lbl ship-lbl">Customer P/N</span><span class="fi-select fi-input" id="ship-cpn">' + E(d.get("cpn")) + '</span></div>\n'
        '      <div class="fi-row"><span class="fi-lbl ship-lbl">Customer PO #</span><span class="fi-select fi-input" id="ship-po">' + E(d.get("po")) + '</span></div>\n'
        '      <div class="fi-row"><span class="fi-lbl ship-lbl">Work Order/Lot</span><span class="fi-select fi-input" id="ship-wo">' + E(wo) + '</span></div>\n'
        '      <div class="fi-row"><span class="fi-lbl ship-lbl">Engineering</span><span class="fi-select fi-input" id="ship-eng">' + E(d.get("eng")) + '</span></div>\n'
        '      <div class="fi-row"><span class="fi-lbl ship-lbl">Revision</span><span class="fi-input" id="ship-rev">' + E(d.get("rev")) + '</span></div>\n'
        '      <div class="fi-row qty"><span class="fi-lbl ship-lbl">Qty Shipped</span><span class="fi-input" id="ship-qty">' + E(d.get("qty")) + '</span></div>\n'
        '      <div class="fi-row"><span class="fi-lbl ship-lbl">Date</span><span class="fi-date" id="ship-date">' + E(_slip_date_display(d.get("date"))) + '</span></div>\n'
        '      <div class="fi-row"><span class="fi-lbl ship-lbl">Date Code</span><span class="fi-input" id="ship-dc" aria-label="Date Code in WWYY format">' + E(wwyy) + '</span></div>\n'
        '    </div>\n'
        '    <div style="padding:4px 8px;flex-shrink:1;overflow:hidden;min-height:0;display:flex;flex-direction:column;">\n'
        '      <div class="ship-section-lbl">Boxes</div>\n'
        '      <div style="overflow-y:auto;overflow-x:auto;flex:1;min-height:0;">\n'
        '        <table class="box-table">\n'
        '          <thead><tr><th class="col-box">Box</th><th class="col-size">Size</th><th class="col-qty">Qty</th><th class="col-wt">Weight (lbs)</th></tr></thead>\n'
        '          <tbody id="ship-box-tbody">' + "".join(rows) + '</tbody>\n'
        '          <tfoot><tr><td colspan="2" style="text-align:right;font-size:10px;font-weight:700;font-family:Arial,sans-serif;color:#555;padding-right:4px;">Total:</td><td id="ship-total-qty" style="font-weight:900;font-size:10px;font-family:Arial,sans-serif;">' + E(tq) + '</td><td id="ship-total-weight" style="font-weight:900;font-size:10px;font-family:Arial,sans-serif;">' + E(tw) + '</td></tr></tfoot>\n'
        '        </table>\n'
        '      </div>\n'
        '    </div>\n'
        '    <div style="padding:3px 8px;flex-shrink:0;"><div class="ship-section-lbl">Customer Parts</div><div style="display:flex;gap:16px;"><label class="opt opt-ship"><input type="radio" name="ship-custparts" id="ship-cp-no" value="no"' + ('' if cp_yes else ' checked') + ' onclick="return false"> No</label><label class="opt opt-ship"><input type="radio" name="ship-custparts" id="ship-cp-yes" value="yes"' + (' checked' if cp_yes else '') + ' onclick="return false"> Yes</label></div></div>\n'
        '    <div style="padding:3px 8px;flex-shrink:0;"><div class="ship-section-lbl">Additional Requirements</div><div style="display:flex;gap:12px;flex-wrap:wrap;"><label class="opt opt-ship"><input type="checkbox" id="ship-fair"' + chk("fair") + ' onclick="return false"> FAIR</label><label class="opt opt-ship"><input type="checkbox" id="ship-si"' + chk("si") + ' onclick="return false"> SI</label><label class="opt opt-ship"><input type="checkbox" id="ship-ppap"' + chk("ppap") + ' onclick="return false"> PPAP</label><label class="opt opt-ship"><input type="checkbox" id="ship-cofc"' + chk("cofc") + ' onclick="return false"> C of C</label></div></div>\n'
        '    <div style="padding:3px 8px;flex-shrink:0;"><div class="ship-section-lbl">Ship Via</div><div class="fi-select" id="ship-via" style="width:100%;border-bottom:1px solid #555;font-size:12px;font-weight:700;font-family:Arial,sans-serif;padding:2px 3px;">' + E(d.get("ship_via")) + '</div></div>\n'
        '    <div style="padding:3px 8px;flex-shrink:0;"><div class="fi-row"><span class="fi-lbl ship-lbl">Prepared By</span><span class="fi-select" id="ship-operator">' + E(d.get("prepared_by")) + '</span></div><div class="fi-row"><span class="fi-lbl ship-lbl">Notes</span><span class="fi-input" id="ship-notes">' + E(d.get("notes")) + '</span></div></div>\n'
        '    <div class="ship-approved-footer" style="margin-top:auto;"><div style="display:flex;align-items:center;gap:8px;flex:1;"><span style="font-size:9px;font-weight:700;font-family:Arial,sans-serif;text-transform:uppercase;color:#1a4a6e;white-space:nowrap;">Approved By:</span><div style="flex:1;border-bottom:1.5px solid #555;height:18px;min-width:80px;"></div></div><div class="footer-qr" id="qr-ship"></div></div>\n'
        '  </div>\n'
        '</div>')
    sheet = '<div class="sheet">\n' + lbl + '\n</div>'
    # Apex refreshFooterQR('ship'): QR of WO + '-1' into #qr-ship, 38 px, level M; no WO -> no QR.
    qr_js = ('var WO=' + json.dumps(wo) + ';\n'
             "(function(){var el=document.getElementById('qr-ship');if(!el)return;el.innerHTML='';if(!WO)return;"
             "try{new QRCode(el,{text:WO+'-1',width:38,height:38,colorDark:'#000',colorLight:'#fff',correctLevel:QRCode.CorrectLevel.M});}catch(e){}})();\n")
    return sheet, qr_js

def build_slip_html(data, print_copy=False):
    """The Shipping Slip page for one shipment (_slip_sheet + the hint bar).
    print_copy=False: the FILED page — plain auto-print, stays open when opened
    again from History. print_copy=True: the copy the app opens for printing —
    the same page, whose tab closes itself after the print dialog
    (_slip_print_js, Dhruvil 2026-09-25)."""
    d = data or {}
    sheet, qr_js = _slip_sheet(d)
    wo = str(d.get("wo") or "")
    body = _slip_bar("Shipping Slip", 1) + sheet
    return _slip_page(f"Shipping Slip - {wo or d.get('eng') or ''}".rstrip(" -"), body, qr_js + _slip_print_js(print_copy))


# Barcodes under their values (Dhruvil 2026-09-25: "barcode is not aligned with
# the actual value") — the ONE deliberate change from Apex's label. Apex centres
# each barcode on the whole row, while the value above it is centred in the space
# to the right of its title, so a barcode never sat under its value. Each barcode
# row now carries an invisible copy of its row's title (same classes, same text,
# so the same width, plus the row's 4 px gap) and centres the barcode in the
# space left: exactly under the value (rules in _SLIP_PAGE_CSS). A barcode wider
# than that space keeps its full size and slides left just enough (its box's
# min-width is its own width + the box's 8 px either side, so the invisible title
# gives way first); only one wider than the whole row shrinks, to exactly Apex's
# width, as Apex's max-width:100% does. The 8 px either side is Apex's own
# barcode-row padding kept inside the box: with the label's 6 px it leaves 14 px
# from the paper edge, over Code 128's quiet zone (10 modules = 13 px) — a slid
# barcode must never sit flush on the edge (review round 2).
def _slip_barcode_row(value, svg_id, title):
    svg = code128_svg(value, svg_id)
    m = re.search(r'\swidth="(\d+)px"', svg)
    w = int(m.group(1)) if m else 0
    return ('<div class="barcode-row bc-under"><span class="bc-sp" aria-hidden="true">'
            '<span class="fi-lbl shipl-lbl">' + _slip_esc(title) + '</span></span>'
            '<span class="bc-box" style="min-width:min(' + str(w + 16) + 'px,100%)">' + svg + '</span></div>')

def _labels_sheets(data):
    """The Shipping Labels' 4 x 6 sheets (the labels page's body, and the pages
    after the slip in the "Print both" print copy): Apex's #lbl-shipl markup once PER BOX (box 2+
    carries box 1's fields, as Apex's shiplLoadBox copies them; qty and weight
    are per box; 'Box i of N'), each on its own 4 x 6 page so one print job
    prints them all. Barcodes (Code128) for Customer P/N, PO, WO, Engineering,
    Revision, Date and Qty — an empty value gives no barcode. The form code is
    data['labels_code']. No boxes recorded -> one label with the slip qty. Each barcode sits under
    its value (_slip_barcode_row). Returns (sheets html, number of boxes)."""
    d = data or {}
    E = _slip_esc
    wo = str(d.get("wo") or "")
    boxes = [b or {} for b in (d.get("boxes") or [])] or [{"qty": d.get("qty"), "weight": ""}]
    n = len(boxes)
    date_iso = _slip_date_iso(d.get("date"))
    date_disp = _slip_date_display(d.get("date"))
    code = E(d.get("labels_code"))
    shared = (
        '  <div class="lbl-hdr-split shipl-hdr"><div><div class="hdr-co" style="color:#0a2a4e;">Printed Circuits Corp.</div></div><div class="hdr-right"><div class="hdr-title">SHIPPING LABEL</div><div class="label-code" id="shipl-code-display">' + code + '</div></div></div>\n'
        '  <div id="shipl-form" style="display:flex;flex:1;flex-direction:column;overflow:hidden;">\n'
        '    <div class="job-info" style="border-bottom:none;flex:1;overflow:hidden;">\n'
        '      <div class="fi-row"><span class="fi-lbl shipl-lbl">Customer</span><span class="fi-select fi-input" id="shipl-cust" style="text-align:center;">' + E(d.get("customer")) + '</span></div>\n'
        '      <div class="fi-row"><span class="fi-lbl shipl-lbl">Customer P/N</span><span class="fi-select fi-input" id="shipl-cpn" style="text-align:center;">' + E(d.get("cpn")) + '</span></div>\n'
        '      ' + _slip_barcode_row(d.get("cpn"), "bc-cpn", "Customer P/N") + '\n'
        '      <div class="fi-row"><span class="fi-lbl shipl-lbl">Customer PO #</span><span class="fi-input" id="shipl-po" style="text-align:center;">' + E(d.get("po")) + '</span></div>\n'
        '      ' + _slip_barcode_row(d.get("po"), "bc-po", "Customer PO #") + '\n'
        '      <div class="fi-row"><span class="fi-lbl shipl-lbl">Work Order/Lot</span><span class="fi-select fi-input" id="shipl-wo" style="text-align:center;">' + E(wo) + '</span></div>\n'
        '      ' + _slip_barcode_row(wo, "bc-wo", "Work Order/Lot") + '\n'
        '      <div class="fi-row"><span class="fi-lbl shipl-lbl">Engineering</span><span class="fi-select fi-input" id="shipl-eng" style="text-align:center;">' + E(d.get("eng")) + '</span></div>\n'
        '      ' + _slip_barcode_row(d.get("eng"), "bc-eng", "Engineering") + '\n'
        '      <div class="fi-row"><span class="fi-lbl shipl-lbl">Revision</span><span class="fi-input" id="shipl-rev" style="text-align:center;">' + E(d.get("rev")) + '</span></div>\n'
        '      ' + _slip_barcode_row(d.get("rev"), "bc-rev", "Revision") + '\n'
        '      <div class="fi-row"><span class="fi-lbl shipl-lbl">Date</span><span class="fi-date" id="shipl-date" style="text-align:center;">' + E(date_disp) + '</span></div>\n'
        '      ' + _slip_barcode_row(date_iso, "bc-date", "Date") + '\n')
    pages = []
    for i, b in enumerate(boxes, 1):
        qty = b.get("qty")
        pages.append('<div class="sheet">\n<div class="label active" id="lbl-shipl">\n' + shared
            + '      <div class="fi-row qty"><span class="fi-lbl shipl-lbl">Qty</span><span class="fi-input" id="shipl-qty" style="text-align:center;">' + E(qty) + '</span></div>\n'
            '      ' + _slip_barcode_row(qty, "bc-qty", "Qty") + '\n'
            '      <div class="fi-row"><span class="fi-lbl shipl-lbl">Notes</span><span class="fi-input" id="shipl-note" style="text-align:center;">' + E(d.get("label_notes")) + '</span></div>\n'
            '    </div>\n'
            '    <div style="padding:3px 8px;border-top:1px solid #ddd;flex-shrink:0;display:flex;align-items:center;gap:6px;"><span style="font-size:9px;font-weight:700;text-transform:uppercase;color:#0a2a4e;font-family:Arial,sans-serif;min-width:72px;">Box</span><span style="font-size:16px;font-weight:900;color:#0a2a4e;font-family:Arial,sans-serif;" id="shipl-box-current-print">' + str(i) + '</span><span style="font-size:11px;font-weight:400;color:#888;font-family:Arial,sans-serif;">of</span><span style="font-size:16px;font-weight:900;color:#0a2a4e;font-family:Arial,sans-serif;" id="shipl-box-total-print">' + str(n) + '</span></div>\n'
            '    <div style="border-top:3px solid #0a2a4e;padding:5px 8px;display:flex;align-items:center;gap:10px;flex-shrink:0;"><span style="font-size:9px;font-weight:700;text-transform:uppercase;color:#0a2a4e;font-family:Arial,sans-serif;">Weight (lbs):</span><span id="shipl-weight" style="width:80px;border:none;border-bottom:1.5px solid #555;font-size:18px;font-weight:900;font-family:Arial,sans-serif;outline:none;text-align:center;">' + E(b.get("weight")) + '</span></div>\n'
            '  </div>\n'
            '</div>\n</div>')
    return "\n".join(pages), n

def build_labels_html(data, print_copy=False):
    """The Shipping Labels page: one 4 x 6 label per box (_labels_sheets) + the
    hint bar. print_copy as for build_slip_html."""
    d = data or {}
    sheets, n = _labels_sheets(d)
    wo = str(d.get("wo") or "")
    body = _slip_bar("Shipping Labels", n) + sheets
    # Apex's label has no footer QR element (refreshFooterQR('shipl') finds none).
    return _slip_page(f"Shipping Labels - {wo or d.get('eng') or ''}".rstrip(" -"), body, _slip_print_js(print_copy))

def build_slip_labels_html(data):
    """The PRINT COPY for 'Print both' (Dhruvil 2026-09-25): the slip, then every
    box label, in ONE page, so one Print in Chrome's dialog prints them all —
    Chrome shows one print dialog at a time, and with two tabs each asking to
    print, the slip waited behind the labels and never printed by itself. Its
    tab closes after the print dialog. The FILED record stays two pages
    (build_slip_html + build_labels_html)."""
    d = data or {}
    slip_sheet, qr_js = _slip_sheet(d)
    lab_sheets, n = _labels_sheets(d)
    wo = str(d.get("wo") or "")
    body = _slip_bar("Shipping Slip + Labels", 1 + n) + slip_sheet + "\n" + lab_sheets
    return _slip_page(f"Shipping Slip + Labels - {wo or d.get('eng') or ''}".rstrip(" -"), body,
                      qr_js + _slip_print_js(True))


# ── PACKING LABEL (Dhruvil 2026-10-19: "add packing label section from this to
# our shipping app … I need to replicate whatever is in apex") ────────────────
# Apex 3.11's #lbl-pack, 3 x 2 in, printed on the 3 x 2 Shipping printer: one
# label per print (Apex has no box queue for it), the shipment's Customer P/N,
# QTY (Apex starts it at 1; the operator edits it), CAGE 1WXP2 and the date code
# (YYWW, Apex calcDateCode) on ITAR/MAG jobs, today's date on Industrial jobs,
# and the WO, each with a Code128 barcode (JsBarcode width 1, height 12). Apex's
# ITAR/MAG-Industrial buttons and WO / Eng pickers are screen-only there and are
# left out; the mode comes from the loaded job's DIVISION instead of a click —
# the one intended difference. Its markup, CSS and print rules are Apex's own;
# the date and date code are frozen into the record when the label prints.
_PACK_APEX_CSS = r"""/* Apex 3.11 lines 42-42 */
.label.label-kit,.label.label-pack,.label.label-spack{width:3in;height:2in;}
/* Apex 3.11 lines 192-212 */
.pack-hdr{padding:2px 5px;border-bottom:2px solid #c9a84c;flex-shrink:0;display:flex;align-items:center;justify-content:space-between;}
.pack-co{font-size:12px;font-weight:900;font-family:Arial,sans-serif;color:#2c4a6e;}
.pack-mode{display:flex;gap:4px;}
.pack-mbtn{padding:2px 7px;border:1.5px solid #ccc;border-radius:4px;font-size:9px;font-weight:700;font-family:Arial,sans-serif;cursor:pointer;background:#fff;}
.pack-mbtn.active{background:#1a4a1a;color:#fff;border-color:#1a4a1a;}
.pack-fields{padding:1px 5px;border-bottom:1px solid #eee;flex-shrink:0;background:#fafafa;}
.pack-frow{display:flex;align-items:center;gap:3px;min-height:13px;border-bottom:1px solid #f5f5f5;}
.pack-frow:last-child{border-bottom:none;}
.pack-flbl{font-size:7px;font-weight:700;text-transform:uppercase;color:#2c4a6e;font-family:Arial,sans-serif;min-width:22px;}
.pack-fsel{flex:1;border:none;outline:none;font-size:9px;font-weight:700;font-family:Arial,sans-serif;background:transparent;color:#000;padding:0;cursor:pointer;}
.pack-body{flex:1;padding:2px 5px;display:flex;flex-direction:column;overflow:hidden;}
.pack-fld{display:flex;flex-direction:column;border-bottom:1px solid #f0f0f0;padding:1px 0;}
.pack-r{display:flex;align-items:center;gap:3px;min-height:14px;}
.pack-k{font-size:8px;font-weight:700;text-transform:uppercase;color:#2c4a6e;font-family:Arial,sans-serif;min-width:36px;flex-shrink:0;}
.pack-v{font-size:11px;font-weight:700;font-family:Arial,sans-serif;color:#000;flex:1;border:none;outline:none;background:transparent;padding:0;}
.pack-vbig{font-size:13px;font-weight:900;}
.pack-vdisp{font-size:11px;font-weight:700;font-family:Arial,sans-serif;color:#000;}
.pack-fld-bc{width:100%;height:14px;display:flex;justify-content:flex-start;overflow:hidden;padding-left:36px;}
.pack-fld-bc svg{max-width:100%;height:14px;}
#lbl-pack[data-mode="itar"] .pack-ind-only{display:none!important;}
#lbl-pack[data-mode="industrial"] .pack-itar-only{display:none!important;}"""
_PACK_PAGE_SIZE_CSS = "@media print{@page{size:3in 2in;margin:0}}"
# App-side: a stand-in for Apex's <input class="pack-v"> clips on one line like
# the control; the page wrapper is the label's 3 x 2 (the slip pages' is 4 x 6).
_PACK_PAGE_CSS = """
span.pack-v{white-space:nowrap;overflow:hidden;}
@media print{.sheet{width:3in;height:2in;}}
"""
PACK_CAGE_DEFAULT = "1WXP2"           # Apex's value="1WXP2"

def pack_form_code(eng, wo, date=None, rand=None):
    """Apex generateCode('pack'): ENG-WO-PACK-YYYYMMDD-RAND4 (no initials)."""
    r = slip_code_rand() if rand is None else int(rand)
    return f"{_slip_code_token(eng, 'ENG')}-{_slip_code_token(wo, 'WO')}-PACK-{_slip_code_date(date)}-{r}"

def pack_datecode_yyww(date_str=None):
    """Apex calcDateCode() as the packing label shows it: YY (calendar year) +
    ISO week, e.g. 2026-10-19 -> '2643' (the slip shows the same digits WWYY)."""
    w = slip_datecode_wwyy(date_str)
    return w[2:] + w[:2]

def pack_mode_for(division):
    """Apex's ITAR/MAG vs Industrial packing label, from the job's division."""
    return "industrial" if str(division or "").strip().lower() == "industrial" else "itar"

def _pack_sheet(data):
    d = data or {}
    E = _slip_esc
    mode = "industrial" if str(d.get("pack_mode") or "").lower() == "industrial" else "itar"
    wo = str(d.get("wo") or ""); cpn = d.get("cpn"); qty = d.get("pack_qty"); cage = d.get("pack_cage"); po = d.get("po")
    dc = str(d.get("pack_datecode") or "") or pack_datecode_yyww()
    date_disp = str(d.get("pack_date") or "") or datetime.date.today().strftime("%m/%d/%Y")
    bc = lambda v, i: code128_svg(v, i, 1, 12)
    lbl = ('<div class="label label-pack active" id="lbl-pack" data-mode="' + mode + '">\n'
        '  <div class="pack-hdr"><div class="pack-co">Printed Circuits Corp.</div></div>\n'
        '  <div class="pack-body">\n'
        '    <div class="pack-fld"><div class="pack-r"><span class="pack-k">P/N</span><span class="pack-v pack-vbig fi-input" id="pack-cpn">' + E(cpn) + '</span></div><div class="pack-fld-bc">' + bc(cpn, "bc-pack-cpn") + '</div></div>\n'
        '    <div style="display:flex;border-bottom:1px solid #f0f0f0;gap:4px;padding:1px 0;">\n'
        '      <div style="flex:1;display:flex;flex-direction:column;"><div class="pack-r"><span class="pack-k">QTY</span><span class="pack-v" id="pack-qty">' + E(qty) + '</span></div><div class="pack-fld-bc">' + bc(qty, "bc-pack-qty") + '</div></div>\n'
        '      <div style="flex:1;display:flex;flex-direction:column;margin-left:-0.5cm;" class="pack-itar-only"><div class="pack-r"><span class="pack-k">CAGE</span><span class="pack-v" id="pack-cage">' + E(cage) + '</span></div><div class="pack-fld-bc">' + bc(cage, "bc-pack-cage") + '</div></div>\n'
        '    </div>\n'
        '    <div class="pack-fld"><div class="pack-r"><span class="pack-k">PO</span><span class="pack-v fi-input" id="pack-po">' + E(po) + '</span></div><div class="pack-fld-bc">' + bc(po, "bc-pack-po") + '</div></div>\n'
        '    <div class="pack-fld pack-itar-only"><div class="pack-r"><span class="pack-k">D/C</span><span class="pack-vdisp" id="pack-dc">' + E(dc) + '</span></div><div class="pack-fld-bc">' + bc(dc, "bc-pack-dc") + '</div></div>\n'
        '    <div class="pack-fld pack-ind-only"><div class="pack-r"><span class="pack-k">DATE</span><span class="pack-vdisp" id="pack-date-disp">' + E(date_disp) + '</span></div><div class="pack-fld-bc">' + bc(date_disp, "bc-pack-date") + '</div></div>\n'
        '    <div class="pack-fld"><div class="pack-r"><span class="pack-k">WO#/Lot</span><span class="pack-vdisp" id="pack-wo-disp">' + (E(wo) or '\u2014') + '</span></div><div class="pack-fld-bc">' + bc(wo, "bc-pack-wo") + '</div></div>\n'
        '  </div>\n'
        '  <div style="text-align:center;font-size:7px;font-family:Arial,sans-serif;color:#bbb;padding:1px;" id="pack-code-display">' + E(d.get("pack_code")) + '</div>\n'
        '</div>')
    return '<div class="sheet">\n' + lbl + '\n</div>'

def build_pack_html(data, print_copy=False):
    """The Packing Label page (Apex #lbl-pack, 3 x 2 in). print_copy as for
    build_slip_html. The form code is in the markup but, as in Apex, never
    printed (Apex's print rule hides #pack-code-display)."""
    d = data or {}
    wo = str(d.get("wo") or "")
    body = _slip_bar("Packing Label", 1, _PACK_PRINTER_HINT) + _pack_sheet(d)
    return _slip_page(f"Packing Label - {wo or d.get('eng') or ''}".rstrip(" -"), body, _slip_print_js(print_copy),
                      apex_extra=_PACK_APEX_CSS, page_size_css=_PACK_PAGE_SIZE_CSS, page_css_extra=_PACK_PAGE_CSS)

# The three printed page kinds: record key -> (printed/pages key, file kind, label, code key).
_SLIP_KINDS = {"slip_html": ("slip", "Shipping_Slip", "Shipping Slip", "slip_code"),
               "labels_html": ("labels", "Shipping_Labels", "Shipping Labels", "labels_code"),
               "pack_html": ("pack", "Packing_Label", "Packing Label", "pack_code")}
def _slip_builder(key):
    return {"slip_html": build_slip_html, "labels_html": build_labels_html, "pack_html": build_pack_html}[key]

def _write_new_file(path, text):
    """Create `path` — never overwrite (exclusive create; the unique name from
    unique_file_path makes a clash practically impossible, 'x' makes it
    impossible)."""
    with open(path, "x", encoding="utf-8") as f:
        f.write(text)
    return str(path)

def write_slip_files(wo_folder, data, which="both"):
    """Write the Shipping Slip and/or Shipping Labels page(s) for one print.
    `which` = 'slip' | 'labels' | 'both'. A fresh form code is drawn for each
    document printed and stamped into data['slip_code'] / data['labels_code'];
    the paths go into data['outputs'].

    Folder rule (same as the COC/Packing List, Dhruvil's standing rule): the
    files go to <WO folder>\\Shipping Slip\\ — that subfolder is created as a
    LEAF only (mkdir(exist_ok=True), never parents=True) and only when the WO
    folder itself already exists. Names: <eng>-<m-d-yy>-Shipping_Slip.html and
    <eng>-<m-d-yy>-Shipping_Labels.html (sanitize_filename_part / fmt_file_date
    / unique_file_path — a reprint the same day gets a _HHMMSS suffix, nothing
    is ever overwritten). WO folder unreachable (offline / no access) — or the
    write itself refused (share dropped between the check and the open, a
    read-only ACL, disk full) -> the temp print folder 'PCC Shipping App -
    Print' (like the 4x2 labels) and 'fallback': True so the caller can tell
    the operator the copy was not saved to the share; printing still works.
    A temp copy is NOT recorded in data['outputs'] (that stays None: it is on
    no other PC and Windows cleans it up) — data['printed'] is what marks the
    page as printed, and reemit_paperwork files it into the WO folder later.
    Each page also freezes a copy of the fields it was rendered from in
    data['pages'] so a re-emit reproduces the paper, not a later edit.
    Once a page of this call is on disk the call no longer raises: a later
    page that can't be written anywhere (the share refused it AND the temp
    folder failed) is reported in 'errors' [(label, error)] with its code left
    as it was, so the page already written keeps its record (otherwise the
    popup dropped the dict and a slip with an unrecorded code sat in the WO
    folder). A first page that fails everywhere still raises — nothing to lose.
    Returns {slip_html, labels_html, fallback, folder, errors}."""
    which = str(which or "both").strip().lower()
    want_slip = which in ("slip", "both")
    want_labels = which in ("labels", "both")
    want_pack = which == "pack"                      # the packing label prints on its own (3 x 2 printer)
    if not (want_slip or want_labels or want_pack):
        raise ValueError(f"write_slip_files: which must be slip/labels/both/pack, not {which!r}")
    folder = None; fallback = False
    if wo_folder:
        try:
            wf = Path(str(wo_folder))
            if wf.is_dir():
                leaf = wf / SLIP_FOLDER
                leaf.mkdir(exist_ok=True)              # leaf only — never parents=True
                if leaf.is_dir():
                    folder = leaf
        except Exception:
            folder = None
    if folder is None:
        folder = _slip_temp_print_folder()
        folder.mkdir(parents=True, exist_ok=True)      # the app's own temp folder, not a stored path
        fallback = True
    eng_safe = sanitize_filename_part(data.get("eng") or "", "unknown")
    fd = fmt_file_date()
    outputs = data.get("outputs")
    if not isinstance(outputs, dict):
        outputs = {"slip_html": None, "labels_html": None, "pack_html": None}
        data["outputs"] = outputs
    pages = data.get("pages")
    if not isinstance(pages, dict):
        pages = {"slip": None, "labels": None, "pack": None}
        data["pages"] = pages
    res = {"slip_html": None, "labels_html": None, "pack_html": None, "fallback": fallback, "folder": str(folder), "errors": []}

    def _emit(key, pkey, kind, label, code_key, code, build):
        nonlocal folder, fallback
        prev_code = data.get(code_key)
        data[code_key] = code                       # one draw per print (Apex applyCode)
        text = build(data)
        try:
            try:
                path = _write_new_file(unique_file_path(folder, f"{eng_safe}-{fd}-{kind}", ".html"), text)
            except OSError:
                if fallback:
                    raise                           # the temp folder itself failed: nothing left to try
                # The leaf resolved but the write was refused: same answer as an
                # unreachable folder. A page already written to the share above
                # stays recorded; this one goes to the temp folder unrecorded.
                folder = _slip_temp_print_folder()
                folder.mkdir(parents=True, exist_ok=True)
                fallback = True; res["fallback"] = True; res["folder"] = str(folder)
                path = _write_new_file(unique_file_path(folder, f"{eng_safe}-{fd}-{kind}", ".html"), text)
        except OSError as e:
            if not (res["slip_html"] or res["labels_html"] or res["pack_html"]):
                raise                               # nothing of this call on disk yet
            data[code_key] = prev_code              # no page printed with the new code
            res["errors"].append((label, f"{type(e).__name__}: {e}"))
            return
        res[key] = path
        outputs[key] = None if fallback else path
        pages[pkey] = json.loads(json.dumps({k: v for k, v in data.items() if k not in ("pages", "printed", "outputs")},
                                            default=str))

    if want_slip:
        _emit("slip_html", "slip", "Shipping_Slip", "Shipping Slip", "slip_code",
              slip_form_code(data.get("eng"), data.get("wo"), data.get("prepared_by")), build_slip_html)
    if want_labels:
        _emit("labels_html", "labels", "Shipping_Labels", "Shipping Labels", "labels_code",
              labels_form_code(data.get("eng"), data.get("wo")), build_labels_html)
    if want_pack:
        # the label shows TODAY (Apex renders it at print time): freeze it into the
        # record so History and Re-emit reproduce the paper
        data["pack_date"] = datetime.date.today().strftime("%m/%d/%Y")
        data["pack_datecode"] = pack_datecode_yyww()
        _emit("pack_html", "pack", "Packing_Label", "Packing Label", "pack_code",
              pack_form_code(data.get("eng"), data.get("wo")), build_pack_html)
    return res

def _slip_temp_print_folder():
    """The app's own temp print folder (shared with the 4x2 labels and the
    COC / Packing print renders) — where a page goes when the WO folder can't
    take it. Never a stored path."""
    import tempfile
    return Path(tempfile.gettempdir()) / "PCC Shipping App - Print"

def _is_temp_print_path(p):
    """True for a page path inside a temp print folder (any PC's): a copy that
    never reached the share, which History and Re-emit must not point at."""
    try:
        return Path(str(p)).parent.name == "PCC Shipping App - Print"
    except Exception:
        return False

def _slip_page_source(slip, key):
    """What ONE printed page ('slip_html' | 'labels_html') was rendered from:
    the page's snapshot in slip['pages'] (the form can be edited and the OTHER
    page reprinted after this one; the paper doesn't change), else the recorded
    slip_data itself for a record without snapshots (older build)."""
    pkey = _SLIP_KINDS[key][0]
    pages = slip.get("pages") if isinstance(slip.get("pages"), dict) else {}
    src = pages.get(pkey) if isinstance(pages.get(pkey), dict) else None
    return src or slip

def _slip_page_code(slip, key):
    """The form code printed on that page (drawn once per print, so it
    identifies the paper); '' when the record has none."""
    src = _slip_page_source(slip, key)
    return str(src.get(_SLIP_KINDS[key][3]) or "").strip()

def _reemit_filed_page(folder, kind, code):
    """For a page printed while the share was down (no recorded path): the file
    in `folder` that already holds this print — a <eng>-<date>-<kind>[_HHMMSS]
    .html whose text carries the page's form code — or None. A read that fails
    counts as no match (the page is then written again under a fresh unique
    name; nothing is ever overwritten)."""
    code = str(code or "").strip()
    if not code:
        return None
    try:
        for p in sorted(Path(folder).glob(f"*-{kind}*.html")):
            try:
                if code in p.read_text(encoding="utf-8", errors="replace"):
                    return p
            except Exception:
                pass
    except Exception:
        pass
    return None

def _reemit_write_slip_html(folder, slip, key):
    """reemit_paperwork's writer for the slip pages: rebuild ONE page from what
    it was printed from (_slip_page_source) with its recorded form code, not a
    new one, under a today-dated unique name; the caller renames it to the
    ORIGINAL file name (or the slip-dated unique name for a page that was never
    filed)."""
    eng = sanitize_filename_part(slip.get("eng") or "", "unknown")
    kind = _SLIP_KINDS[key][1]
    src = _slip_page_source(slip, key)
    p = unique_file_path(folder, f"{eng}-{fmt_file_date()}-{kind}", ".html")
    html_text = _slip_builder(key)(src)
    return _write_new_file(p, html_text)

# ── SlipLabelsPopup's Tk-free half (Dhruvil 2026-09-24) ──────────────────────
# Pre-fill, Apex's required-field rules and the print action live here as plain
# functions so the test suite drives them without a window; the popup only
# wires Tk vars to them.

# Apex getRequiredFields: ship = cust/wo/qty/operator; shipl = cust/wo/qty (per box).
SLIP_REQUIRED_SLIP = (("customer", "Customer"), ("wo", "Work Order / Lot"),
                      ("qty", "Qty Shipped"), ("prepared_by", "Prepared By"))
SLIP_REQUIRED_LABELS = (("customer", "Customer"), ("wo", "Work Order / Lot"))
SLIP_REQUIRED_PACK = (("wo", "Work Order / Lot"), ("cpn", "Customer P/N"))   # Apex: pack-wo, pack-cpn

def _slip_pick(*vals):
    """First non-blank value that isn't the folder parser's 'N/A' placeholder."""
    for v in vals:
        s = str(v or "").strip()
        if s and s.upper() != "N/A":
            return s
    return ""

def slip_prefill(parsed_info=None, monday_hit=None, coc_data=None, packing_data=None,
                 operator="", wo="", serial_count=0, existing=None, division=None):
    """The popup's opening values. A saved slip_data (`existing`) wins outright
    — re-opening shows exactly what was saved/printed, every key normalised to
    the template shape. Otherwise the form fills from the job the way the COC
    / Packing popups do: parsed_info (folder, authoritative) + the Monday hit
    (PO / customer P/N / rev / customer) + what was already typed into the COC
    / Packing List (PO, P/N, rev, qty) + the main-window operator + today.
    Qty Shipped = the scanned serial count; one box row, its qty = that count
    (Apex's box table starts blank, but a one-box shipment is the common case)."""
    d = slip_data_template()
    if existing:
        for k in d:
            if k in existing and existing[k] is not None and k not in ("boxes", "printed", "outputs", "pages"):
                d[k] = existing[k]
        d["boxes"] = [{"size": str((b or {}).get("size") or ""), "qty": str((b or {}).get("qty") or ""),
                       "weight": str((b or {}).get("weight") or "")} for b in (existing.get("boxes") or [])]
        d["printed"].update({k: v for k, v in (existing.get("printed") or {}).items() if k in d["printed"]})
        d["outputs"].update({k: v for k, v in (existing.get("outputs") or {}).items() if k in d["outputs"]})
        pg = existing.get("pages") if isinstance(existing.get("pages"), dict) else {}
        d["pages"].update({k: v for k, v in pg.items() if k in d["pages"] and isinstance(v, dict)})
        for k in ("fair", "si", "ppap", "cofc"):
            d[k] = bool(d[k])
        d["customer_parts"] = "Yes" if str(d["customer_parts"]).strip().lower() == "yes" else "No"
        if not d["date"]:
            d["date"] = datetime.date.today().strftime("%Y-%m-%d")
        if division is not None:
            d["pack_mode"] = pack_mode_for(division)
        return d
    pi = parsed_info or {}; mh = monday_hit or {}; coc = coc_data or {}; pk = packing_data or {}
    d["customer"] = _slip_pick(pi.get("customer"), pk.get("customer"), mh.get("customer"))
    d["cpn"] = _slip_pick(pi.get("cpn"), mh.get("cpn"), coc.get("pn"), pk.get("pn"), pi.get("pn"))
    d["po"] = _slip_pick(coc.get("po"), pk.get("po"), pi.get("po"), mh.get("po"))
    d["wo"] = _slip_pick(wo, pk.get("wo"), mh.get("wo"))
    d["eng"] = _slip_pick(pi.get("eng"), coc.get("job"), pk.get("eng"), mh.get("eng"))
    d["rev"] = _slip_pick(pi.get("rev"), coc.get("rev"), monday_bom_rev(mh.get("rev")))   # the BOM rev (2026-10-05)
    d["qty"] = str(serial_count) if serial_count else _slip_pick(coc.get("qty"))
    d["datecode_wwyy"] = slip_datecode_wwyy(d["date"])
    d["prepared_by"] = str(operator or "").strip()
    d["boxes"] = [{"size": "", "qty": d["qty"], "weight": ""}]
    if division is not None:
        d["pack_mode"] = pack_mode_for(division)    # Apex's ITAR/MAG vs Industrial label (2026-10-19)
    return d

def _slip_size_value(v):
    """What a box row's Size dropdown holds, as stored: Apex's blank option
    ('--') and the '+ Add new size…' entry are never a size."""
    s = str(v or "").strip()
    return "" if s in (_SLIP_SIZE_NONE, _SLIP_ADD_SIZE) else s

def slip_required_missing(data, which="both"):
    """Apex's validateRequired for the slip ('ship') and the labels ('shipl'):
    the labels of the required fields that are blank, in form order. Labels
    need a qty on EVERY box (each box is its own label); with no box rows the
    slip qty stands in. 'both' = the union."""
    which = str(which or "both").strip().lower()
    d = data or {}
    def blank(k): return not str(d.get(k) or "").strip()
    missing = []
    if which == "pack":
        return [lbl for k, lbl in SLIP_REQUIRED_PACK if blank(k)]
    if which in ("slip", "both"):
        missing += [lbl for k, lbl in SLIP_REQUIRED_SLIP if blank(k)]
    if which in ("labels", "both"):
        missing += [lbl for k, lbl in SLIP_REQUIRED_LABELS if blank(k) and lbl not in missing]
        boxes = d.get("boxes") or []
        if boxes:
            for i, b in enumerate(boxes, 1):
                if not str((b or {}).get("qty") or "").strip():
                    missing.append(f"Box {i} qty")
        elif blank("qty") and "Qty Shipped" not in missing:
            missing.append("Qty Shipped")
    return missing

# What one print opens: the print copy's file-name stem and its label in messages.
_SLIP_PRINT_KINDS = {("slip",): ("Print_Slip", "Shipping Slip"),
                     ("labels",): ("Print_Labels", "Shipping Labels"),
                     ("slip", "labels"): ("Print_Slip_and_Labels", "Shipping Slip + Labels"),
                     ("pack",): ("Print_Packing_Label", "Packing Label")}

def _slip_print_copy_html(data, kinds):
    if kinds == ("slip", "labels"):
        return build_slip_labels_html(data)
    if kinds == ("slip",):
        return build_slip_html(data, print_copy=True)
    if kinds == ("pack",):
        return build_pack_html(data, print_copy=True)
    return build_labels_html(data, print_copy=True)

def _tidy_print_copies(folder, max_age_h=24):
    """Best-effort: delete the app's OWN throw-away print pages (slip / labels
    print copies, 4x2 board labels) older than a day from the temp print
    folder. Filed pages never live there under these names; a fresh copy a
    browser may still be loading is never touched."""
    try:
        cutoff = time.time() - max_age_h * 3600
        for pat in ("*-Print_Slip*.html", "*-Print_Labels*.html", "*-Print_Packing_Label*.html", "labels4x2_*.html"):
            for f in Path(folder).glob(pat):
                try:
                    if f.is_file() and f.stat().st_mtime < cutoff:
                        f.unlink()
                except Exception:
                    pass
    except Exception:
        pass

def slip_print_pages(wo_folder, data, which="both", open_file=None):
    """The Print buttons' non-Tk half. First the FILED page(s): write_slip_files
    draws the fresh form code(s), writes the record copy into the WO folder and
    stamps data['outputs']. Then ONE print copy of everything this print wrote
    is opened (Dhruvil 2026-09-25): a throw-away page in the temp print folder,
    the same page(s) as filed, whose Chrome tab closes itself after the print
    dialog; 'Print both' is ONE page (the slip, then every box label) so a
    single Print prints all of it. The filed pages are no longer opened here
    (History > Open still shows them, and they stay open). data['printed']
    [slip|labels] gets the app's date + time for each page the opened copy
    carries — only once it opened. If the print copy can't be written, the
    filed page(s) are opened one by one, as before 2026-09-25.
    Returns write_slip_files' dict plus 'print_html' (the print copy, or None),
    'opened' (paths) and 'open_errors' ([(label, error)]). open_file defaults
    to os.startfile; tests pass a stub."""
    opener = open_file or os.startfile
    res = write_slip_files(wo_folder, data, which)
    printed = data.get("printed")
    if not isinstance(printed, dict):
        printed = {"slip": None, "labels": None, "pack": None}
        data["printed"] = printed
    res["opened"] = []; res["open_errors"] = []; res["print_html"] = None
    kinds = tuple(k for key, k in (("slip_html", "slip"), ("labels_html", "labels"), ("pack_html", "pack")) if res.get(key))
    if not kinds:
        return res
    stem, label = _SLIP_PRINT_KINDS[kinds]
    try:
        folder = _slip_temp_print_folder()
        folder.mkdir(parents=True, exist_ok=True)      # the app's own temp folder, not a stored path
        _tidy_print_copies(folder)
        eng_safe = sanitize_filename_part(data.get("eng") or "", "unknown")
        path = _write_new_file(unique_file_path(folder, f"{eng_safe}-{fmt_file_date()}-{stem}", ".html"),
                               _slip_print_copy_html(data, kinds))
        res["print_html"] = path
        targets = [(path, kinds, label)]
    except Exception:
        targets = [(res[key], (k,), lbl) for key, (k, _kind, lbl, _ck) in _SLIP_KINDS.items() if res.get(key)]
    for p, ks, lbl in targets:
        try:
            opener(str(p))
            stamp = fmt_date() + " " + fmt_time()
            for k in ks:
                printed[k] = stamp
            res["opened"].append(str(p))
        except Exception as e:
            res["open_errors"].append((lbl, f"{type(e).__name__}: {e}"))
    return res

def slip_printed_summary(data):
    """One line for the popup's status row: 'Slip printed 9/24/2026 10:12 AM
    ·  Labels not printed' (fmt_date() + fmt_time(), no zero padding)."""
    pr = (data or {}).get("printed") or {}
    bits = []
    for k, lbl in (("slip", "Slip"), ("labels", "Labels"), ("pack", "Packing label")):
        bits.append(f"{lbl} printed {pr[k]}" if pr.get(k) else f"{lbl} not printed")
    return "  ·  ".join(bits)

def _slip_box_count_ok(proposed):
    """validatecommand for the popup's box-count spinbox: blank or one/two
    digits only. A scanner that fires into the spinbox ('16452' + Return)
    otherwise built 99 rows at once; the arrows / Return / Print still clamp
    the value to 1..99 (review round 2)."""
    s = str(proposed or "")
    return s == "" or (s.isdigit() and len(s) <= 2)

def _slip_missing_text(missing, cap=8):
    """The 'Missing required fields' bullet list, capped: the first `cap`
    labels and '… and N more' — a 99-box form must not raise a 100-line
    message box whose OK button is off the screen."""
    items = [str(m) for m in (missing or [])]
    shown = items[:cap]
    if len(items) > cap:
        shown.append(f"… and {len(items) - cap} more")
    return "\n  •  ".join(shown)

def _slip_scroll_fraction(top, height, view_top, view_height, total):
    """Where the popup's body canvas must scroll (yview_moveto fraction) so a
    field spanning [top, top+height] (px from the body's top) is fully inside
    the visible [view_top, view_top+view_height]; None when it already is or
    nothing scrolls. A keyboard / scanner operator Tabs into 'Ship Via' below
    the fold — the row must follow the focus (review round 2)."""
    try:
        top = float(top); height = float(height); view_top = float(view_top)
        view_height = float(view_height); total = float(total)
    except Exception:
        return None
    if total <= 0 or view_height <= 0 or total <= view_height:
        return None
    if top < view_top:
        want = top
    elif top + height > view_top + view_height:
        want = top + height - view_height
    else:
        return None
    return max(0.0, min(1.0, want / total))


###############################################################################
# DRAFT PERSISTENCE
###############################################################################
# Drafts let operators save partial shipment progress for supervisor review,
# without generating final files yet. A draft is a JSON sidecar in the WO folder
# named `.shipping_draft.json`. On WO load, the app auto-detects and resumes.

DRAFT_FILENAME = ".shipping_draft.json"

def _draft_path(wo_folder, draft_id=None):
    """File for one draft. Multiple drafts per WO coexist (added 2026-08-21 so a
    second batch can start while a supervisor still has an earlier draft open):
    each shipment session carries a draft_id (timestamp string) -> its own file
    '.shipping_draft_<id>.json'. A blank/None id means the pre-multi-draft legacy
    file '.shipping_draft.json' (still fully supported)."""
    if draft_id:
        return Path(wo_folder) / f".shipping_draft_{sanitize_filename_part(str(draft_id), fallback='d')}.json"
    return Path(wo_folder) / DRAFT_FILENAME

def save_shipment_draft(wo_folder, pending_photos, serials, coc_data, packing_data,
                        operator="", notes="", draft_id=None, slip_data=None,
                        started_at="", started_by=""):
    """
    Persist current shipment state to this shipment's own draft file (see
    _draft_path). Photo paths are stored, NOT moved — actual file moves happen
    only on Close. Returns the path written, or None on error.
    slip_data (Dhruvil 2026-09-24): the Shipping Slip & Labels form travels in
    the draft exactly like coc_data / packing_data (None when never opened).
    """
    try:
        wf = Path(wo_folder)
        if not wf.exists(): return None
        # Filter to photos that still exist on disk (operator may have deleted some)
        valid_photos = [p for p in (pending_photos or []) if Path(p).exists()]
        draft = {
            "version": 2,
            "draft_id": str(draft_id or ""),
            "saved_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "saved_by": operator or "",
            # When THIS shipment was first started (WO load), so a draft resumed in a
            # later app session keeps its real start for History grouping.
            "started_at": started_at or "",
            "started_by": started_by or "",
            "pending_photos": valid_photos,
            "serials": list(serials or []),
            "coc_data": coc_data or None,
            "packing_data": packing_data or None,
            "slip_data": slip_data or None,
            "notes": notes or "",
        }
        out = _draft_path(wf, draft_id)
        _atomic_write_json(out, draft)
        return str(out)
    except Exception:
        return None

def load_shipment_draft(wo_folder, path=None):
    """
    Load one draft — the legacy file by default, or the specific draft file given
    in `path` (from list_shipment_drafts). Returns the dict or None. Photo paths
    that no longer exist on disk are dropped silently.
    """
    try:
        df = Path(path) if path else (Path(wo_folder) / DRAFT_FILENAME)
        if not df.exists(): return None
        with open(df, encoding="utf-8") as f:
            d = json.load(f)
        # Drop missing photos so we don't reference dead paths
        d["pending_photos"] = [p for p in d.get("pending_photos", []) if Path(p).exists()]
        return d
    except Exception:
        return None

def list_shipment_drafts(wo_folder):
    """Every saved draft in this WO folder, NEWEST first. Returns
    [{'path','draft_id','saved_at','saved_by','serial_count','photo_count'}...];
    unreadable files are skipped. Includes the legacy single-draft file."""
    out = []
    try:
        wf = Path(wo_folder)
        for f in wf.iterdir():
            n = f.name.lower()
            if not (n.startswith(".shipping_draft") and n.endswith(".json")):
                continue
            try:
                with open(f, encoding="utf-8") as fh:
                    d = json.load(fh)
                if not isinstance(d, dict): continue
                out.append({"path": str(f),
                            "draft_id": str(d.get("draft_id") or ""),
                            "saved_at": str(d.get("saved_at") or ""),
                            "saved_by": str(d.get("saved_by") or ""),
                            "serial_count": len(d.get("serials") or []),
                            "photo_count": len(d.get("pending_photos") or [])})
            except Exception:
                continue
    except Exception:
        pass
    out.sort(key=lambda r: r["saved_at"], reverse=True)
    return out

def delete_shipment_draft(wo_folder, draft_id=None):
    """Remove ONE draft file — this shipment's own (by draft_id; blank/None means
    the legacy file). Used after Save and Close completes successfully; other
    shipments' drafts in the same WO folder are never touched.
    Returns (True, None) on success (or if file didn't exist),
    or (False, "<error description>") on failure.

    Retries up to 3 times with short delays in case the file is transiently
    held (Windows Explorer preview, antivirus scan, etc.).
    """
    import time
    try:
        df = _draft_path(wo_folder, draft_id)
        if not df.exists():
            return (True, None)
    except Exception as e:
        return (False, f"path error: {type(e).__name__}: {e}")

    last_err = None
    for attempt in range(3):
        try:
            df.unlink()
            return (True, None)
        except FileNotFoundError:
            return (True, None)
        except Exception as e:
            last_err = f"{type(e).__name__}: {e}"
            time.sleep(0.3)
    return (False, last_err or "delete failed")

def log_shipment_event(wo_folder, message):
    """No-op. Was previously appending to .shipping_log.txt for diagnostics, but
    operators don't want extra files in the WO folder. Kept as a callable so
    existing call sites in _fin / _save_draft continue to work without changes."""
    pass


###############################################################################
# UI HELPERS
###############################################################################

def _btn(parent,text,cmd,bg=C_GREEN,fg=C_WHITE,width=None,padx=16,pady=8,font_size=10):
    kw=dict(text=text,command=cmd,bg=bg,fg=fg,font=("Segoe UI",font_size,"bold"),relief="flat",cursor="hand2",padx=padx,pady=pady,activebackground=bg,activeforeground=fg)
    if width: kw["width"]=width
    b=tk.Button(parent,**kw)
    # Hover darkening reads the button's CURRENT base color (b._base_bg) instead of
    # closing over the creation-time hex: a runtime b.config(bg=...) now survives
    # mouse-over (previously Leave silently reverted it), and the hex is parsed
    # once per hover instead of being re-split character-wise from a closure.
    b._base_bg=bg
    b.bind("<Enter>",lambda e: b.config(bg=_darken(b._base_bg)))
    b.bind("<Leave>",lambda e: b.config(bg=b._base_bg))
    return b

class _HoverTip:
    """A small name tag shown while the mouse rests on an icon-only button
    (the footer's printer button, Dhruvil 2026-09-25). Binds with add='+' so
    the button's own hover colouring (_btn) keeps working; hidden on leave or
    click."""
    def __init__(self,widget,text,delay=400):
        self.w=widget; self.text=text; self.delay=delay; self._job=None; self._tip=None
        widget.bind("<Enter>",self._schedule,add="+")
        widget.bind("<Leave>",self._hide,add="+")
        widget.bind("<ButtonPress>",self._hide,add="+")
    def _schedule(self,e=None):
        self._cancel(); self._job=self.w.after(self.delay,self._show)
    def _cancel(self):
        if self._job:
            try: self.w.after_cancel(self._job)
            except Exception: pass
            self._job=None
    def _show(self):
        self._job=None
        if self._tip is not None: return
        try:
            t=tk.Toplevel(self.w); t.wm_overrideredirect(True)
            try: t.attributes("-topmost",True)
            except Exception: pass
            lb=tk.Label(t,text=self.text,bg="#FFFBE6",fg=C_DARK,relief="solid",bd=1,font=("Segoe UI",9),padx=8,pady=3)
            lb.pack(); t.update_idletasks()
            x=self.w.winfo_rootx(); y=self.w.winfo_rooty()-t.winfo_reqheight()-4
            t.wm_geometry(f"+{x}+{max(0,y)}"); self._tip=t
        except Exception:
            self._tip=None
    def _hide(self,e=None):
        self._cancel()
        if self._tip is not None:
            try: self._tip.destroy()
            except Exception: pass
            self._tip=None

def _darken(hx,amt=25):
    """#rrggbb darkened by amt per channel (used for button hover states)."""
    try:
        r,g,b=(int(hx[i:i+2],16) for i in (1,3,5))
        return f"#{max(0,r-amt):02x}{max(0,g-amt):02x}{max(0,b-amt):02x}"
    except Exception:
        return hx

def _section_header(parent,num,title):
    row=tk.Frame(parent,bg=C_BG); row.pack(fill="x",pady=(5,2))
    badge=tk.Frame(row,bg=C_GREEN,width=26,height=26); badge.pack(side="left",padx=(0,8)); badge.pack_propagate(False)
    tk.Label(badge,text=str(num),bg=C_GREEN,fg=C_WHITE,font=("Segoe UI",10,"bold")).place(relx=.5,rely=.5,anchor="center")
    tk.Label(row,text=title,bg=C_BG,fg=C_DARK,font=F_SECTION).pack(side="left")
    tk.Frame(parent,bg=C_BORDER,height=1).pack(fill="x")

def _card(parent):
    f=tk.Frame(parent,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); f.pack(fill="x",pady=2); return f

def _info_row(card,label,val="—",fg=C_DARK,bold=False):
    o=tk.Frame(card,bg=C_WHITE); o.pack(fill="x"); tk.Frame(card,bg="#EBEBEB",height=1).pack(fill="x")
    tk.Label(o,text=label,bg=C_WHITE,fg=C_SUBTLE,font=F_CAPTION,width=12,anchor="e",padx=8,pady=5).pack(side="left")
    tk.Frame(o,bg=C_BORDER,width=1).pack(side="left",fill="y",pady=2)
    l=tk.Label(o,text=val,bg=C_WHITE,fg=fg,font=(F_VALUE_BOLD if bold else F_VALUE),anchor="w",padx=10,wraplength=530,justify="left")
    l.pack(side="left",fill="x",expand=True); return l

def _center(win,parent):
    win.update_idletasks()
    px=parent.winfo_x()+(parent.winfo_width()-win.winfo_width())//2
    py=parent.winfo_y()+(parent.winfo_height()-win.winfo_height())//2
    win.geometry(f"+{max(0,px)}+{max(0,py)}")

###############################################################################
# POPUP DIALOGS
###############################################################################

class FileInUseDialog(tk.Toplevel):
    """
    Dialog shown when a write fails because the file is locked (open in Excel
    on this or another machine). Three outcomes:
      - 'retry': try the same write again (user has closed the file)
      - 'skip':  skip this file and continue closing the rest of the shipment
      - 'cancel': abort the whole close
    Result available as self.result.
    """
    def __init__(self, parent, file_label, file_path, error_text):
        super().__init__(parent); self.title("File is in use")
        self.resizable(False, False); self.transient(parent); self.grab_set(); self.configure(bg=C_BG)
        self.result = "cancel"
        h = tk.Frame(self, bg=C_AMBER); h.pack(fill="x")
        tk.Label(h, text=f"{file_label} file is open on another computer",
                 bg=C_AMBER, fg=C_WHITE, font=("Segoe UI", 12, "bold"),
                 wraplength=480, justify="left").pack(padx=24, pady=14, anchor="w")
        body = tk.Frame(self, bg=C_BG, padx=24, pady=14); body.pack(fill="x")
        tk.Label(body,
                 text=("Someone may have this file open in Excel.\n\n"
                       "To fix:\n"
                       "  1. Find who has the file open (often shown in Excel's title bar)\n"
                       "  2. Ask them to close it (or close it yourself)\n"
                       "  3. Click Retry below"),
                 bg=C_BG, fg=C_DARK, font=("Segoe UI", 10),
                 wraplength=480, justify="left").pack(anchor="w", pady=(0, 8))
        path_frame = tk.Frame(body, bg="#F4F4F4",
                              highlightbackground=C_BORDER, highlightthickness=1)
        path_frame.pack(fill="x", pady=(0, 8))
        tk.Label(path_frame, text="File:", bg="#F4F4F4", fg=C_SUBTLE,
                 font=("Segoe UI", 8)).pack(anchor="w", padx=8, pady=(6, 0))
        tk.Label(path_frame, text=str(file_path), bg="#F4F4F4", fg=C_DARK,
                 font=("Consolas", 9), wraplength=460, justify="left"
                 ).pack(anchor="w", padx=8, pady=(0, 6))
        tk.Label(body, text=f"Error: {error_text}", bg=C_BG, fg=C_RED,
                 font=("Segoe UI", 9), wraplength=480, justify="left").pack(anchor="w")
        btn_row = tk.Frame(self, bg=C_BG, pady=14); btn_row.pack()
        retry_btn = _btn(btn_row, "🔄  Retry", self._retry,
                          bg=C_GREEN, fg=C_WHITE,
                          font_size=11, padx=24, pady=10)
        retry_btn.pack(side="left", padx=6)
        _btn(btn_row, "Skip this file", self._skip,
             bg=C_AMBER, fg=C_WHITE, font_size=11, padx=18, pady=10
             ).pack(side="left", padx=6)
        _btn(btn_row, "Cancel close", self._cancel,
             bg="#DDD", fg=C_MID, font_size=11, padx=18, pady=10
             ).pack(side="left", padx=6)
        self.bind("<Return>", lambda e: self._retry())
        self.bind("<Escape>", lambda e: self._cancel())
        retry_btn.focus_set()
        _center(self, parent); self.wait_window()
    def _retry(self): self.result = "retry"; self.destroy()
    def _skip(self): self.result = "skip"; self.destroy()
    def _cancel(self): self.result = "cancel"; self.destroy()

class ConfirmActionDialog(tk.Toplevel):
    """
    Generic confirmation dialog for irreversible or important actions.
    Args:
      title: window title
      heading: large bold heading (e.g., "Close this shipment?")
      message: explanatory body text
      confirm_label: label for the confirm button (e.g., "Yes, Close")
      cancel_label: label for the cancel button (default "Cancel")
      heading_color: bg color for the heading bar (defaults to amber)
      confirm_color: bg color for the confirm button (defaults to green)
    Result available as self.result (True if confirmed, False otherwise).
    """
    def __init__(self, parent, title, heading, message,
                 confirm_label="Yes", cancel_label="Cancel",
                 heading_color=None, confirm_color=None,
                 checkbox_label=None, checkbox_default=False):
        super().__init__(parent); self.title(title); self.resizable(False, False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG); self.result = False
        self.checked = bool(checkbox_default)   # final state of the optional checkbox
        heading_color = heading_color or C_AMBER
        confirm_color = confirm_color or C_GREEN
        h = tk.Frame(self, bg=heading_color); h.pack(fill="x")
        tk.Label(h, text=heading, bg=heading_color, fg=C_WHITE,
                 font=("Segoe UI", 12, "bold"), wraplength=440, justify="left"
                 ).pack(padx=24, pady=14, anchor="w")
        body = tk.Frame(self, bg=C_BG, padx=24, pady=18); body.pack(fill="x")
        tk.Label(body, text=message, bg=C_BG, fg=C_DARK,
                 font=("Segoe UI", 10), wraplength=440, justify="left"
                 ).pack(anchor="w")
        if checkbox_label:
            self._cbv = tk.BooleanVar(value=bool(checkbox_default))
            tk.Checkbutton(body, text=checkbox_label, variable=self._cbv,
                           bg=C_BG, fg=C_DARK, activebackground=C_BG,
                           font=("Segoe UI", 10), anchor="w",
                           command=lambda: setattr(self, "checked", self._cbv.get())
                           ).pack(anchor="w", pady=(10, 0))
        btn_row = tk.Frame(self, bg=C_BG, pady=14); btn_row.pack()
        confirm_btn = _btn(btn_row, confirm_label, self._yes,
                            bg=confirm_color, fg=C_WHITE,
                            font_size=11, padx=24, pady=10)
        confirm_btn.pack(side="left", padx=8)
        _btn(btn_row, cancel_label, self._no, bg="#DDD", fg=C_MID,
             font_size=11, padx=24, pady=10).pack(side="left", padx=8)
        # Scanner-safety: a keyboard-wedge barcode scanner terminates every scan
        # with Enter. Bind Return on the confirm BUTTON only (not the window) and
        # arm it after a short delay, so a scan burst or double-tapped Enter that
        # lands as the dialog opens can't trigger an irreversible action. A human
        # pressing Enter 0.4s after the dialog appears is unaffected. Escape
        # (never sent by scanners) cancels immediately.
        self._armed = False
        self.after(400, self._arm)
        confirm_btn.bind("<Return>", lambda e: self._yes())
        self.bind("<Escape>", lambda e: self._no())
        confirm_btn.focus_set()
        _center(self, parent); self.wait_window()
    def _arm(self): self._armed = True
    def _yes(self):
        if not self._armed: return
        self.result = True; self.destroy()
    def _no(self): self.result = False; self.destroy()

class PhotoPopup(tk.Toplevel):
    def __init__(self,parent,pending):
        super().__init__(parent); self.title("CCA Photos"); self.geometry("960x560")
        self.transient(parent); self.resizable(True,True); self.grab_set(); self.configure(bg=C_BG)
        self.result=None; self._photos=list(pending); self._tkimg=None; self._build(); _center(self,parent)
        # Closing via the window's X keeps the photos added (same as Done) rather
        # than silently discarding the operator's selections.
        self.protocol("WM_DELETE_WINDOW", self._done)
        if not self._photos: self.after(150,self._add)
        self.wait_window()
    def _build(self):
        # Orange header keeps the accent contract the main screen teaches (orange = photos).
        h=tk.Frame(self,bg=C_ORANGE); h.pack(fill="x")
        tk.Label(h,text="CCA Photos",bg=C_ORANGE,fg=C_WHITE,font=("Segoe UI",14,"bold")).pack(side="left",padx=20,pady=14)
        self._ctr=tk.Label(h,text="0 photos",bg=C_ORANGE,fg="#FFE1C7",font=("Segoe UI",13,"bold")); self._ctr.pack(side="right",padx=20)
        t=tk.Frame(self,bg=C_BG,padx=20,pady=12); t.pack(fill="x")
        _btn(t,"+ Add photos",self._add,bg=C_GREEN,fg=C_WHITE,font_size=10).pack(side="left")
        self._del_btn=_btn(t,"Delete selected",self._del,bg=C_RED,fg=C_WHITE,font_size=10)
        self._del_btn.pack(side="left",padx=(8,0)); self._del_btn.config(state="disabled")
        # The destructive side-effect gets a real notice bar (was a 9pt inline label
        # that read as an error and was easy to miss).
        w=tk.Frame(self,bg=C_ORANGE_LITE,highlightbackground=C_ORANGE,highlightthickness=1); w.pack(fill="x",padx=20,pady=(0,8))
        tk.Label(w,text="⚠  Photos will be MOVED (cut) from their source folder when the shipment is saved",
                 bg=C_ORANGE_LITE,fg=C_ORANGE,font=("Segoe UI",10),padx=12,pady=6,anchor="w").pack(fill="x")
        mid=tk.Frame(self,bg=C_BG); mid.pack(fill="both",expand=True,padx=20,pady=(0,12))
        lf=tk.Frame(mid,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); lf.pack(side="left",fill="both",expand=True)
        sb=ttk.Scrollbar(lf); sb.pack(side="right",fill="y")
        self._lb=tk.Listbox(lf,font=("Segoe UI",11),bg=C_WHITE,fg=C_MID,selectbackground=C_GREEN_LITE,selectforeground=C_GREEN,relief="flat",bd=0,selectmode="extended",activestyle="none",yscrollcommand=sb.set)
        self._lb.pack(fill="both",expand=True,padx=8,pady=8); sb.config(command=self._lb.yview)
        self._lb.bind("<<ListboxSelect>>",self._preview)
        # Click-to-preview pane: verifies the RIGHT photos are attached. Loads one
        # image per click (a thumbnail strip would open every file up-front).
        pv=tk.Frame(mid,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1,width=310)
        pv.pack(side="left",fill="y",padx=(10,0)); pv.pack_propagate(False)
        self._pv_cap=tk.Label(pv,text="Select a photo to preview",bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",9),wraplength=290)
        self._pv_cap.pack(side="bottom",pady=(0,10))
        self._pv_img=tk.Label(pv,bg=C_WHITE); self._pv_img.pack(expand=True)
        f=tk.Frame(self,bg="#E8E8E8",pady=14); f.pack(fill="x")
        _btn(f,"✓  Done",self._done,bg=C_GREEN,fg=C_WHITE,font_size=11,padx=24,pady=10).pack()
        self._ref()
    def _preview(self,event=None):
        sel=self._lb.curselection()
        # destructive Delete only enabled when something is actually selected
        self._del_btn.config(state="normal" if sel else "disabled")
        if not sel:
            self._tkimg=None; self._pv_img.config(image="",text="")
            self._pv_cap.config(text="Select a photo to preview"); return
        p=self._photos[sel[-1]]
        try:
            from PIL import Image,ImageTk
            img=Image.open(p); img.thumbnail((290,380))
            self._tkimg=ImageTk.PhotoImage(img)
            self._pv_img.config(image=self._tkimg,text="")
        except Exception:
            self._tkimg=None
            self._pv_img.config(image="",text="(preview unavailable)",fg=C_SUBTLE,font=("Segoe UI",9))
        self._pv_cap.config(text=Path(p).name)
    def _add(self):
        files=filedialog.askopenfilenames(parent=self,title="Select CCA photos",filetypes=[("Images","*.jpg *.jpeg *.png *.bmp *.tif *.tiff"),("All","*.*")])
        for f in files:
            if f not in self._photos: self._photos.append(f)
        self._ref()
    def _del(self):
        sel=list(self._lb.curselection())
        if not sel: return
        if messagebox.askyesno("Confirm",f"Remove {len(sel)} photo(s)?",parent=self):
            for i in reversed(sel): self._photos.pop(i)
            self._ref()
    def _ref(self):
        self._lb.delete(0,"end")
        for i,p in enumerate(self._photos,1): self._lb.insert("end",f"   {i:>3}.   {Path(p).name}")
        self._ctr.config(text=f"{len(self._photos)} photo{'s' if len(self._photos)!=1 else ''}")
        self._preview()   # listbox rebuild clears selection — reset the pane too
    def _done(self): self.result=list(self._photos); self.destroy()

class DuplicateSerialDialog(tk.Toplevel):
    """Blocking duplicate-serial alert, hardened against barcode scanners.

    A keyboard-wedge scanner types the serial then presses Enter. With a native
    messagebox, scanning the next unit before reading the alert makes that
    trailing Enter press OK — the alert closes unread and the new serial is
    silently lost. Here the OK button stays disabled (and Return unarmed) for
    750ms, so a scan burst bounces off harmlessly and the operator must
    consciously acknowledge the duplicate before continuing.
    """
    def __init__(self, parent, serial):
        super().__init__(parent); self.title("Duplicate serial"); self.resizable(False, False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG)
        # Pull keyboard focus onto this (bind-free) toplevel immediately so a
        # scan fired during the unarmed window dies harmlessly here instead of
        # typing into the serial entry behind the dialog.
        self.focus_set()
        # Title-bar X goes through the same arming gate as OK/Escape.
        self.protocol("WM_DELETE_WINDOW", self._ack)
        h = tk.Frame(self, bg=C_RED); h.pack(fill="x")
        tk.Label(h, text="Duplicate — NOT added", bg=C_RED, fg=C_WHITE,
                 font=("Segoe UI", 12, "bold")).pack(padx=24, pady=14, anchor="w")
        body = tk.Frame(self, bg=C_BG, padx=24, pady=18); body.pack(fill="x")
        tk.Label(body, text=(f"{serial} is already in the list.\n\n"
                             "This serial was NOT added. Check for a double-scan "
                             "or a reused/mislabeled barcode."),
                 bg=C_BG, fg=C_DARK, font=("Segoe UI", 10), wraplength=420,
                 justify="left").pack(anchor="w")
        br = tk.Frame(self, bg=C_BG, pady=14); br.pack()
        self._ok = _btn(br, "OK", self._ack, bg=C_RED, fg=C_WHITE,
                        font_size=11, padx=32, pady=10)
        self._ok.pack()
        self._ok.config(state="disabled")
        self._armed = False
        self.after(750, self._arm)
        self._ok.bind("<Return>", lambda e: self._ack())
        self.bind("<Escape>", lambda e: self._ack())
        self.bell()
        _center(self, parent); self.wait_window()
    def _arm(self):
        self._armed = True
        try: self._ok.config(state="normal"); self._ok.focus_set()
        except Exception: pass
    def _ack(self):
        if not self._armed: return
        self.destroy()

class HistoricalDuplicateDialog(tk.Toplevel):
    """Blocking alert when a scanned serial was already shipped in a PAST closed
    shipment of the same job (Eng# + P/N). Scanner-proof like DuplicateSerialDialog:
    both buttons stay disabled for 750ms, and ONLY the safe 'Skip' is bound to
    Enter/Escape — so a scan burst can never dismiss it and can never trigger the
    override. 'Supervisor override…' is a deliberate click, gated by admin login in
    the caller. self.result is 'skip' (default) or 'override'."""
    def __init__(self, parent, serial, occ):
        super().__init__(parent); self.title("Serial already shipped")
        self.resizable(False, False); self.transient(parent); self.grab_set()
        self.configure(bg=C_BG); self.result = "skip"; self.focus_set()
        self.protocol("WM_DELETE_WINDOW", self._skip)
        h = tk.Frame(self, bg=C_AMBER); h.pack(fill="x")
        tk.Label(h, text="⚠  Serial already shipped before", bg=C_AMBER, fg=C_WHITE,
                 font=("Segoe UI", 12, "bold")).pack(padx=24, pady=14, anchor="w")
        where = []
        if occ.get("wo"):       where.append(f"WO {occ['wo']}")
        if occ.get("ts"):       where.append(str(occ["ts"]))
        if occ.get("operator"): where.append(f"by {occ['operator']}")
        wtxt = "  ·  ".join(where) if where else "a previous closed shipment"
        body = tk.Frame(self, bg=C_BG, padx=24, pady=18); body.pack(fill="x")
        tk.Label(body, text=str(serial), bg=C_BG, fg=C_DARK,
                 font=("Courier New", 15, "bold")).pack(anchor="w")
        tk.Label(body, text=(f"was already scanned and shipped in:\n    {wtxt}\n\n"
                             "It was NOT added. Check for a wrong board, a reused or "
                             "mislabeled barcode, or a unit that already shipped. A "
                             "supervisor can override to add it anyway."),
                 bg=C_BG, fg=C_MID, font=("Segoe UI", 10), wraplength=440,
                 justify="left").pack(anchor="w", pady=(8, 0))
        br = tk.Frame(self, bg=C_BG, pady=14); br.pack()
        self._skip_btn = _btn(br, "Skip — don't add", self._skip, bg=C_GREEN, fg=C_WHITE,
                              font_size=11, padx=22, pady=10)
        self._skip_btn.pack(side="left", padx=(0, 10))
        self._ovr_btn = _btn(br, "Supervisor override…", self._override, bg=C_AMBER, fg=C_WHITE,
                             font_size=11, padx=22, pady=10)
        self._ovr_btn.pack(side="left")
        self._skip_btn.config(state="disabled"); self._ovr_btn.config(state="disabled")
        self._armed = False; self.after(750, self._arm)
        self.bind("<Return>", lambda e: self._skip())
        self.bind("<Escape>", lambda e: self._skip())
        self.bell(); _center(self, parent); self.wait_window()
    def _arm(self):
        self._armed = True
        for b in (self._skip_btn, self._ovr_btn):
            try: b.config(state="normal")
            except Exception: pass
        try: self._skip_btn.focus_set()
        except Exception: pass
    def _skip(self):
        if not self._armed: return
        self.result = "skip"; self.destroy()
    def _override(self):
        if not self._armed: return
        self.result = "override"; self.destroy()

class OutlierReviewDialog(tk.Toplevel):
    """Advisory review shown at submit: lists serials that look odd vs. the batch,
    each with a Keep/Remove choice (default Keep — never silently drops a serial).
    self.result is the list of serials the operator chose to REMOVE (possibly
    empty). Cancel / Escape / window-X keep everything."""
    def __init__(self, parent, outliers):
        super().__init__(parent); self.title("Check these serial numbers")
        self.resizable(False, True); self.transient(parent); self.grab_set()
        self.configure(bg=C_BG); self.result = []; self._vars = {}
        self._items = list(outliers)
        self.protocol("WM_DELETE_WINDOW", self._keep_all)
        h = tk.Frame(self, bg=C_AMBER); h.pack(fill="x")
        tk.Label(h, text="⚠  Some serials look different from the rest", bg=C_AMBER,
                 fg=C_WHITE, font=("Segoe UI", 12, "bold")).pack(padx=24, pady=13, anchor="w")
        intro = tk.Frame(self, bg=C_BG, padx=22); intro.pack(fill="x", pady=(14, 4))
        tk.Label(intro, text=(f"{len(self._items)} of the scanned serials don’t match the pattern of "
                              "the rest of the batch. Review each one — tick “Remove” to drop a bad "
                              "scan, or leave it to keep it as-is. Nothing is removed unless you tick it."),
                 bg=C_BG, fg=C_MID, font=("Segoe UI", 9), wraplength=520,
                 justify="left").pack(anchor="w")
        lf = tk.Frame(self, bg=C_WHITE, highlightbackground=C_BORDER, highlightthickness=1)
        lf.pack(fill="both", expand=True, padx=22, pady=10)
        cv = tk.Canvas(lf, bg=C_WHITE, highlightthickness=0,
                       height=min(340, max(70, 66 * len(self._items))))
        sb = ttk.Scrollbar(lf, orient="vertical", command=cv.yview); cv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); cv.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(cv, bg=C_WHITE); wid = cv.create_window((0, 0), window=inner, anchor="nw")
        cv.bind("<Configure>", lambda e: cv.itemconfigure(wid, width=e.width))
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        for serial, reason in self._items:
            row = tk.Frame(inner, bg=C_WHITE); row.pack(fill="x", padx=10, pady=6)
            tk.Frame(inner, bg="#EEEEEE", height=1).pack(fill="x")
            v = tk.BooleanVar(value=False); self._vars[serial] = v
            tk.Checkbutton(row, text="Remove", variable=v, bg=C_WHITE, fg=C_RED,
                           activebackground=C_WHITE, activeforeground=C_RED,
                           font=("Segoe UI", 9, "bold")).pack(side="right", padx=(8, 0))
            txt = tk.Frame(row, bg=C_WHITE); txt.pack(side="left", fill="x", expand=True)
            tk.Label(txt, text=str(serial), bg=C_WHITE, fg=C_DARK,
                     font=("Courier New", 12, "bold")).pack(anchor="w")
            tk.Label(txt, text=str(reason), bg=C_WHITE, fg=C_SUBTLE, font=("Segoe UI", 9),
                     wraplength=430, justify="left").pack(anchor="w")
        br = tk.Frame(self, bg="#E8E8E8", pady=12); br.pack(fill="x")
        _btn(br, "Remove all flagged", self._remove_all, bg=C_RED, fg=C_WHITE,
             font_size=9, padx=12, pady=8).pack(side="left", padx=(22, 0))
        _btn(br, "Keep all", self._keep_all, bg="#DDD", fg=C_MID,
             font_size=9, padx=12, pady=8).pack(side="left", padx=(8, 0))
        self._done_btn = _btn(br, "✓  Apply and continue", self._apply, bg=C_GREEN, fg=C_WHITE,
                              font_size=10, padx=18, pady=8)
        self._done_btn.pack(side="right", padx=(0, 22))
        self.bind("<Escape>", lambda e: self._keep_all())
        self._armed = False; self.after(400, lambda: setattr(self, "_armed", True))
        _center(self, parent); self.wait_window()
    def _remove_all(self):
        for v in self._vars.values(): v.set(True)
    def _keep_all(self):
        self.result = []; self.destroy()
    def _apply(self):
        if not self._armed: return
        self.result = [s for s, v in self._vars.items() if v.get()]
        self.destroy()

class PNPromptDialog(tk.Toplevel):
    """Small prompt to confirm the P/N when a job loads without one. The serial
    duplicate check is scoped by P/N, so it must be known before scanning. Returns
    the entered P/N in self.result, or '' if the operator skips."""
    def __init__(self, parent, wo, eng):
        super().__init__(parent); self.title("Confirm P/N"); self.resizable(False, False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG); self.result = ""
        h = tk.Frame(self, bg=C_BLUE); h.pack(fill="x")
        tk.Label(h, text="Confirm the Part Number", bg=C_BLUE, fg=C_WHITE,
                 font=("Segoe UI", 12, "bold")).pack(padx=22, pady=12, anchor="w")
        bd = tk.Frame(self, bg=C_BG, padx=24, pady=18); bd.pack(fill="x")
        sub = f"WO {wo}" + (f"  ·  Eng# {eng}" if eng else "")
        tk.Label(bd, text=(sub + "\n\nThis job loaded without a P/N. Enter the P/N so the app can "
                           "check serials against past shipments of the same part.\n\n"
                           "(You can skip — the duplicate check will simply be off for this shipment.)"),
                 bg=C_BG, fg=C_DARK, font=("Segoe UI", 10), wraplength=420,
                 justify="left").pack(anchor="w")
        self._v = tk.StringVar()
        e = ttk.Entry(bd, textvariable=self._v, font=("Segoe UI", 13)); e.pack(fill="x", pady=(12, 0), ipady=5)
        e.bind("<Return>", lambda ev: self._ok())
        br = tk.Frame(self, bg=C_BG, pady=12); br.pack()
        _btn(br, "Use this P/N", self._ok, bg=C_GREEN, fg=C_WHITE, padx=20, pady=9).pack(side="left", padx=(0, 10))
        _btn(br, "Skip", self._skip, bg="#DDD", fg=C_MID, padx=20, pady=9).pack(side="left")
        self.bind("<Escape>", lambda ev: self._skip())
        _center(self, parent); e.focus_set(); self.wait_window()
    def _ok(self):
        self.result = self._v.get().strip(); self.destroy()
    def _skip(self):
        self.result = ""; self.destroy()

class SerialPopup(tk.Toplevel):
    def __init__(self,parent,serials,notes="",dup_ctx=None,job_pn=""):
        super().__init__(parent); self.title("CCA Serial Numbers"); self.geometry("760x700")
        self.transient(parent); self.resizable(True,True); self.grab_set(); self.configure(bg=C_BG)
        self.result=None; self._sn=list(serials)
        self.notes_result=notes or ""; self._notes_init=notes or ""
        # Job P/N (from the loaded WO) — anchors the odd-serial check so it works
        # even for a tiny batch. Falls back to the dup context's pn if not passed.
        self._job_pn=job_pn or (dup_ctx or {}).get("pn","") or ""
        # dup_ctx (or None to disable the historical check): a dict with keys
        #   lookup(serial)->occ|None, admins, approved (shared set), eng, pn, wo,
        #   division, operator. See ShippingApp._dup_ctx_for_scan().
        self._dup=dup_ctx or None
        # 'approved' is a set of normalized serials a supervisor cleared. It is
        # SHARED with the app so approvals survive closing/reopening the popup and
        # the submit sweep never re-prompts for an already-approved serial.
        self._approved=(self._dup or {}).get("approved")
        if self._approved is None: self._approved=set()
        # keys the approved set with the SAME normalization as the history index —
        # exact match, so a different Rev counts as a different serial.
        self._idfn=(self._dup or {}).get("identity") or _norm_serial
        self.overrides=[]   # [{serial, approved_by, prior}] collected this session
        # Batch override: ONE supervisor login (the 🔓 button) authorizes EVERY
        # duplicate re-scan this session — boards legitimately ship twice sometimes,
        # and a login per board was unusable for a big batch. None = off.
        self._dup_session_admin=None
        self._build(); _center(self,parent)
        # Closing via the window's X keeps scanned serials (same as Done) instead
        # of throwing away everything scanned in this session.
        self.protocol("WM_DELETE_WINDOW", self._done)
        self._entry.focus_set(); self.wait_window()
    def _build(self):
        h=tk.Frame(self,bg=C_BLUE); h.pack(fill="x")
        tk.Label(h,text="CCA Serial Numbers",bg=C_BLUE,fg=C_WHITE,font=("Segoe UI",14,"bold")).pack(side="left",padx=20,pady=14)
        self._ctr=tk.Label(h,text="0 serials",bg=C_BLUE,fg="#D6EFFF",font=("Segoe UI",16,"bold")); self._ctr.pack(side="right",padx=20)
        s=tk.Frame(self,bg=C_BG,padx=20,pady=14); s.pack(fill="x")
        tk.Label(s,text="Scan barcode or type serial — press Enter to add:",bg=C_BG,fg=C_MID,font=("Segoe UI",9)).pack(anchor="w",pady=(0,6))
        sr=tk.Frame(s,bg=C_BG); sr.pack(fill="x")
        self._var=tk.StringVar()
        self._entry=ttk.Entry(sr,textvariable=self._var,font=("Segoe UI",20)); self._entry.pack(side="left",fill="x",expand=True,ipady=10)
        self._entry.bind("<Return>",lambda e: self._add())
        _btn(sr,"Add",self._add,bg=C_BLUE,fg=C_WHITE,font_size=12,padx=24,pady=10).pack(side="left",padx=(10,0))
        _btn(sr,"Paste List",self._paste,bg=C_GREEN_MED,fg=C_WHITE,font_size=10,padx=16,pady=10).pack(side="left",padx=(8,0))
        t=tk.Frame(self,bg=C_BG,padx=20,pady=4); t.pack(fill="x")
        self._del_btn=_btn(t,"Delete selected",self._del,bg=C_RED,fg=C_WHITE,font_size=10)
        self._del_btn.pack(side="left"); self._del_btn.config(state="disabled")
        # 4x2 board labels on the HPKE200 (same layout as the Apex spack-4x2 label):
        # prints every scanned serial — or only the SELECTED rows, e.g. to reprint one.
        _btn(t,"🖨 Print 4×2 labels",self._print_labels,bg=C_TEAL,fg=C_WHITE,font_size=10).pack(side="right")
        # Batch re-ship authorization (one login for the whole session).
        if self._dup:
            self._sess_btn=_btn(t,"🔓 Supervisor: allow re-ship…",self._toggle_dup_session,
                                bg=C_AMBER,fg=C_WHITE,font_size=10)
            self._sess_btn.pack(side="left",padx=(8,0))
        lf=tk.Frame(self,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); lf.pack(fill="both",expand=True,padx=20,pady=(4,12))
        sb=ttk.Scrollbar(lf); sb.pack(side="right",fill="y")
        self._lb=tk.Listbox(lf,font=("Courier New",12),bg=C_WHITE,fg=C_MID,selectbackground=C_BLUE_LITE,selectforeground=C_BLUE,relief="flat",bd=0,selectmode="extended",activestyle="none",yscrollcommand=sb.set)
        self._lb.pack(fill="both",expand=True,padx=8,pady=8); sb.config(command=self._lb.yview)
        self._lb.bind("<<ListboxSelect>>",lambda e: self._del_btn.config(state="normal" if self._lb.curselection() else "disabled"))
        nf=tk.Frame(self,bg=C_BG,padx=20); nf.pack(fill="x",pady=(0,10))
        tk.Label(nf,text="Shipment notes (optional) — where it's going / special instructions:",bg=C_BG,fg=C_MID,font=("Segoe UI",9)).pack(anchor="w",pady=(0,4))
        self._notes=tk.Text(nf,height=3,font=("Segoe UI",10),wrap="word",bg=C_WHITE,relief="flat",highlightbackground=C_BORDER,highlightthickness=1,padx=6,pady=4)
        self._notes.pack(fill="x")
        if self._notes_init: self._notes.insert("1.0", self._notes_init)
        f=tk.Frame(self,bg="#E8E8E8",pady=14); f.pack(fill="x")
        _btn(f,"✓  Done",self._done,bg=C_GREEN,fg=C_WHITE,font_size=11,padx=24,pady=10).pack()
        self._ref()
    def _add(self):
        v=self._var.get().strip()
        if not v: return
        # Handle multi-line paste directly into the entry
        lines=[l.strip() for l in v.replace('\r','').split('\n') if l.strip()]
        if len(lines)>1:
            added,dupes=self._add_many(lines)
            self._var.set(""); self._entry.focus_set()
            self._show_paste_result(added,dupes)
            return
        if v in self._sn:
            # Duplicates are a real signal (double-scan / reused or mislabeled
            # barcode) — stop the operator with a blocking alert so they
            # investigate. DuplicateSerialDialog is scanner-proof: the trailing
            # Enter of the NEXT scan can't dismiss it unread, unlike a native
            # messagebox whose OK button fires on Return.
            DuplicateSerialDialog(self, v)
            # Tk grabs don't stack: the child dialog's grab_set stole our modal
            # grab and destroying it merely releases the grab — re-assert it so
            # SerialPopup stays modal after the alert is acknowledged.
            self.grab_set()
            self._var.set(""); self._entry.focus_set(); return
        # Historical duplicate: was this serial already shipped in a PAST closed
        # shipment of the same job (Eng#+P/N)? If so, block (scanner-proof) unless
        # a supervisor overrides. Best-effort while the index is still loading; the
        # submit sweep in _done is the guaranteed backstop.
        if not self._check_history(v):
            self._var.set(""); self._entry.focus_set(); return
        self._sn.append(v); self._ref(); self._lb.see("end"); self._var.set(""); self._entry.focus_set()
    def _paste(self):
        try: clip=self.clipboard_get()
        except Exception: messagebox.showinfo("Paste","Nothing in clipboard.",parent=self); return
        lines=[l.strip() for l in clip.replace('\r','').split('\n') if l.strip()]
        if not lines: messagebox.showinfo("Paste","No serial numbers found in clipboard.",parent=self); return
        added,dupes=self._add_many(lines)
        self._show_paste_result(added,dupes)
    def _add_many(self,lines):
        """Add each serial not already present. Returns (added_count, dupes) where
        dupes is the list of serials that were skipped — includes both serials
        already in the list AND repeats within the pasted batch."""
        added=0; dupes=[]
        for sn in lines:
            if sn in self._sn: dupes.append(sn)
            else: self._sn.append(sn); added+=1
        self._ref(); self._lb.see("end")
        return added,dupes
    def _show_paste_result(self,added,dupes):
        """Report a paste, LISTING the actual duplicates (with repeat counts) so the
        operator can see which serials were skipped, not just how many."""
        msg=f"Added {added} serial(s)."
        if dupes:
            counts={}
            for d in dupes: counts[d]=counts.get(d,0)+1
            uniq=list(counts.keys()); shown=uniq[:40]
            body="\n".join(f"   • {d}"+(f"  (x{counts[d]})" if counts[d]>1 else "") for d in shown)
            more=f"\n   … and {len(uniq)-40} more" if len(uniq)>40 else ""
            msg+=f"\n\n{len(dupes)} duplicate(s) skipped:\n{body}{more}"
        messagebox.showinfo("Paste result",msg,parent=self)
    def _del(self):
        sel=list(self._lb.curselection())
        if not sel: return
        if messagebox.askyesno("Confirm",f"Remove {len(sel)} serial(s)?",parent=self):
            for i in reversed(sel): self._sn.pop(i)
            self._ref()
    def _ref(self):
        self._lb.delete(0,"end")
        for i,s in enumerate(self._sn,1): self._lb.insert("end",f"   {i:>4}.   {s}")
        self._ctr.config(text=f"{len(self._sn)} serial{'s' if len(self._sn)!=1 else ''}")
        self._del_btn.config(state="disabled")   # rebuild clears the selection
    def _print_labels(self):
        """Generate 4×2 board labels (P/N + QR on the left, Serial + QR on the
        right — the Apex spack-4x2 layout) for the scanned serials and open them
        in the browser: the print dialog pops automatically, operator picks the
        HPKE200 and everything prints as one job (one 4x2 page per serial).
        With rows SELECTED, only those are printed (handy for reprinting one)."""
        if not self._sn:
            messagebox.showwarning("No serials","Scan at least one serial first.",parent=self); return
        sel=list(self._lb.curselection())
        serials=[self._sn[i] for i in sel] if sel else list(self._sn)
        pn=(self._job_pn or "").strip()
        if not _norm_key(pn):            # blank or N/A — ask, and remember for this popup
            from tkinter import simpledialog
            pn=(simpledialog.askstring("Part number",
                "P/N to print on the labels:",parent=self) or "").strip()
            self.grab_set()              # simpledialog releases our modal grab
            if not pn:
                return
            self._job_pn=pn
        import tempfile
        try:
            folder=Path(tempfile.gettempdir())/"PCC Shipping App - Print"
            path=write_label_html_4x2(folder, pn, serials,
                                      wo=str((self._dup or {}).get("wo") or ""))
            os.startfile(str(path))
        except Exception as e:
            messagebox.showerror("Labels",
                f"Could not create the labels:\n{type(e).__name__}: {e}",parent=self)
        self._entry.focus_set()
    # ── historical-duplicate + outlier checks ────────────────────────────────
    def _check_history(self, v):
        """True if serial v may be added; False to skip. Blocks (scanner-proof)
        when v was shipped before in this job, unless a supervisor overrides."""
        if not self._dup: return True
        k=self._idfn(v)
        if k in self._approved: return True
        try: occ=self._dup["lookup"](v)
        except Exception: occ=None
        if not occ: return True
        return self._resolve_dup(v, occ)
    def _toggle_dup_session(self):
        """One supervisor login authorizes re-shipping EVERY duplicate scanned in
        this session (each board still logged individually with the approver's id).
        Click again to end the authorization."""
        if self._dup_session_admin:
            self._dup_session_admin=None
            try: self._sess_btn.config(text="🔓 Supervisor: allow re-ship…",bg=C_AMBER)
            except Exception: pass
            return
        admins=(self._dup or {}).get("admins") or []
        if not admins:
            try: admins=load_admins()
            except Exception: admins=[]
        if not admins:
            messagebox.showerror("Admin roster unavailable",
                "The admin list could not be loaded from the share, so the "
                "authorization can't be verified right now.",parent=self)
            return
        lg=AdminLoginDialog(self, admins); self.grab_set()
        if not getattr(lg,"result",False): return
        self._dup_session_admin=getattr(lg,"admin_id","") or "admin"
        try: self._sess_btn.config(
                text=f"🔒 RE-SHIP ALLOWED ({self._dup_session_admin}) — click to end",bg=C_RED)
        except Exception: pass
    def _resolve_dup(self, v, occ):
        """Show the historical-dup dialog for v; return True to keep (supervisor
        approved), False to skip/remove. Shared by live-add and the submit sweep.
        While the batch authorization (🔓) is active, duplicates are approved and
        logged WITHOUT a dialog — one login covers the whole batch."""
        k=self._idfn(v)
        if k in self._approved: return True
        if self._dup_session_admin:
            self._approved.add(k)
            self._log_override(v, occ, self._dup_session_admin)
            return True
        dlg=HistoricalDuplicateDialog(self, v, occ); self.grab_set()
        if dlg.result!="override": return False
        admins=(self._dup or {}).get("admins") or []
        lg=AdminLoginDialog(self, admins); self.grab_set()
        if not getattr(lg,"result",False): return False
        self._approved.add(k)
        self._log_override(v, occ, getattr(lg,"admin_id",""))
        return True
    def _log_override(self, serial, occ, admin_id):
        """Record a supervisor override — kept both as a durable standalone history
        event AND collected for this shipment's close/draft record (audit trail)."""
        entry={"serial": str(serial), "approved_by": admin_id or "",
               "prior": {"wo": occ.get("wo",""), "ts": occ.get("ts",""),
                         "operator": occ.get("operator",""), "event_id": occ.get("event_id","")}}
        self.overrides.append(entry)
        d=self._dup or {}
        try:
            write_history_event("dup_override", {
                "wo": d.get("wo",""), "division": d.get("division",""),
                "operator": d.get("operator",""), "eng": d.get("eng",""), "pn": d.get("pn",""),
                "serial": str(serial), "approved_by": admin_id or "", "prior": entry["prior"]})
        except Exception: pass
    def _historical_sweep(self):
        """At submit, re-check the WHOLE list against the history index for serials
        that slipped in (pasted, or scanned before the index finished loading) and
        aren't supervisor-approved. Each is resolved with the same per-serial dialog
        (skip -> removed, override -> kept)."""
        if not self._dup: return
        remove=[]
        for s in list(self._sn):
            if self._idfn(s) in self._approved: continue
            try: occ=self._dup["lookup"](s)
            except Exception: occ=None
            if not occ: continue
            if not self._resolve_dup(s, occ):
                remove.append(s)
        if remove:
            rm=set(remove); self._sn=[s for s in self._sn if s not in rm]; self._ref()
    def _outlier_review(self):
        """At submit, flag serials that look odd vs. the batch and let the operator
        drop the bad scans. Advisory — never blocks."""
        try: outliers=detect_serial_outliers(self._sn, pn=self._job_pn)
        except Exception: outliers=[]
        if not outliers: return
        dlg=OutlierReviewDialog(self, outliers); self.grab_set()
        rm=set(dlg.result or [])
        if rm:
            self._sn=[s for s in self._sn if s not in rm]; self._ref()
    def _done(self):
        # Final safety checks before returning: historical duplicates first
        # (correctness — may remove/keep-with-override), then the advisory outlier
        # review on what remains.
        self._historical_sweep()
        self._outlier_review()
        self.result=list(self._sn)
        try: self.notes_result=self._notes.get("1.0","end").strip()
        except Exception: pass
        self.destroy()

class COCPopup(tk.Toplevel):
    TO=["CCA","Box Build","Cable and Wire Harness","Bare PCB"]
    AO=["Lead free","Leaded","Special","N/A"]
    def __init__(self,parent,wo,sc,ed=None,pi=None):
        super().__init__(parent); self.title("Certificate of Compliance"); self.geometry("600x600")
        self.resizable(False,False); self.transient(parent); self.grab_set(); self.configure(bg=C_BG)
        self.result=None; self.print_requested=False; self._wo=wo; self._sc=sc; pi=pi or {}
        self._sd=tk.StringVar(value=fmt_date()); self._po=tk.StringVar(value=pi.get("po",""))
        self._pn=tk.StringVar(value=pi.get("pn","")); self._rv=tk.StringVar(value=pi.get("rev",""))
        self._qty=tk.StringVar(value=str(sc)); self._dc=tk.StringVar(value=fmt_datecode())
        self._job=tk.StringVar(value=pi.get("eng","") or wo); self._ty=tk.StringVar(value=self.TO[0])
        self._ap=tk.StringVar(value=self.AO[0])
        if ed:
            for k,a in [('ship_date','_sd'),('po','_po'),('pn','_pn'),('rev','_rv'),('qty','_qty'),('datecode','_dc'),('job','_job')]:
                v=ed.get(k,'')
                if v: getattr(self,a).set(v)
            if ed.get("type"): self._ty.set(ed["type"])
            if ed.get("assembly_process"): self._ap.set(ed["assembly_process"])
        self._build(); _center(self,parent); self.wait_window()
    def _build(self):
        h=tk.Frame(self,bg=C_PURPLE); h.pack(fill="x")
        tk.Label(h,text="Certificate of Compliance",bg=C_PURPLE,fg=C_WHITE,font=("Segoe UI",14,"bold")).pack(side="left",padx=20,pady=14)
        tk.Label(h,text=f"WO {self._wo}",bg=C_PURPLE,fg="#D4B8E8",font=("Segoe UI",11,"bold")).pack(side="right",padx=20)
        # Reserve the button bar at the BOTTOM first (side="bottom") so a tall form
        # can never push Save/Cancel off this fixed-height, non-resizable window.
        # (That was the bug: the required-field row + themed field heights overflowed
        # the window and clipped the buttons.)
        f=tk.Frame(self,bg="#E8E8E8",pady=14); f.pack(side="bottom",fill="x")
        bf=tk.Frame(f,bg="#E8E8E8"); bf.pack()
        save_b=_btn(bf,"✓  Save",lambda: self._save(False),bg=C_PURPLE,fg=C_WHITE,font_size=11,padx=16,pady=10); save_b.pack(side="left",padx=5)
        print_b=_btn(bf,"🖨  Save and Print",lambda: self._save(True),bg=C_TEAL,fg=C_WHITE,font_size=11,padx=16,pady=10); print_b.pack(side="left",padx=5)
        cancel_b=_btn(bf,"Cancel",self._cancel,bg="#DDD",fg=C_MID,font_size=11,padx=16,pady=10); cancel_b.pack(side="left",padx=5)
        bd=tk.Frame(self,bg=C_BG,padx=28,pady=12); bd.pack(fill="both",expand=True)
        def fr(p,l,v,cv=None,required=False):
            r=tk.Frame(p,bg=C_BG); r.pack(fill="x",pady=3)
            tk.Label(r,text=l,bg=C_BG,fg=C_MID,font=("Segoe UI",9),width=18,anchor="e").pack(side="left")
            # fixed-width star column keeps required and optional rows aligned
            tk.Label(r,text="*" if required else " ",bg=C_BG,fg=C_RED,font=("Segoe UI",10,"bold"),width=1).pack(side="left",padx=(0,6))
            if cv:
                w=ttk.Combobox(r,textvariable=v,values=cv,state="readonly",font=("Segoe UI",11),width=34); w.pack(side="left",fill="x",expand=True,ipady=3)
            else:
                w=ttk.Entry(r,textvariable=v,font=("Segoe UI",11),width=36); w.pack(side="left",fill="x",expand=True,ipady=3)
            return w
        fr(bd,"Ship Date:",self._sd); self._po_e=fr(bd,"PO #:",self._po,required=True); self._pn_e=fr(bd,"P/N:",self._pn,required=True)
        fr(bd,"Rev:",self._rv); fr(bd,"QTY:",self._qty); fr(bd,"Date Code (YYWW):",self._dc)
        fr(bd,"PCC Job # (Eng):",self._job); fr(bd,"Type:",self._ty,cv=self.TO); fr(bd,"Assembly Process:",self._ap,cv=self.AO)
        tk.Label(bd,text="*  required     ·     Serial numbers are copied automatically from scanned serials.",bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",8)).pack(pady=(10,0))
        # Keyboard flow, scanner-safe: Return only activates the FOCUSED button
        # (Tab reaches them) — a stray barcode scan into a text field can no
        # longer commit the form with a serial number in the PO field. Escape
        # (never sent by scanners) cancels anywhere; focus starts on the first
        # empty required field so typing can begin immediately.
        save_b.bind("<Return>",lambda e: self._save(False))
        print_b.bind("<Return>",lambda e: self._save(True))
        cancel_b.bind("<Return>",lambda e: self._cancel())
        self.bind("<Escape>",lambda e: self._cancel())
        (self._po_e if not self._po.get().strip() else
         self._pn_e if not self._pn.get().strip() else self._po_e).focus_set()
    def _save(self, do_print=False):
        missing=[n for n,v in (("PO #",self._po),("P/N",self._pn)) if not v.get().strip()]
        if missing:
            messagebox.showwarning("Missing required fields",
                "Please fill in:\n\n  •  "+"\n  •  ".join(missing),parent=self)
            (self._po_e if not self._po.get().strip() else self._pn_e).focus_set(); return
        # QTY sanity vs scanned serials — catches the stale-draft case (COC filled
        # day 1, more serials scanned later). Warns, never blocks a real mismatch.
        q=self._qty.get().strip()
        if self._sc>0 and q!=str(self._sc):
            if not messagebox.askyesno("QTY check",
                f"QTY is '{q or '0'}' but {self._sc} serial(s) are scanned.\n\nSave anyway?",
                parent=self): return
        self.result={"ship_date":self._sd.get(),"po":self._po.get().strip(),"pn":self._pn.get().strip(),
            "rev":self._rv.get().strip(),"qty":self._qty.get(),"datecode":self._dc.get(),
            "job":self._job.get(),"type":self._ty.get(),"assembly_process":self._ap.get()}
        self.print_requested=do_print
        self.destroy()
    def _cancel(self): self.destroy()

class PackingListPopup(tk.Toplevel):
    """Packing list popup. All fields auto-fill but remain operator-editable."""
    def __init__(self,parent,wo,sc,ed=None,pi=None,coc=None):
        super().__init__(parent); self.title("Packing List"); self.geometry("620x540")
        self.resizable(False,False); self.transient(parent); self.grab_set(); self.configure(bg=C_BG)
        self.result=None; self.print_requested=False; self._wo=wo; pi=pi or {}; coc=coc or {}
        # Auto-fill priorities: existing edit > COC value > parsed_info > sensible default
        # Customer & P/N come from parsed_info, derived from the job folder path.
        self._customer = tk.StringVar(value=pi.get("customer",""))
        self._packing  = tk.StringVar()
        self._po       = tk.StringVar(value=coc.get("po","") or pi.get("po",""))
        self._pn       = tk.StringVar(value=pi.get("pn","") or coc.get("pn",""))
        self._wo_var   = tk.StringVar(value=wo)
        self._eng      = tk.StringVar(value=pi.get("eng","") or coc.get("job",""))
        self._date     = tk.StringVar(value=fmt_date())
        self._dc       = tk.StringVar(value=coc.get("datecode","") or fmt_datecode())
        if ed:
            for k,a in [('customer','_customer'),('packing_no','_packing'),('po','_po'),
                        ('pn','_pn'),('wo','_wo_var'),('eng','_eng'),
                        ('date','_date'),('datecode','_dc')]:
                v=ed.get(k,'')
                if v: getattr(self,a).set(v)
        self._build(); _center(self,parent); self.wait_window()
    def _build(self):
        h=tk.Frame(self,bg=C_TEAL); h.pack(fill="x")
        tk.Label(h,text="Packing List",bg=C_TEAL,fg=C_WHITE,font=("Segoe UI",14,"bold")).pack(side="left",padx=20,pady=14)
        tk.Label(h,text=f"WO {self._wo}",bg=C_TEAL,fg=C_WHITE,font=("Segoe UI",11,"bold")).pack(side="right",padx=20)
        # Reserve button bar at bottom first so it can never get squeezed
        btnbar=tk.Frame(self,bg="#E8E8E8",pady=18); btnbar.pack(side="bottom",fill="x")
        bf=tk.Frame(btnbar,bg="#E8E8E8"); bf.pack()
        save_b=_btn(bf,"✓  Save",lambda: self._save(False),bg=C_TEAL,fg=C_WHITE,font_size=13,padx=26,pady=14); save_b.pack(side="left",padx=6)
        print_b=_btn(bf,"🖨  Save and Print",lambda: self._save(True),bg=C_GREEN,fg=C_WHITE,font_size=13,padx=26,pady=14); print_b.pack(side="left",padx=6)
        cancel_b=_btn(bf,"Cancel",self._cancel,bg="#DDD",fg=C_MID,font_size=13,padx=26,pady=14); cancel_b.pack(side="left",padx=6)
        save_b.bind("<Return>",lambda e: self._save(False))
        print_b.bind("<Return>",lambda e: self._save(True))
        cancel_b.bind("<Return>",lambda e: self._cancel())
        bd=tk.Frame(self,bg=C_BG,padx=28,pady=16); bd.pack(fill="both",expand=True)
        tk.Label(bd,text="All fields are pre-filled but fully editable.",bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",9,"italic")).pack(anchor="w",pady=(0,8))
        def fr(p,l,v):
            r=tk.Frame(p,bg=C_BG); r.pack(fill="x",pady=4)
            tk.Label(r,text=l,bg=C_BG,fg=C_MID,font=("Segoe UI",9),width=20,anchor="e").pack(side="left",padx=(0,10))
            w=ttk.Entry(r,textvariable=v,font=("Segoe UI",11),width=36); w.pack(side="left",fill="x",expand=True,ipady=3)
            return w
        # Order matches the xlsx output: row 1 then row 2 of the info table
        fr(bd,"Customer:",self._customer)
        packing_e=fr(bd,"Packing List #:",self._packing)
        fr(bd,"P.O. #:",self._po)
        fr(bd,"P/N:",self._pn)
        fr(bd,"WO / LOT #:",self._wo_var)
        fr(bd,"PCC Engineering #:",self._eng)
        fr(bd,"Date:",self._date)
        fr(bd,"Date Code (YYWW):",self._dc)
        # Scanner-safe keyboard flow: Return only activates the focused button
        # (bound above); Escape cancels anywhere; focus starts on the one field
        # that is never prefilled (Packing List #).
        self.bind("<Escape>",lambda e: self._cancel())
        packing_e.focus_set()
    def _save(self, do_print=False):
        self.result={
            "customer":   self._customer.get().strip(),
            "packing_no": self._packing.get().strip(),
            "po":         self._po.get().strip(),
            "pn":         self._pn.get().strip(),
            "wo":         self._wo_var.get().strip(),
            "eng":        self._eng.get().strip(),
            "date":       self._date.get().strip(),
            "datecode":   self._dc.get().strip(),
        }
        self.print_requested=do_print
        self.destroy()
    def _cancel(self): self.destroy()

class SlipLabelsPopup(tk.Toplevel):
    """Shipping Slip & Shipping Labels (Dhruvil 2026-09-24): the Apex 'Shipping
    Slip' / 'Shipping Label' forms, pre-filled from the job (folder + Monday +
    COC / Packing List) and printed through the browser in Apex's exact format
    (build_slip_html / build_labels_html). Modelled on PackingListPopup; the
    Tk-free half (slip_prefill / slip_required_missing / slip_print_pages) is
    what the buttons call. Print writes the page(s) into <WO folder>\\Shipping
    Slip\\ (or the temp print folder when the share isn't reachable — the
    operator is told), opens them (the print dialog pops by itself) and stamps
    the codes / outputs / printed times; the popup STAYS open so the other page
    can be printed next, and reprints are allowed. self.result = the slip_data
    to keep on the app, None when nothing was saved or printed; after a print
    Cancel/Escape still keeps the printed record — History must match the paper."""
    FIELDS=("customer","cpn","po","wo","eng","rev","qty","date","datecode_wwyy",
            "prepared_by","notes","ship_via","label_notes","pack_qty","pack_cage")
    def __init__(self,parent,wo,sc,ed=None,pi=None,monday=None,coc=None,packing=None,
                 operator="",operators=None,wo_folder=None,box_sizes=None,online=None,division=None):
        super().__init__(parent); self.title("Shipping Labels and Slip")   # Dhruvil 2026-09-25
        # Tall enough for the whole one-box form (Options included) on a 1080
        # screen; smaller screens get the scrollbar + focus-follows-scroll.
        try: sh=int(self.winfo_screenheight())
        except Exception: sh=1080
        self.geometry(f"780x{max(600,min(980,sh-90))}")
        self.resizable(False,True); self.transient(parent); self.grab_set(); self.configure(bg=C_BG)
        self.result=None; self._wo=wo; self._sc=sc; self._wo_folder=wo_folder
        self._operators=[o for o in (operators or []) if o and not str(o).startswith("—")]
        self._sizes=_clean_box_sizes(list(box_sizes or [])) or list(SLIP_BOXSIZE_OPTIONS)   # the shared list (load_box_sizes)
        self._online_hint=online          # False = the app knows PCC2 is down: no 2-s probe per added size
        self._base=slip_prefill(pi,monday,coc,packing,operator,wo,sc,existing=ed,division=division)
        self._v={k: tk.StringVar(value=str(self._base.get(k) or "")) for k in self.FIELDS}
        self._v["customer_parts"]=tk.StringVar(value=self._base.get("customer_parts") or "No")
        for k in ("fair","si","ppap","cofc"): self._v[k]=tk.BooleanVar(value=bool(self._base.get(k)))
        self._box_rows=[]                      # [(size_var, qty_var, weight_var)] in box order
        self._box_widgets=[]                   # [(row frame, i/N label, size combobox)] parallel to _box_rows
        self._nbox=tk.IntVar(value=max(1,min(99,len(self._base.get("boxes") or []))))
        self._build(); _center(self,parent); self.wait_window()
    def _build(self):
        h=tk.Frame(self,bg=C_BLUE); h.pack(fill="x")
        tk.Label(h,text="Shipping Labels and Slip",bg=C_BLUE,fg=C_WHITE,font=("Segoe UI",14,"bold")).pack(side="left",padx=20,pady=14)
        tk.Label(h,text=f"WO {self._wo}",bg=C_BLUE,fg=C_WHITE,font=("Segoe UI",11,"bold")).pack(side="right",padx=20)
        # Reserve the button bar at the BOTTOM first so the tall form can never
        # push the buttons off the window (PackingListPopup's lesson).
        btnbar=tk.Frame(self,bg="#E8E8E8",pady=10); btnbar.pack(side="bottom",fill="x")
        self._status=tk.Label(btnbar,text=slip_printed_summary(self._base),bg="#E8E8E8",fg=C_TEAL,font=("Segoe UI",9))
        self._status.pack()
        bf=tk.Frame(btnbar,bg="#E8E8E8"); bf.pack(pady=(4,0))
        ps_b=_btn(bf,"🖨  Print Slip",lambda: self._print("slip"),bg=C_BLUE,fg=C_WHITE,font_size=11,padx=14,pady=10); ps_b.pack(side="left",padx=4)
        self._pl_btn=_btn(bf,self._labels_btn_text(),lambda: self._print("labels"),bg=C_BLUE,fg=C_WHITE,font_size=11,padx=14,pady=10); self._pl_btn.pack(side="left",padx=4)
        pb_b=_btn(bf,"🖨  Print both",lambda: self._print("both"),bg=C_GREEN,fg=C_WHITE,font_size=11,padx=14,pady=10); pb_b.pack(side="left",padx=4)
        save_b=_btn(bf,"✓  Save",self._save,bg=C_TEAL,fg=C_WHITE,font_size=11,padx=14,pady=10); save_b.pack(side="left",padx=4)
        cancel_b=_btn(bf,"Cancel",self._cancel,bg="#DDD",fg=C_MID,font_size=11,padx=14,pady=10); cancel_b.pack(side="left",padx=4)
        # Scanner-safe keyboard flow (as the COC / Packing popups): Return only
        # activates the FOCUSED button; Escape / the window X cancel anywhere.
        ps_b.bind("<Return>",lambda e: self._print("slip"))
        self._pl_btn.bind("<Return>",lambda e: self._print("labels"))
        pb_b.bind("<Return>",lambda e: self._print("both"))
        save_b.bind("<Return>",lambda e: self._save())
        cancel_b.bind("<Return>",lambda e: self._cancel())
        self.bind("<Escape>",lambda e: self._cancel())
        self.protocol("WM_DELETE_WINDOW",self._cancel)
        # Scrollable body (the box rows can make the form tall).
        wrap=tk.Frame(self,bg=C_BG); wrap.pack(fill="both",expand=True)
        cv=tk.Canvas(wrap,bg=C_BG,highlightthickness=0)
        sb=ttk.Scrollbar(wrap,orient="vertical",command=cv.yview); cv.configure(yscrollcommand=sb.set)
        sb.pack(side="right",fill="y"); cv.pack(side="left",fill="both",expand=True)
        bd=tk.Frame(cv,bg=C_BG,padx=24,pady=8); wid=cv.create_window((0,0),window=bd,anchor="nw")
        cv.bind("<Configure>",lambda e: cv.itemconfigure(wid,width=e.width))
        bd.bind("<Configure>",lambda e: cv.configure(scrollregion=cv.bbox("all")))
        self._cv=cv; self._bd=bd
        # Wheel scrolls the form — except over a dropdown / spinbox, which
        # would otherwise change the value AND scroll.
        self.bind("<MouseWheel>",lambda e: None if isinstance(e.widget,(ttk.Combobox,tk.Spinbox,ttk.Spinbox))
                  else cv.yview_scroll(int(-e.delta/120),"units"))
        # Keyboard focus below the fold (Tab from the last box field lands in the
        # Options rows) scrolls the row into view — no typing blind.
        self.bind("<FocusIn>",lambda e: self._scroll_into_view(e.widget))
        tk.Label(bd,text="Pre-filled from the job (folder · Monday · COC / Packing List) — everything is editable.  *  required to print",
                 bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",9,"italic")).pack(anchor="w",pady=(0,6))
        def sec(title):
            tk.Label(bd,text=title,bg=C_BG,fg=C_BLUE,font=("Segoe UI",10,"bold")).pack(anchor="w",pady=(8,2))
            tk.Frame(bd,bg=C_BORDER,height=1).pack(fill="x",pady=(0,4))
        def fr(l,v,required=False,cv_=None,readonly=False):
            r=tk.Frame(bd,bg=C_BG); r.pack(fill="x",pady=2)
            tk.Label(r,text=l,bg=C_BG,fg=C_MID,font=("Segoe UI",9),width=18,anchor="e").pack(side="left")
            tk.Label(r,text="*" if required else " ",bg=C_BG,fg=C_RED,font=("Segoe UI",10,"bold"),width=1).pack(side="left",padx=(0,6))
            if cv_ is not None:
                # Pick-only like Apex's <select>s (Dhruvil 2026-09-25): a click
                # ANYWHERE in the field opens the list, as the main window's
                # Operator box does. An empty list (no operator roster offline)
                # falls back to typing, so a required field can still be filled.
                vals=list(cv_)
                w=ttk.Combobox(r,textvariable=v,values=vals,font=("Segoe UI",11),width=34,
                               state="readonly" if vals else "normal")
                self._no_wheel(w)
            else:
                w=ttk.Entry(r,textvariable=v,font=("Segoe UI",11),width=36,state="readonly" if readonly else "normal")
            w.pack(side="left",fill="x",expand=True,ipady=2)
            return w
        sec("Shipping Slip")
        self._cust_e=fr("Customer:",self._v["customer"],required=True)
        fr("Customer P/N:",self._v["cpn"]); fr("Customer PO #:",self._v["po"])
        fr("Work Order / Lot:",self._v["wo"],required=True); fr("Engineering #:",self._v["eng"]); fr("Revision:",self._v["rev"])
        fr("Qty Shipped:",self._v["qty"],required=True)
        fr("Date:",self._v["date"])
        fr("Date Code (WWYY):",self._v["datecode_wwyy"],readonly=True)
        self._v["date"].trace_add("write",lambda *a: self._v["datecode_wwyy"].set(slip_datecode_wwyy(self._v["date"].get())))
        fr("Prepared By:",self._v["prepared_by"],required=True,cv_=self._operators)
        fr("Notes:",self._v["notes"])
        sec("Boxes")
        r=tk.Frame(bd,bg=C_BG); r.pack(fill="x",pady=2)
        tk.Label(r,text="Number of boxes:",bg=C_BG,fg=C_MID,font=("Segoe UI",9),width=18,anchor="e").pack(side="left",padx=(0,8))
        spn=tk.Spinbox(r,from_=1,to=99,textvariable=self._nbox,width=4,font=("Segoe UI",11),command=self._rebuild_boxes,
                       validate="key",validatecommand=(self.register(_slip_box_count_ok),"%P"))   # blank or 1-2 digits: a stray scan can't build 99 rows
        spn.pack(side="left")
        spn.bind("<Return>",lambda e: self._rebuild_boxes()); spn.bind("<FocusOut>",lambda e: self._rebuild_boxes())
        self._tot=tk.Label(r,text="",bg=C_BG,fg=C_DARK,font=("Segoe UI",10,"bold")); self._tot.pack(side="left",padx=(18,0))
        hdr=tk.Frame(bd,bg=C_BG); hdr.pack(fill="x",padx=(26,0))
        for t,w in (("Box",7),("Size",22),("Qty",8),("Weight (lbs)",12)):   # Size 22: lines up over the 20-wide dropdown
            tk.Label(hdr,text=t,bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",8,"bold"),width=w,anchor="w").pack(side="left",padx=(0,6))
        self._box_frame=tk.Frame(bd,bg=C_BG); self._box_frame.pack(fill="x",padx=(26,0))
        self._rebuild_boxes(initial=self._base.get("boxes") or [])
        sec("Options")
        r=tk.Frame(bd,bg=C_BG); r.pack(fill="x",pady=2)
        tk.Label(r,text="Customer Parts:",bg=C_BG,fg=C_MID,font=("Segoe UI",9),width=18,anchor="e").pack(side="left",padx=(0,8))
        for t in ("No","Yes"):
            tk.Radiobutton(r,text=t,value=t,variable=self._v["customer_parts"],bg=C_BG,fg=C_DARK,activebackground=C_BG,
                           font=("Segoe UI",10),cursor="hand2").pack(side="left",padx=(0,10))
        r=tk.Frame(bd,bg=C_BG); r.pack(fill="x",pady=2)
        tk.Label(r,text="Paperwork:",bg=C_BG,fg=C_MID,font=("Segoe UI",9),width=18,anchor="e").pack(side="left",padx=(0,8))
        for k,t in (("fair","FAIR"),("si","SI"),("ppap","PPAP"),("cofc","C of C")):
            tk.Checkbutton(r,text=t,variable=self._v[k],bg=C_BG,fg=C_DARK,activebackground=C_BG,
                           font=("Segoe UI",10),cursor="hand2").pack(side="left",padx=(0,10))
        fr("Ship Via:",self._v["ship_via"],cv_=[_SLIP_VIA_NONE]+SLIP_SHIPVIA_OPTIONS)
        fr("Label notes:",self._v["label_notes"])
        # Packing Label (Dhruvil 2026-10-19): Apex's 3 x 2 #lbl-pack — P/N, PO and
        # WO come from the form above; QTY starts at 1 as in Apex; CAGE and the
        # date code print on ITAR/MAG jobs, today's date on Industrial ones.
        sec("Packing Label  (3 × 2 in · one label per print · 3 × 2 Shipping printer)")
        itar=str(self._base.get("pack_mode") or "itar")!="industrial"
        r=tk.Frame(bd,bg=C_BG); r.pack(fill="x",pady=2)
        tk.Label(r,text="Label type:",bg=C_BG,fg=C_MID,font=("Segoe UI",9),width=18,anchor="e").pack(side="left",padx=(0,8))
        tk.Label(r,text=("ITAR / MAG  —  prints CAGE and D/C" if itar else "Industrial  —  prints today's DATE")+"  (from the job's division)",
                 bg=C_BG,fg=C_DARK,font=("Segoe UI",10,"bold")).pack(side="left")
        fr("Packing QTY:",self._v["pack_qty"])
        if itar: fr("CAGE:",self._v["pack_cage"])
        r=tk.Frame(bd,bg=C_BG); r.pack(fill="x",pady=(6,2))
        tk.Label(r,text="",bg=C_BG,width=18).pack(side="left",padx=(0,8))
        pk_b=_btn(r,"🖨  Print Packing Label",lambda: self._print("pack"),bg=C_BLUE,fg=C_WHITE,font_size=11,padx=14,pady=8); pk_b.pack(side="left")
        pk_b.bind("<Return>",lambda e: self._print("pack"))
        # Initial focus: the Customer entry when it is blank, else Box 1's Size —
        # the first field left to fill. Never the spinbox: a barcode scanned
        # there ('16452' + Return) built 99 rows (review round 2).
        try: first=self._cust_e if not self._v["customer"].get().strip() else self._box_widgets[0][2]
        except Exception: first=self._cust_e
        first.focus_set()
    def _no_wheel(self,w):
        """The mouse wheel over a readonly dropdown re-picks its value (ttk's
        own binding) — over this scrolling form that silently changed a Size or
        Ship Via, and on the Size list could land on '+ Add new size…'. Over a
        dropdown the wheel scrolls the form instead and never changes it."""
        w.bind("<MouseWheel>",lambda e: (self._cv.yview_scroll(int(-e.delta/120),"units"),"break")[1])
    def _size_values(self):
        return [_SLIP_SIZE_NONE]+list(self._sizes)+[_SLIP_ADD_SIZE]
    def _size_picked(self,sv,sz):
        """A Size picked. '+ Add new size…' is never kept as the value: the row
        goes back to what it showed and the add dialog opens (after the list has
        closed and released its grab)."""
        v=sv.get()
        if v!=_SLIP_ADD_SIZE:
            sz._last=v; return
        sv.set(getattr(sz,"_last","") or "")
        self.after_idle(lambda: self._add_size(sv,sz))
    def _add_size(self,sv,sz):
        """'+ Add new size…' (Dhruvil 2026-09-25: "we keep getting new sizes
        every time"): type the size once — it goes into the list SHARED by every
        station (add_box_size) and into this row. Share unreachable: the size is
        used for this shipment only, and the operator is told."""
        from tkinter import simpledialog
        try:
            raw=simpledialog.askstring("New box size","Box size to add to the list  (for example 18x12x6):",parent=self)
            self._regrab()                    # simpledialog releases our modal grab
            size=normalize_box_size(raw)
            if not size: return
            bad=box_size_problem(size)
            if bad:
                messagebox.showwarning("Box size",bad,parent=self); self._regrab(); return
            have=next((x for x in self._sizes if x.lower()==size.lower()),None)
            saved,err=True,None
            if have is None:
                try:
                    self._sizes,saved,err=add_box_size(size,online=self._online_hint)
                except Exception as e:
                    saved,err=False,f"{type(e).__name__}: {e}"
                    self._sizes=_box_sizes_with(self._sizes,size)
                have=next((x for x in self._sizes if x.lower()==size.lower()),size)
            vals=self._size_values()
            for _f,_l,cb in self._box_widgets:
                try: cb.config(values=vals,height=min(25,len(vals)))
                except Exception: pass
            sv.set(have); sz._last=have
            if not saved:
                messagebox.showwarning("Box size not saved",
                    f"'{have}' is used for this shipment, but it could not be added to the shared list:\n\n{err}\n\n"
                    "Add it again once PCC2 is reachable, or ask an admin (Admin ▸ Box sizes).",parent=self)
                self._regrab()
        finally:
            # simpledialog hands the keyboard focus to the popup ITSELF: the next
            # Tab then landed on 'Print Slip' and Return printed (review round 2).
            try: sz.focus_set()
            except Exception: pass
    def _labels_btn_text(self):
        n=len(self._box_rows) or int(self._nbox.get() or 1)
        return f"🖨  Print Labels ({n} box{'es' if n!=1 else ''})"
    def _add_box_row(self,size,qty,wt):
        """Append one box row (i/N label · Size combobox · Qty · Weight) to the
        box frame; the label is renumbered by _rebuild_boxes."""
        sv=tk.StringVar(value=size); qv=tk.StringVar(value=qty); wv=tk.StringVar(value=wt)
        r=tk.Frame(self._box_frame,bg=C_BG); r.pack(fill="x",pady=1)
        l=tk.Label(r,text="",bg=C_BG,fg=C_DARK,font=("Segoe UI",10,"bold"),width=7,anchor="w"); l.pack(side="left",padx=(0,6))
        # 20 wide + as many rows as entries (up to 25): '+ Add new size…' is never
        # cut off (the list font is Segoe UI 14) nor hidden below the fold of the
        # default 10-row list (review round 2).
        vals=self._size_values()
        sz=ttk.Combobox(r,textvariable=sv,values=vals,font=("Segoe UI",10),width=20,height=min(25,len(vals)),state="readonly")
        sz.pack(side="left",padx=(0,6)); sz._last=size
        sz.bind("<<ComboboxSelected>>",lambda e,sv=sv,sz=sz: self._size_picked(sv,sz))
        self._no_wheel(sz)
        ttk.Entry(r,textvariable=qv,font=("Segoe UI",10),width=8).pack(side="left",padx=(0,6))
        ttk.Entry(r,textvariable=wv,font=("Segoe UI",10),width=12).pack(side="left")
        qv.trace_add("write",lambda *a: self._update_totals()); wv.trace_add("write",lambda *a: self._update_totals())
        self._box_rows.append((sv,qv,wv)); self._box_widgets.append((r,l,sz))
    def _rebuild_boxes(self,initial=None):
        """Bring the box rows in line with the spinbox (1..99) — INCREMENTALLY:
        rows that stay are not touched (what was typed stays typed, and the
        widget the keyboard focus is moving into survives: Tk sets the new focus
        BEFORE it delivers the spinbox's FocusOut, so tearing every row down here
        killed the very entry that was just clicked and stranded the keyboard on
        the popup — review rounds 1 and 2); rows past the count go, missing rows
        are appended blank (Apex's box table starts blank), the i/N labels are
        renumbered. `initial` (the saved boxes) seeds the rows on open. Live
        totals follow every keystroke."""
        try: n=int(str(self._nbox.get()).strip() or 1)
        except Exception: n=1
        n=max(1,min(99,n))
        try: self._nbox.set(n)
        except Exception: pass
        if initial is not None:
            for fr_,_l,_s in self._box_widgets: fr_.destroy()
            self._box_rows=[]; self._box_widgets=[]
            for b in initial[:n]:
                b=b or {}
                self._add_box_row(str(b.get("size") or ""),str(b.get("qty") or ""),str(b.get("weight") or ""))
        try: focus=self.focus_get()
        except Exception: focus=None
        lost=False
        while len(self._box_rows)>n:
            self._box_rows.pop(); fr_,_l,_s=self._box_widgets.pop()
            lost=lost or (focus is not None and str(focus).startswith(str(fr_)+"."))
            fr_.destroy()
        while len(self._box_rows)<n:
            self._add_box_row("","","")
        for i,(_f,l,_s) in enumerate(self._box_widgets,1):
            l.config(text=f"{i}/{n}")
        if lost:
            # the focused row was removed: land on the last row's Size, not on the window
            try: self.after_idle(self._box_widgets[-1][2].focus_set)
            except Exception: pass
        self._update_totals()
        try: self._pl_btn.config(text=self._labels_btn_text())
        except Exception: pass
    def _scroll_into_view(self,w):
        """<FocusIn> on any widget inside the scrollable body: move the canvas
        so the focused row is fully visible (_slip_scroll_fraction)."""
        try:
            cv=self._cv; bd=self._bd
            if not str(w).startswith(str(bd)+"."): return
            row=w
            while row.master is not bd: row=row.master     # the field's row frame
            f=_slip_scroll_fraction(row.winfo_y(),row.winfo_height(),cv.canvasy(0),cv.winfo_height(),bd.winfo_height())
            if f is not None: cv.yview_moveto(f)
        except Exception:
            pass
    def _update_totals(self):
        q,w=calc_box_totals([{"qty":qv.get(),"weight":wv.get()} for _s,qv,wv in self._box_rows])
        try: self._tot.config(text=f"Totals:  qty {q}   ·   weight {w}")
        except Exception: pass
    def _collect(self):
        """Read the form into a fresh slip_data. The codes / printed / outputs
        come from the base record — they change only when a page prints. Uses
        nothing but .get() on the vars so the test suite drives it on a stand-in."""
        d=slip_data_template(); base=self._base or {}
        for k in SlipLabelsPopup.FIELDS:
            d[k]=str(self._v[k].get() or "").strip()
        if d["ship_via"]==_SLIP_VIA_NONE: d["ship_via"]=""            # Apex's blank option
        d["date"]=_parse_user_date(d["date"]) or d["date"]          # store YYYY-MM-DD when parseable
        d["datecode_wwyy"]=d["datecode_wwyy"] or slip_datecode_wwyy(d["date"])
        d["customer_parts"]="Yes" if str(self._v["customer_parts"].get() or "").strip().lower()=="yes" else "No"
        for k in ("fair","si","ppap","cofc"): d[k]=bool(self._v[k].get())
        d["boxes"]=[{"size":_slip_size_value(s.get()),"qty":str(q.get() or "").strip(),"weight":str(w.get() or "").strip()}
                    for s,q,w in self._box_rows]
        d["slip_code"]=str(base.get("slip_code") or ""); d["labels_code"]=str(base.get("labels_code") or "")
        for k in ("pack_mode","pack_code","pack_date","pack_datecode"):
            d[k]=str(base.get(k) or d.get(k) or "")
        d["printed"].update({k:v for k,v in (base.get("printed") or {}).items() if k in d["printed"]})
        d["outputs"].update({k:v for k,v in (base.get("outputs") or {}).items() if k in d["outputs"]})
        d["pages"].update({k:v for k,v in (base.get("pages") or {}).items() if k in d["pages"]})
        return d
    def _regrab(self):
        # Tk grabs don't stack: re-assert ours after a child dialog closed.
        try: self.grab_set()
        except Exception: pass
    def _print(self,which):
        self._rebuild_boxes()   # a count TYPED into the spinbox: a tk.Button click takes no focus, so no FocusOut applied it
        d=self._collect()
        missing=slip_required_missing(d,which)
        if missing:
            messagebox.showwarning("Missing required fields",
                "Please fill in:\n\n  •  "+_slip_missing_text(missing),parent=self)
            self._regrab(); return
        # QTY sanity vs scanned serials (same warning as the COC — warns, never blocks).
        q=d["qty"]
        if which!="pack" and self._sc>0 and q!=str(self._sc):     # the packing label's QTY is its own (Apex: starts at 1)
            ok=messagebox.askyesno("QTY check",
                f"Qty Shipped is '{q or '0'}' but {self._sc} serial(s) are scanned.\n\nPrint anyway?",parent=self)
            self._regrab()
            if not ok: return
        try:
            res=slip_print_pages(self._wo_folder,d,which)
        except Exception as e:
            messagebox.showerror("Shipping Labels and Slip",
                f"Could not create the page(s):\n{type(e).__name__}: {e}",parent=self)
            self._regrab(); return
        # From here the record reflects paper: keep it even if Cancel follows.
        self._base=d; self.result=d
        try: self._status.config(text=slip_printed_summary(d))
        except Exception: pass
        if res.get("errors"):
            # a LATER page of this print could be written nowhere (share refused
            # it, temp folder failed too); the page already written stays recorded
            messagebox.showerror("Shipping Labels and Slip",
                "Could not create:\n\n"+"\n".join(f"  •  {lbl}: {err}" for lbl,err in res["errors"])
                +"\n\nThe other page was written and opened as usual — print the missing one again.",parent=self)
            self._regrab()
        if res.get("open_errors"):
            messagebox.showerror("Print",
                "The page(s) were written but could not be opened for printing:\n\n"
                +"\n".join(f"  •  {lbl}: {err}" for lbl,err in res["open_errors"])
                +f"\n\nOpen them by hand from:\n{res.get('folder')}",parent=self)
            self._regrab()
        if res.get("fallback"):
            lost=[lbl for key,lbl in (("slip_html","Shipping Slip"),("labels_html","Shipping Labels"),("pack_html","Packing Label"))
                  if res.get(key) and not (d.get("outputs") or {}).get(key)]
            messagebox.showwarning("Printed — NOT saved to the share",
                f"{' + '.join(lost) or 'The page(s)'} opened for printing, but the WO folder isn't "
                "reachable (or couldn't be written) from here, so the copy was NOT saved to the share.\n\n"
                f"It was saved locally instead in:\n  {res.get('folder')}\n\n"
                "The slip data itself still goes into the shipment's History record — once the "
                "share is back, History ▸ Re-emit paperwork files the page into the WO folder.",parent=self)
            self._regrab()
    def _save(self):
        self._rebuild_boxes()   # as in _print: apply a typed box count
        self.result=self._collect(); self.destroy()
    def _cancel(self): self.destroy()

class SearchingDialog(tk.Toplevel):
    def __init__(self,parent,wo,div):
        super().__init__(parent); self.overrideredirect(True); self.configure(bg=C_GREEN)
        f=tk.Frame(self,bg=C_GREEN,padx=36,pady=24); f.pack()
        tk.Label(f,text="Searching server...",bg=C_GREEN,fg=C_WHITE,font=("Segoe UI",13,"bold")).pack()
        tk.Label(f,text=f"{div}  ·  WO {wo}",bg=C_GREEN,fg="#9FE1CB",font=("Segoe UI",10)).pack(pady=(4,0))
        self.grab_set(); _center(self,parent); self.update()

class WOPickerDialog(tk.Toplevel):
    def __init__(self,parent,matches):
        super().__init__(parent); self.title("Multiple matches"); self.resizable(False,False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG); self.result=None; self._m=matches
        h=tk.Frame(self,bg=C_AMBER); h.pack(fill="x")
        tk.Label(h,text="Multiple Work Orders Found — Select One",bg=C_AMBER,fg=C_WHITE,font=("Segoe UI",12,"bold")).pack(padx=24,pady=14)
        bd=tk.Frame(self,bg=C_BG,padx=24,pady=12); bd.pack(fill="x")
        for i,m in enumerate(matches):
            f=tk.Frame(bd,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1,cursor="hand2"); f.pack(fill="x",pady=4)
            f.bind("<Button-1>",lambda e,x=i: self._pick(x))
            inn=tk.Frame(f,bg=C_WHITE,padx=16,pady=12); inn.pack(fill="x"); inn.bind("<Button-1>",lambda e,x=i: self._pick(x))
            nl=tk.Label(inn,text=m.name,bg=C_WHITE,fg=C_GREEN,font=("Segoe UI",12,"bold"),cursor="hand2"); nl.pack(anchor="w"); nl.bind("<Button-1>",lambda e,x=i: self._pick(x))
            pl=tk.Label(inn,text=str(m.parent),bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",8),cursor="hand2",wraplength=500,justify="left"); pl.pack(anchor="w"); pl.bind("<Button-1>",lambda e,x=i: self._pick(x))
        ft=tk.Frame(self,bg=C_BG,pady=10); ft.pack(fill="x"); _btn(ft,"Cancel",self._cancel,bg="#DDD",fg=C_MID,padx=20,pady=8).pack()
        _center(self,parent); self.wait_window()
    def _pick(self,i): self.result=self._m[i]; self.destroy()
    def _cancel(self): self.destroy()

class DraftPickerDialog(tk.Toplevel):
    """Shown when a loading WO already has saved draft(s): the operator chooses to
    RESUME one (e.g. a supervisor closing a specific batch) or START A NEW shipment
    alongside them (a second batch no longer waits for the first to close). Rows
    show date · saver · counts so drafts are easy to tell apart. Buttons arm after
    400ms (scanner-proof). self.result: ('resume', rec) | ('new', None) | None."""
    def __init__(self, parent, wo, drafts):
        super().__init__(parent); self.title("Saved drafts found"); self.resizable(False, False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG)
        self.result = None; self._drafts = list(drafts)
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        h = tk.Frame(self, bg=C_TEAL); h.pack(fill="x")
        tk.Label(h, text=f"📋  WO {wo} has {len(drafts)} saved draft(s)", bg=C_TEAL, fg=C_WHITE,
                 font=("Segoe UI", 12, "bold")).pack(padx=22, pady=12, anchor="w")
        bd = tk.Frame(self, bg=C_BG, padx=22, pady=12); bd.pack(fill="both", expand=True)
        tk.Label(bd, text=("Resume a draft to continue/close it — or start a NEW shipment; "
                           "the saved drafts stay untouched for later."),
                 bg=C_BG, fg=C_MID, font=("Segoe UI", 9), wraplength=470, justify="left").pack(anchor="w", pady=(0, 8))
        lf = tk.Frame(bd, bg=C_WHITE, highlightbackground=C_BORDER, highlightthickness=1)
        lf.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(lf); sb.pack(side="right", fill="y")
        self._lb = tk.Listbox(lf, font=("Segoe UI", 11), height=min(8, max(3, len(drafts))),
                              bg=C_WHITE, fg=C_DARK, selectbackground=C_TEAL_LITE,
                              selectforeground=C_TEAL, relief="flat", bd=0, activestyle="none",
                              yscrollcommand=sb.set, width=58)
        self._lb.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        sb.config(command=self._lb.yview)
        for d in self._drafts:
            self._lb.insert("end",
                f"  {d.get('saved_at') or '(no date)'}   ·   by {d.get('saved_by') or '?'}   ·   "
                f"{d.get('serial_count', 0)} serial(s), {d.get('photo_count', 0)} photo(s)")
        self._lb.selection_set(0)
        self._lb.bind("<Double-1>", lambda e: self._resume())
        br = tk.Frame(self, bg="#E8E8E8", pady=12); br.pack(fill="x")
        self._res_btn = _btn(br, "▶  Resume selected draft", self._resume, bg=C_TEAL, fg=C_WHITE,
                             font_size=10, padx=16, pady=9)
        self._res_btn.pack(side="left", padx=(22, 8))
        self._new_btn = _btn(br, "✚  Start NEW shipment", self._new, bg=C_GREEN, fg=C_WHITE,
                             font_size=10, padx=16, pady=9)
        self._new_btn.pack(side="left")
        _btn(br, "Cancel", self._cancel, bg="#DDD", fg=C_MID, font_size=10,
             padx=16, pady=9).pack(side="right", padx=(0, 22))
        # Scanner-proof: a trailing Enter from a barcode scan can't pick for you.
        self._res_btn.config(state="disabled"); self._new_btn.config(state="disabled")
        self._armed = False
        self.after(400, self._arm)
        self.bind("<Escape>", lambda e: self._cancel())
        _center(self, parent); self._lb.focus_set(); self.wait_window()
    def _arm(self):
        self._armed = True
        for b in (self._res_btn, self._new_btn):
            try: b.config(state="normal")
            except Exception: pass
    def _resume(self):
        if not self._armed: return
        sel = self._lb.curselection()
        i = sel[0] if sel else 0
        self.result = ("resume", self._drafts[i]); self.destroy()
    def _new(self):
        if not self._armed: return
        self.result = ("new", None); self.destroy()
    def _cancel(self):
        self.result = None; self.destroy()

class ConfirmDialog(tk.Toplevel):
    def __init__(self,parent,wo,wf,div,monday=None):
        super().__init__(parent); self.title("Confirm job"); self.resizable(False,False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG); self.result=False
        info=parse_wo_path(wf); root=Path(PRODUCTION_ROOTS[div])
        try: rel=wf.relative_to(root)
        except ValueError: rel=wf
        h=tk.Frame(self,bg=C_GREEN); h.pack(fill="x")
        tk.Label(h,text="Job Found — Please Confirm",bg=C_GREEN,fg=C_WHITE,font=("Segoe UI",13,"bold")).pack(padx=24,pady=14)
        bd=tk.Frame(self,bg=C_BG,padx=24,pady=16); bd.pack(fill="x")
        cd=tk.Frame(bd,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); cd.pack(fill="x")
        def row(l,v,fg=C_DARK):
            r=tk.Frame(cd,bg=C_WHITE); r.pack(fill="x"); tk.Frame(cd,bg="#EBEBEB",height=1).pack(fill="x")
            tk.Label(r,text=l,bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",9),width=15,anchor="e",padx=10,pady=8).pack(side="left")
            tk.Frame(r,bg=C_BORDER,width=1).pack(side="left",fill="y",pady=3)
            tk.Label(r,text=v,bg=C_WHITE,fg=fg,font=("Segoe UI",10),anchor="w",padx=12,wraplength=420,justify="left").pack(side="left",fill="x",expand=True)
        row("Division",div,fg=C_GREEN if div=="Industrial" else C_AMBER)
        row("Workorder #",wf.name,fg=C_GREEN); row("Customer",info["customer"])
        row("P/N",info["pn"]); row("Engineering #",info["eng"])
        if monday and monday.get("po"):
            row("Customer PO",monday["po"],fg=C_BLUE)   # Monday-only field, shown for verification
        row("Full path",str(rel),fg=C_MID)
        if monday:
            tk.Label(bd,text="✓  Matched to a Monday work order — Customer PO / details will be pre-filled.",
                     bg=C_BG,fg=C_BLUE,font=("Segoe UI",8,"italic")).pack(anchor="w",pady=(8,0))
        bt=tk.Frame(self,bg=C_BG,pady=16); bt.pack()
        yb=_btn(bt,"✓   Yes",self._y,bg=C_GREEN,fg=C_WHITE,padx=20,pady=10); yb.pack(side="left",padx=8)
        _btn(bt,"✗   No",self._n,bg="#DDD",fg=C_MID,padx=20,pady=10).pack(side="left",padx=8)
        self.bind("<Return>",lambda e: self._y()); self.bind("<Escape>",lambda e: self._n()); yb.focus_set()
        _center(self,parent); self.wait_window()
    def _y(self): self.result=True; self.destroy()
    def _n(self): self.result=False; self.destroy()

class _MondaySearchingDialog(tk.Toplevel):
    """Tiny modal shown while the Monday lookup runs on a worker thread — blue,
    to distinguish it from the green 'Searching server...' share walk."""
    def __init__(self,parent,term):
        super().__init__(parent); self.overrideredirect(True); self.configure(bg=C_BLUE)
        f=tk.Frame(self,bg=C_BLUE,padx=36,pady=24); f.pack()
        tk.Label(f,text="Searching Monday.com…",bg=C_BLUE,fg=C_WHITE,font=("Segoe UI",13,"bold")).pack()
        tk.Label(f,text=f"Looking up {term}",bg=C_BLUE,fg="#CFE4FF",font=("Segoe UI",10)).pack(pady=(4,0))
        self.grab_set(); _center(self,parent); self.update()

class MondayPickerDialog(tk.Toplevel):
    """Shown when a WO/Eng# matches more than one Monday item. Picking one uses it
    to pre-fill; 'Skip' proceeds to the normal folder search with no autofill."""
    def __init__(self,parent,hits):
        super().__init__(parent); self.title("Multiple Monday matches"); self.resizable(False,False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG); self.result=None; self._h=hits
        h=tk.Frame(self,bg=C_BLUE); h.pack(fill="x")
        tk.Label(h,text="Multiple Monday Matches — Select One",bg=C_BLUE,fg=C_WHITE,font=("Segoe UI",12,"bold")).pack(padx=24,pady=14)
        bd=tk.Frame(self,bg=C_BG,padx=24,pady=12); bd.pack(fill="x")
        for i,m in enumerate(hits):
            f=tk.Frame(bd,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1,cursor="hand2"); f.pack(fill="x",pady=4)
            f.bind("<Button-1>",lambda e,x=i: self._pick(x))
            inn=tk.Frame(f,bg=C_WHITE,padx=16,pady=12); inn.pack(fill="x"); inn.bind("<Button-1>",lambda e,x=i: self._pick(x))
            title=f"WO {m.get('wo') or '?'}   ·   Eng {m.get('eng') or '?'}"
            bits=[m.get('customer',''),(f"PO {m['po']}" if m.get('po') else ''),
                  (f"Rev {m['rev']}" if m.get('rev') else ''),(f"[{m['group']}]" if m.get('group') else '')]
            sub=" · ".join([b for b in bits if b])
            nl=tk.Label(inn,text=title,bg=C_WHITE,fg=C_BLUE,font=("Segoe UI",12,"bold"),cursor="hand2"); nl.pack(anchor="w"); nl.bind("<Button-1>",lambda e,x=i: self._pick(x))
            pl=tk.Label(inn,text=sub,bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",9),cursor="hand2",wraplength=520,justify="left"); pl.pack(anchor="w"); pl.bind("<Button-1>",lambda e,x=i: self._pick(x))
        ft=tk.Frame(self,bg=C_BG,pady=10); ft.pack(fill="x")
        _btn(ft,"Skip Monday — search folders",self._cancel,bg="#DDD",fg=C_MID,padx=20,pady=8).pack()
        _center(self,parent); self.wait_window()
    def _pick(self,i): self.result=self._h[i]; self.destroy()
    def _cancel(self): self.destroy()

class ExportDialog(tk.Toplevel):
    def __init__(self,parent,wo,count):
        super().__init__(parent); self.title("Export"); self.resizable(False,False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG); self.result=None
        h=tk.Frame(self,bg=C_GREEN); h.pack(fill="x")
        tk.Label(h,text="Save serial numbers as:",bg=C_GREEN,fg=C_WHITE,font=("Segoe UI",12,"bold")).pack(padx=24,pady=12)
        tk.Label(h,text=f"{count} serial{'s' if count!=1 else ''} · WO {wo}",bg=C_GREEN,fg="#9FE1CB",font=("Segoe UI",10)).pack(pady=(0,10))
        bd=tk.Frame(self,bg=C_BG,padx=24,pady=16); bd.pack(fill="x")
        def oc(bg,t,s,cmd):
            f=tk.Frame(bd,bg=bg,highlightbackground=C_BORDER,highlightthickness=1,cursor="hand2"); f.pack(fill="x",pady=6)
            f.bind("<Button-1>",lambda e: cmd())
            i=tk.Frame(f,bg=bg,padx=18,pady=14); i.pack(fill="x"); i.bind("<Button-1>",lambda e: cmd())
            tl=tk.Label(i,text=t,bg=bg,fg=C_WHITE,font=("Segoe UI",12,"bold"),anchor="w"); tl.pack(fill="x"); tl.bind("<Button-1>",lambda e: cmd())
            sl=tk.Label(i,text=s,bg=bg,fg="#C8E8D8" if bg==C_GREEN else "#B5D4F4",font=("Segoe UI",9),anchor="w"); sl.pack(fill="x"); sl.bind("<Button-1>",lambda e: cmd())
        oc(C_GREEN,"Excel only","WO serials.xlsx",self._x); oc(C_BLUE,"Excel + Word","Excel + Word document",self._b)
        # Explicit Cancel that goes back without closing the shipment. Escape and
        # the window-X do the same — they must NOT silently proceed.
        cf=tk.Frame(bd,bg=C_BG); cf.pack(fill="x",pady=(8,0))
        _btn(cf,"Cancel — don't close yet",self._cancel,bg="#DDD",fg=C_MID,font_size=10,padx=18,pady=8).pack()
        self.bind("<Escape>",lambda e: self._cancel())
        self.protocol("WM_DELETE_WINDOW",self._cancel)
        _center(self,parent); self.wait_window()
    def _x(self): self.result="xlsx"; self.destroy()
    def _b(self): self.result="both"; self.destroy()
    def _cancel(self): self.result=None; self.destroy()

class AdminLoginDialog(tk.Toplevel):
    def __init__(self,parent,admins):
        super().__init__(parent); self.title("Admin login"); self.resizable(False,False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG); self.result=False; self._adm=admins; self._default_pw=False; self.admin_id=""
        h=tk.Frame(self,bg=C_DARK); h.pack(fill="x")
        tk.Label(h,text="Administrator Login",bg=C_DARK,fg=C_WHITE,font=("Segoe UI",13,"bold")).pack(padx=24,pady=14)
        bd=tk.Frame(self,bg=C_BG,padx=36,pady=24); bd.pack()
        tk.Label(bd,text="Admin ID",bg=C_BG,fg=C_MID,font=("Segoe UI",9)).pack(anchor="w")
        self._id=tk.StringVar(); ie=ttk.Entry(bd,textvariable=self._id,font=("Segoe UI",12),width=28); ie.pack(fill="x",pady=(2,14),ipady=5)
        tk.Label(bd,text="Password",bg=C_BG,fg=C_MID,font=("Segoe UI",9)).pack(anchor="w")
        self._pw=tk.StringVar(); pe=ttk.Entry(bd,textvariable=self._pw,font=("Segoe UI",12),width=28,show="●"); pe.pack(fill="x",pady=(2,6),ipady=5)
        pe.bind("<Return>",lambda e: self._login())
        sv=tk.BooleanVar(value=False)
        tk.Checkbutton(bd,text="Show password",variable=sv,bg=C_BG,fg=C_SUBTLE,activebackground=C_BG,
                       font=("Segoe UI",8),command=lambda: pe.config(show="" if sv.get() else "●")
                       ).pack(anchor="w",pady=(0,12))
        # Keyboard flow: Enter on the ID field moves to the password; Escape closes.
        ie.bind("<Return>",lambda e: pe.focus_set())
        self.bind("<Escape>",lambda e: self.destroy())
        self._err=tk.Label(bd,text="",bg=C_BG,fg=C_RED,font=("Segoe UI",9)); self._err.pack()
        _btn(bd,"Login →",self._login,bg=C_DARK,fg=C_WHITE,padx=32,pady=10).pack(pady=(10,0))
        _center(self,parent); ie.focus_set(); self.wait_window()
    def _login(self):
        ok,_upg,is_default=authenticate_admin(self._id.get().strip(),self._pw.get(),self._adm)
        if ok: self._default_pw=is_default; self.admin_id=self._id.get().strip(); self.result=True; self.destroy()
        else: self._err.config(text="Incorrect ID or password.")

class AdminPanel(tk.Toplevel):
    def __init__(self,parent,admins,operators,on_save):
        super().__init__(parent); self.title("Admin Panel"); self.geometry("680x580")
        self.resizable(False,False); self.transient(parent); self.grab_set(); self.configure(bg=C_BG)
        self._a=[a.copy() for a in admins]; self._o=list(operators); self._os=on_save
        h=tk.Frame(self,bg=C_DARK); h.pack(fill="x")
        tk.Label(h,text="Admin Panel — Printed Circuits Corp",bg=C_DARK,fg=C_WHITE,font=("Segoe UI",13,"bold")).pack(side="left",padx=20,pady=13)
        tb=tk.Frame(self,bg="#DCDCDC"); tb.pack(fill="x"); self._tb={}
        for k,l in [("op","  Operators  "),("ad","  Administrators  "),("bx","  Box sizes  "),("pu","  Portal upload  ")]:
            b=tk.Button(tb,text=l,relief="flat",cursor="hand2",font=("Segoe UI",10,"bold"),pady=10,command=lambda x=k: self._st(x))
            b.pack(side="left"); self._tb[k]=b
        tk.Label(self,text="Changes are saved immediately",bg=C_BG,fg=C_SUBTLE,
                 font=("Segoe UI",8)).pack(anchor="e",padx=24,pady=(4,0))
        self._bd=tk.Frame(self,bg=C_BG); self._bd.pack(fill="both",expand=True,padx=24,pady=(6,16))
        self._st("op"); _center(self,parent); self.wait_window()
    def _st(self,k):
        for x,b in self._tb.items(): b.config(bg=C_GREEN if x==k else "#DCDCDC",fg=C_WHITE if x==k else C_MID)
        for w in self._bd.winfo_children(): w.destroy()
        if k=="op": self._top()
        elif k=="bx": self._tbx()
        elif k=="pu": self._tpu()
        else: self._tad()
    def _tpu(self):
        """Customer Portal upload settings for THIS PC (saved locally, never on the share)."""
        f=self._bd; c=portal_upload.load_config()
        tk.Label(f,text="Send CoC, Packing List and Serial list PDFs to the Customer Portal",bg=C_BG,fg=C_GREEN,font=("Segoe UI",11,"bold")).pack(anchor="w",pady=(0,4))
        tk.Label(f,text="Industrial and MAG - ITAR jobs. Anything marked export controlled / USML / CUI is never sent.\n"
                 "The upload key comes from the portal admin (Supabase secret SHIPPING_UPLOAD_KEY). It's saved only on this PC.",
                 bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",9),justify="left").pack(anchor="w",pady=(0,10))
        self._pu_en=tk.BooleanVar(value=c["enabled"])
        self._pu_key=tk.StringVar(value=c["key"]); self._pu_url=tk.StringVar(value=c["url"])
        tk.Checkbutton(f,text="Turn on portal upload for this PC",variable=self._pu_en,bg=C_BG,font=("Segoe UI",10)).pack(anchor="w",pady=(0,6))
        g=tk.Frame(f,bg=C_BG); g.pack(fill="x")
        for r,(l,v,kw) in enumerate([("Upload key",self._pu_key,{"show":"●"}),("Portal address",self._pu_url,{})]):
            tk.Label(g,text=l,bg=C_BG,fg=C_MID,font=("Segoe UI",9),width=14,anchor="e").grid(row=r,column=0,padx=(0,8),pady=4)
            ttk.Entry(g,textvariable=v,font=("Segoe UI",10),width=52,**kw).grid(row=r,column=1,sticky="w",ipady=3)
        self._pu_msg=tk.Label(f,text="",bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",9),wraplength=600,justify="left"); self._pu_msg.pack(anchor="w",pady=(10,4))
        br=tk.Frame(f,bg=C_BG); br.pack(fill="x",pady=4)
        _btn(br,"Test connection",self._pu_test,bg=C_DARK,fg=C_WHITE).pack(side="left")
        _btn(br,"Save",self._pu_save,bg=C_GREEN,fg=C_WHITE).pack(side="right")
    def _pu_values(self):
        return {"enabled":self._pu_en.get(),"key":self._pu_key.get().strip(),
                "url":self._pu_url.get().strip() or portal_upload.DEFAULT_URL}
    def _pu_test(self):
        c=self._pu_values()
        if len(c["key"])<24:
            self._pu_msg.config(text="Enter the upload key (at least 24 characters).",fg=C_RED); return
        self._pu_msg.config(text="Checking…",fg=C_SUBTLE); self.update()
        try: ok,msg=portal_upload.ping(c)
        except Exception as e: ok,msg=False,f"{type(e).__name__}: {e}"
        self._pu_msg.config(text=("✓ " if ok else "✗ ")+msg,fg=C_GREEN if ok else C_RED)
    def _pu_save(self):
        c=self._pu_values()
        if c["enabled"] and not portal_upload.config_ready(c):
            self._pu_msg.config(text="Can't turn it on yet: an upload key of at least 24 characters and an https address are needed.",fg=C_RED); return
        try: portal_upload.save_config(c)
        except Exception as e: self._pu_msg.config(text=f"Couldn't save: {e}",fg=C_RED); return
        self._pu_msg.config(text="Saved. "+("Portal upload is ON for this PC." if c["enabled"] else "Portal upload is OFF for this PC."),fg=C_GREEN)
    def _top(self):
        f=self._bd; tk.Label(f,text="Operator roster",bg=C_BG,fg=C_GREEN,font=("Segoe UI",11,"bold")).pack(anchor="w",pady=(0,8))
        lf=tk.Frame(f,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); lf.pack(fill="both",expand=True)
        sb=ttk.Scrollbar(lf); sb.pack(side="right",fill="y")
        self._olb=tk.Listbox(lf,font=("Segoe UI",11),bg=C_WHITE,fg=C_DARK,selectbackground=C_GREEN_LITE,selectforeground=C_GREEN,relief="flat",bd=0,activestyle="none",yscrollcommand=sb.set)
        self._olb.pack(fill="both",expand=True,padx=6,pady=6); sb.config(command=self._olb.yview)
        for o in self._o: self._olb.insert("end",f"   {o}")
        ar=tk.Frame(f,bg=C_BG); ar.pack(fill="x",pady=(10,4))
        tk.Label(ar,text="New name:",bg=C_BG,fg=C_MID,font=("Segoe UI",9)).pack(side="left")
        self._nv=tk.StringVar(); e=ttk.Entry(ar,textvariable=self._nv,font=("Segoe UI",11),width=26); e.pack(side="left",padx=(8,8),ipady=4)
        e.bind("<Return>",lambda ev: self._ao()); _btn(ar,"+ Add",self._ao,bg=C_GREEN,fg=C_WHITE).pack(side="left")
        br=tk.Frame(f,bg=C_BG); br.pack(fill="x",pady=4)
        _btn(br,"Remove selected",self._ro,bg=C_RED,fg=C_WHITE).pack(side="left")
        _btn(br,"Close",self._sc,bg=C_GREEN,fg=C_WHITE).pack(side="right")
    def _ao(self):
        n=self._nv.get().strip()
        if not n: return
        if n in self._o: messagebox.showwarning("Dup",f"'{n}' exists.",parent=self); return
        self._o.append(n); self._olb.insert("end",f"   {n}"); self._nv.set(""); self._p()
    def _ro(self):
        s=self._olb.curselection()
        if not s: return
        i=s[0]
        if messagebox.askyesno("Remove",f"Remove '{self._o[i]}'?",parent=self): self._o.pop(i); self._olb.delete(i); self._p()
    def _tbx(self):
        """Box sizes for the Shipping Slip's Size list (Dhruvil 2026-09-25).
        Operators add sizes from the slip form ('+ Add new size…'); here an admin
        can also add one, or remove a mistyped one. Saved to the share at once,
        like the operator roster; every station sees the list next time the
        form opens."""
        self._bx=load_box_sizes()
        f=self._bd; tk.Label(f,text="Box sizes (Shipping Slip)",bg=C_BG,fg=C_GREEN,font=("Segoe UI",11,"bold")).pack(anchor="w",pady=(0,8))
        lf=tk.Frame(f,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); lf.pack(fill="both",expand=True)
        sb=ttk.Scrollbar(lf); sb.pack(side="right",fill="y")
        self._blb=tk.Listbox(lf,font=("Segoe UI",11),bg=C_WHITE,fg=C_DARK,selectbackground=C_GREEN_LITE,selectforeground=C_GREEN,relief="flat",bd=0,activestyle="none",yscrollcommand=sb.set)
        self._blb.pack(fill="both",expand=True,padx=6,pady=6); sb.config(command=self._blb.yview)
        for s in self._bx: self._blb.insert("end",f"   {s}")
        ar=tk.Frame(f,bg=C_BG); ar.pack(fill="x",pady=(10,4))
        tk.Label(ar,text="New size:",bg=C_BG,fg=C_MID,font=("Segoe UI",9)).pack(side="left")
        self._bxv=tk.StringVar(); e=ttk.Entry(ar,textvariable=self._bxv,font=("Segoe UI",11),width=26); e.pack(side="left",padx=(8,8),ipady=4)
        e.bind("<Return>",lambda ev: self._abx()); _btn(ar,"+ Add",self._abx,bg=C_GREEN,fg=C_WHITE).pack(side="left")
        br=tk.Frame(f,bg=C_BG); br.pack(fill="x",pady=4)
        _btn(br,"Remove selected",self._rbx,bg=C_RED,fg=C_WHITE).pack(side="left")
        _btn(br,"Close",self._sc,bg=C_GREEN,fg=C_WHITE).pack(side="right")
    # Add / remove go through add_box_size / remove_box_size, which read the
    # shared list FRESH and merge — the list this tab showed when it opened may
    # be stale (a station added a size since), and writing it back erased that
    # size (review round 2). The tab then shows the list as saved.
    def _abx(self):
        s=normalize_box_size(self._bxv.get())
        if not s: return
        bad=box_size_problem(s)
        if bad: messagebox.showwarning("Box size",bad,parent=self); return
        fresh=load_box_sizes()
        if any(x.lower()==s.lower() for x in fresh):
            self._bx=fresh; self._blb_refresh()
            messagebox.showwarning("Dup",f"'{s}' exists.",parent=self); return
        try: sizes,saved,err=add_box_size(s)
        except Exception as e: sizes,saved,err=fresh,False,f"{type(e).__name__}: {e}"
        if not saved: self._bx_not_saved(err); return
        self._bx=sizes; self._blb_refresh(); self._bxv.set("")
    def _rbx(self):
        sel=self._blb.curselection()
        if not sel: return
        s=self._bx[sel[0]]
        if len(self._bx)==1: messagebox.showerror("Cannot","Keep at least one box size.",parent=self); return
        if not messagebox.askyesno("Remove",f"Remove '{s}' from the box sizes?",parent=self): return
        sizes,saved,err=remove_box_size(s)
        if not saved: self._bx_not_saved(err); return
        self._bx=sizes; self._blb_refresh()
    def _blb_refresh(self):
        self._blb.delete(0,"end")
        for s in self._bx: self._blb.insert("end",f"   {s}")
    def _bx_not_saved(self,err):
        messagebox.showerror("Not saved",
            "Could not save the box sizes to the server:\n\n"
            f"{BOX_SIZES_FILE}\n\n{err}\n\n"
            "Your change was NOT saved — try again when the server is reachable.", parent=self)
    def _tad(self):
        f=self._bd; tk.Label(f,text="Administrator accounts",bg=C_BG,fg=C_DARK,font=("Segoe UI",11,"bold")).pack(anchor="w",pady=(0,8))
        lf=tk.Frame(f,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); lf.pack(fill="both",expand=True)
        sb=ttk.Scrollbar(lf); sb.pack(side="right",fill="y")
        self._alb=tk.Listbox(lf,font=("Segoe UI",11),bg=C_WHITE,fg=C_DARK,selectbackground="#FFE8E8",selectforeground=C_RED,relief="flat",bd=0,activestyle="none",yscrollcommand=sb.set)
        self._alb.pack(fill="both",expand=True,padx=6,pady=6); sb.config(command=self._alb.yview)
        for a in self._a: self._alb.insert("end",f"   {a['name']}   [{a['id']}]")
        g=tk.Frame(f,bg=C_BG); g.pack(fill="x",pady=(10,4))
        self._an=tk.StringVar(); self._ai=tk.StringVar(); self._ap=tk.StringVar()
        for c,(l,v,kw) in enumerate([("Name",self._an,{}),("ID",self._ai,{}),("Password",self._ap,{"show":"●"})]):
            tk.Label(g,text=l,bg=C_BG,fg=C_MID,font=("Segoe UI",9)).grid(row=0,column=c*2,sticky="w",padx=(0,4))
            e=ttk.Entry(g,textvariable=v,font=("Segoe UI",10),width=14,**kw)
            e.grid(row=1,column=c*2,padx=(0,10),ipady=3)
            if kw.get("show"): self._ap_e=e
        _btn(g,"+ Add",self._aa,bg=C_DARK,fg=C_WHITE).grid(row=1,column=6,padx=(4,0))
        self._ap_show=tk.BooleanVar(value=False)
        tk.Checkbutton(g,text="Show password",variable=self._ap_show,bg=C_BG,fg=C_SUBTLE,
                       activebackground=C_BG,font=("Segoe UI",8),
                       command=lambda: self._ap_e.config(show="" if self._ap_show.get() else "●")
                       ).grid(row=2,column=4,sticky="w")
        br=tk.Frame(f,bg=C_BG); br.pack(fill="x",pady=4)
        _btn(br,"Remove selected",self._ra,bg=C_RED,fg=C_WHITE).pack(side="left")
        _btn(br,"Set password…",self._chpw,bg=C_DARK,fg=C_WHITE).pack(side="left",padx=(8,0))
        _btn(br,"Close",self._sc,bg=C_GREEN,fg=C_WHITE).pack(side="right")
    def _aa(self):
        n,i,p=self._an.get().strip(),self._ai.get().strip(),self._ap.get()
        if not n or not i or not p: messagebox.showwarning("Missing","Fill all fields.",parent=self); return
        if any(a["id"].lower()==i.lower() for a in self._a): messagebox.showwarning("Dup",f"ID '{i}' exists.",parent=self); return
        self._a.append({"id":i,"name":n,"pw_hash":hash_password(p)})
        self._alb.insert("end",f"   {n}   [{i}]"); self._an.set(""); self._ai.set(""); self._ap.set(""); self._p()
    def _ra(self):
        s=self._alb.curselection()
        if not s: return
        if len(self._a)==1: messagebox.showerror("Cannot","At least one admin required.",parent=self); return
        i=s[0]
        if messagebox.askyesno("Remove",f"Remove '{self._a[i]['name']}'?",parent=self): self._a.pop(i); self._alb.delete(i); self._p()
    def _chpw(self):
        """Set a new password for the selected admin — previously impossible
        without deleting and re-adding the account (and a sole admin could never
        rotate their own password at all)."""
        s=self._alb.curselection()
        if not s:
            messagebox.showwarning("No selection","Select an administrator first.",parent=self); return
        i=s[0]; adm=self._a[i]
        dlg=tk.Toplevel(self); dlg.title("Set password"); dlg.resizable(False,False)
        dlg.transient(self); dlg.grab_set(); dlg.configure(bg=C_BG)
        h=tk.Frame(dlg,bg=C_DARK); h.pack(fill="x")
        tk.Label(h,text=f"Set password — {adm['name']} [{adm['id']}]",bg=C_DARK,fg=C_WHITE,
                 font=("Segoe UI",11,"bold")).pack(padx=20,pady=12)
        bd=tk.Frame(dlg,bg=C_BG,padx=28,pady=16); bd.pack()
        v1=tk.StringVar(); v2=tk.StringVar()
        tk.Label(bd,text="New password",bg=C_BG,fg=C_MID,font=("Segoe UI",9)).pack(anchor="w")
        e1=ttk.Entry(bd,textvariable=v1,font=("Segoe UI",11),width=26,show="●"); e1.pack(fill="x",pady=(2,10),ipady=4)
        tk.Label(bd,text="Confirm password",bg=C_BG,fg=C_MID,font=("Segoe UI",9)).pack(anchor="w")
        e2=ttk.Entry(bd,textvariable=v2,font=("Segoe UI",11),width=26,show="●"); e2.pack(fill="x",pady=(2,4),ipady=4)
        sv=tk.BooleanVar(value=False)
        tk.Checkbutton(bd,text="Show passwords",variable=sv,bg=C_BG,fg=C_SUBTLE,activebackground=C_BG,
                       font=("Segoe UI",8),command=lambda: [e.config(show="" if sv.get() else "●") for e in (e1,e2)]
                       ).pack(anchor="w",pady=(0,6))
        err=tk.Label(bd,text="",bg=C_BG,fg=C_RED,font=("Segoe UI",9)); err.pack(anchor="w")
        def ok():
            p1,p2=v1.get(),v2.get()
            if not p1: err.config(text="Password cannot be empty."); return
            if p1!=p2: err.config(text="Passwords do not match."); return
            adm["pw_hash"]=hash_password(p1)
            self._p()
            dlg.destroy()
            messagebox.showinfo("Password set",f"Password updated for {adm['name']}.",parent=self)
        bf=tk.Frame(dlg,bg=C_BG,pady=12); bf.pack()
        _btn(bf,"Set password",ok,bg=C_DARK,fg=C_WHITE,font_size=10,padx=20,pady=8).pack(side="left",padx=6)
        _btn(bf,"Cancel",dlg.destroy,bg="#DDD",fg=C_MID,font_size=10,padx=20,pady=8).pack(side="left",padx=6)
        dlg.bind("<Return>",lambda e: ok()); dlg.bind("<Escape>",lambda e: dlg.destroy())
        _center(dlg,self); e1.focus_set(); dlg.wait_window()
        # Tk grabs don't stack: re-assert AdminPanel's modal grab released by the
        # child dialog's destroy.
        try: self.grab_set()
        except Exception: pass
    def _p(self):
        # Persist to the share. If it fails (e.g. PCC2 offline) report clearly and
        # return False instead of letting the exception escape the Tk callback —
        # that used to silently jam the panel (the change looked saved but wasn't,
        # and Save & Close appeared dead because destroy() never ran).
        try:
            save_admins(self._a); save_operators(self._o); self._os(self._a,self._o)
            return True
        except Exception as e:
            messagebox.showerror("Not saved",
                "Could not save roster changes to the server:\n\n"
                f"{ADMINS_FILE}\n\n{type(e).__name__}: {e}\n\n"
                "PCC2 may be offline. Your change was NOT saved — try again when the "
                "server is reachable.", parent=self)
            return False
    def _sc(self):
        if self._p(): self.destroy()

###############################################################################
# HISTORY VIEWER
###############################################################################
# Read-only browser over the JSON event files under HISTORY_DIR. Loads only
# the selected month folders, filters in memory, and never writes anything
# except a user-requested CSV export.

def _hist_event_time(r):
    """Event time as an epoch float (ts_epoch, else parse ts). 0.0 if unknown."""
    try:
        return float(r.get("ts_epoch"))
    except Exception:
        pass
    try:
        return time.mktime(time.strptime(str(r.get("ts") or ""), "%Y-%m-%d %H:%M:%S"))
    except Exception:
        return 0.0

class PortalQueueDialog(tk.Toplevel):
    """This PC's Customer Portal uploads (newest first) with a Retry now button."""
    LABELS={"pending":"Waiting","done":"Sent","failed":"FAILED","blocked":"Not sent (blocked)"}
    def __init__(self,parent,uploader):
        super().__init__(parent); self.title("Portal uploads"); self.geometry("780x420"); self.configure(bg=C_BG)
        self.transient(parent); self._up=uploader
        h=tk.Frame(self,bg=C_GREEN); h.pack(fill="x")
        tk.Label(h,text="Customer Portal uploads on this PC",bg=C_GREEN,fg=C_WHITE,font=("Segoe UI",12,"bold")).pack(side="left",padx=18,pady=10)
        bar=tk.Frame(self,bg=C_BG); bar.pack(side="bottom",fill="x",padx=18,pady=10)
        _btn(bar,"Retry now",self._retry,bg=C_TEAL,fg=C_WHITE).pack(side="left")
        _btn(bar,"Close",self.destroy,bg="#DDD",fg=C_MID).pack(side="right")
        cols=("when","what","status","note")
        self._tv=ttk.Treeview(self,columns=cols,show="headings",height=12)
        for c,w,t in [("when",120,"Queued"),("what",230,"Document"),("status",140,"Status"),("note",270,"Details")]:
            self._tv.heading(c,text=t); self._tv.column(c,width=w,anchor="w")
        self._tv.pack(fill="both",expand=True,padx=18,pady=(10,0))
        self._fill(); _center(self,parent)
    def _fill(self):
        self._tv.delete(*self._tv.get_children())
        jobs=sorted(portal_upload.load_jobs(),key=lambda j: j.get("created_at",0),reverse=True)
        for j in jobs[:200]:
            when=datetime.datetime.fromtimestamp(j.get("created_at",0)).strftime("%m/%d %I:%M %p")
            self._tv.insert("","end",values=(when,portal_upload.describe(j),self.LABELS.get(j.get("status"),j.get("status")),j.get("last_error","")))
    def _retry(self):
        if self._up: self._up.retry_now()
        self.after(1500,self._fill)

def group_shipment_history(records, voided_ids):
    """Collapse raw History events into ONE item per shipment for the viewer
    (Dhruvil 2026-10-24: 'one row per shipment, not per save'). A shipment is the
    draft_save + close + delete events that share a draft_id; events with no
    draft_id, and every receive / backfill_shipped / void, each stand alone so
    nothing disappears from History. Returns a list of item dicts with a common
    flat display shape (caller sorts):
      kind='shipment' -> draft_id, status, wo/eng/pn/customer/division/operator,
        started_by/at, saved_by/at (last save), closed_by/at/station, serial_count,
        ts (newest event epoch), timeline (events oldest->newest), rep (the close,
        else the newest event).
      kind='event'    -> rec (raw), same flat keys."""
    groups = {}
    standalone = []
    for r in records:
        if not isinstance(r, dict):
            continue
        et = r.get("event_type") or ""
        did = str(r.get("draft_id") or "")
        if did and et in ("draft_save", "close", "delete"):
            groups.setdefault(did, []).append(r)
        else:
            standalone.append(r)
    items = []
    for did, evs in groups.items():
        evs = sorted(evs, key=_hist_event_time)
        closes = [e for e in evs if e.get("event_type") == "close"]
        dels = [e for e in evs if e.get("event_type") == "delete"]
        saves = [e for e in evs if e.get("event_type") == "draft_save"]
        close = closes[-1] if closes else None
        latest = evs[-1]
        if close and str(close.get("event_id") or "") in voided_ids:
            status = "🚫 Voided"
        elif dels and not closes:
            status = "✕ Deleted"
        elif close:
            status = "✓ Closed"
        else:
            status = "✎ Draft"
        started_by = started_at = ""
        for e in evs:                       # an explicit WO-load stamp wins
            if e.get("started_by") or e.get("started_at"):
                started_by = str(e.get("started_by") or ""); started_at = str(e.get("started_at") or ""); break
        if not started_at:                  # fall back to the first save / its resumed-from
            first = saves[0] if saves else evs[0]
            rd = first.get("resumed_from_draft") or {}
            started_by = started_by or str(first.get("operator") or rd.get("saved_by") or "")
            started_at = str(rd.get("saved_at") or first.get("ts") or "")
        last_save = saves[-1] if saves else None
        rep = close or latest
        items.append({
            "kind": "shipment", "draft_id": did, "status": status,
            "wo": str(rep.get("wo") or ""), "eng": str(rep.get("eng") or ""),
            "pn": str(rep.get("pn") or ""), "customer": str(rep.get("customer") or ""),
            "division": str(rep.get("division") or ""), "operator": str(rep.get("operator") or ""),
            "started_by": started_by, "started_at": started_at,
            "saved_by": str((last_save or {}).get("operator") or ""),
            "saved_at": str((last_save or {}).get("ts") or ""),
            "closed_by": str((close or {}).get("operator") or ""),
            "closed_at": str((close or {}).get("ts") or ""),
            "station": str((close or {}).get("station") or ""),
            "serial_count": (rep.get("serial_count") if rep.get("serial_count") is not None
                             else len(rep.get("serials") or [])),
            "ts": _hist_event_time(latest), "timeline": evs, "rep": rep,
        })
    _EVL = {"receive": "📦 Received", "backfill_shipped": "⏪ Imported", "void": "🚫 Void",
            "draft_save": "✎ Draft", "close": "✓ Closed", "delete": "✕ Deleted"}
    for r in standalone:
        et = r.get("event_type") or ""
        posted = et in ("receive", "backfill_shipped")
        items.append({
            "kind": "event", "rec": r, "status": _EVL.get(et, et or "?"),
            "wo": str(r.get("wo") or ""), "eng": str(r.get("eng") or ""), "pn": str(r.get("pn") or ""),
            "customer": str(r.get("customer") or ""), "division": str(r.get("division") or ""),
            "operator": str(r.get("operator") or ""),
            "started_by": "", "started_at": "", "saved_by": "", "saved_at": "",
            "closed_by": str(r.get("operator") or "") if posted else "",
            "closed_at": str(r.get("ts") or "") if (posted or et == "void") else "",
            "station": str(r.get("station") or ""),
            "serial_count": (r.get("serial_count") if r.get("serial_count") is not None
                             else len(r.get("serials") or [])),
            "ts": _hist_event_time(r), "timeline": [r], "rep": r,
        })
    return items

class HistoryDetail(tk.Toplevel):
    """Read-only detail card for a single shipment history record."""
    LABELS={"close":"Closed","draft_save":"Draft","delete":"Deleted",
            "receive":"Received","backfill_shipped":"Imported baseline",
            "void":"Void marker"}
    COLORS={"close":C_GREEN,"draft_save":C_TEAL,"delete":C_RED,
            "receive":C_ORANGE,"backfill_shipped":C_MID,"void":C_SUBTLE}
    # Record types a supervisor may VOID (serials stop counting; record kept for audit).
    VOIDABLE=("close","receive","backfill_shipped")
    def __init__(self,parent,rec,voided=None):
        super().__init__(parent); self._rec=rec or {}
        self._voided_info=voided   # the void marker record, when this record is voided
        et=self._rec.get("event_type") or "?"; label=self.LABELS.get(et,et)
        self.title(f"History — {label} — WO {self._rec.get('wo','?')}")
        self.geometry("820x680"); self.transient(parent); self.configure(bg=C_BG)
        self._build(label); _center(self,parent)
        self.bind("<Escape>",lambda e: self.destroy())
    def _build(self,label):
        r=self._rec; hc=self.COLORS.get(r.get("event_type"),C_GREEN)
        h=tk.Frame(self,bg=hc); h.pack(fill="x")
        tk.Label(h,text=f"{label} — WO {r.get('wo','?')}",bg=hc,fg=C_WHITE,font=("Segoe UI",13,"bold")).pack(side="left",padx=20,pady=12)
        tk.Label(h,text=str(r.get("ts") or ""),bg=hc,fg="#DDEEE4",font=("Segoe UI",10)).pack(side="right",padx=20)
        bb=tk.Frame(self,bg=C_BG); bb.pack(side="bottom",fill="x",padx=16,pady=10)
        _btn(bb,"Close",self.destroy,bg="#DDD",fg=C_MID).pack(side="right")
        if r.get("wo_folder"):
            _btn(bb,"Open WO Folder",lambda: self._open(r.get("wo_folder")),bg=C_GREEN,fg=C_WHITE).pack(side="left",padx=(0,8))
        # Supervisor void — only for countable records that aren't already voided.
        if (r.get("event_type") in self.VOIDABLE and r.get("event_id")
                and not self._voided_info):
            _btn(bb,"🚫 Void record…",self._void,bg=C_RED,fg=C_WHITE).pack(side="left",padx=(0,8))
        # Supervisor re-emit — regenerate a closed shipment's COC/Packing List from
        # History (identical format) when the folder copy was deleted or lost.
        if (r.get("event_type")=="close" and (r.get("coc") or r.get("packing") or r.get("slip"))
                and not self._voided_info):
            _btn(bb,"📄 Re-emit paperwork…",self._reemit,bg=C_TEAL,fg=C_WHITE).pack(side="left",padx=(0,8))
        out=r.get("outputs") or {}
        # Shipping Slip / Labels pages (Dhruvil 2026-09-24): recorded inside the
        # slip data (and mirrored into outputs by the close) — either works.
        sout=(r.get("slip") or {}).get("outputs") or {}
        # The "Open …" buttons get a row of their own above the actions: a full
        # close record has five of them, more than one row of this window holds
        # (pack silently dropped the last ones off the right edge).
        ob=None
        for lbl,key in [("Open COC","coc_xlsx"),("Open Packing List","packing_xlsx"),("Open Serials","serials_xlsx"),
                        ("Open Shipping Slip","slip_html"),("Open Shipping Labels","labels_html"),("Open Packing Label","pack_html")]:
            p=out.get(key) or sout.get(key)
            if not p: continue
            if ob is None: ob=tk.Frame(self,bg=C_BG); ob.pack(side="bottom",fill="x",padx=16,pady=(0,2))
            _btn(ob,lbl,lambda x=p: self._open(x),bg=C_BLUE,fg=C_WHITE).pack(side="left",padx=(0,8))
        wrap=tk.Frame(self,bg=C_BG); wrap.pack(fill="both",expand=True,padx=16,pady=(10,0))
        cv=tk.Canvas(wrap,bg=C_BG,highlightthickness=0)
        sb=ttk.Scrollbar(wrap,orient="vertical",command=cv.yview); cv.configure(yscrollcommand=sb.set)
        sb.pack(side="right",fill="y"); cv.pack(side="left",fill="both",expand=True)
        bd=tk.Frame(cv,bg=C_BG); wid=cv.create_window((0,0),window=bd,anchor="nw")
        cv.bind("<Configure>",lambda e: cv.itemconfigure(wid,width=e.width))
        bd.bind("<Configure>",lambda e: cv.configure(scrollregion=cv.bbox("all")))
        self.bind("<MouseWheel>",lambda e: None if isinstance(e.widget,tk.Listbox) else cv.yview_scroll(int(-e.delta/120),"units"))
        def sec(title):
            tk.Label(bd,text=title,bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",8,"bold")).pack(anchor="w",pady=(6,0))
            return _card(bd)
        if self._voided_info:
            vi=self._voided_info
            vb=tk.Frame(bd,bg="#FBEAE8",highlightbackground=C_RED,highlightthickness=2)
            vb.pack(fill="x",pady=(0,6))
            tk.Label(vb,text=("🚫 VOIDED by "+(vi.get("approved_by") or "?")+" on "+(vi.get("ts") or "?")
                     +(("  —  reason: "+str(vi.get("reason"))) if vi.get("reason") else "")
                     +"\nThis record no longer counts — duplicate checks and the Receiving tracker ignore it."),
                     bg="#FBEAE8",fg=C_RED,font=("Segoe UI",10,"bold"),justify="left",
                     wraplength=660,padx=12,pady=8).pack(anchor="w")
        c=sec("Event")
        _info_row(c,"Event ID",str(r.get("event_id") or "—"))
        _info_row(c,"Station",f"{r.get('station') or '—'}  ·  user {r.get('win_user') or '—'}  ·  app {r.get('app_version') or '—'}")
        c=sec("Job")
        _info_row(c,"Workorder #",str(r.get("wo") or "—"))
        _info_row(c,"Division",str(r.get("division") or "—"))
        _info_row(c,"Customer",str(r.get("customer") or "—"))
        _info_row(c,"P/N",str(r.get("pn") or "—"))
        _info_row(c,"Engineering #",str(r.get("eng") or "—"))
        _info_row(c,"WO folder",str(r.get("wo_folder") or "—"),fg=C_MID)
        c=sec("Operator")
        _info_row(c,"Operator",str(r.get("operator") or "—"))
        rd=r.get("resumed_from_draft") or {}
        if rd: _info_row(c,"Resumed from",f"Started by {rd.get('saved_by') or '?'} at {rd.get('saved_at') or '?'}",fg=C_TEAL)
        if str(r.get("notes") or "").strip():
            c=sec("Notes"); _info_row(c,"Notes",str(r.get("notes")))
        ov=r.get("dup_overrides") or []
        if ov:
            c=sec("Supervisor overrides")
            for o in ov:
                pr=(o.get("prior") or {})
                why=(str(o.get("reason")) if o.get("reason")
                     else (f"prior WO {pr.get('wo') or '?'} {pr.get('ts') or ''}").strip())
                _info_row(c,str(o.get("serial") or "?"),
                          f"approved by {o.get('approved_by') or '?'} — {why}",fg=C_AMBER)
        sl=[str(s) for s in (r.get("serials") or [])]
        c=sec("Serials")
        _info_row(c,"Count",str(r.get("serial_count",len(sl))))
        if sl:
            lf=tk.Frame(c,bg=C_WHITE); lf.pack(fill="x",padx=8,pady=6)
            ssb=ttk.Scrollbar(lf); ssb.pack(side="right",fill="y")
            lb=tk.Listbox(lf,font=("Courier New",10),height=min(8,len(sl)),bg=C_WHITE,fg=C_MID,relief="flat",bd=0,activestyle="none",yscrollcommand=ssb.set)
            lb.pack(side="left",fill="x",expand=True); ssb.config(command=lb.yview)
            for s in sl: lb.insert("end",f"  {s}")
        coc=r.get("coc") or {}
        if coc:
            c=sec("COC")
            for k,lbl in [("ship_date","Ship Date"),("po","PO #"),("pn","P/N"),("rev","Rev"),("qty","QTY"),("datecode","Date Code"),("job","PCC Job #"),("type","Type"),("assembly_process","Assembly")]:
                _info_row(c,lbl,str(coc.get(k) or "—"))
        pk=r.get("packing") or {}
        if pk:
            c=sec("Packing List")
            for k,lbl in [("customer","Customer"),("packing_no","Packing #"),("po","P.O. #"),("pn","P/N"),("wo","WO / LOT #"),("eng","Eng #"),("date","Date"),("datecode","Date Code")]:
                _info_row(c,lbl,str(pk.get(k) or "—"))
        slip=r.get("slip") or {}
        if slip:
            c=sec("Shipping Slip & Labels")
            for k,lbl in [("customer","Customer"),("cpn","Customer P/N"),("po","Customer PO"),("wo","WO / Lot"),("eng","Eng #"),("rev","Rev"),
                          ("qty","Qty Shipped"),("date","Date"),("datecode_wwyy","Date Code"),("prepared_by","Prepared By"),("notes","Notes")]:
                _info_row(c,lbl,str(slip.get(k) or "—"))
            bx=slip.get("boxes") or []
            for i,b in enumerate(bx,1):
                b=b or {}
                _info_row(c,f"Box {i}/{len(bx)}",f"{b.get('size') or '—'}   ·   qty {b.get('qty') or '—'}   ·   {b.get('weight') or '—'} lbs")
            tq,tw=calc_box_totals(bx)
            _info_row(c,"Totals",f"qty {tq}   ·   {tw}")
            _info_row(c,"Customer Parts",str(slip.get("customer_parts") or "No"))
            flags=[t for k,t in (("fair","FAIR"),("si","SI"),("ppap","PPAP"),("cofc","C of C")) if slip.get(k)]
            _info_row(c,"Paperwork",", ".join(flags) if flags else "—")
            _info_row(c,"Ship Via",str(slip.get("ship_via") or "—"))
            _info_row(c,"Label notes",str(slip.get("label_notes") or "—"))
            _info_row(c,"Slip code",str(slip.get("slip_code") or "—"),fg=C_MID)
            _info_row(c,"Labels code",str(slip.get("labels_code") or "—"),fg=C_MID)
            pr=slip.get("printed") or {}
            _info_row(c,"Slip printed",str(pr.get("slip") or "not printed"),fg=C_TEAL if pr.get("slip") else C_SUBTLE)
            _info_row(c,"Labels printed",str(pr.get("labels") or "not printed"),fg=C_TEAL if pr.get("labels") else C_SUBTLE)
            if pr.get("pack") or slip.get("pack_code"):
                _info_row(c,"Packing label",f"QTY {slip.get('pack_qty') or '—'}  ·  "
                          +("CAGE "+str(slip.get("pack_cage") or "—")+"  ·  D/C "+str(slip.get("pack_datecode") or "—")
                            if str(slip.get("pack_mode") or "itar")!="industrial" else "DATE "+str(slip.get("pack_date") or "—")))
                _info_row(c,"Packing label code",str(slip.get("pack_code") or "—"),fg=C_MID)
                _info_row(c,"Packing label printed",str(pr.get("pack") or "not printed"),fg=C_TEAL if pr.get("pack") else C_SUBTLE)
        ph=r.get("photos") or {}
        if ph:
            c=sec("Photos")
            _info_row(c,"Moved",str(ph.get("moved_count",0)))
            mv=[Path(str(p)).name for p in (ph.get("moved") or [])]
            if mv: _info_row(c,"Files",", ".join(mv))
            if ph.get("photo_dir"): _info_row(c,"Photo dir",str(ph.get("photo_dir")),fg=C_MID)
            if ph.get("missing"): _info_row(c,"Missing",", ".join(str(x) for x in ph["missing"]),fg=C_AMBER)
            if ph.get("failed"): _info_row(c,"Failed",", ".join(str(x) for x in ph["failed"]),fg=C_RED)
        fs=r.get("file_status") or {}
        spr=slip.get("printed") or {}
        if out or fs or any(sout.values()) or spr.get("slip") or spr.get("labels") or spr.get("pack"):
            # a draft that only carries an unprinted slip form has no output files
            # (sout is the template's {None, None} then) — no section for it
            c=sec("Output files")
            rows=[("Shipping Slip","slip_html",None,"slip"),("Shipping Labels","labels_html",None,"labels"),("Packing Label","pack_html",None,"pack")]
            if out or fs:
                rows=[("Serials xlsx","serials_xlsx","serials",None),("Serials Word","serials_docx","word",None),
                      ("COC","coc_xlsx","coc",None),("Packing List","packing_xlsx","packing",None)]+rows
            for lbl,pkey,skey,prkey in rows:
                p=out.get(pkey) or sout.get(pkey); st=str(fs.get(skey) or "") if skey else ""
                if prkey and not p:
                    # slip pages only when printed; printed with the share down
                    # = no filed copy yet (Re-emit paperwork files it)
                    if spr.get(prkey):
                        _info_row(c,lbl,"printed while the share was down — not filed yet (Re-emit paperwork files it into the WO folder)",fg=C_AMBER)
                    continue
                _info_row(c,lbl,(str(p) if p else "—")+(f"   [{st}]" if st else ""),fg=C_DARK if p else C_SUBTLE)
        if r.get("draft_path"):
            c=sec("Draft")
            _info_row(c,"Draft file",str(r.get("draft_path")),fg=C_MID)
        if r.get("event_type")=="close":
            c=sec("Result")
            ok=bool(r.get("critical_ok"))
            _info_row(c,"Status","OK" if ok else "ERRORS",fg=C_GREEN if ok else C_RED)
            for err in (r.get("errors") or []): _info_row(c,"Error",str(err).strip(),fg=C_RED)
        dl=r.get("deleted") or {}
        if dl:
            c=sec("Deletion")
            _info_row(c,"Had COC","yes" if dl.get("had_coc") else "no")
            _info_row(c,"Had Packing List","yes" if dl.get("had_packing") else "no")
            if "had_slip" in dl:
                _info_row(c,"Had Shipping Slip","yes" if dl.get("had_slip") else "no")
                # pages printed before the delete stay in the WO folder (never touched)
                for p in [p for p in (dl.get("slip_outputs") or {}).values() if p]:
                    _info_row(c,"Slip page left in WO folder",str(p),fg=C_AMBER)
            _info_row(c,"Draft removed","yes" if dl.get("draft_removed") else f"NO — {dl.get('draft_error') or ''}",fg=C_DARK if dl.get("draft_removed") else C_RED)
    def _open(self,p):
        """Open a file/folder with the Windows default handler; warn, never crash."""
        try:
            if not p or not Path(str(p)).exists():
                messagebox.showwarning("Not found",f"Path does not exist (moved or renamed?):\n\n{p}",parent=self); return
            os.startfile(str(p))
        except Exception as e:
            messagebox.showerror("Open failed",f"Could not open:\n{p}\n\n{type(e).__name__}: {e}",parent=self)
    def _void(self):
        """Supervisor-only: void this record (admin login + reason). The record stays
        in history marked VOIDED, but its serials stop counting — for shipments that
        were recorded by mistake (wrong WO, wrong boards, double entry)."""
        r=self._rec
        n=r.get("serial_count",len(r.get("serials") or []))
        if not messagebox.askyesno("Void record",
                f"Void this {self.LABELS.get(r.get('event_type'),'record')} record?\n\n"
                f"WO {r.get('wo','?')}  ·  {n} serial(s)  ·  {r.get('ts','?')}\n\n"
                f"The record STAYS in history marked VOIDED (audit), but its serials "
                f"no longer count as shipped/received — duplicate checks and the "
                f"Receiving tracker will ignore it.",parent=self):
            return
        admins=load_admins()
        if not admins:
            messagebox.showerror("Admin roster unavailable",
                "The admin list could not be loaded from the share, so the void "
                "can't be verified right now. Check the network and try again.",parent=self)
            return
        lg=AdminLoginDialog(self,admins)
        if not getattr(lg,"result",False): return
        from tkinter import simpledialog
        reason=(simpledialog.askstring("Reason",
                "Why is this record being voided? (kept in the audit trail)",
                parent=self) or "").strip()
        p=write_void_event(r,reason,getattr(lg,"admin_id",""))
        if not p:
            messagebox.showerror("Void failed",
                "Could not write the void record to the share history — nothing "
                "was changed. Check the network and try again.",parent=self)
            return
        # Take effect IMMEDIATELY on this station: the duplicate check caches the
        # shipped-serials index per loaded job, so drop that cache (keeping any
        # in-progress supervisor approvals) and rebuild; same for an open
        # Receiving window. Other stations pick the void up on their next WO load.
        try:
            app=self.master.master if isinstance(self.master,HistoryViewer) else None
            if app is not None:
                try:
                    app._dup_index={}; app._dup_index_ready=False
                    app._dup_index_key=None; app._dup_index_building=None
                    if getattr(app,"wo_folder",None): app._kickoff_dup_index()
                except Exception: pass
                try:
                    rw=getattr(app,"_receiving_win",None)
                    if rw is not None and rw.winfo_exists():
                        job=rw._current_job()
                        if job:
                            lbl=rw._shown_pn(rw._state) if rw._state else _job_display_pn(job)   # the label the window shows
                            rw._state_key=None
                            rw._start_state_build(job["eng"],job["pn"],label=lbl)
                except Exception: pass
        except Exception: pass
        messagebox.showinfo("Voided",
            "✓ Record voided. Its serials no longer count — effective immediately "
            "on this PC (duplicate check + Receiving tracker refreshed). Other "
            "stations update when they next load a work order.",parent=self)
        try:
            if isinstance(self.master,HistoryViewer): self.master._search()
        except Exception: pass
        self.destroy()
    def _reemit(self):
        """Supervisor-only: re-generate this closed shipment's COC + Packing List
        from the History record (barcodes/logo/format identical to the app's
        originals; the original ship-date filenames). For paperwork deleted or
        lost from the WO folder. Never overwrites a document still on disk."""
        r=self._rec
        n=r.get("serial_count",len(r.get("serials") or []))
        docs=[]
        if r.get("coc"): docs.append("COC")
        if r.get("packing"): docs.append("Packing List")
        # Slip pages only re-emit when they were actually printed: a recorded
        # path, or a printed stamp alone (printed with the share down — the
        # re-emit is what files that page into the WO folder).
        sout=(r.get("slip") or {}).get("outputs") or {}; out=r.get("outputs") or {}; spr=(r.get("slip") or {}).get("printed") or {}
        if sout.get("slip_html") or out.get("slip_html") or spr.get("slip"): docs.append("Shipping Slip")
        if sout.get("labels_html") or out.get("labels_html") or spr.get("labels"): docs.append("Shipping Labels")
        if sout.get("pack_html") or out.get("pack_html") or spr.get("pack"): docs.append("Packing Label")
        if not docs:
            messagebox.showinfo("Nothing to re-emit",
                "This record has no COC / Packing List"+(" / printed Shipping Slip" if r.get("slip") else "")
                +" data to rebuild from.",parent=self); return
        if not messagebox.askyesno("Re-emit paperwork",
                f"Re-generate {' + '.join(docs)} for this shipment from History?\n\n"
                f"WO {r.get('wo','?')}  ·  {n} serial(s)  ·  {r.get('ts','?')}\n\n"
                f"Files go into the shipment's folder with the ORIGINAL ship-date "
                f"name, in the app's exact format (barcodes and logo included). A "
                f"document still on disk is left untouched (never overwritten). Use "
                f"this to restore paperwork that was deleted or lost.",parent=self):
            return
        admins=load_admins()
        if not admins:
            messagebox.showerror("Admin roster unavailable",
                "The admin list could not be loaded from the share, so this can't be "
                "verified right now. Check the network and try again.",parent=self)
            return
        lg=AdminLoginDialog(self,admins)
        if not getattr(lg,"result",False): return
        res=reemit_paperwork(r)
        parts=[]
        if res.get("wrote"):
            parts.append("Re-created:\n"+"\n".join("  • "+Path(p).name for p in res["wrote"]))
        if res.get("kept"):
            parts.append("Left as-is:\n"+"\n".join("  • "+k for k in res["kept"]))
        if res.get("errors"):
            parts.append("Problems:\n"+"\n".join("  • "+e for e in res["errors"]))
        msg="\n\n".join(parts) or "Nothing was done."
        if res.get("wrote") and not res.get("errors"):
            messagebox.showinfo("Re-emit paperwork — done",msg,parent=self)
        elif res.get("wrote"):
            messagebox.showwarning("Re-emit paperwork — partial",msg,parent=self)
        else:
            messagebox.showerror("Re-emit paperwork",msg,parent=self)

class ShipmentDetail(tk.Toplevel):
    """One shipment's full save TIMELINE (Dhruvil 2026-10-24): the draft-saves,
    close and delete that share a draft_id, oldest->newest. Double-click a row to
    open that single event's full card (HistoryDetail)."""
    def __init__(self, parent, item, voided=None):
        super().__init__(parent); self._item=item; self._voided=voided or {}
        self.title(f"Shipment — WO {item.get('wo','?')}")
        self.geometry("800x600"); self.transient(parent); self.configure(bg=C_BG)
        self.bind("<Escape>", lambda e: self.destroy())
        self._build(); _center(self, parent)
    def _openf(self, p):
        try:
            if p and Path(str(p)).exists(): os.startfile(str(p))
            else: messagebox.showwarning("Not found", f"Path doesn't exist:\n{p}", parent=self)
        except Exception as e:
            messagebox.showerror("Open failed", str(e), parent=self)
    def _build(self):
        it=self._item; rep=it.get("rep") or {}
        hc={"✓ Closed":C_GREEN,"✎ Draft":C_TEAL,"✕ Deleted":C_RED,"🚫 Voided":C_SUBTLE}.get(it.get("status"),C_GREEN)
        h=tk.Frame(self,bg=hc); h.pack(fill="x")
        tk.Label(h,text=f"{it.get('status','')} — WO {it.get('wo','?')}",bg=hc,fg=C_WHITE,font=("Segoe UI",13,"bold")).pack(side="left",padx=20,pady=12)
        bb=tk.Frame(self,bg=C_BG); bb.pack(side="bottom",fill="x",padx=16,pady=10)
        _btn(bb,"Close",self.destroy,bg="#DDD",fg=C_MID).pack(side="right")
        if rep.get("wo_folder"):
            _btn(bb,"Open WO Folder",lambda:self._openf(rep.get("wo_folder")),bg=C_GREEN,fg=C_WHITE).pack(side="left",padx=(0,8))
        info=tk.Frame(self,bg=C_BG); info.pack(fill="x",padx=16,pady=(10,0))
        for lbl,val in [("Eng #",it.get("eng")),("P/N",it.get("pn")),("Customer",it.get("customer")),
                        ("Division",it.get("division")),
                        ("Started",HistoryViewer._byline(it.get("started_by"),it.get("started_at"))),
                        ("Last saved",HistoryViewer._byline(it.get("saved_by"),it.get("saved_at"))),
                        ("Closed",HistoryViewer._byline(it.get("closed_by"),it.get("closed_at"),it.get("station","")))]:
            row=tk.Frame(info,bg=C_BG); row.pack(fill="x",pady=1)
            tk.Label(row,text=str(lbl)+":",bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",9),width=11,anchor="w").pack(side="left")
            tk.Label(row,text=str(val or "—"),bg=C_BG,fg=C_DARK,font=("Segoe UI",10),anchor="w").pack(side="left")
        tk.Label(self,text="Save timeline — double-click a row for its full details:",bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",9,"bold")).pack(anchor="w",padx=16,pady=(12,2))
        tf=tk.Frame(self,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); tf.pack(fill="both",expand=True,padx=16,pady=(0,4))
        sb=ttk.Scrollbar(tf); sb.pack(side="right",fill="y")
        tv=ttk.Treeview(tf,columns=("time","event","by","station","serials"),show="headings",yscrollcommand=sb.set); sb.config(command=tv.yview)
        for c,t,w,a in [("time","Date/Time",150,"w"),("event","Event",120,"w"),("by","By",120,"w"),("station","Station",120,"w"),("serials","S/N",52,"center")]:
            tv.heading(c,text=t); tv.column(c,width=w,anchor=a)
        tv.pack(side="left",fill="both",expand=True)
        self._tl_tv=tv; self._tl_recs={}
        for e in (it.get("timeline") or []):
            et=e.get("event_type") or ""; lab=HistoryViewer.EVENT_LABEL.get(et,et)
            if str(e.get("event_id") or "") in self._voided: lab+=" (VOIDED)"
            sc=e.get("serial_count")
            if sc is None: sc=len(e.get("serials") or [])
            iid=tv.insert("","end",values=(str(e.get("ts") or ""),lab,str(e.get("operator") or ""),str(e.get("station") or ""),sc))
            self._tl_recs[iid]=e
        tv.bind("<Double-1>",self._open_event)
        tv.bind("<Return>",lambda ev:self._open_event(None))
    def _open_event(self,event=None):
        iid=self._tl_tv.identify_row(event.y) if event else (self._tl_tv.selection() or [None])[0]
        e=self._tl_recs.get(iid)
        if e: HistoryDetail(self,e,voided=self._voided.get(str(e.get("event_id") or "")))

class HistoryViewer(tk.Toplevel):
    """Non-modal browser for shipment history records (close / draft / delete)."""
    COLS=("status","wo","eng","pn","customer","started","saved","closed","serials")
    # Glyph prefixes encode the event type by SHAPE as well as color, so rows stay
    # distinguishable for colorblind users, in print-outs and in the CSV export.
    EVENT_LABEL={"close":"✓ Closed","draft_save":"✎ Draft","delete":"✕ Deleted",
                 "receive":"📦 Received","backfill_shipped":"⏪ Imported",
                 "void":"🚫 Void"}
    def __init__(self,parent):
        super().__init__(parent); self.title("Shipment History"); self.geometry("1100x680")
        # NOT transient: a transient Toplevel loses its maximize/minimize buttons on
        # Windows (Dhruvil 2026-10-28 couldn't maximize it). This is a non-modal
        # browser, so it gets a normal resizable/maximizable window + a taskbar entry.
        self.configure(bg=C_BG); self.minsize(900,520)
        self._all=[]; self._shown=[]; self._records={}; self._skipped=0
        self._voided={}   # target_event_id -> the void marker record
        self._sort_col="ts"; self._sort_desc=True   # ts = newest shipment activity
        self._months=self._list_months()
        self._build(); _center(self,parent)
        self.bind("<Escape>",lambda e: self.destroy())
        self.after(80,self._search)
    def _list_months(self):
        """Return sorted YYYY-MM subfolder names found under HISTORY_DIR."""
        months=[]
        try:
            root=Path(HISTORY_DIR)
            if not root.exists(): return []
            for y in root.iterdir():
                if not (y.is_dir() and re.fullmatch(r"\d{4}",y.name)): continue
                for m in y.iterdir():
                    if m.is_dir() and re.fullmatch(r"\d{4}-\d{2}",m.name): months.append(m.name)
        except Exception:
            pass
        return sorted(months)
    def _build(self):
        h=tk.Frame(self,bg=C_GREEN); h.pack(fill="x")
        tk.Label(h,text="Shipment History",bg=C_GREEN,fg=C_WHITE,font=("Segoe UI",13,"bold")).pack(side="left",padx=20,pady=13)
        tk.Label(h,text="Closed shipments · drafts · deletions",bg=C_GREEN,fg="#9FE1CB",font=("Segoe UI",9)).pack(side="left")
        self._status=tk.Label(self,text="",bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",9),anchor="w")
        self._status.pack(side="bottom",fill="x",padx=18,pady=(0,8))
        fc=tk.Frame(self,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); fc.pack(fill="x",padx=16,pady=(12,4))
        fi=tk.Frame(fc,bg=C_WHITE,padx=12,pady=10); fi.pack(fill="x")
        r1=tk.Frame(fi,bg=C_WHITE); r1.pack(fill="x")
        def fld(label,width):
            f=tk.Frame(r1,bg=C_WHITE); f.pack(side="left",padx=(0,12))
            tk.Label(f,text=label,bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",8)).pack(anchor="w")
            v=tk.StringVar(); e=ttk.Entry(f,textvariable=v,font=("Segoe UI",10),width=width); e.pack(ipady=2)
            e.bind("<Return>",lambda ev: self._search()); return v
        self._f_wo=fld("WO contains",11); self._f_eng=fld("Eng# contains",11); self._f_sn=fld("Serial # contains",15); self._f_cust=fld("Customer contains",15)
        def cbo(label,values,width):
            f=tk.Frame(r1,bg=C_WHITE); f.pack(side="left",padx=(0,12))
            tk.Label(f,text=label,bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",8)).pack(anchor="w")
            c=ttk.Combobox(f,values=values,state="readonly",font=("Segoe UI",10),width=width); c.pack(ipady=1); return c
        self._f_op=cbo("Operator",["All"],13); self._f_op.set("All")
        self._f_div=cbo("Division",["All","Industrial","MAG - ITAR"],12); self._f_div.set("All")
        r2=tk.Frame(fi,bg=C_WHITE); r2.pack(fill="x",pady=(10,0))
        tk.Label(r2,text="Show:",bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",9)).pack(side="left")
        self._t_closed=tk.BooleanVar(value=True); self._t_draft=tk.BooleanVar(value=True); self._t_del=tk.BooleanVar(value=True)
        self._t_recv=tk.BooleanVar(value=True)
        for txt,var in [("Closed",self._t_closed),("Drafts",self._t_draft),("Deleted",self._t_del),("Received",self._t_recv)]:
            tk.Checkbutton(r2,text=txt,variable=var,bg=C_WHITE,fg=C_MID,font=("Segoe UI",9),activebackground=C_WHITE,command=self._apply_filters).pack(side="left",padx=(6,0))
        tk.Label(r2,text="From:",bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",9)).pack(side="left",padx=(20,4))
        self._f_from=ttk.Combobox(r2,values=self._months,state="readonly",font=("Segoe UI",10),width=9); self._f_from.pack(side="left",ipady=1)
        tk.Label(r2,text="To:",bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",9)).pack(side="left",padx=(10,4))
        self._f_to=ttk.Combobox(r2,values=self._months,state="readonly",font=("Segoe UI",10),width=9); self._f_to.pack(side="left",ipady=1)
        today=datetime.date.today(); cur=f"{today:%Y-%m}"
        prev=f"{(today.replace(day=1)-datetime.timedelta(days=1)):%Y-%m}"
        if self._months:
            self._f_from.set(prev if prev in self._months else (cur if cur in self._months else self._months[-1]))
            self._f_to.set(cur if cur in self._months else self._months[-1])
        _btn(r2,"Search",self._search,bg=C_GREEN,fg=C_WHITE,font_size=9,padx=18,pady=5).pack(side="left",padx=(20,0))
        _btn(r2,"Export CSV",self._export_csv,bg=C_TEAL,fg=C_WHITE,font_size=9,padx=14,pady=5).pack(side="right")
        tf=tk.Frame(self,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); tf.pack(fill="both",expand=True,padx=16,pady=(6,4))
        sb=ttk.Scrollbar(tf); sb.pack(side="right",fill="y")
        self._tv=ttk.Treeview(tf,columns=self.COLS,show="headings",yscrollcommand=sb.set)
        sb.config(command=self._tv.yview)
        self._col_titles={}
        for c,t,wd,anc,stretch in [("status","Status",92,"w",False),("wo","WO",82,"w",False),("eng","Eng#",70,"w",False),("pn","P/N",120,"w",True),("customer","Customer",120,"w",True),("started","Started (by · when)",150,"w",False),("saved","Last saved (by · when)",150,"w",False),("closed","Closed (by · when · PC)",175,"w",False),("serials","S/N",46,"center",False)]:
            self._col_titles[c]=t
            self._tv.heading(c,text=t,command=lambda x=c: self._sort(x))
            self._tv.column(c,width=wd,anchor=anc,stretch=stretch)
        self._tv.pack(side="left",fill="both",expand=True)
        # Light background tints + neutral text (a fully red row used to read as an
        # ERROR rather than simply "deleted", and colored the whole grid noisy).
        self._tv.tag_configure("close",background=C_WHITE,foreground=C_DARK)
        self._tv.tag_configure("draft_save",background=C_TEAL_LITE,foreground=C_DARK)
        self._tv.tag_configure("delete",background="#FBEAE8",foreground=C_DARK)
        self._tv.tag_configure("receive",background=C_ORANGE_LITE,foreground=C_DARK)
        self._tv.tag_configure("backfill_shipped",background="#F0F0F0",foreground=C_DARK)
        self._tv.tag_configure("void",background="#F0F0F0",foreground=C_SUBTLE)
        self._tv.tag_configure("voided",background="#EDEDED",foreground="#9A9A9A")
        self._tv.bind("<Double-1>",self._detail)
        self._tv.bind("<Return>",lambda e: self._detail(None))   # keyboard opens detail too
        # Filters apply as soon as a choice is made (checkboxes already did).
        self._f_op.bind("<<ComboboxSelected>>",lambda e: self._apply_filters())
        self._f_div.bind("<<ComboboxSelected>>",lambda e: self._apply_filters())
        self._f_from.bind("<<ComboboxSelected>>",lambda e: self._search())
        self._f_to.bind("<<ComboboxSelected>>",lambda e: self._search())
    def _months_selected(self):
        """List of existing YYYY-MM folders between From/To (inclusive)."""
        a,b=self._f_from.get(),self._f_to.get()
        if a and b and a>b: a,b=b,a
        return [m for m in self._months if (not a or m>=a) and (not b or m<=b)]
    def _load_records(self,months):
        """Read every *.json event file in the given month folders, file names
        newest-first. Returns (records, skipped_count); unreadable files are
        counted and skipped, never fatal. Updates the footer while loading."""
        recs=[]; skipped=0; n=0
        for ym in months:
            folder=Path(HISTORY_DIR)/ym[:4]/ym
            try:
                files=sorted((f for f in folder.iterdir() if f.name.lower().endswith(".json")),key=lambda f: f.name,reverse=True)
            except Exception:
                continue
            for f in files:
                try:
                    with open(f,encoding="utf-8") as fh: rec=json.load(fh)
                    if isinstance(rec,dict): recs.append(rec)
                    else: skipped+=1
                except Exception:
                    skipped+=1
                n+=1
                if n%50==0:
                    try:
                        self._status.config(text=f"Loading… {len(recs)} records"); self.update_idletasks()
                    except Exception: pass
        return recs,skipped
    def _search(self):
        """Reload records for the selected month range, then filter."""
        if not self.winfo_exists(): return
        self._months=self._list_months()
        self._f_from["values"]=self._months; self._f_to["values"]=self._months
        if self._months and (not self._f_from.get() or not self._f_to.get()):
            today=datetime.date.today(); cur=f"{today:%Y-%m}"
            prev=f"{(today.replace(day=1)-datetime.timedelta(days=1)):%Y-%m}"
            if not self._f_from.get(): self._f_from.set(prev if prev in self._months else (cur if cur in self._months else self._months[-1]))
            if not self._f_to.get(): self._f_to.set(cur if cur in self._months else self._months[-1])
        months=self._months_selected()
        self._all=[]; self._skipped=0
        if months:
            self._status.config(text="Loading…"); self.update_idletasks()
            self._all,self._skipped=self._load_records(months)
        # Map voided targets so their rows can be grayed and their detail banner shown.
        self._voided={}
        for r in self._all:
            if isinstance(r,dict) and r.get("event_type")=="void":
                t=str(r.get("target_event_id") or "")
                if t: self._voided[t]=r
        ops=sorted({str(r.get("operator")) for r in self._all if r.get("operator")})
        cur=self._f_op.get(); self._f_op["values"]=["All"]+ops
        if cur not in ["All"]+ops: self._f_op.set("All")
        self._apply_filters()
        if not self._months: self._status.config(text="No history yet")
    def _apply_filters(self):
        """Group loaded events into ONE item per shipment (Dhruvil 2026-10-24),
        filter by the filter-bar values, and repopulate."""
        woq=self._f_wo.get().strip().lower(); enq=self._f_eng.get().strip().lower()
        snq=self._f_sn.get().strip().lower(); cuq=self._f_cust.get().strip().lower()
        opq=self._f_op.get(); dvq=self._f_div.get()
        show_closed=self._t_closed.get(); show_draft=self._t_draft.get()
        show_del=self._t_del.get(); show_recv=self._t_recv.get()
        items=group_shipment_history(self._all,{str(k) for k in self._voided})
        shown=[]
        for it in items:
            status=it.get("status","")
            if status=="✓ Closed" and not show_closed: continue
            if status=="✎ Draft" and not show_draft: continue
            if status in ("✕ Deleted","🚫 Voided","🚫 Void") and not show_del: continue
            if status in ("📦 Received","⏪ Imported") and not show_recv: continue
            if woq and woq not in it.get("wo","").lower(): continue
            if enq and enq not in it.get("eng","").lower(): continue
            if cuq and cuq not in it.get("customer","").lower(): continue
            if opq not in ("","All") and opq not in (it.get("operator",""),it.get("closed_by",""),it.get("saved_by",""),it.get("started_by","")): continue
            if dvq not in ("","All") and it.get("division","")!=dvq: continue
            if snq and not any(snq in str(s).lower() for e in it.get("timeline",[]) for s in (e.get("serials") or [])): continue
            shown.append(it)
        self._shown=shown
        self._shown.sort(key=lambda it: self._sort_key(it,self._sort_col),reverse=self._sort_desc)
        self._populate()
    def _rec_time(self,r):
        """Best-effort event time as epoch float (ts_epoch, else parse ts)."""
        return _hist_event_time(r)
    @staticmethod
    def _byline(by,when,extra=""):
        when=str(when or "")[:16]
        if not by and not when and not extra: return "—"
        s=str(by or "")
        if when: s=(s+"  ·  "+when) if s else when
        if extra: s=(s+"  ·  "+extra) if s else extra
        return s
    def _row_vals(self,it):
        """Display tuple for one grouped shipment / standalone event."""
        return (it.get("status",""),it.get("wo",""),it.get("eng",""),it.get("pn",""),
                it.get("customer",""),
                self._byline(it.get("started_by"),it.get("started_at")),
                self._byline(it.get("saved_by"),it.get("saved_at")),
                self._byline(it.get("closed_by"),it.get("closed_at"),it.get("station","")),
                it.get("serial_count",""))
    def _sort_key(self,it,col):
        """Sort key for one item (default 'ts' = newest shipment activity)."""
        if col=="ts": return it.get("ts",0.0)
        if col=="started": return str(it.get("started_at") or "")
        if col=="saved": return str(it.get("saved_at") or "")
        if col=="closed": return str(it.get("closed_at") or "")
        try: v=self._row_vals(it)[self.COLS.index(col)]
        except Exception: return it.get("ts",0.0)
        if col=="serials":
            try: return float(v)
            except Exception: return -1.0
        return str(v).lower()
    def _sort(self,col):
        """Toggle asc/desc sort on a column header click."""
        if col==self._sort_col: self._sort_desc=not self._sort_desc
        else: self._sort_col=col; self._sort_desc=(col in ("started","saved","closed"))
        self._shown.sort(key=lambda it: self._sort_key(it,col),reverse=self._sort_desc)
        self._populate()
    def _populate(self):
        """Refill the tree from self._shown (one row per shipment) + footer count."""
        for iid in self._tv.get_children(): self._tv.delete(iid)
        self._records={}
        _tag={"✓ Closed":"close","✎ Draft":"draft_save","✕ Deleted":"delete",
              "🚫 Voided":"voided","🚫 Void":"void","📦 Received":"receive","⏪ Imported":"backfill_shipped"}
        for it in self._shown:
            iid=self._tv.insert("","end",values=self._row_vals(it),tags=(_tag.get(it.get("status"),""),))
            self._records[iid]=it
        for c,t in self._col_titles.items():
            self._tv.heading(c,text=t+("  ▼" if self._sort_desc else "  ▲") if c==self._sort_col else t)
        n_ship=sum(1 for it in self._shown if it.get("kind")=="shipment")
        self._status.config(text=f"{len(self._shown)} rows · {n_ship} shipment(s) ({self._skipped} files skipped)  —  double-click for the full timeline")
    def _detail(self,event=None):
        """Open the full timeline (a grouped shipment) or the single-event card."""
        iid=self._tv.identify_row(event.y) if event else (self._tv.selection() or [None])[0]
        if not iid: return
        it=self._records.get(iid)
        if not it: return
        if it.get("kind")=="shipment":
            ShipmentDetail(self,it,self._voided)
        else:
            rec=it.get("rec") or {}
            HistoryDetail(self,rec,voided=self._voided.get(str(rec.get("event_id") or "")))
    def _export_csv(self):
        """Write the currently filtered rows to a CSV chosen by the user."""
        if not self._shown:
            messagebox.showwarning("Nothing to export","No records to export — run a search first.",parent=self); return
        fym=self._f_from.get() or "all"; tym=self._f_to.get() or "all"
        path=filedialog.asksaveasfilename(parent=self,defaultextension=".csv",filetypes=[("CSV files","*.csv")],initialfile=f"shipment_history_{fym}_{tym}.csv")
        if not path: return
        coc_keys=["ship_date","po","pn","rev","qty","datecode","job","type","assembly_process"]
        pk_keys=["customer","packing_no","po","pn","wo","eng","date","datecode"]
        hdr=(["date_time","type","wo","division","operator","customer","pn","eng","station","win_user","app_version","event_id","status","serial_count","serials","photos_moved","photos_missing","photos_failed"]
             +[f"coc_{k}" for k in coc_keys]+[f"packing_{k}" for k in pk_keys]
             +["out_serials_xlsx","out_serials_docx","out_coc_xlsx","out_packing_xlsx","errors","wo_folder","resumed_saved_by","resumed_saved_at","notes","dup_override_count"])
        try:
            with open(path,"w",newline="",encoding="utf-8-sig") as f:
                w=csv.writer(f); w.writerow(hdr)
                for r in (e for it in self._shown for e in (it.get("timeline") or [it.get("rep") or {}])):
                    et=r.get("event_type") or ""
                    coc=r.get("coc") or {}; pk=r.get("packing") or {}
                    ph=r.get("photos") or {}; out=r.get("outputs") or {}
                    rd=r.get("resumed_from_draft") or {}
                    sc=r.get("serial_count")
                    if sc is None: sc=len(r.get("serials") or [])
                    st=("OK" if r.get("critical_ok") else "ERRORS") if et=="close" else ""
                    row=([str(r.get("ts") or ""),HistoryDetail.LABELS.get(et,et),str(r.get("wo") or ""),str(r.get("division") or ""),str(r.get("operator") or ""),str(r.get("customer") or ""),str(r.get("pn") or ""),str(r.get("eng") or ""),str(r.get("station") or ""),str(r.get("win_user") or ""),str(r.get("app_version") or ""),str(r.get("event_id") or ""),st,sc,";".join(str(s) for s in (r.get("serials") or [])),ph.get("moved_count",0),len(ph.get("missing") or []),len(ph.get("failed") or [])]
                         +[str(coc.get(k) or "") for k in coc_keys]+[str(pk.get(k) or "") for k in pk_keys]
                         +[str(out.get("serials_xlsx") or ""),str(out.get("serials_docx") or ""),str(out.get("coc_xlsx") or ""),str(out.get("packing_xlsx") or ""),"; ".join(str(e).strip() for e in (r.get("errors") or [])),str(r.get("wo_folder") or ""),str(rd.get("saved_by") or ""),str(rd.get("saved_at") or ""),str(r.get("notes") or "").replace("\n"," ").replace("\r"," "),len(r.get("dup_overrides") or [])])
                    w.writerow(row)
            messagebox.showinfo("Export complete",f"Exported to:\n\n{path}",parent=self)
        except Exception as e:
            messagebox.showerror("Export failed",f"Could not write CSV:\n\n{type(e).__name__}: {e}",parent=self)

###############################################################################
# RECEIVING UI — window + dialogs
###############################################################################

class ReceiveIssueDialog(tk.Toplevel):
    """Scanner-proof blocking alert for a problem scan at Receiving (already
    received / doesn't match the job P/N). Per Dhruvil: NOTHING is a hard block —
    every stop offers 'Supervisor override…' (admin login handled by the caller).
    Buttons stay disabled 750ms; Enter/Escape only Skip. result: 'skip'|'override'."""
    def __init__(self, parent, serial, heading, message):
        super().__init__(parent); self.title("Check this board")
        self.resizable(False, False); self.transient(parent); self.grab_set()
        self.configure(bg=C_BG); self.result = "skip"; self.focus_set()
        self.protocol("WM_DELETE_WINDOW", self._skip)
        h = tk.Frame(self, bg=C_AMBER); h.pack(fill="x")
        tk.Label(h, text="⚠  " + heading, bg=C_AMBER, fg=C_WHITE,
                 font=("Segoe UI", 12, "bold")).pack(padx=24, pady=13, anchor="w")
        body = tk.Frame(self, bg=C_BG, padx=24, pady=16); body.pack(fill="x")
        tk.Label(body, text=str(serial), bg=C_BG, fg=C_DARK,
                 font=("Courier New", 14, "bold")).pack(anchor="w")
        tk.Label(body, text=message + "\n\nIt was NOT added. A supervisor can override "
                 "to record it anyway.", bg=C_BG, fg=C_MID, font=("Segoe UI", 10),
                 wraplength=430, justify="left").pack(anchor="w", pady=(8, 0))
        br = tk.Frame(self, bg=C_BG, pady=12); br.pack()
        self._skip_btn = _btn(br, "Skip — don't add", self._skip, bg=C_GREEN, fg=C_WHITE,
                              font_size=11, padx=20, pady=9)
        self._skip_btn.pack(side="left", padx=(0, 10))
        self._ovr_btn = _btn(br, "Supervisor override…", self._override, bg=C_AMBER,
                             fg=C_WHITE, font_size=11, padx=20, pady=9)
        self._ovr_btn.pack(side="left")
        self._skip_btn.config(state="disabled"); self._ovr_btn.config(state="disabled")
        self._armed = False; self.after(750, self._arm)
        self.bind("<Return>", lambda e: self._skip())
        self.bind("<Escape>", lambda e: self._skip())
        self.bell(); _center(self, parent); self.wait_window()
    def _arm(self):
        self._armed = True
        for b in (self._skip_btn, self._ovr_btn):
            try: b.config(state="normal")
            except Exception: pass
        try: self._skip_btn.focus_set()
        except Exception: pass
    def _skip(self):
        if self._armed: self.result = "skip"; self.destroy()
    def _override(self):
        if self._armed: self.result = "override"; self.destroy()

class _BusyDialog(tk.Toplevel):
    """Tiny modal 'please wait' box; the CALLER pumps events and destroys it."""
    def __init__(self, parent, text):
        super().__init__(parent); self.title("Please wait"); self.resizable(False, False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG)
        self.protocol("WM_DELETE_WINDOW", lambda: None)
        tk.Label(self, text=text, bg=C_BG, fg=C_DARK,
                 font=("Segoe UI", 11)).pack(padx=34, pady=22)
        _center(self, parent)

class ProgressDialog(tk.Toplevel):
    """Modal progress box for LONG SYNCHRONOUS work on the UI thread — moving
    photos to the share and writing the documents can take 30s+, and the window
    used to simply freeze with no feedback. step(text) updates the status line
    AND pumps events so the box stays painted; the grab swallows stray clicks and
    scanner bursts meanwhile. There is no close button — the operation owns the
    box, and callers close it in try/finally so it can never be left stuck."""
    def __init__(self, parent, title, heading):
        super().__init__(parent); self.title(title); self.resizable(False, False)
        self.transient(parent); self.configure(bg=C_BG)
        self.protocol("WM_DELETE_WINDOW", lambda: None)   # not closable by hand
        h = tk.Frame(self, bg=C_TEAL); h.pack(fill="x")
        tk.Label(h, text="⏳  " + heading, bg=C_TEAL, fg=C_WHITE,
                 font=("Segoe UI", 12, "bold")).pack(padx=24, pady=12, anchor="w")
        self._lbl = tk.Label(self, text="Starting…", bg=C_BG, fg=C_MID,
                             font=("Segoe UI", 10), padx=26, pady=16, anchor="w",
                             justify="left", width=54)
        self._lbl.pack(fill="x")
        try: self.grab_set()
        except Exception: pass
        _center(self, parent)
        try: self.update()
        except Exception: pass
    def step(self, text):
        """Update the status line, re-assert the grab (a child dialog may have
        taken it), and pump events so the UI repaints between slow file steps."""
        try:
            self._lbl.config(text=text)
            self.grab_set()
            self.update()
        except Exception:
            pass
    def regrab(self):
        try: self.grab_set()
        except Exception: pass
    def close(self):
        try: self.grab_release()
        except Exception: pass
        try: self.destroy()
        except Exception: pass

def _picker_default_ticks(rows):
    """ReportPickerDialog's opening tick state (Dhruvil 2026-09-23, OPT-IN): a
    row is pre-ticked only when its job already HAS a report AND can be rebuilt
    (tickable) — an existing report keeps updating anyway, so re-ticking it only
    adds folder truth; a job without one is never ticked for the operator. Tk-free
    so the rule is testable (review round 2)."""
    return [bool(r.get("has")) and bool(r.get("tickable")) for r in rows]

def _picker_set(rows, ticked, i, val):
    """Tick / untick row i in place; an un-tickable row (no live -900/-970 WO
    folder, unreadable history, folders not readable from this PC) is refused —
    returns True when the state changed."""
    if not (0 <= i < len(rows)) or not rows[i].get("tickable"):
        return False
    ticked[i] = bool(val)
    return True

def _picker_picked(rows, ticked):
    """The rows 'Rebuild selected' hands to _rebuild_report_rows: ticked AND
    tickable — never anything else, whatever the tick list says ("create reports
    only for those user selects")."""
    return [r for i, r in enumerate(rows) if ticked[i] and r.get("tickable")]

class ReportPickerDialog(tk.Toplevel):
    """🔁 Rebuild ALL reports → job picker (Dhruvil 2026-09-23, OPT-IN: "show all
    the jobs first but create reports only for those user selects"). One row per
    job from list_testing_jobs (rows come pre-computed by _collect_report_rows on
    a worker thread): tick · Eng# · P/N (coated-only tagged) · last shipment ·
    status. Pre-ticked = already has a report (those keep updating anyway; a
    rebuild adds folder truth). Rows with no work-order folder or unreadable
    history cannot be ticked. Click / Space toggle a row; Enter is INERT (a
    barcode scan's trailing Enter must not flip a tick — review round 2). Buttons
    arm after 400ms (scanner-proof); Escape / titlebar close = Cancel. The tick
    rules live in _picker_default_ticks / _picker_set / _picker_picked (Tk-free).
    self.result: list of ticked rows, or None."""
    def __init__(self, parent, rows):
        super().__init__(parent); self.title("Rebuild ALL reports — pick the jobs")
        self.resizable(True, True); self.transient(parent); self.grab_set(); self.configure(bg=C_BG)
        self.result = None; self._rows = list(rows)
        self._ticked = _picker_default_ticks(self._rows)
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        h = tk.Frame(self, bg=C_TEAL); h.pack(fill="x")
        tk.Label(h, text=f"🔁  Rebuild ALL reports — {len(self._rows)} job(s) on record", bg=C_TEAL,
                 fg=C_WHITE, font=("Segoe UI", 12, "bold")).pack(padx=22, pady=12, anchor="w")
        bd = tk.Frame(self, bg=C_BG, padx=22, pady=12); bd.pack(fill="both", expand=True)
        tk.Label(bd, text=("Tick the jobs whose Shipping Report should be (re)built from shipment history "
                           "+ the job's COC / Packing List / Shipping Serials folders. Jobs that already "
                           "have a report are pre-ticked. A ticked job WITHOUT a report gets one for the "
                           "first time; unticked jobs are left alone — no file, no folder."),
                 bg=C_BG, fg=C_MID, font=("Segoe UI", 9), wraplength=720, justify="left").pack(anchor="w", pady=(0, 8))
        lf = tk.Frame(bd, bg=C_WHITE, highlightbackground=C_BORDER, highlightthickness=1)
        lf.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(lf); sb.pack(side="right", fill="y")
        cols = ("eng", "pn", "last", "status")
        self._tv = ttk.Treeview(lf, columns=cols, show="tree headings", selectmode="browse",
                                height=min(18, max(6, len(self._rows))), yscrollcommand=sb.set)
        self._tv.pack(side="left", fill="both", expand=True)
        sb.config(command=self._tv.yview)
        self._tv.heading("#0", text="", anchor="center"); self._tv.column("#0", width=44, stretch=False, anchor="center")
        self._tv.heading("eng", text="Eng #"); self._tv.column("eng", width=90, stretch=False, anchor="w")
        self._tv.heading("pn", text="P/N"); self._tv.column("pn", width=330, anchor="w")
        self._tv.heading("last", text="Last shipment"); self._tv.column("last", width=110, stretch=False, anchor="center")
        # wide enough for 'report found — no live -900/-970 work-order folder' (~270 px)
        self._tv.heading("status", text="Report"); self._tv.column("status", width=290, stretch=False, anchor="w")
        try:
            self._tv.tag_configure("off", foreground="#9A9A9A")
            self._tv.tag_configure("coat", foreground=C_TEAL)
        except Exception: pass
        for i, r in enumerate(self._rows):
            j = r.get("job") or {}
            pn = _job_display_pn(j)
            if r.get("coated_only"): pn += _COATED_ONLY_TAG
            tags = () if r.get("tickable") else ("off",)
            if r.get("coated_only") and r.get("tickable"): tags = ("coat",)
            self._tv.insert("", "end", iid=str(i), text=self._mark(i),
                            values=(j.get("eng", ""), pn, r.get("last_ts") or "", r.get("status") or ""),
                            tags=tags)
        self._tv.bind("<Button-1>", self._click)
        self._tv.bind("<space>", lambda e: self._toggle_selected())
        # no <Return> binding: a clicked row stays SELECTED, and the Enter a barcode
        # scanner sends after every scan would silently flip its tick (review round 2)
        # wraps at the Treeview's width: the count sentence with the un-tickable
        # clause is ~175 chars and would otherwise set the dialog's width
        self._count = tk.Label(bd, text="", bg=C_BG, fg=C_SUBTLE, font=("Segoe UI", 9), anchor="w",
                               wraplength=860, justify="left")
        self._count.pack(fill="x", pady=(6, 0))
        sel = tk.Frame(bd, bg=C_BG); sel.pack(fill="x", pady=(6, 0))
        for txt, fn in (("Select all", self._sel_all), ("Select none", self._sel_none),
                        ("Select with report", self._sel_with)):
            _btn(sel, txt, fn, bg="#DDD", fg=C_MID, font_size=9, padx=10, pady=5).pack(side="left", padx=(0, 6))
        br = tk.Frame(self, bg="#E8E8E8", pady=12); br.pack(fill="x")
        self._go_btn = _btn(br, "Rebuild selected (0)", self._go, bg=C_TEAL, fg=C_WHITE,
                            font_size=10, padx=16, pady=9)
        self._go_btn.pack(side="left", padx=(22, 8))
        self._cancel_btn = _btn(br, "Cancel", self._cancel, bg="#DDD", fg=C_MID, font_size=10,
                                padx=16, pady=9)
        self._cancel_btn.pack(side="right", padx=(0, 22))
        # Scanner-proof: a trailing Enter from a barcode scan can't fire the rebuild.
        self._go_btn.config(state="disabled")
        self._armed = False
        self.after(400, self._arm)
        self.bind("<Escape>", lambda e: self._cancel())
        self._refresh_count()
        _center(self, parent); self._tv.focus_set(); self.wait_window()
    def _mark(self, i):
        if not self._rows[i].get("tickable"): return "—"
        return "☑" if self._ticked[i] else "☐"
    def _arm(self):
        self._armed = True; self._refresh_count()
    def _refresh_count(self):
        n = sum(1 for t in self._ticked if t)
        # (an Archive-only copy is re-homed, not created — review round 4)
        new = sum(1 for i, t in enumerate(self._ticked)
                  if t and not self._rows[i].get("has") and not self._rows[i].get("found"))
        off = sum(1 for r in self._rows if not r.get("tickable"))
        try:
            self._count.config(text=(f"{n} ticked" + (f"  ·  {new} of them have no report yet (will be created)" if new else "")
                                     + (f"  ·  {off} job(s) can't be ticked (no -900/-970 work-order folder found, "
                                        f"history unreadable, or folders not readable from this PC)" if off else "")))
            self._go_btn.config(text=f"Rebuild selected ({n})",
                                state=("normal" if (self._armed and n) else "disabled"))
        except Exception: pass
    def _set(self, i, val):
        if not _picker_set(self._rows, self._ticked, i, val): return
        try: self._tv.item(str(i), text=self._mark(i))
        except Exception: pass
    def _click(self, event):
        iid = self._tv.identify_row(event.y)
        if not iid: return
        try: i = int(iid)
        except Exception: return
        if 0 <= i < len(self._rows):
            self._set(i, not self._ticked[i]); self._refresh_count()
    def _toggle_selected(self):
        sel = self._tv.selection()
        if not sel: return
        try: i = int(sel[0])
        except Exception: return
        if 0 <= i < len(self._rows):
            self._set(i, not self._ticked[i]); self._refresh_count()
    def _sel_all(self):
        for i in range(len(self._rows)): self._set(i, True)
        self._refresh_count()
    def _sel_none(self):
        for i in range(len(self._rows)): self._set(i, False)
        self._refresh_count()
    def _sel_with(self):
        for i, r in enumerate(self._rows): self._set(i, bool(r.get("has")))
        self._refresh_count()
    def _go(self):
        if not self._armed: return
        picked = _picker_picked(self._rows, self._ticked)
        if not picked: return
        self.result = picked; self.destroy()
    def _cancel(self):
        self.result = None; self.destroy()

class BackfillDialog(tk.Toplevel):
    """Import older testing shipments (boards shipped before History existed) by
    reading every WO folder's 'Shipping Serials' files for one job. Writes ONE
    'backfill_shipped' History event per WO folder; serials already known to the
    job are skipped, so re-running is harmless (idempotent)."""
    def __init__(self, parent, app, state=None):
        super().__init__(parent); self._app = app; self._state = state
        self.title("Import older testing shipments"); self.resizable(False, False)
        self.transient(parent); self.grab_set(); self.configure(bg=C_BG)
        self.imported = 0
        self._pd = None; self._scan = None   # scan = [(wo, serials, ship_date, files)]
        self._known = set()                  # fresh already-recorded set (authoritative)
        self._known_ok = False               # False -> couldn't verify; Import disabled
        self._scan_key = None                # (eng,pn) the scan/known were built for
        self._coated_scan = False            # scan found ONLY sibling (…970) serials
        self._typed_coated = False           # the P/N typed/auto-filled is the coated side
        h = tk.Frame(self, bg=C_TEAL); h.pack(fill="x")
        tk.Label(h, text="⏪  Import older testing shipments", bg=C_TEAL, fg=C_WHITE,
                 font=("Segoe UI", 12, "bold")).pack(padx=22, pady=12, anchor="w")
        bd = tk.Frame(self, bg=C_BG, padx=24, pady=14); bd.pack(fill="both", expand=True)
        tk.Label(bd, text=("Boards shipped for testing BEFORE the app started recording history "
                 "aren't in the tracker yet. Type any of this job's testing WO numbers — the app "
                 "finds the job folder and reads every WO's serials files to fill in the baseline."),
                 bg=C_BG, fg=C_MID, font=("Segoe UI", 9), wraplength=470, justify="left").pack(anchor="w")
        r1 = tk.Frame(bd, bg=C_BG); r1.pack(fill="x", pady=(10, 2))
        tk.Label(r1, text="Testing WO #:", bg=C_BG, fg=C_MID, font=("Segoe UI", 10)).pack(side="left")
        self._wo = tk.StringVar()
        we = ttk.Entry(r1, textvariable=self._wo, font=("Segoe UI", 12), width=16)
        we.pack(side="left", padx=8, ipady=3); we.bind("<Return>", lambda e: self._find())
        _btn(r1, "Find job folder", self._find, bg=C_GREEN, fg=C_WHITE, font_size=9,
             padx=14, pady=6).pack(side="left")
        r2 = tk.Frame(bd, bg=C_BG); r2.pack(fill="x", pady=2)
        tk.Label(r2, text="Eng #:", bg=C_BG, fg=C_MID, font=("Segoe UI", 10)).pack(side="left")
        self._eng = tk.StringVar()
        ttk.Entry(r2, textvariable=self._eng, font=("Segoe UI", 11), width=12).pack(side="left", padx=(8, 16), ipady=2)
        tk.Label(r2, text="Testing P/N:", bg=C_BG, fg=C_MID, font=("Segoe UI", 10)).pack(side="left")
        self._pn = tk.StringVar()
        ttk.Entry(r2, textvariable=self._pn, font=("Segoe UI", 11), width=18).pack(side="left", padx=8, ipady=2)
        self._status = tk.Label(bd, text="", bg=C_BG, fg=C_SUBTLE, font=("Segoe UI", 9),
                                wraplength=470, justify="left")
        self._status.pack(anchor="w", pady=(6, 4))
        lf = tk.Frame(bd, bg=C_WHITE, highlightbackground=C_BORDER, highlightthickness=1)
        lf.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(lf); sb.pack(side="right", fill="y")
        self._lb = tk.Listbox(lf, font=("Courier New", 10), height=9, bg=C_WHITE, fg=C_MID,
                              relief="flat", bd=0, activestyle="none", yscrollcommand=sb.set)
        self._lb.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        sb.config(command=self._lb.yview)
        br = tk.Frame(self, bg="#E8E8E8", pady=12); br.pack(fill="x")
        self._imp_btn = _btn(br, "⏪  Import these shipments", self._import, bg=C_TEAL,
                             fg=C_WHITE, font_size=10, padx=16, pady=8)
        self._imp_btn.pack(side="left", padx=(22, 0)); self._imp_btn.config(state="disabled")
        _btn(br, "Close", self.destroy, bg="#DDD", fg=C_MID, font_size=10,
             padx=16, pady=8).pack(side="right", padx=(0, 22))
        self.bind("<Escape>", lambda e: self.destroy())
        if state and state.get("pn"):
            self._eng.set(state.get("eng", "")); self._pn.set(state.get("pn", ""))
            if state.get("pd_path"):
                self._pd = Path(state["pd_path"])
                self._status.config(text=f"Job folder: {self._pd}\nClick 'Find job folder' to re-detect, "
                                    "or scan folders directly below.")
                self._scan_folders()
        _center(self, parent); we.focus_set(); self.wait_window()
    def _find(self):
        wo = self._wo.get().strip()
        if not wo:
            messagebox.showwarning("No WO", "Type one of this job's testing WO numbers first.", parent=self); return
        div = self._app.division_var.get()
        sd = SearchingDialog(self, wo, div); self.update()
        stats = {}
        try: m = search_wo_folder(div, wo, progress=self.update, stats=stats)
        finally:
            sd.destroy()
            # SearchingDialog took the grab; Tk grabs don't stack, so destroying it
            # released OURS too. Re-grab on EVERY path (incl. "Not found") or this
            # modal dialog goes non-modal and the window behind accepts clicks.
            self.grab_set()
        if not m:
            messagebox.showerror("Not found", f"No WO folder matching '{wo}' found under "
                                 f"{PRODUCTION_ROOTS[div]}\n(searched {stats.get('scanned', 0)} job folders).",
                                 parent=self); return
        wf = m[0] if len(m) == 1 else (WOPickerDialog(self, m).result)
        self.grab_set()
        if not wf: return
        self._pd = wf.parent
        info = parse_wo_path(wf)
        if not self._eng.get().strip() and info.get("eng") not in ("", "N/A"): self._eng.set(info["eng"])
        if not self._pn.get().strip() and info.get("pn") not in ("", "N/A"): self._pn.set(info["pn"])
        self._scan_folders()
    def _refresh_known(self, eng, pn):
        """Fresh, authoritative 'already recorded' set for exactly (eng, pn) — the
        window's cached state may be stale or None, which used to make every serial
        look 'new' and break the promised idempotency."""
        self._known = set(); self._known_ok = False
        self._scan_key = (_norm_key(eng), _norm_key(pn))
        if not (_norm_key(eng) and _norm_key(pn)):
            return
        import threading
        box = {}
        def work():
            try: box["state"] = build_receiving_state(eng, pn)
            except Exception: box["state"] = None
        th = threading.Thread(target=work, daemon=True); th.start()
        busy = _BusyDialog(self, "Checking what's already recorded…")
        start = time.monotonic()
        try:
            while th.is_alive() and (time.monotonic() - start) < 60.0:
                try: self.update()
                except Exception: break
                time.sleep(0.05)
        finally:
            try: busy.destroy()
            except Exception: pass
            self.grab_set()
        stn = box.get("state")
        if stn is not None and not stn.get("walk_error"):
            self._known = set(stn.get("shipped") or {}); self._known_ok = True
    def _scan_folders(self):
        pn = self._pn.get().strip(); eng = self._eng.get().strip()
        # a previous scan's ⚠ colour / coated verdict must not linger — reset
        # before the early return too (review round 6)
        self._coated_scan = False; self._typed_coated = False
        try: self._status.config(fg=C_SUBTLE)
        except Exception: pass
        if not self._pd or not pn:
            self._status.config(text="Fill in the Testing P/N, then find the job folder."); return
        self._refresh_known(eng, pn)
        busy = _BusyDialog(self, "Reading serials files from the WO folders…"); self.update()
        scan = []
        try:
            try: wos = [d for d in sorted(self._pd.iterdir()) if d.is_dir()
                        and not _is_report_folder(d.name)]
            except Exception: wos = []
            for d in wos:
                try: self.update()
                except Exception: pass
                serials, sdate, files = parse_serials_from_shipping_files(d, pn)
                if serials:
                    scan.append((d.name, serials, sdate, files))
        finally:
            busy.destroy(); self.grab_set()
        self._scan = scan
        known = self._known
        self._lb.delete(0, "end")
        total = new = 0
        for wo, serials, sdate, files in scan:
            fresh = [s for s in serials if _serial_identity(s, pn) not in known]
            total += len(serials); new += len(fresh)
            self._lb.insert("end", f" WO {wo}:  {len(serials)} serial(s) in {files} file(s)"
                            + (f", {len(fresh)} new" if len(fresh) != len(serials) else "")
                            + (f"   shipped {sdate}" if sdate else ""))
        if not scan:
            self._status.config(text="No serials files found in this job's WO folders.")
            self._imp_btn.config(state="disabled"); return
        if not self._known_ok:
            self._status.config(text=f"Job folder: {self._pd}\nFound {total} serial(s) across "
                                f"{len(scan)} WO(s) — but the history check FAILED, so what's "
                                f"already recorded couldn't be verified. Fix Eng #/P/N (or retry) "
                                f"and scan again before importing.")
            self._imp_btn.config(state="disabled"); return
        # Sibling-P/N guard (Dhruvil 2026-09-23): every serial found carries the
        # SIBLING P/N of the one typed. A -900 typed with …970 serials is a coated
        # WO, not a testing shipment (a -900-labelled import of -970 serials is the
        # one History shape that plants a coated folder as a TESTING anchor); a
        # coated label typed (auto-filled from a -970 WO folder) with …900-form
        # serials is the in-progress batch, which is then recorded as COATED
        # shipments (review round 2). Warn here and confirm again at Import.
        self._coated_scan = _backfill_serials_look_coated(pn, [s for _, ss, _, _ in scan for s in ss])
        self._typed_coated = _tr_pn_parts(pn)[1] == "970"
        if not self._coated_scan:
            warn = ""
        elif self._typed_coated:
            warn = ("\n⚠ These serials carry the TESTING (…900) P/N, but the P/N typed above is the "
                    "COATED one — under that label they are recorded as coated shipments (Note: "
                    "scanned with a -900 testing label), not as shipped for testing.")
        else:
            warn = ("\n⚠ These serials carry the COATED (…970) P/N, not the testing P/N typed above. "
                    "Coated shipments are recorded automatically by the Shipping App; importing them "
                    "here would label them as testing shipments.")
        self._status.config(text=f"Job folder: {self._pd}\nFound {total} serial(s) across "
                            f"{len(scan)} WO(s) — {new} not yet in the tracker." + warn,
                            fg=(C_ORANGE if self._coated_scan else C_SUBTLE))
        self._imp_btn.config(state="normal" if new else "disabled")
    def _import(self):
        eng, pn = self._eng.get().strip(), self._pn.get().strip()
        if not eng or not pn or not self._scan or not self._pd:
            messagebox.showwarning("Missing info", "Eng #, Testing P/N and a folder scan are needed first.",
                                   parent=self); return
        # Eng#/P/N edited AFTER the folders were scanned -> the preview/known-set
        # describe different values; re-scan with the new ones instead of importing
        # a stale serial set under freshly typed identifiers.
        if self._scan_key != (_norm_key(eng), _norm_key(pn)) or not self._known_ok:
            messagebox.showinfo("Re-scan needed",
                "The Eng # / P/N changed after the folders were scanned (or the history "
                "check didn't complete). Re-scanning with the current values — review "
                "the preview, then click Import again.", parent=self)
            self._scan_folders(); return
        known = set(self._known)
        per_wo = []
        for wo, serials, sdate, files in self._scan:
            fresh = [s for s in serials if _serial_identity(s, pn) not in known]
            if fresh: per_wo.append((wo, fresh, sdate))
        n = sum(len(s) for _, s, _ in per_wo)
        if not n:
            messagebox.showinfo("Nothing new", "Every serial found is already in the tracker.", parent=self); return
        typed_coated = self._typed_coated
        if self._coated_scan:
            if typed_coated:
                ok = messagebox.askyesno("Testing-form serials?",
                    "The scanned serials carry the testing (…900) P/N but the P/N typed above is the "
                    "coated one — they will be recorded as COATED shipments under that label. "
                    "Import them anyway?", parent=self)
            else:
                ok = messagebox.askyesno("Coated serials?",
                    "The scanned serials carry the coated (…970) P/N — are you sure they were "
                    "shipped FOR TESTING?", parent=self)
            self.grab_set()
            if not ok: return
        c = ConfirmActionDialog(self, title="Import baseline",
            heading="Import these older shipments?",
            message=(f"Eng # {eng}  ·  {'Coated' if typed_coated else 'Testing'} P/N {pn}\n\n"
                     f"{n} serial(s) across {len(per_wo)} WO(s) will be "
                     + ("recorded as COATED shipments (this P/N is the coated side)"
                        if typed_coated else "recorded as SHIPPED FOR TESTING")
                     + " (baseline only — no documents are generated, nothing "
                     f"is marked received).\n\nRe-running this import later is safe: already-imported serials "
                     f"are skipped."),
            confirm_label="Import", cancel_label="Cancel",
            heading_color=C_TEAL, confirm_color=C_TEAL)
        self.grab_set()
        if not c.result: return
        div = _division_of_path(self._pd)
        ok = 0
        for wo, serials, sdate in per_wo:
            p = write_history_event("backfill_shipped", {
                "wo": wo, "division": div, "eng": eng, "pn": pn,
                "operator": (self._app.operator_var.get() if not self._app.operator_var.get().startswith("—") else ""),
                "serial_count": len(serials), "serials": list(serials),
                "ship_date": sdate, "production_data": str(self._pd)})
            if p: ok += len(serials)
        self.imported = ok
        # the closing line follows the typed side too — the confirm just said
        # COATED shipments, the summary must not say 'shipped-for-testing' (review round 3)
        messagebox.showinfo("Imported", f"{ok} serial(s) imported as "
                            + ("COATED shipments (this P/N is the coated side)." if typed_coated
                               else "the shipped-for-testing baseline."), parent=self)
        self.destroy()

class ReceivingWindow(tk.Toplevel):
    """📦 Receive boards returned from customer testing. Scan the TESTING labels
    (the board comes back wearing the label it shipped with); each save writes a
    'receive' History event (source of truth) and, when the job already has a
    Shipping Report (OPT-IN, Dhruvil 2026-09-23), refreshes that one file in its
    work-order folder's 'Shipping Reports\\' subfolder (renamed 2026-09-15 from
    'Received after testing'). Non-modal, one instance (like HistoryViewer).
    In-house testing (Dhruvil 2026-09-23): the customer's test fixture is now
    in-house, so NEW jobs ship only coated (-970) boards. Those jobs are listed
    (tagged) but nothing is receivable on them; the report is OPT-IN per job —
    nothing is created until 🔁 Build report (or a ticked row in Rebuild ALL's
    picker), and only an EXISTING report is updated by a receiving save."""
    # exact wording, shared by scan / paste / the pre-save sweep. Under the OPT-IN
    # model nothing is tracked until 🔁 Build report has been clicked once, so the
    # texts say 'once it has been built' — never 'automatically' (review round 7:
    # the chips of the same job said 'No report for this job yet' at the same time)
    _COATED_HINT = ("\n\nThis looks like a COATED (…970) label. Coated boards are not received here — "
                    "their shipments are tracked in the job's Shipping Report once it has been built "
                    "(🔁 Build report; the file lives in the work-order folder, next to the COC).")
    _NOTHING_MSG = ("Nothing to receive on this job — its boards are tested in-house; coated shipments "
                    "are tracked in the job's Shipping Report once it has been built (🔁 Build report).")
    _NOT_IN_RECORDS_MSG = ("was never shipped for testing according to the records. In-house tested boards "
                           "must NOT be received here — their coated shipment is tracked in the job's "
                           "Shipping Report (once built). "
                           "If this really is an OLD testing shipment made before the app kept records, a "
                           "supervisor can override (or run 'Import older shipments…' first).")
    _PAPERWORK_MSG = ("wears a -900 testing label, but its ONLY shipping record is coated paperwork in the job "
                      "folder (a COC / Packing List / Shipping Serials file under the -970 work order) — History "
                      "never saw that shipment. Receiving it here would make the next history-only update treat "
                      "this job as a TESTING job (report renamed to the -900 spelling, dropdown tag dropped, the "
                      "same boards listed as 'still at customer'). Run 'Import older shipments…' for that coated "
                      "shipment first, then receive; or Skip. A supervisor can override.")
    def __init__(self, app):
        super().__init__(app); self._app = app
        self.title("Receiving — boards returned from testing")
        self.geometry("880x720"); self.minsize(780, 600)
        # NOT transient: a transient Toplevel loses its maximize/minimize buttons on
        # Windows (Dhruvil 2026-10-28 couldn't maximize it). Non-modal window -> normal
        # resizable/maximizable title bar + taskbar entry.
        self.configure(bg=C_BG)
        self._jobs = []; self._jobs_ready = False
        self._jobs_thread = None; self._jobs_box = None
        self._state = None; self._state_key = None; self._state_thread = None
        self._state_box = None; self._state_box_key = None
        self._session = []            # [{'serial','status'}] this session's scans
        self._approved = {}           # norm serial -> {stop headings} a supervisor cleared this session
        self._sess_overrides = []     # audit entries folded into the receive event
        self._saving = False          # True while _save runs (blocks titlebar close)
        self._backfill_open = False   # re-entry guard for the Import dialog
        self._build(); _center(self, app)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Escape>", lambda e: self._on_close())
        self.after(120, self._load_jobs_bg)
    # ── UI ────────────────────────────────────────────────────────────────
    def _build(self):
        h = tk.Frame(self, bg=C_TEAL); h.pack(fill="x")
        tk.Label(h, text="📦  Receiving — boards returned from testing", bg=C_TEAL, fg=C_WHITE,
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=20, pady=13)
        self._sess_ctr = tk.Label(h, text="0 scanned", bg=C_TEAL, fg="#CFF1E2",
                                  font=("Segoe UI", 14, "bold"))
        self._sess_ctr.pack(side="right", padx=20)
        jb = tk.Frame(self, bg=C_WHITE, highlightbackground=C_BORDER, highlightthickness=1)
        jb.pack(fill="x", padx=16, pady=(12, 4))
        ji = tk.Frame(jb, bg=C_WHITE, padx=12, pady=8); ji.pack(fill="x")
        tk.Label(ji, text="Job:", bg=C_WHITE, fg=C_MID, font=("Segoe UI", 10, "bold")).pack(side="left")
        # width 60: room for 'Eng 4006  ·  P/N 125-1030-970C  (in-house tested — coated only)'
        # (~430 px) — at 44 the tag was clipped in the field AND in the open list
        self._job_cb = ttk.Combobox(ji, values=[], state="readonly", font=("Segoe UI", 11), width=60)
        self._job_cb.pack(side="left", padx=(8, 10), ipady=2)
        self._job_cb.bind("<<ComboboxSelected>>", lambda e: self._select_job_idx(self._job_cb.current()))
        self._job_status = tk.Label(ji, text="Loading jobs from history…", bg=C_WHITE, fg=C_SUBTLE,
                                    font=("Segoe UI", 9))
        self._job_status.pack(side="left")
        self._chips = tk.Label(jb, text="Scan a returned board, or pick the job above.",
                               bg=C_WHITE, fg=C_MID, font=("Segoe UI", 10, "bold"), padx=12, pady=4,
                               anchor="w", justify="left")
        self._chips.pack(fill="x")
        # Wrap (never clip) the tracker chips when the window is narrow.
        self._chips.bind("<Configure>", lambda e: self._chips.config(wraplength=max(300, e.width - 30)))
        s = tk.Frame(self, bg=C_BG, padx=18, pady=8); s.pack(fill="x")
        lr = tk.Frame(s, bg=C_BG); lr.pack(fill="x", pady=(0, 5))
        tk.Label(lr, text="Scan the returned board's TESTING label — press Enter to add:",
                 bg=C_BG, fg=C_MID, font=("Segoe UI", 9)).pack(side="left")
        # PASS/FAIL toggle: boards arrive sorted into passed/failed piles — scan the
        # passed pile (default), flip to FAILED for the failed pile.
        mf = tk.Frame(lr, bg=C_BG); mf.pack(side="right")
        tk.Label(mf, text="Receiving as:", bg=C_BG, fg=C_SUBTLE,
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
        self._mode = "pass"
        self._btn_pass = tk.Button(mf, text="✔ PASSED", relief="flat", cursor="hand2",
                                   font=("Segoe UI", 9, "bold"), padx=12, pady=3,
                                   command=lambda: self._set_mode("pass"))
        self._btn_pass.pack(side="left", padx=(0, 4))
        self._btn_fail = tk.Button(mf, text="✘ FAILED", relief="flat", cursor="hand2",
                                   font=("Segoe UI", 9, "bold"), padx=12, pady=3,
                                   command=lambda: self._set_mode("fail"))
        self._btn_fail.pack(side="left")
        self._set_mode("pass")
        sr = tk.Frame(s, bg=C_BG); sr.pack(fill="x")
        self._var = tk.StringVar()
        self._entry = ttk.Entry(sr, textvariable=self._var, font=("Segoe UI", 18))
        self._entry.pack(side="left", fill="x", expand=True, ipady=8)
        self._entry.bind("<Return>", lambda e: self._add())
        _btn(sr, "Add", self._add, bg=C_TEAL, fg=C_WHITE, font_size=12, padx=22, pady=9).pack(side="left", padx=(10, 0))
        _btn(sr, "Paste List", self._paste, bg=C_GREEN_MED, fg=C_WHITE, font_size=10,
             padx=14, pady=9).pack(side="left", padx=(8, 0))
        t = tk.Frame(self, bg=C_BG, padx=18); t.pack(fill="x")
        self._del_btn = _btn(t, "Delete selected", self._del, bg=C_RED, fg=C_WHITE, font_size=9)
        self._del_btn.pack(side="left"); self._del_btn.config(state="disabled")
        tk.Label(t, text="⚠ = not found in testing-shipment records (kept — import older shipments to fix)",
                 bg=C_BG, fg=C_SUBTLE, font=("Segoe UI", 8)).pack(side="right")
        lf = tk.Frame(self, bg=C_WHITE, highlightbackground=C_BORDER, highlightthickness=1)
        lf.pack(fill="both", expand=True, padx=18, pady=(4, 8))
        sb = ttk.Scrollbar(lf); sb.pack(side="right", fill="y")
        self._lb = tk.Listbox(lf, font=("Courier New", 12), bg=C_WHITE, fg=C_MID,
                              selectbackground=C_TEAL_LITE, selectforeground=C_TEAL, relief="flat",
                              bd=0, selectmode="extended", activestyle="none", yscrollcommand=sb.set)
        self._lb.pack(fill="both", expand=True, padx=8, pady=8); sb.config(command=self._lb.yview)
        self._lb.bind("<<ListboxSelect>>",
                      lambda e: self._del_btn.config(state="normal" if self._lb.curselection() else "disabled"))
        df = tk.Frame(self, bg=C_BG, padx=18); df.pack(fill="x", pady=(0, 4))
        tk.Label(df, text="Received on (optional — for importing OLD records; blank = today):",
                 bg=C_BG, fg=C_MID, font=("Segoe UI", 9)).pack(side="left")
        self._rdate = tk.StringVar()
        ttk.Entry(df, textvariable=self._rdate, font=("Segoe UI", 10), width=13).pack(side="left", padx=8, ipady=2)
        tk.Label(df, text="e.g. 2026-05-14 or 5/14/26", bg=C_BG, fg=C_SUBTLE,
                 font=("Segoe UI", 8)).pack(side="left")
        nf = tk.Frame(self, bg=C_BG, padx=18); nf.pack(fill="x", pady=(0, 6))
        tk.Label(nf, text="Notes for this receipt (optional) — carrier, boxes, condition:",
                 bg=C_BG, fg=C_MID, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 3))
        self._notes = tk.Text(nf, height=2, font=("Segoe UI", 10), wrap="word", bg=C_WHITE,
                              relief="flat", highlightbackground=C_BORDER, highlightthickness=1,
                              padx=6, pady=4)
        self._notes.pack(fill="x")
        f = tk.Frame(self, bg="#E8E8E8", pady=11); f.pack(fill="x")
        _btn(f, "✓  Save received boards", self._save, bg=C_GREEN, fg=C_WHITE, font_size=11,
             padx=20, pady=9).pack(side="left", padx=(18, 8))
        _btn(f, "📄 Open file", self._open_file, bg=C_BLUE, fg=C_WHITE, font_size=9,
             padx=12, pady=8).pack(side="left", padx=(0, 6))
        # OPT-IN (Dhruvil 2026-09-23): reads "🔁 Build report" while the picked job
        # has no report yet (this button is the ONLY per-job way to create one),
        # "🔁 Rebuild report" once it exists — relabelled by _refresh_chips.
        self._rebuild_btn = _btn(f, "🔁 Rebuild report", self._rebuild, bg=C_TEAL, fg=C_WHITE,
                                 font_size=9, padx=12, pady=8)
        self._rebuild_btn.pack(side="left", padx=(0, 6))
        _btn(f, "🔁 Rebuild ALL reports", self._rebuild_all, bg=C_GREEN_MED, fg=C_WHITE, font_size=9,
             padx=12, pady=8).pack(side="left", padx=(0, 6))
        _btn(f, "⏪ Import older shipments…", self._backfill, bg="#6C7A6C", fg=C_WHITE, font_size=9,
             padx=12, pady=8).pack(side="left")
        _btn(f, "Close", self._on_close, bg="#DDD", fg=C_MID, font_size=10,
             padx=16, pady=8).pack(side="right", padx=(0, 18))
        self._entry.focus_set()
    # ── background loads ──────────────────────────────────────────────────
    # Worker threads only fill a plain box dict; the RESULT is adopted into Tk
    # state either by the 200ms after()-poll OR directly by a pumping waiter —
    # never keyed off thread-liveness alone (a thread that dies between polls
    # used to make _wait_state report failure after a successful build).
    def _on_close(self):
        if self._saving:
            self.bell(); return       # never tear the window down mid-save
        if self._session:
            if not messagebox.askyesno("Close Receiving",
                    f"{len(self._session)} scanned board(s) are NOT saved.\n\n"
                    f"Close anyway and discard them?", parent=self):
                return
        self.destroy()
    def _load_jobs_bg(self):
        import threading
        box = {}
        def work():
            try: box["jobs"] = list_testing_jobs()
            except Exception: box["jobs"] = []
        th = threading.Thread(target=work, daemon=True); th.start()
        self._jobs_thread = th; self._jobs_box = box
        def poll():
            try:
                if not self.winfo_exists(): return
            except Exception: return
            if th.is_alive(): self.after(200, poll); return
            self._adopt_jobs(box)
        poll()
    def _adopt_jobs(self, box):
        if self._jobs_ready or box is None: return
        self._jobs = box.get("jobs") or []; self._jobs_ready = True
        try:
            # In-house-tested (coated-only) jobs are LISTED (Dhruvil 2026-09-23, Q2)
            # — appended after the old-flow entries, shown under their REAL coated
            # label and tagged; scanning into one stops with 'Nothing to receive'.
            self._job_cb["values"] = [self._job_label(j) for j in self._jobs]
            self._job_status.config(text=(f"{len(self._jobs)} job(s) — pick one, or just scan a board"
                                          if self._jobs else
                                          "No testing jobs in history yet — scan a board or use Import."))
        except Exception: pass
    @staticmethod
    def _job_label(j):
        if j.get("testing", True):
            return f"Eng {j['eng']}  ·  P/N {j['pn']}"
        return f"Eng {j['eng']}  ·  P/N {_job_display_pn(j)}{_COATED_ONLY_TAG}"
    @classmethod
    def _shown_pn(cls, st):
        """The P/N the operator knows this job by: the real coated label for an
        in-house-tested job (the dropdown / chips show that one), else the -900 key."""
        return _coated_label(st) if cls._coated_only(st) else str((st or {}).get("pn") or "")
    @staticmethod
    def _coated_only(st):
        """True for an in-house-tested job: History holds no -900 label, no
        testing shipment and no receive that PROVES one (see _has_testing — a
        supervisor-overridden receive of a coated-form or Case-B board does not,
        decision (a); same predicate as the report's 'In-house tested' header).
        Nothing is receivable on such a job."""
        return bool(st) and not st.get("walk_error") and not _has_testing(st)
    def _dup_checks_due(self, k):
        """Do the DUPLICATE stops ('Scanned twice' / 'Already received') still apply
        to this board? An override answers the question it was asked: clearing
        the coated-only 'Nothing to receive here' stop must not also wave through
        a board that is already in-house (review round 4 — on an in-house-tested
        job that stop is the entry ticket for every scan, so the duplicate check
        had effectively vanished there). The same for the 'Not in shipping
        records' stop (review round 6): on an old-flow job whose testing shipment
        is folder-only / pre-History that stop is routinely overridden, and before
        it existed such a board had no stop at all — its 'Scanned twice' check
        was live, so a scanner double-fire was caught; one records override must
        not put the board into the receive event twice. The same again for the
        Save sweep's 'Coated shipment on paperwork only' stop (review round 7).
        Any OTHER (duplicate) override keeps today's one-approval-clears-all
        behaviour."""
        return not (set(self._approved.get(k) or ())
                    - {"Nothing to receive here", "Not in shipping records", "Coated shipment on paperwork only"})
    def _pn_check_due(self, k):
        """Does the P/N sanity stop ('Doesn't match this job') still apply to this
        board? The same rule as the duplicate stops (review round 5): the coated-
        only 'Nothing to receive here' override answers only its own question, so
        a WO / PO barcode scanned by mistake on an in-house-tested job still gets
        the P/N stop — unchecked it reached History as a receive and, carrying
        neither P/N, counted as a genuine -900 board come back (the job flipped to
        testing, lost its tag / home / real-label file name). Any OTHER override
        keeps today's one-approval-clears-all behaviour; old-flow jobs never have
        the coated-only heading, so for them this is `k not in self._approved`."""
        return self._dup_checks_due(k)
    def _wait_jobs(self, timeout=30.0):
        """Pump until the job list lands. True when ready."""
        if self._jobs_ready: return True
        th = self._jobs_thread
        if th is not None and th.is_alive():
            busy = _BusyDialog(self, "Loading the job list from history…")
            start = time.monotonic()
            try:
                while th.is_alive() and (time.monotonic() - start) < timeout:
                    try: self.update()
                    except Exception: break
                    time.sleep(0.05)
            finally:
                try: busy.destroy()
                except Exception: pass
        if not self._jobs_ready and self._jobs_thread is not None and not self._jobs_thread.is_alive():
            self._adopt_jobs(self._jobs_box)
        return self._jobs_ready
    def _select_job_idx(self, idx):
        if idx is None or idx < 0 or idx >= len(self._jobs): return
        j = self._jobs[idx]
        new_key = (_norm_key(j["eng"]), _norm_key(j["pn"]))
        # Switching jobs mid-session would silently re-attribute every scanned
        # board to the new job (same-P/N repeat builds make this undetectable
        # later) — make the operator decide, and clear the approvals either way.
        if self._session and self._state_key is not None and new_key != self._state_key:
            if not messagebox.askyesno("Discard scans?",
                    f"{len(self._session)} scanned board(s) belong to the previous job and "
                    f"will be DISCARDED if you switch.\n\nSwitch job and discard them?",
                    parent=self):
                try:    # revert the dropdown to the current job
                    for i, jj in enumerate(self._jobs):
                        if (_norm_key(jj["eng"]), _norm_key(jj["pn"])) == self._state_key:
                            self._job_cb.current(i); break
                except Exception: pass
                return
            self._session = []; self._approved = {}; self._sess_overrides = []
            self._ref()
        self._start_state_build(j["eng"], j["pn"], label=_job_display_pn(j))
    def _start_state_build(self, eng, pn, label=None):
        key = (_norm_key(eng), _norm_key(pn))
        # Skip only if this job's state is already good or still building — a
        # walk_error state may be rebuilt by re-picking the job (the retry path).
        if self._state_key == key and ((self._state is not None and not self._state.get("walk_error"))
                                       or (self._state_thread and self._state_thread.is_alive())):
            return
        self._state_key = key; self._state = None
        # `label` = the P/N the dropdown showed (the real coated label for an in-house-
        # tested job) — the loading line must not flash an invented -900 spelling
        self._chips.config(text=f"Loading history for Eng {eng} · P/N {label or pn} …", fg=C_SUBTLE)
        try: self._rebuild_btn.config(text="🔁 Rebuild report")   # neutral until the walk lands
        except Exception: pass
        import threading
        box = {}
        def work():
            try: box["state"] = _annotate_report_exists(build_receiving_state(eng, pn))
            except Exception: box["state"] = {"eng": eng, "pn": pn, "shipped": {}, "received": {},
                                              "coated": {}, "pd_path": None, "wo_path": None,
                                              "walk_error": True}
        th = threading.Thread(target=work, daemon=True); th.start()
        self._state_thread = th; self._state_box = box; self._state_box_key = key
        def poll():
            try:
                if not self.winfo_exists(): return
            except Exception: return
            if th.is_alive(): self.after(200, poll); return
            self._adopt_state(key, box)
        poll()
    def _adopt_state(self, key, box):
        """Move a finished build's result into self._state (poll AND waiter path).
        `box is self._state_box` rejects a SUPERSEDED build: after a void or
        backfill restarts the walk for the same job, an older still-running
        thread must not install its pre-change result over the fresh one."""
        if box is None or box is not self._state_box \
                or self._state_key != key or self._state is not None:
            return
        self._state = box.get("state")
        self._fold_cached_scan()
        self._refresh_chips()
    def _refresh_chips(self):
        st = self._state
        if not st:
            return
        try:
            if st.get("walk_error"):
                self._chips.config(fg=C_ORANGE, text=(
                    f"⚠ Could not read ALL of this job's history (network hiccup) — counts may be "
                    f"incomplete and saving is disabled. Re-pick the job to retry."))
                try: self._rebuild_btn.config(text="🔁 Rebuild report")   # not the previous job's label
                except Exception: pass
                return
            m = receiving_summary(st)
            if m.get("has_testing", True):
                line = (f"Eng {st['eng']} · P/N {st['pn']}     Shipped for testing: {m['shipped']}   ·   "
                        f"Still at customer: {m['outstanding']}   ·   "
                        f"Received back: {m['received']}  =  in-house {m['in_house'] + m['failed']} "
                        f"(awaiting coating {m['in_house']}, failed {m['failed']})"
                        f"  +  coated & shipped {m['coated']}   ·   "
                        f"Coated boards on record: {m['coated_all']}")
            else:
                # In-house-tested job (Dhruvil 2026-09-23): coated counts only — the
                # same line the report's header carries, under the REAL coated label.
                lbl = _coated_label(st)
                line = (f"Eng {st['eng']} · P/N {lbl}     In-house tested — not shipped for testing"
                        f"   ·   Coated boards shipped (P/N {lbl}): {m['coated_all']}")
            # OPT-IN status line (R4): the report FILE is the switch. The answer was
            # computed on the worker thread right after the walk (never here). It is
            # TRI-STATE: None = this PC cannot see the job's folders, which is NOT
            # "no report" (review round 1 — the desk PC has no MAG-ITAR access).
            has = st.get("_report_exists")
            found = st.get("_report_found_at")
            blk = st.get("_report_blocked_by")      # (target, Eng#) of another job's file at our target
            try: bp, be = blk if blk else (None, "")
            except Exception: bp, be = blk, ""
            build = False                     # the button reads 🔁 Build report only for a true first build
            if has:
                if st.get("_report_at_home") is False:
                    # name the copy the answer rests on: it is what keeps the job
                    # opted in, so an operator opting OUT (R6) knows what to delete.
                    # Whichever tree the copy sits in, the very next write brings it
                    # to the target: this line only shows while the HOME tree has a
                    # live WO (otherwise the resolver answers the holder WO itself
                    # and the report is simply 'at home'), and then a Save / the
                    # after-close update migrates any older-WO copy — a sibling-tree
                    # one included — exactly as a Rebuild would (review round 5;
                    # round 4 promised a Rebuild for the sibling case, wrongly)
                    tgt = st.get("_report_path")
                    how = "the next Save / after-shipment update (or 🔁 Rebuild report) brings it to"
                    line += (f"\n📄 Report exists under an older name or folder — {how}: {tgt}"
                             f"\n     (the copy that keeps this job opted in: {found}"
                             f" — delete it too if this job should stop getting a report)")
                else:
                    line += f"\n📄 Report: {st.get('_report_path')}"
                if blk:
                    # another job's same-named file sits at this job's report
                    # location: the path above is OUR copy; every update is
                    # refused until that file is moved away (review round 6)
                    line += (f"\n     ⚠ another job's file{f' (Engineering # {be})' if be else ''} sits at this job's "
                             f"report location: {bp} — move it away; every update is refused until then.")
            elif has is None:
                line += ("\n📄 Can't check for this job's report from this PC (no access to its work-order "
                         "folders) — nothing is written from here; save / rebuild on a PC that can open them.")
            elif blk:
                # another job's file at the target and NO copy of our own: a Build
                # is bound to be refused, so say what to move away — never 'Build
                # report creates it' (review round 7; the button reads Build report,
                # which is what works once the file is gone)
                build = True
                line += (f"\n📄 No report of this job's own yet — another job's file{f' (Engineering # {be})' if be else ''} "
                         f"sits at its report location: {bp}. Move that file away, then use 🔁 Build report.")
            elif not st.get("wo_path"):
                # no writer can place a file for this job (every writer gates on wo_path):
                # the 📄 Open file text, not a Build promise (review round 4)
                line += ("\n📄 No -900/-970 work-order folder is known for this job yet — close a shipment "
                         "for it first (or run 'Import older shipments…'), then use 🔁 Build report.")
            elif found:
                # R6: an Archive-only copy is NOT opted in — but it is not "no report"
                # either: a Rebuild re-homes it with that copy as the base
                line += (f"\n📄 Report exists only under Archive — 🔁 Rebuild report re-homes it into the live "
                         f"work-order folder (nothing updates it until then): {found}")
            elif has is False:
                build = True
                line += ("\n📄 No report for this job yet — 🔁 Build report creates it (first time only; "
                         "afterwards it updates itself after every shipment / receiving).")
            self._chips.config(fg=C_DARK, text=line)
            try: self._rebuild_btn.config(text="🔁 Build report" if build else "🔁 Rebuild report")
            except Exception: pass
        except Exception: pass
    def _current_job(self):
        if self._state is not None:
            return {"eng": self._state["eng"], "pn": self._state["pn"]}
        i = self._job_cb.current()
        if i is not None and 0 <= i < len(self._jobs):
            return self._jobs[i]
        return None
    # ── scanning ──────────────────────────────────────────────────────────
    def _add(self):
        v = self._var.get().strip()
        if not v: return
        self._var.set("")
        if self._current_job() is None:
            # First scan decides the job — but the job list may still be loading
            # from the share; wait for it instead of bouncing the scan with a
            # misleading "unknown job" message.
            if not self._jobs_ready and not self._wait_jobs():
                self._var.set(v); self._entry.select_range(0, "end"); self._entry.focus_set()
                messagebox.showwarning("Still loading",
                    "The job list is still loading from the share — give it a moment, "
                    "then press Enter again (your scan is kept in the box).", parent=self)
                return
            j = detect_job_for_serial(v, self._jobs)
            if j is None:
                # keep the scan in the entry so it isn't lost while they pick a job
                self._var.set(v); self._entry.select_range(0, "end")
                # a …970 label gets the coated explanation; a -900 mis-scan keeps
                # today's text verbatim (Dhruvil 2026-09-23)
                hint = self._COATED_HINT if re.search(r'\d{6,7}970', _alnum(v)) else ""
                messagebox.showwarning("Which job?",
                    "Couldn't match this scan to a known testing job.\n\n"
                    "Pick the job in the dropdown first (or run 'Import older shipments…' "
                    "if this job has never been recorded), then press Enter again." + hint, parent=self)
                self._entry.focus_set(); return
            try: self._job_cb.current(self._jobs.index(j))
            except Exception: pass
            self._start_state_build(j["eng"], j["pn"], label=_job_display_pn(j))
        # multi-line paste directly into the entry -> batch flow
        lines = [l.strip() for l in v.replace("\r", "").split("\n") if l.strip()]
        if len(lines) > 1:
            self._add_many(lines); self._entry.focus_set(); return
        k = self._sid(v)
        st = self._state
        # In-house-tested job (Dhruvil 2026-09-23, Q2): NOTHING is receivable on
        # it — every scan stops (blocking; Supervisor override stays, audited).
        if self._coated_only(st) and k not in self._approved:
            if not self._issue(v, "Nothing to receive here", self._NOTHING_MSG):
                self._entry.focus_set(); return
        # duplicate within this session (the coated-only override above does not
        # clear the duplicate stops — review round 4)
        if any(self._sid(e["serial"]) == k for e in self._session) and self._dup_checks_due(k):
            if not self._issue(v, "Scanned twice", "was already scanned in THIS receiving session."):
                self._entry.focus_set(); return
        if st is not None:
            prior = st["received"].get(k)
            # Cycle-aware: only an alert if the board is CURRENTLY received — a
            # failed board that was re-shipped and is coming back AGAIN is normal.
            if prior and _currently_received(st, k):
                if self._dup_checks_due(k):
                    when = (prior.get("ts", "") + (f" by {prior['operator']}" if prior.get("operator") else "")).strip()
                    if not self._issue(v, "Already received", f"was already received back {when}."):
                        self._entry.focus_set(); return
            elif not self._pn_check_due(k):
                pass                    # the remaining checks were answered by the override
            elif not self._carries_pn(v):
                # (the P/N SHOWN is the one the operator sees everywhere else — the
                #  real coated label on an in-house-tested job, whose boards
                #  _carries_pn accepts since round 5; the -900 key otherwise)
                if not self._issue(v, "Doesn't match this job",
                                   f"does not carry this job's P/N ({self._shown_pn(st)}) — wrong board or bad scan."):
                    self._entry.focus_set(); return
            elif k in self._approved:
                pass                    # coated-only override alone: the records question IS that stop (review round 5)
            elif not st.get("walk_error") and k not in st["shipped"] and not st["received"].get(k):
                # Never shipped for testing on record (Dhruvil 2026-09-23: boards are
                # now tested in-house — a -900 label with no testing shipment is most
                # likely one of those, not a return). Used to be accepted with a ⚠;
                # now the usual stop. Paste keeps the ⚠ (it imports old lists). A
                # walk_error state is NOT the records (its maps are empty/partial):
                # no stop from it — the scan stays 'pending' and the Save sweep
                # applies the stop against the good walk (review round 3).
                if not self._issue(v, "Not in shipping records", self._NOT_IN_RECORDS_MSG):
                    self._entry.focus_set(); return
        status = ("pending" if st is None or st.get("walk_error")
                  else ("ok" if k in st["shipped"] else "not_shipped"))
        self._session.append({"serial": v, "status": status, "result": self._mode})
        self._ref(); self._lb.see("end"); self._entry.focus_set()
    def _set_mode(self, mode):
        """PASS/FAIL toggle for subsequent scans/pastes (boards arrive sorted)."""
        self._mode = mode
        try:
            if mode == "fail":
                self._btn_fail.config(bg=C_RED, fg=C_WHITE, activebackground=C_RED, activeforeground=C_WHITE)
                self._btn_pass.config(bg="#E4E4E4", fg=C_MID, activebackground="#E4E4E4", activeforeground=C_MID)
            else:
                self._btn_pass.config(bg=C_GREEN, fg=C_WHITE, activebackground=C_GREEN, activeforeground=C_WHITE)
                self._btn_fail.config(bg="#E4E4E4", fg=C_MID, activebackground="#E4E4E4", activeforeground=C_MID)
        except Exception: pass
    def _paste(self):
        """Paste a serial list from the clipboard (e.g. your old manual tracking
        Excel — copy the serial column). Uses the current PASS/FAIL toggle."""
        try: clip = self.clipboard_get()
        except Exception:
            messagebox.showinfo("Paste", "Nothing in clipboard.", parent=self); return
        lines = [l.strip() for l in clip.replace("\r", "").split("\n") if l.strip()]
        if not lines:
            messagebox.showinfo("Paste", "No serial numbers found in clipboard.", parent=self); return
        self._add_many(lines)
    def _add_many(self, lines):
        """Batch add (paste). No per-serial popups — the whole list is processed
        with the SAME rules as scanning, then ONE summary reports what was skipped:
        already-received and wrong-P/N lines are skipped automatically (so pasting
        an overlapping historical list is repeat-safe); serials not found in the
        shipping records are KEPT with a ⚠ (import the shipped baseline to clear)."""
        if self._current_job() is None:
            if not self._jobs_ready and not self._wait_jobs():
                messagebox.showwarning("Still loading",
                    "The job list is still loading from the share — try Paste again in a moment.",
                    parent=self); return
            j = None
            for ln in lines:
                j = detect_job_for_serial(ln, self._jobs)
                if j: break
            if j is None:
                hint = (self._COATED_HINT
                        if any(re.search(r'\d{6,7}970', _alnum(ln)) for ln in lines) else "")
                messagebox.showwarning("Which job?",
                    "Couldn't match the pasted serials to a known testing job.\n\n"
                    "Pick the job in the dropdown first, then Paste again." + hint, parent=self); return
            try: self._job_cb.current(self._jobs.index(j))
            except Exception: pass
            self._start_state_build(j["eng"], j["pn"], label=_job_display_pn(j))
        if not self._wait_state():
            messagebox.showerror("History unavailable",
                "Could not read this job's history from the share — try again.", parent=self); return
        st = self._state
        if st.get("walk_error"):
            messagebox.showerror("History incomplete",
                "This job's history could not be read completely — re-pick the job to retry, "
                "then Paste again.", parent=self); return
        # In-house-tested job (Dhruvil 2026-09-23, Q2): ONE blocking stop for the
        # whole paste; a supervisor override approves (and audits) every line.
        if self._coated_only(st):
            if not self._issue(f"{len(lines)} pasted serial(s)", "Nothing to receive here",
                               self._NOTHING_MSG, approve=lines):
                return
        added = 0; unk = []
        sk_sess, sk_recv, sk_pn = [], [], []
        sess = {self._sid(e["serial"]) for e in self._session}
        for v in lines:
            k = self._sid(v)
            if not k: continue
            if k in sess:
                sk_sess.append(v); continue
            if _currently_received(st, k) and self._dup_checks_due(k):
                sk_recv.append(v); continue     # (a coated-only override does not clear this — review round 4)
            if self._pn_check_due(k):           # (…nor the P/N check — review round 5)
                if not self._carries_pn(v):
                    sk_pn.append(v); continue
            status = "ok" if k in st["shipped"] else "not_shipped"
            if status == "not_shipped": unk.append(v)
            self._session.append({"serial": v, "status": status, "result": self._mode})
            sess.add(k); added += 1
        self._ref(); self._lb.see("end")
        def _sample(lst, cap=15):
            body = "\n".join(f"   • {x}" for x in lst[:cap])
            return body + (f"\n   … and {len(lst) - cap} more" if len(lst) > cap else "")
        msg = [f"Added {added} serial(s) as {'FAILED' if self._mode == 'fail' else 'PASSED'}."]
        if unk:
            msg.append(f"\n⚠ {len(unk)} of them are not in the shipping records (kept — if they are OLD "
                       f"shipments run 'Import older shipments…' to seed the baseline; in-house tested "
                       f"boards should not be received here):\n{_sample(unk)}")
        if sk_recv:
            msg.append(f"\n{len(sk_recv)} skipped — already received before (re-pasting an "
                       f"overlapping list is safe):\n{_sample(sk_recv)}")
        if sk_pn:
            msg.append(f"\n{len(sk_pn)} skipped — don't carry this job's P/N ({self._shown_pn(st)}):\n{_sample(sk_pn)}")
        if sk_sess:
            msg.append(f"\n{len(sk_sess)} skipped — already in this session's list:\n{_sample(sk_sess)}")
        messagebox.showinfo("Paste result", "\n".join(msg), parent=self)
    def _sid(self, s):
        """Board-identity key for the current job — Eng scope + P/N + seq, REV
        IGNORED (mid-job revision changes must not split one board in two)."""
        return _serial_identity(s, (self._state or {}).get("pn", "")
                                or (self._current_job() or {}).get("pn", ""))
    def _carries_pn(self, serial):
        """Does the scan embed this job's P/N? The -900 key as always; on an
        in-house-tested (coated-only) job its boards wear the COATED label, so the
        digit-form coated P/N counts too (review round 5: the P/N stop now runs
        after the coated-only override, and must not reject the job's own -970
        boards). Old-flow jobs: byte-identical. The keys are the walk's own digit-
        only embed keys (review round 6: under an only-lettered -900 key the raw
        _alnum(pn) rejected the job's own rev-D board that sits in `shipped`)."""
        st = self._state or {}
        pa, ca = _tr_embed_keys(st.get("pn", ""))
        if len(pa) >= 6 and pa in _alnum(serial):
            return True
        if self._coated_only(st):
            return len(ca) >= 6 and ca in _alnum(serial)
        return False
    def _issue(self, serial, heading, message, approve=None):
        """Blocking issue dialog; True = supervisor overrode (approved + audited).
        `approve` = the serials the override clears (default: just `serial`) —
        a paste into an in-house-tested job shows ONE stop for the whole list."""
        d = ReceiveIssueDialog(self, serial, heading, message)
        if d.result != "override":
            return False
        # The admin roster loads once at app start and fails CLOSED to [] on a bad
        # share read — with an empty list every login would say "incorrect password"
        # forever, silently turning the override into a hard block. Reload once,
        # and if it's still empty say the real reason instead of blaming the typing.
        if not self._app.admins:
            try: self._app.admins = load_admins()
            except Exception: pass
        if not self._app.admins:
            messagebox.showerror("Admin roster unavailable",
                "The admin list could not be loaded from the share, so the override "
                "can't be verified right now.\n\nThe board stays SKIPPED. Check the "
                "network / admins.json and try again.", parent=self)
            return False
        lg = AdminLoginDialog(self, self._app.admins)
        if not getattr(lg, "result", False):
            return False
        for s in (approve if approve is not None else [serial]):
            k = self._sid(s)
            if not k: continue
            self._approved.setdefault(k, set()).add(heading)   # per reason (review round 4)
            self._sess_overrides.append({"serial": str(s), "reason": heading,
                                         "approved_by": getattr(lg, "admin_id", "") or ""})
        return True
    def _del(self):
        sel = list(self._lb.curselection())
        if not sel: return
        if messagebox.askyesno("Confirm", f"Remove {len(sel)} scan(s) from this session?", parent=self):
            for i in reversed(sel): self._session.pop(i)
            self._ref()
    def _ref(self):
        self._lb.delete(0, "end")
        nfail = 0
        for i, e in enumerate(self._session, 1):
            is_fail = e.get("result") == "fail"
            if is_fail: nfail += 1
            mark = "✘ FAIL" if is_fail else "✔ pass"
            ovr = "  (override)" if self._sid(e["serial"]) in self._approved else ""
            note = "   ⚠ not in shipping records" if e["status"] == "not_shipped" else ""
            self._lb.insert("end", f"  {i:>4}. {mark}  {e['serial']}{note}{ovr}")
            if is_fail:
                try: self._lb.itemconfig("end", fg=C_RED)
                except Exception: pass
        n = len(self._session)
        self._sess_ctr.config(text=f"{n} scanned" + (f" ({nfail} failed)" if nfail else ""))
        self._del_btn.config(state="disabled")
    # ── saving ────────────────────────────────────────────────────────────
    def _wait_state(self, timeout=45.0):
        """Block (pumping the UI) until the history walk lands. True if ready.
        The worker only fills a box; ADOPT the result here directly — waiting for
        the 200ms after()-poll to fire was a race that reported failure right
        after a successful build."""
        if self._state is not None: return True
        th = self._state_thread
        if th is not None and th.is_alive():
            busy = _BusyDialog(self, "Reading this job's history…")
            start = time.monotonic()
            try:
                while th.is_alive() and (time.monotonic() - start) < timeout:
                    try: self.update()
                    except Exception: break
                    time.sleep(0.05)
            finally:
                try: busy.destroy()
                except Exception: pass
        if self._state is None and th is not None and not th.is_alive():
            self._adopt_state(self._state_box_key, self._state_box)
        return self._state is not None
    def _save(self):
        if self._saving: return
        self._saving = True
        try:
            self._do_save()
        finally:
            self._saving = False
    def _do_save(self):
        if not self._session:
            messagebox.showwarning("Nothing scanned", "Scan the returned boards first.", parent=self); return
        job = self._current_job()
        if job is None:
            messagebox.showwarning("No job", "Pick the job in the dropdown first.", parent=self); return
        op = self._app.operator_var.get()
        if op.startswith("—"):
            messagebox.showwarning("No operator",
                "Select your name in the MAIN window (section 1) before saving received boards.",
                parent=self); return
        if not self._wait_state():
            messagebox.showerror("History unavailable",
                "Could not read this job's history from the share — try again.", parent=self); return
        if self._state.get("walk_error"):
            messagebox.showerror("History incomplete",
                "This job's history could not be read completely (network hiccup), so the "
                "duplicate checks and the report can't be trusted right now.\n\nYour scans are "
                "kept — re-pick the job in the dropdown to retry, then Save again.", parent=self)
            return
        st = self._state
        # Final sweep: entries scanned before the state landed get the same checks
        # (and the same Skip / Supervisor-override choice) as live scans.
        coated_only = self._coated_only(st)
        kept = []
        seen_now = set()
        cb_folder = st.get("_caseb_folder") or set()          # Case-B boards known from folder paperwork ONLY
        cpn_cb = _coat_ident_pn(st.get("pn")) if cb_folder else ""
        for e in self._session:
            v = e["serial"]; k = self._sid(v)
            if coated_only and k not in self._approved:
                if not self._issue(v, "Nothing to receive here", self._NOTHING_MSG):
                    continue
            # A Case-B board the REPORT files from folder paperwork alone (a Rebuild
            # with folder truth; no close / import in History): the daemon's history-
            # only walk cannot know it is Case B, so this receive would flip the job to
            # testing on the next update (review round 7 — a supervisor may override)
            if cb_folder and _serial_identity(v, cpn_cb) in cb_folder \
                    and "Coated shipment on paperwork only" not in (self._approved.get(k) or ()):
                if not self._issue(v, "Coated shipment on paperwork only", self._PAPERWORK_MSG):
                    continue
            # (the coated-only override does not clear the duplicate stops — review round 4)
            if k in seen_now and self._dup_checks_due(k):
                if not self._issue(v, "Scanned twice", "appears more than once in this session."):
                    continue
            prior = st["received"].get(k)
            # Cycle-aware: a failed board re-shipped and now coming back AGAIN
            # is normal — only a CURRENTLY-received board is a duplicate.
            if prior and _currently_received(st, k):
                if self._dup_checks_due(k):
                    when = (prior.get("ts", "") + (f" by {prior['operator']}" if prior.get("operator") else "")).strip()
                    if not self._issue(v, "Already received", f"was already received back {when}."):
                        continue
            elif not self._pn_check_due(k):
                pass                    # the remaining checks were answered by the override
            elif not self._carries_pn(v):
                if not self._issue(v, "Doesn't match this job",
                                   f"does not carry this job's P/N ({self._shown_pn(st)}) — wrong board or bad scan."):
                    continue
            elif k in self._approved:
                pass                    # coated-only override alone: the records question IS that stop (review round 5)
            elif e["status"] == "pending" and k not in st["shipped"] and not st["received"].get(k):
                # a single scan taken before the history landed gets the same
                # 'Not in shipping records' stop a live scan gets (pasted lines
                # never carry 'pending' — paste keeps its ⚠ behaviour)
                if not self._issue(v, "Not in shipping records", self._NOT_IN_RECORDS_MSG):
                    continue
            e["status"] = "ok" if k in st["shipped"] else "not_shipped"
            seen_now.add(k); kept.append(e)
        self._session = kept; self._ref()
        if not kept:
            messagebox.showinfo("Nothing to save", "No boards left to save after review.", parent=self); return
        # Optional 'Received on' date for importing OLD manual records.
        rdate = ""
        rdate_raw = self._rdate.get().strip()
        if rdate_raw:
            rdate = _parse_user_date(rdate_raw)
            if rdate is None:
                messagebox.showwarning("Date not understood",
                    f"Couldn't read the 'Received on' date:  {rdate_raw}\n\n"
                    f"Use a format like 2026-05-14 or 5/14/26 — or leave it blank for today.",
                    parent=self)
                return
        n = len(kept)
        nfail = sum(1 for e in kept if e.get("result") == "fail")
        unk = sum(1 for e in kept if e["status"] == "not_shipped")
        if coated_only:
            # an in-house-tested job (Dhruvil 2026-09-23): these boards were never
            # shipped for testing — the confirm must not say they came back from it
            # (review round 6); the old-flow sentence is byte-identical
            head = (f"Record {n} board(s) as RECEIVED on this in-house-tested job\n"
                    f"(supervisor override — these boards were not shipped for testing)?\n")
        else:
            head = f"Record {n} board(s) as RECEIVED back from testing?\n"
        msg = (head +
               f"({n - nfail} PASSED  ·  {nfail} FAILED)\n\n"
               f"Eng # {st['eng']}  ·  P/N {self._shown_pn(st)}\nOperator: {op}")
        if rdate: msg += f"\nReceived on: {rdate} (imported old record)"
        if unk: msg += f"\n\n⚠ {unk} of them are not in the testing-shipment records (kept anyway)."
        if self._sess_overrides: msg += f"\n⚠ {len(self._sess_overrides)} supervisor override(s) this session."
        c = ConfirmActionDialog(self, title="Save received boards",
                                heading="Save received boards?", message=msg,
                                confirm_label="Save", cancel_label="Cancel",
                                heading_color=C_TEAL, confirm_color=C_GREEN)
        if not c.result: return
        notes = ""
        try: notes = self._notes.get("1.0", "end").strip()
        except Exception: pass
        rec = {"wo": st["pn"], "division": _division_of_path(st.get("pd_path")),
               "operator": op, "eng": st["eng"], "pn": st["pn"],
               "serial_count": n, "serials": [e["serial"] for e in kept],
               "failed": [e["serial"] for e in kept if e.get("result") == "fail"],
               "received_date": rdate,
               "not_in_shipped": [e["serial"] for e in kept if e["status"] == "not_shipped"],
               "dup_overrides": list(self._sess_overrides),
               "notes": notes, "production_data": st.get("pd_path")}
        path = write_history_event("receive", rec)
        if not path:
            messagebox.showerror("Save failed",
                "Could not write the receive record to the share history.\n\n"
                "Nothing was saved — check the network and try again.", parent=self)
            return
        # Fold this session in immediately (display fallback if the re-walk fails) —
        # the walk's own rules apply (a coated-form override stays out of `received`)
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        _fold_receive_locally(st, kept, op, rdate, ts)
        self._session = []; self._approved = {}; self._sess_overrides = []
        try: self._notes.delete("1.0", "end")
        except Exception: pass
        try: self._rdate.set("")
        except Exception: pass
        self._ref()
        # …then RE-WALK history before regenerating the report, so the file also
        # carries receives saved by OTHER stations/sessions since this window's
        # state was built (regenerating from the cached state silently dropped
        # their rows and re-listed their boards as "still at customer").
        fresh_ok = self._rebuild_state_sync("Updating the tracker from history…")
        cur = self._state or st
        if not fresh_ok:
            # the re-walk failed and the OLDER state is what we have: its cached
            # opt-in answer dates from the job pick, and a report deleted since
            # (opt-out, R6) must not be re-created from it — recompute from disk:
            # a full read-only _report_lookup (customer / sibling folder listings
            # + a stat per candidate + a workbook head read per hit) ON THE TK
            # THREAD — acceptable on this failure path only, never on the success
            # path or in _refresh_chips (review rounds 2 + 6)
            _forget_report_cache(cur)
        # OPT-IN (Dhruvil 2026-09-23, R3): UPDATE-ONLY — a receiving save refreshes
        # an EXISTING report and never creates one ("Once report is created app
        # needs to update that report only"). The answer came with the fresh walk;
        # _receive_save_gate is the one decision point (tests exercise it directly).
        gate = _receive_save_gate(cur)
        found = cur.get("_report_found_at")       # the copy the answer rests on (before the write)
        self._refresh_chips()
        if gate:
            out = self._regen_xlsx()
            tail = (f"\n\nFile updated:\n{out}" if out else
                    "\n\n(The xlsx file will update next time — the boards ARE saved.)")
            try:
                # the report was brought home from an older copy the merge could not
                # remove (user content in it): say so, or the operator can't opt out
                if out and found and Path(found) != Path(out) and Path(found).exists():
                    tail += (f"\n\nIt was rebuilt from an older copy that is still on disk:\n{found}\n"
                             f"Delete that copy too if this job should stop getting a report.")
            except Exception: pass
        elif gate is None:
            tail = ("\n\nThis job's Shipping Report could not be checked from this PC (no access to its "
                    "work-order folders), so nothing was written to Excel — the boards ARE saved in "
                    "History. A save or 🔁 Rebuild report on a PC that can open the job's folders "
                    "(e.g. the shipping PC) brings the report up to date.")
        elif not cur.get("wo_path"):
            # no writer can place the file (the same precondition every writer has —
            # review round 4): say that, not "no report yet"
            tail = ("\n\nNo -900/-970 work-order folder is known for this job (no live shipment record "
                    "anchors it), so nothing was written to Excel — the boards ARE saved in History. Close "
                    "a shipment for it or run 'Import older shipments…' first.")
        elif cur.get("_report_blocked_by"):
            # another job's file sits at this job's report location and this job has
            # no report of its own: a Build would be refused — say what to move away,
            # not 'no report yet … use Build report' (review round 7)
            try: bp, be = cur.get("_report_blocked_by")
            except Exception: bp, be = cur.get("_report_blocked_by"), ""
            tail = (f"\n\nAnother job's file{f' (Engineering # {be})' if be else ''} sits at this job's report "
                    f"location, so nothing was written to Excel — the boards ARE saved in History:\n{bp}\n"
                    f"Move that file away, then use 🔁 Build report.")
        elif found:
            # the only copy sits under Archive: not opted in (R6), but not "no report" either
            tail = (f"\n\nThis job's Shipping Report exists only under Archive, so nothing was written to Excel — "
                    f"the boards ARE saved in History:\n{found}\nUse 🔁 Rebuild report to re-home it into the "
                    f"live work-order folder.")
        else:
            tail = ("\n\nNo Shipping Report exists for this job yet, so nothing was written to Excel — "
                    "the boards ARE saved in History. Use 🔁 Build report when you want this job's report.")
        m = receiving_summary(cur)
        if coated_only:
            # an in-house-tested job (override path): nothing was shipped for testing,
            # so no 'N of M shipped (K still at customer)' — the confirm two dialogs
            # earlier said the same (review round 7)
            sofar = ("Recorded as a supervisor-overridden receive on this in-house-tested job "
                     "(its boards were not shipped for testing).")
        else:
            sofar = (f"Received so far: {m['received']} of {m['shipped']} shipped "
                     f"({m['outstanding']} still at customer).")
        messagebox.showinfo("Received", f"✓ {n} board(s) recorded as received.\n\n" + sofar + tail,
                            parent=self)
        self._entry.focus_set()
    def _rebuild_state_sync(self, text="Updating the tracker from history…"):
        """Fresh history walk for the current job, pumping the UI. On success the
        fresh state replaces the cache and True is returned; on failure the prior
        (locally folded) state is kept for display and False is returned."""
        job = self._current_job()
        if job is None: return False
        old = self._state
        key = (_norm_key(job["eng"]), _norm_key(job["pn"]))
        import threading
        box = {}
        def work():
            try: box["state"] = _annotate_report_exists(build_receiving_state(job["eng"], job["pn"]))
            except Exception: box["state"] = None
        th = threading.Thread(target=work, daemon=True); th.start()
        busy = _BusyDialog(self, text)
        start = time.monotonic()
        try:
            while th.is_alive() and (time.monotonic() - start) < 60.0:
                try: self.update()
                except Exception: break
                time.sleep(0.05)
        finally:
            try: busy.destroy()
            except Exception: pass
        stn = box.get("state")
        if stn is not None and not stn.get("walk_error"):
            self._state = stn; self._state_key = key
            self._state_thread = None; self._state_box = box; self._state_box_key = key
            self._fold_cached_scan()
            return True
        if old is not None:
            self._state = old
        return False
    # ── the xlsx report ───────────────────────────────────────────────────
    def _regen_xlsx(self):
        st = self._state
        if not st: return None
        if not st.get("wo_path"):
            messagebox.showwarning("No job folder",
                "This job's work-order folder isn't known yet — no live shipment record (a -900 close or "
                "import, or a coated close on an in-house-tested job) anchors it: its -900 shipments "
                "were voided or predate History, so the xlsx can't be placed. The received boards "
                "ARE saved to history; close a shipment for this job or run 'Import older shipments…', "
                "then try again.", parent=self)
            return None
        while True:
            try:
                out = write_received_xlsx(st)
                # the file exists now, at home -> opted in (chips line + button label follow)
                _mark_report_written(st, out)
                self._refresh_chips()
                self._warn_conflicts(st)
                return out
            except PermissionError as e:
                # write_received_xlsx has already resolved st['wo_path'] to the REAL
                # folder (tree switch included) before the swap that raised — name
                # the file actually locked there, not the walk's cached guess
                dlg = FileInUseDialog(self, "Shipping Report",
                                      received_xlsx_path(st["wo_path"], _report_pn(st)),
                                      f"{type(e).__name__}: {e}")
                if dlg.result == "retry":
                    continue
                messagebox.showinfo("File busy",
                    "The file is open on another PC, so it wasn't rewritten now.\n\n"
                    "The received boards ARE saved to history — the file updates "
                    "automatically on the next save (or use 🔁 Rebuild report).", parent=self)
                return None
            except Exception as e:
                messagebox.showerror("File error",
                    f"Could not write the received-boards file:\n{type(e).__name__}: {e}\n\n"
                    f"The received boards ARE saved to history.", parent=self)
                return None
    def _warn_conflicts(self, st):
        """Overlap alert (Dhruvil 2026-09-02: 'if some data is overlapping then
        raise an error'): a board already in the file with DIFFERENT data. The
        file's row is always kept; the user reconciles."""
        conf = (st or {}).get("xlsx_conflicts") or []
        if not conf:
            return
        lines = [f"• {c['serial']}   [{c['sheet']}] {c['field']}: "
                 f"file has '{c['file']}', {c.get('note') or 'shipping history'} says '{c['history']}'"
                 for c in conf[:12]]
        if len(conf) > 12:
            lines.append(f"… and {len(conf) - 12} more")
        try:
            messagebox.showwarning("⚠ Overlapping data",
                "The file and shipping history DISAGREE for these boards:\n\n"
                + "\n".join(lines) +
                "\n\nThe file's rows were KEPT exactly as they are. Fix the row in the "
                "file (or void the wrong record in 🕘 History), then use 🔁 Rebuild report.",
                parent=self)
        except Exception:
            pass
    def _rebuild(self):
        if self._current_job() is None:
            messagebox.showwarning("No job", "Pick the job first.", parent=self); return
        if not self._wait_state():
            messagebox.showerror("History unavailable",
                "Could not read this job's history from the share — try again.", parent=self); return
        # Always a FRESH walk — the button's whole point is to true-up the file
        # with every station's events, not to rewrite it from this window's cache.
        if not self._rebuild_state_sync("Rebuilding from history…"):
            messagebox.showerror("History incomplete",
                "Could not read this job's history completely — the file was left "
                "unchanged. Try again in a moment.", parent=self); return
        self._refresh_chips()      # the fresh state is on screen even if we stop below
        st = self._state
        # OPT-IN (Dhruvil 2026-09-23): this button is the ONLY per-job way to CREATE
        # a report — say so the first time (a report deleted by hand is rebuilt from
        # History; manual edits that lived only in that file are gone). A None
        # answer (folders not readable from here, or a job folder that could not be
        # listed) is not "no report" — and it is not left to the write either: when
        # the home WO itself resolves and only a sibling folder blipped, the write
        # would have built a FRESH file next to the real one (review round 5), so
        # this button refuses up front like the picker does. Neither is an
        # ARCHIVE-only copy (_report_found_at with a False answer) "no report": the
        # write re-homes it with that copy as its base, nothing is built from scratch.
        found0 = st.get("_report_found_at")          # (before the write overwrites the cache)
        if st.get("_report_exists") is None:
            messagebox.showwarning("Can't check",
                "This job's report can't be checked from this PC (no access to its work-order "
                "folders, or a folder could not be listed) — nothing was written. Run 🔁 Rebuild "
                "report on the shipping PC, or try again in a moment.", parent=self)
            return
        blk0 = st.get("_report_blocked_by")
        if st.get("_report_exists") is False and blk0:
            # another job's file sits at this job's report location and this job has
            # no copy of its own: the write is bound to refuse it — say so NOW, before
            # the 'Build report?' question, the share-bound folder scan and the second
            # history walk (review round 7)
            try: bp0, be0 = blk0
            except Exception: bp0, be0 = blk0, ""
            messagebox.showwarning("Another job's file",
                f"Another job's report{f' (Engineering # {be0})' if be0 else ''} sits at this job's report "
                f"location:\n{bp0}\n\nThis job has no report of its own yet — nothing was written. Move that "
                f"file away, then use 🔁 Build report.", parent=self)
            return
        first = bool(st.get("wo_path")) and st.get("_report_exists") is False and not found0
        if st.get("wo_path") and st.get("_report_exists") is False and not found0:
            if not messagebox.askyesno("Build report?",
                    "This job has no report yet. Build it now from shipment history (and the "
                    "job's COC / Packing List / Shipping Serials folders)?", parent=self):
                return
        # FOLDER TRUTH (Dhruvil 2026-09-03): 🔁 Rebuild also reads the actual
        # -900/-970 job folders (COC / Packing List / Shipping Serials), so
        # shipments that never went through the app or an import still count.
        scan = None
        try:
            prog = ProgressDialog(self, "Scanning job folders…",
                                  "Reading COC / Packing List / Shipping Serials files")
            try:
                scan = scan_job_folders(
                    self._state,
                    progress=lambda t=None: prog.step(t or "Reading shipment files…"))
            finally:
                prog.close()
        except Exception:
            scan = None
        if scan and scan.get("ok"):
            # keep the scan: every SAVE re-walks history into a fresh state, and
            # without re-folding it would silently drop the folder-found
            # shipments until the next Rebuild (review round 3)
            self._scan_cache = (self._state_key, scan)
        # Walk history AGAIN now that the (slow, share-bound) folder scan — and the
        # 'Build report?' question — are behind us: a shipment another station
        # closed meanwhile was skipped by its own after-close update when this is
        # the job's FIRST report (no file yet), and would otherwise be missing from
        # the file written below (review round 3). On success the cached scan is
        # re-folded by _rebuild_state_sync; on failure the walk above stays.
        if not self._rebuild_state_sync("Rebuilding from history…"):
            if scan and scan.get("ok"):
                fold_folder_scan(self._state, scan)
        st = self._state
        self._refresh_chips()
        # A human 🔁 Rebuild is the ONE place the report may move between the -900
        # and -970 work-order trees (Dhruvil 2026-09-23) — never the after-close
        # thread or a Save. Flag dropped again so a later Save can't inherit it.
        st["allow_tree_switch"] = True
        try:
            out = self._regen_xlsx()
        finally:
            st.pop("allow_tree_switch", None)
        if out:
            # a FIRST build says 'built', not 'rebuilt' — the operator was asked
            # 'Build report?' and must not go looking for a file that was replaced
            # (review round 7; the Rebuild-ALL summary says 'created (first time)')
            verb = "built" if first else "rebuilt"
            if scan and scan.get("ok"):
                msg = (f"File {verb} from history + the job folders "
                       f"({scan['files']} shipment file(s) read across "
                       f"{len(scan['jobs'])} job folder(s)).")
            else:
                why = (scan or {}).get("error") or "no access from this PC"
                msg = (f"File {verb} from HISTORY ONLY — the job folders could not "
                       f"be read here ({why}).\nOld shipments that only exist as "
                       "folder files are NOT included; run Rebuild on a PC that can "
                       "open the job folders (e.g. the shipping PC).")
            if st.get("tree_switched"):
                # (tree_switched = the file was PLACED / MOVED into the sibling tree
                #  this write; a report already living there is updated in place and
                #  never flagged — review round 5, so this sentence is never said about
                #  a folder that still exists)
                side = "-970 coated" if st.get("home_side") == "testing" else "-900 testing"
                if found0:
                    msg += (f"\n\nThe report moved to the job's {side} work-order folder — its home "
                            f"tree has no live work-order folder; the old copy was merged in.")
                else:
                    # a FIRST build into the sibling tree: nothing moved, nothing merged
                    msg += (f"\n\nThe report was placed in the job's {side} work-order folder — its "
                            f"home tree has no live work-order folder.")
            messagebox.showinfo("Built" if first else "Rebuilt", msg + f"\n\n{out}", parent=self)
        # the heads-up is about the SCAN, not the file — show it even when the
        # xlsx write was blocked (review round 3: it used to vanish with it)
        heads = list(((self._state or {}).get("folder_scan") or {}).get("discrepancies") or [])
        # coated side (Dhruvil 2026-09-23): a coated board in 2+ COC files, and the
        # -900-labelled boards shipped under a coated WO (Q3) — popup only, never
        # written to the file (their Note on the sheet already says so)
        heads += list(((self._state or {}).get("folder_scan") or {}).get("coated_headsup") or [])[:6]
        caseb = [e for e in ((self._state or {}).get("coated_log") or [])
                 if _CASE_B_NOTE in str(e.get("note") or "")]
        for e in caseb[:6]:
            heads.append(f"{e.get('serial')}:  {_CASE_B_NOTE} (coated WO {e.get('wo')})")
        if len(caseb) > 6:
            heads.append(f"… and {len(caseb) - 6} more board(s) {_CASE_B_NOTE}")
        notes = receiving_anomaly_notes(self._state or {})
        if notes:
            shipped_map = (self._state or {}).get("shipped") or {}
            for k, txt in list(notes.items())[:6]:
                sn = (shipped_map.get(k) or {}).get("serial") or k
                heads.append(f"{sn}:  {txt}")
        if heads:
            lines = "\n".join("• " + h for h in heads[:12])
            more = f"\n… and {len(heads) - 12} more" if len(heads) > 12 else ""
            messagebox.showinfo("Heads-up (not an error)",
                "Worth a look — the shipment records don't fully agree:\n\n"
                + lines + more + "\n\nCounts use the COC files; nothing was "
                "blocked. Fix the odd file (or record) when you get a chance.",
                parent=self)
    def _pump_worker(self, box, title, heading, idle_text, timeout=1800):
        """Run a box-filling worker to completion behind a ProgressDialog — the
        Rebuild-ALL pump loop (worker fills the dict; only this thread touches Tk).
        True when the worker set box['done'] within the timeout."""
        prog = ProgressDialog(self, title, heading)
        start = time.monotonic()
        try:
            while not box.get("done") and (time.monotonic() - start) < timeout:
                try:
                    prog.step(f"Job {box.get('i', 0)} of {box.get('n', 0)}: {box.get('cur', '')}"
                              if box.get("n") else idle_text)
                    self.update()
                except Exception:
                    break
                time.sleep(0.1)
        finally:
            try: prog.close()
            except Exception: pass
        return bool(box.get("done"))
    def _rebuild_all(self):
        """🔁 Rebuild ALL reports (Dhruvil 2026-09-15) — since 2026-09-23 a PICKER
        (OPT-IN: "show all the jobs first but create reports only for those user
        selects"). Step 1: every job's history walk + the read-only opt-in check on
        a worker thread (progress box). Step 2: ReportPickerDialog — one row per
        job, pre-ticked when a report already exists. Step 3: the full rebuild
        (history + the job folders, cross-tree move allowed) for the TICKED jobs
        only, on a worker thread; one job's failure never stops the rest. Unticked
        jobs are left alone — no file, no folder."""
        import threading
        box = {"n": 0, "i": 0, "cur": "", "done": False, "rows": []}
        threading.Thread(target=_collect_report_rows, args=(box,), daemon=True).start()
        done = self._pump_worker(box, "Checking reports…", "Checking which jobs have a report…",
                                 "Loading job list…")
        rows = list(box.get("rows") or [])
        if not done:
            messagebox.showerror("Timed out",
                "Checking the jobs took too long — nothing was written. Try again in a moment.",
                parent=self); return
        if not rows:
            messagebox.showinfo("Rebuild ALL reports",
                "No test-and-return or in-house-tested jobs in history yet."
                + (f"\n\n{box['error']}" if box.get("error") else ""), parent=self); return
        picked = ReportPickerDialog(self, rows).result
        if not picked:
            return
        box2 = {"n": len(picked), "i": 0, "cur": "", "done": False, "ok": 0, "created": 0,
                "folders": 0, "coated_only": 0, "moved": 0, "failed": []}
        threading.Thread(target=_rebuild_report_rows, args=(picked, box2), daemon=True).start()
        self._pump_worker(box2, "Rebuilding selected reports…",
                          "Reading shipment history + job folders…", "Starting…")
        # This window's current job may have been (re)built — a fresh walk (with the
        # opt-in check) replaces the cached view; the folder scan cache re-folds.
        try:
            job = self._current_job()
            if job is not None:
                # (no state yet -> the dropdown entry's own label, never the -900 key — review round 5)
                lbl = self._shown_pn(self._state) if self._state else _job_display_pn(job)
                self._state_key = None; self._state = None
                self._start_state_build(job["eng"], job["pn"], label=lbl)
        except Exception:
            pass
        n_hist = sum(1 for r in rows if r.get("status") == "history unreadable")
        n_home = sum(1 for r in rows if r.get("status") in ("no WO folder known",
                                                            "no -900/-970 work-order folder found",
                                                            "report found — no live -900/-970 work-order folder"))
        n_nochk = sum(1 for r in rows if r.get("status") == "can't check from this PC")
        n_blk = sum(1 for r in rows if r.get("status") == "another job's file at its report location")
        n_left = sum(1 for r in rows if r.get("tickable")) - len(picked)
        parts = [f"Rebuilt {box2['ok']} report(s)"
                 + (f", {box2['created']} created (first time)" if box2["created"] else "")
                 + (f", {box2['folders']} with folder-truth" if box2["folders"] else "")
                 + (f", {box2['coated_only']} in-house tested (coated only)" if box2["coated_only"] else "")
                 + (f", {box2['moved']} moved to the sibling (-900/-970) work-order folder" if box2["moved"] else "")
                 + (f"  ·  {n_left} not ticked (left alone)" if n_left > 0 else "")
                 + (f"  ·  {n_hist} skipped (history unreadable)" if n_hist else "")
                 + (f"  ·  {n_home} skipped (no -900/-970 work-order folder found)" if n_home else "")
                 + (f"  ·  {n_nochk} skipped (job folders not readable from this PC)" if n_nochk else "")
                 + (f"  ·  {n_blk} skipped (another job's file at the report location — move it away)" if n_blk else "")]
        if not box2.get("done"):
            parts.append("\n⚠ The rebuild did not finish within the time limit — some ticked jobs "
                         "may not have been written. Run it again.")
        if box2["failed"]:
            parts.append(f"\n{len(box2['failed'])} had errors:\n  "
                         + "\n  ".join(box2["failed"][:8])
                         + (f"\n  … and {len(box2['failed']) - 8} more" if len(box2["failed"]) > 8 else ""))
        messagebox.showinfo("Rebuild ALL reports — done", "\n".join(parts), parent=self)
    def _fold_cached_scan(self):
        """Re-fold the last successful folder scan of THIS job into any freshly
        walked state — a Save's history-only re-walk must not silently revert
        the folder truth a Rebuild established (review round 3)."""
        try:
            sc = getattr(self, "_scan_cache", None)
            if sc and self._state is not None and sc[0] == self._state_key:
                fold_folder_scan(self._state, sc[1])
        except Exception:
            pass
    def _open_file(self):
        st = self._state
        if not st:
            messagebox.showwarning("No file", "Pick the job first.", parent=self); return
        if st.get("walk_error"):
            # the except-fallback state knows no folders — that is "unread", not
            # "no work-order folder" (review round 3)
            messagebox.showwarning("History incomplete",
                "This job's history could not be read — re-pick the job to retry, then open "
                "the file.", parent=self); return
        # The report's REAL place (Dhruvil 2026-09-23): resolved read-only against
        # the folders that exist and named by _report_pn — never the raw recorded
        # wo_path string + the -900 spelling (wrong for a renamed WO or a coated-only
        # job). Cached on the state by the worker thread that walked it.
        if "_report_exists" not in st:
            _annotate_report_exists(st)
        p = st.get("_report_path")
        if st.get("_report_exists") is None:
            messagebox.showwarning("No access",
                "This job's report can't be checked from this PC (no access to its work-order "
                "folders) — open it on a PC that can, e.g. the shipping PC.", parent=self); return
        blk = st.get("_report_blocked_by")
        if blk and not st.get("_report_exists"):
            # another job's file sits at this job's report location and this job has
            # no copy of its own: never open THAT as this job's report (review round 7
            # — the cached path used to be the foreign file and os.startfile got it)
            try: bp, be = blk
            except Exception: bp, be = blk, ""
            messagebox.showwarning("Not this job's file",
                f"Another job's report{f' (Engineering # {be})' if be else ''} sits at this job's report "
                f"location:\n{bp}\n\nThis job has no report of its own yet — move that file away, then use "
                f"🔁 Build report.", parent=self); return
        if not st.get("wo_path") or p is None:      # (_report_lookup: no wo_path -> no path)
            messagebox.showwarning("No file",
                "No -900 or -970 work-order folder is known for this job yet — close a shipment "
                "for it first, then use 🔁 Build report.", parent=self); return
        try: ok = Path(p).exists()
        except Exception: ok = False
        if not ok:
            if st.get("_report_exists") and st.get("_report_at_home") is False:
                # the same sentence the chips print for this state (review round 7:
                # the chips said 'the next Save brings it', this said 'Rebuild first')
                messagebox.showinfo("No file yet",
                    "This job's report exists only under an older name or folder — "
                    "the next Save / after-shipment update (or 🔁 Rebuild report) brings it to "
                    "its current place, then open it.", parent=self)
            elif st.get("_report_exists"):
                # it WAS there when this job was picked: renamed/moved/deleted since
                messagebox.showinfo("No file yet",
                    "The report is not at its expected place any more (renamed, moved or deleted "
                    "since this job was picked) — use 🔁 Rebuild report to (re)create it there, "
                    "then open it.", parent=self)
            elif st.get("_report_found_at"):
                # an Archive-only copy: frozen, not the report — guide, don't open it
                messagebox.showinfo("No file yet",
                    f"This job's report exists only under Archive:\n{st.get('_report_found_at')}\n\n"
                    "Use 🔁 Rebuild report to re-home it into the live work-order folder, then open it.",
                    parent=self)
            else:
                messagebox.showinfo("No file yet", "No report for this job yet — use 🔁 Build report.",
                                    parent=self)
            return
        try: os.startfile(str(p))
        except Exception as e:
            messagebox.showerror("Open failed", f"{type(e).__name__}: {e}", parent=self)
    def _backfill(self):
        if self._backfill_open: return           # re-entry guard (grab loss can't nest dialogs)
        self._backfill_open = True
        try:
            self._wait_state()                   # wait out an in-flight build unconditionally
            d = BackfillDialog(self, self._app, state=self._state)
            if getattr(d, "imported", 0):
                # Baseline changed — reload the job list and rebuild the state.
                self._jobs_ready = False; self._load_jobs_bg()
                job = self._current_job()
                if job:
                    lbl = self._shown_pn(self._state) if self._state else _job_display_pn(job)
                    self._state_key = None
                    self._start_state_build(job["eng"], job["pn"], label=lbl)
        finally:
            self._backfill_open = False


###############################################################################
# MAIN APP
###############################################################################

def _enable_dpi_awareness():
    """Tell Windows this process manages its own DPI, so the UI renders crisp on
    125%/150%-scaled displays instead of being bitmap-stretched (blurry). Uses
    System-DPI-aware (value 1), which is right for a fixed single-monitor shipping
    station and needs no per-monitor DPI-change handling. Must run BEFORE the
    first Tk window is created. No-op off Windows or if already set."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)   # PROCESS_SYSTEM_DPI_AWARE
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()        # pre-8.1 fallback
    except Exception:
        pass

class ShippingApp(tk.Tk):
    def __init__(self):
        _enable_dpi_awareness()   # before Tk() is created
        super().__init__(); self.title("PCC Shipping Station")
        self._apply_dpi_scaling()   # match fonts + window size to the monitor's DPI
        self._init_ttk_style()      # consistent, on-brand ttk widgets
        self.configure(bg=C_BG)
        # Open maximized so all six steps + the status chips fit without scrolling
        # every shipment; the DPI-scaled geometry above is the restore-down size.
        try: self.state("zoomed")
        except Exception: pass
        # One fast reachability probe at startup (no filesystem freeze). Everything
        # else keys off self._online so we don't re-probe repeatedly.
        self._online=pcc2_online()
        self.admins=load_admins(online=self._online); self.operators=load_operators(online=self._online)
        self.current_wo=tk.StringVar(); self.division_var=tk.StringVar(value="Industrial"); self.operator_var=tk.StringVar()
        self.pics_path=None; self.serials_path=None; self.wo_folder=None
        self.pending_photos=[]; self.serials=[]; self.packing_data=None; self.coc_data=None
        self.slip_data=None            # Shipping Slip & Labels form (Dhruvil 2026-09-24) — travels like coc_data
        self.loaded_wo=None; self.parsed_info={}; self.resumed_from_draft=None
        self._started_at=""; self._started_by=""   # when/who started THIS shipment (set at WO load; carried into draft+close for History "Started")
        self._offline_shipment=False   # True while a locally-captured Outbox shipment is loaded
        self._monday_hit=None          # last Monday.com lookup folded into the current load
        self._monday_busy=False        # a Monday lookup worker thread is in flight
        self.shipment_notes=""         # free-text notes for the current shipment
        self.current_draft_id=None     # this shipment's own draft file id (multi-draft; None until first save)
        self._reset_dup_state()        # serial duplicate-history index + supervisor approvals
        self._cta_banners=[]   # (frame,label,accent_fg,accent_bg) for sections 3-6
        # App-wide dropdown-list styling. Lives HERE (not inside a section builder)
        # so reordering sections can never silently change global behavior.
        self.option_add('*TCombobox*Listbox.font',('Segoe UI',14))
        self.option_add('*TCombobox*Listbox.selectBackground',C_GREEN)
        self.option_add('*TCombobox*Listbox.selectForeground',C_WHITE)
        self._load_logo(); self._build_ui(); self._refresh_operators()
        self._uc()   # initialize chips, banner dimming AND the footer readout
        self._start_portal_uploader()   # Customer Portal: background sender + footer status
        # Scanner-first: focus the WO entry immediately so the first barcode
        # scan of the session lands in the field instead of being dropped.
        self.wo_entry.focus_set()
    # ─── Customer Portal upload (portal_upload.py) ─────────────────────────
    # A background thread sends queued PDFs and reports through a thread-safe
    # queue that the UI reads once a second (tkinter is only touched here).
    def _start_portal_uploader(self):
        self._portal_events=_queue.Queue()
        try:
            self._portal=portal_upload.PortalUploader(lambda kind,job: self._portal_events.put((kind,job)))
            self._portal.start()
        except Exception:
            self._portal=None
        self._portal_tick()
    def _portal_tick(self):
        failures=[]
        try:
            while True:
                kind,job=self._portal_events.get_nowait()
                if kind in ("failed","blocked"): failures.append(job)
        except _queue.Empty:
            pass
        self._portal_refresh_status()
        for job in failures:
            messagebox.showwarning("Portal upload failed",
                f"{portal_upload.describe(job)} was NOT added to the Customer Portal.\n\n"
                f"Reason: {job.get('last_error') or 'unknown'}\n\n"
                f"The shipment itself is fine. Tell the office so they can upload it by hand.")
        self.after(1000,self._portal_tick)
    def _portal_refresh_status(self):
        lbl=getattr(self,"portal_lbl",None)
        if lbl is None: return
        try:
            if not portal_upload.config_ready(portal_upload.load_config()):
                lbl.config(text="Portal upload: off",fg=C_SUBTLE); return
            s=portal_upload.summary()
        except Exception:
            return
        parts=[f"✓ {s['sent_today']} sent today"]
        if s["waiting"]: parts.append(f"{s['waiting']} waiting")
        if s["failed"]: parts.append(f"⚠ {s['failed']} failed")
        lbl.config(text="Portal: "+"  ·  ".join(parts),fg=C_RED if s["failed"] else (C_AMBER if s["waiting"] else C_GREEN))
    def _portal_window(self):
        PortalQueueDialog(self,getattr(self,"_portal",None))
    def _write_pdf(self, label, make, log):
        """Writes the PDF copy of a document next to its .xlsx. Never raises: returns (path, error)."""
        try:
            if pdf_docs is None:
                raise RuntimeError("PDF library (reportlab) is not installed")
            path = make()
            log(f"{label} PDF written: {path}")
            return path, None
        except Exception as e:
            import traceback as _tb
            err = f"{type(e).__name__}: {e}"
            log(f"{label} PDF FAILED: {err} | {_tb.format_exc()}")
            return None, err
    def _apply_dpi_scaling(self):
        """With DPI-awareness on, Windows no longer stretches the app, so Tk must
        scale point-based fonts itself. Set tk scaling to the real DPI and grow
        the initial window proportionally. No-op at 100% (computes to Tk's default
        1.333); on a 150% display fonts render sharp AND at correct physical size."""
        g = 1.0
        try:
            s = self.winfo_fpixels("1i") / 72.0
            self.tk.call("tk", "scaling", s)
            g = max(1.0, s / 1.3333333)
        except Exception:
            pass
        self._gscale = g
        self.geometry(f"{int(880*g)}x{int(750*g)}")
        self.minsize(int(720*g), int(500*g))
    def _init_ttk_style(self):
        """One consistent, on-brand look for the ttk widgets (entries, dropdowns,
        the History table) instead of the default Windows-gray theme sitting next
        to the app's custom green buttons."""
        try:
            st = ttk.Style(self); st.theme_use("clam")
        except Exception:
            return
        g = getattr(self, "_gscale", 1.0)
        for name in ("TEntry", "TCombobox"):
            st.configure(name, fieldbackground=C_WHITE, background=C_WHITE,
                         foreground=C_DARK, bordercolor=C_BORDER,
                         lightcolor=C_BORDER, darkcolor=C_BORDER, padding=4)
        st.configure("TCombobox", arrowsize=int(16*g))
        st.map("TCombobox",
               fieldbackground=[("readonly", C_WHITE), ("disabled", "#F0F0F0")],
               foreground=[("readonly", C_DARK)],
               selectbackground=[("readonly", C_WHITE)],
               selectforeground=[("readonly", C_DARK)])
        st.map("TEntry", bordercolor=[("focus", C_GREEN)], lightcolor=[("focus", C_GREEN)])
        st.configure("Treeview", background=C_WHITE, fieldbackground=C_WHITE,
                     foreground=C_DARK, rowheight=int(26*g), font=("Segoe UI", 10),
                     bordercolor=C_BORDER)
        st.map("Treeview",
               background=[("selected", C_GREEN_LITE)],
               foreground=[("selected", C_GREEN)])
        st.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
                     background="#E6E6E6", foreground=C_DARK, relief="flat", padding=4)
        st.map("Treeview.Heading", background=[("active", "#DADADA")])
        st.configure("TScrollbar", background="#CFCFCF", troughcolor=C_BG,
                     bordercolor=C_BG, arrowcolor=C_MID)
    def _load_logo(self):
        self._logo_img=None
        try:
            from PIL import Image,ImageTk
            img=Image.open(LOGO_PATH); img=img.resize((48,48),Image.LANCZOS); self._logo_img=ImageTk.PhotoImage(img)
        except Exception: pass
    def _build_ui(self):
        self._build_header()
        # ─── OFFLINE banner (shown only when PCC2 is unreachable) ──────
        self._offline_bar=tk.Frame(self,bg=C_ORANGE)
        ob=tk.Frame(self._offline_bar,bg=C_ORANGE); ob.pack(fill="x",padx=16,pady=6)
        tk.Label(ob,text="⚠  OFFLINE — PCC2 is not reachable. Work orders can't be looked up; "
                 "you can still ship urgent orders to the local Outbox and sync later.",
                 bg=C_ORANGE,fg=C_WHITE,font=("Segoe UI",10,"bold"),anchor="w").pack(side="left")
        _btn(ob,"Open Outbox",self._open_outbox,bg="#FFFFFF",fg=C_ORANGE,font_size=9,padx=12,pady=4).pack(side="right",padx=(8,0))
        _btn(ob,"Re-check PCC2",self._recheck_online,bg="#FFFFFF",fg=C_ORANGE,font_size=9,padx=12,pady=4).pack(side="right")
        self._refresh_online_banner()
        # ─── FOOTER pinned to the bottom of the window ─────────────────
        # The Save Shipment / Save and Close buttons must always be visible
        # regardless of scroll position. Pack the footer FIRST (with side=bottom)
        # so it claims its space before the scrollable area expands into the rest.
        footer = tk.Frame(self, bg=C_BG)
        footer.pack(side="bottom", fill="x", padx=36, pady=(0, 8))
        self._sf(footer)
        # ─── Scrollable middle area (sections 1-6) ─────────────────────
        # Use a Canvas + Scrollbar + inner Frame pattern. Anything taller than
        # the window scrolls; works on any screen size down to ~600px tall.
        scroll_wrap = tk.Frame(self, bg=C_BG)
        scroll_wrap.pack(side="top", fill="both", expand=True)
        canvas = tk.Canvas(scroll_wrap, bg=C_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(scroll_wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        w = tk.Frame(canvas, bg=C_BG)
        win_id = canvas.create_window((0, 0), window=w, anchor="nw")
        # Make the inner frame's width follow the canvas width so children stretch
        def _on_canvas_configure(event):
            canvas.itemconfigure(win_id, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)
        # Recalculate scroll region whenever the inner frame's size changes
        def _on_inner_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        w.bind("<Configure>", _on_inner_configure)
        # Mouse-wheel support — bind globally so the wheel scrolls the canvas
        # whether the mouse is over the canvas or any of its children.
        def _on_mousewheel(event, step=None):
            # Windows / Mac normalize delta to +-120 per notch
            try:
                if event.widget.winfo_toplevel() is not self: return
                canvas.yview_scroll(step if step is not None else int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass
        self.bind_all("<MouseWheel>", _on_mousewheel)
        # Also Linux button-4/5 wheel events
        self.bind_all("<Button-4>", lambda e: _on_mousewheel(e, -1))
        self.bind_all("<Button-5>", lambda e: _on_mousewheel(e, 1))
        # Inner content frame — slight horizontal padding inside the canvas
        content = tk.Frame(w, bg=C_BG)
        content.pack(fill="both", expand=True, padx=36, pady=8)
        self._s1(content); self._s2(content); self._s3(content); self._s4(content); self._s5(content); self._s6(content)
    def _build_header(self):
        h=tk.Frame(self,bg=C_GREEN,height=60); h.pack(fill="x"); h.pack_propagate(False)
        self._header_frame=h
        l=tk.Frame(h,bg=C_GREEN); l.pack(side="left",padx=16,pady=6)
        if self._logo_img: tk.Label(l,image=self._logo_img,bg=C_GREEN).pack(side="left",padx=(0,12))
        tf=tk.Frame(l,bg=C_GREEN); tf.pack(side="left")
        tk.Label(tf,text="PRINTED CIRCUITS CORP",bg=C_GREEN,fg=C_WHITE,font=("Segoe UI",14,"bold")).pack(anchor="w")
        tk.Label(tf,text="Shipping Station — Quality · Speed · Precision",bg=C_GREEN,fg="#9FE1CB",font=("Segoe UI",9)).pack(anchor="w")
        r=tk.Frame(h,bg=C_GREEN); r.pack(side="right",padx=18)
        a=tk.Frame(r,bg="#252525",highlightbackground="#555",highlightthickness=1,cursor="hand2"); a.pack(side="right",pady=14)
        a.bind("<Button-1>",lambda e: self._oa())
        al=tk.Label(a,text="⚙  Admin",bg="#252525",fg="#CCC",font=("Segoe UI",9,"bold"),padx=14,pady=5,cursor="hand2"); al.pack(); al.bind("<Button-1>",lambda e: self._oa())
        # History as a matching ghost chip (was a blue filled button, which added a
        # third hue to the header and competed with the workflow's serials-blue).
        hb=tk.Frame(r,bg="#252525",highlightbackground="#555",highlightthickness=1,cursor="hand2"); hb.pack(side="right",pady=14,padx=(0,10))
        hb.bind("<Button-1>",lambda e: self._open_history())
        hl=tk.Label(hb,text="🕘  History",bg="#252525",fg="#CCC",font=("Segoe UI",9,"bold"),padx=14,pady=5,cursor="hand2"); hl.pack(); hl.bind("<Button-1>",lambda e: self._open_history())
        # Receiving — boards back from customer testing (same ghost-chip style).
        rb=tk.Frame(r,bg="#252525",highlightbackground="#555",highlightthickness=1,cursor="hand2"); rb.pack(side="right",pady=14,padx=(0,10))
        rb.bind("<Button-1>",lambda e: self._open_receiving())
        rl=tk.Label(rb,text="📦  Receiving",bg="#252525",fg="#CCC",font=("Segoe UI",9,"bold"),padx=14,pady=5,cursor="hand2"); rl.pack(); rl.bind("<Button-1>",lambda e: self._open_receiving())
    def _s1(self,p):
        _section_header(p,1,"Select Division and Operator")
        c=tk.Frame(p,bg=C_BG); c.pack(pady=4)
        dc=tk.Frame(c,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); dc.pack(side="left",padx=(0,10))
        di=tk.Frame(dc,bg=C_WHITE,padx=24,pady=9); di.pack(fill="both",expand=True)
        tk.Label(di,text="Production Division",bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",9)).pack(anchor="w",pady=(0,6))
        bt=tk.Frame(di,bg=C_WHITE); bt.pack(fill="x")
        self.btn_ind=tk.Button(bt,text="Industrial",width=15,relief="flat",cursor="hand2",font=("Segoe UI",14,"bold"),padx=12,pady=9,highlightbackground=C_BORDER,highlightthickness=1,command=lambda: self._sd("Industrial")); self.btn_ind.pack(side="left",padx=(0,8))
        self.btn_itar=tk.Button(bt,text="MAG - ITAR",width=15,relief="flat",cursor="hand2",font=("Segoe UI",14,"bold"),padx=12,pady=9,highlightbackground=C_BORDER,highlightthickness=1,command=lambda: self._sd("MAG - ITAR")); self.btn_itar.pack(side="left")
        self._sd("Industrial")
        oc=tk.Frame(c,bg=C_WHITE,highlightbackground=C_BORDER,highlightthickness=1); oc.pack(side="left")
        oi=tk.Frame(oc,bg=C_WHITE,padx=24,pady=9); oi.pack(fill="both",expand=True)
        tk.Label(oi,text="Operator",bg=C_WHITE,fg=C_SUBTLE,font=("Segoe UI",9)).pack(anchor="w",pady=(0,8))
        self.op_combo=ttk.Combobox(oi,textvariable=self.operator_var,state="readonly",font=("Segoe UI",16),width=24,height=6)
        self.op_combo.pack(fill="x",ipady=8)
    def _s2(self,p):
        _section_header(p,2,"Work order number")
        cd=_card(p); i=tk.Frame(cd,bg=C_WHITE,padx=14,pady=5); i.pack(fill="x")
        wr=tk.Frame(i,bg=C_WHITE); wr.pack(fill="x")
        self.wo_entry=ttk.Entry(wr,textvariable=self.current_wo,font=("Segoe UI",16)); self.wo_entry.pack(side="left",fill="x",expand=True,ipady=5)
        self.wo_entry.bind("<Return>",lambda e: self._lw()); self.wo_entry.bind("<FocusIn>",lambda e: self.wo_entry.select_range(0,"end"))
        _btn(wr,"Search →",self._lw,bg=C_GREEN,fg=C_WHITE,font_size=10,padx=16,pady=7).pack(side="left",padx=(8,0))
        inf=tk.Frame(i,bg=C_GREEN_LITE,highlightbackground=C_BORDER,highlightthickness=1); inf.pack(fill="x",pady=(6,0))
        self.lbl_div=_info_row(inf,"Division","—"); self.lbl_wo=_info_row(inf,"Workorder #","—",bold=True)
        self.lbl_cust=_info_row(inf,"Customer","—",bold=True); self.lbl_pn=_info_row(inf,"P/N","—",bold=True)
        self.lbl_eng=_info_row(inf,"Engineering #","—"); self.lbl_fp=_info_row(inf,"Full path","—")
        self.lbl_st=_info_row(inf,"Status","Scan a work order to begin",fg=C_SUBTLE)
    def _s3(self,p):
        _section_header(p,3,"CCA Photos"); cd=_card(p)
        i=tk.Frame(cd,bg=C_WHITE,padx=14,pady=5); i.pack(fill="x")
        r=tk.Frame(i,bg=C_WHITE); r.pack(fill="x")
        a=tk.Frame(r,bg=C_ORANGE_LITE,highlightbackground=C_ORANGE,highlightthickness=2,cursor="hand2",padx=14,pady=8); a.pack(side="left",fill="x",expand=True)
        a.bind("<Button-1>",lambda e: self._pp()); l=tk.Label(a,text="Click here to add photos →",bg=C_ORANGE_LITE,fg=C_ORANGE,font=("Segoe UI",10,"bold"),cursor="hand2"); l.pack(anchor="w"); l.bind("<Button-1>",lambda e: self._pp())
        self._cta_banners.append((a,l,C_ORANGE,C_ORANGE_LITE))
        self.pcl=tk.Label(r,text="0 photos",bg="#F0F0F0",fg=C_SUBTLE,font=("Segoe UI",11,"bold"),padx=16,pady=6,highlightbackground=C_BORDER,highlightthickness=1,cursor="hand2"); self.pcl.pack(side="right",padx=(8,0))
        self.pcl.bind("<Button-1>",lambda e: self._pp())
    def _s4(self,p):
        _section_header(p,4,"CCA Serial Numbers"); cd=_card(p)
        i=tk.Frame(cd,bg=C_WHITE,padx=14,pady=5); i.pack(fill="x")
        r=tk.Frame(i,bg=C_WHITE); r.pack(fill="x")
        a=tk.Frame(r,bg=C_BLUE_LITE,highlightbackground=C_BLUE,highlightthickness=2,cursor="hand2",padx=14,pady=8); a.pack(side="left",fill="x",expand=True)
        a.bind("<Button-1>",lambda e: self._sp()); l=tk.Label(a,text="Click here to add serial numbers →",bg=C_BLUE_LITE,fg=C_BLUE,font=("Segoe UI",10,"bold"),cursor="hand2"); l.pack(anchor="w"); l.bind("<Button-1>",lambda e: self._sp())
        self._cta_banners.append((a,l,C_BLUE,C_BLUE_LITE))
        self.scl=tk.Label(r,text="0 serials",bg="#F0F0F0",fg=C_SUBTLE,font=("Segoe UI",11,"bold"),padx=16,pady=6,highlightbackground=C_BORDER,highlightthickness=1,cursor="hand2"); self.scl.pack(side="right",padx=(8,0))
        self.scl.bind("<Button-1>",lambda e: self._sp())
    def _s5(self,p):
        _section_header(p,5,"Certificate of Compliance"); cd=_card(p)
        i=tk.Frame(cd,bg=C_WHITE,padx=14,pady=5); i.pack(fill="x")
        r=tk.Frame(i,bg=C_WHITE); r.pack(fill="x")
        a=tk.Frame(r,bg=C_PURPLE_LITE,highlightbackground=C_PURPLE,highlightthickness=2,cursor="hand2",padx=14,pady=8); a.pack(side="left",fill="x",expand=True)
        a.bind("<Button-1>",lambda e: self._cp()); l=tk.Label(a,text="Click here to fill Certificate of Compliance →",bg=C_PURPLE_LITE,fg=C_PURPLE,font=("Segoe UI",10,"bold"),cursor="hand2"); l.pack(anchor="w"); l.bind("<Button-1>",lambda e: self._cp())
        self._cta_banners.append((a,l,C_PURPLE,C_PURPLE_LITE))
        self.csl=tk.Label(r,text="Not filled",bg="#F0F0F0",fg=C_SUBTLE,font=("Segoe UI",11,"bold"),padx=16,pady=6,highlightbackground=C_BORDER,highlightthickness=1,cursor="hand2"); self.csl.pack(side="right",padx=(8,0))
        self.csl.bind("<Button-1>",lambda e: self._cp())
    def _s6(self,p):
        _section_header(p,6,"Packing List"); cd=_card(p)
        i=tk.Frame(cd,bg=C_WHITE,padx=14,pady=5); i.pack(fill="x")
        r=tk.Frame(i,bg=C_WHITE); r.pack(fill="x")
        a=tk.Frame(r,bg=C_TEAL_LITE,highlightbackground=C_TEAL,highlightthickness=2,cursor="hand2",padx=14,pady=8); a.pack(side="left",fill="x",expand=True)
        a.bind("<Button-1>",lambda e: self._plp()); l=tk.Label(a,text="Click here to fill Packing List →",bg=C_TEAL_LITE,fg=C_TEAL,font=("Segoe UI",10,"bold"),cursor="hand2"); l.pack(anchor="w"); l.bind("<Button-1>",lambda e: self._plp())
        self._cta_banners.append((a,l,C_TEAL,C_TEAL_LITE))
        self.plsl=tk.Label(r,text="Not filled",bg="#F0F0F0",fg=C_SUBTLE,font=("Segoe UI",11,"bold"),padx=16,pady=6,highlightbackground=C_BORDER,highlightthickness=1,cursor="hand2"); self.plsl.pack(side="right",padx=(8,0))
        self.plsl.bind("<Button-1>",lambda e: self._plp())
    def _build_footer(self,p):
        tk.Frame(p,bg=C_BORDER,height=1).pack(fill="x",pady=(8,0))
        f=tk.Frame(p,bg=C_BG); f.pack(fill="x",pady=8)
        # Shipping Labels and Slip (Dhruvil 2026-09-25: "move that button to the
        # bottom left corner, keep only the Print logo, separate from everything
        # else"): an icon-only printer button in the footer's bottom-LEFT corner,
        # its name on hover; the Save buttons stay centred on the window. Three
        # grid columns — the two outer ones equal (uniform) so the centre group
        # stays centred; in a narrow window (720 px is allowed) the outer columns
        # give way first, never Save and Close. Enabled only while a WO is loaded
        # (_update_status_chips flips it); printable any time before the close.
        f.grid_columnconfigure(0,weight=1,uniform="footer_side")
        f.grid_columnconfigure(2,weight=1,uniform="footer_side")
        self.slip_btn=_btn(f,"🖨",self._open_slip_popup,bg=C_BLUE,fg=C_WHITE,font_size=18,padx=12,pady=2)
        self.slip_btn.grid(row=0,column=0,sticky="w")
        _HoverTip(self.slip_btn,"Shipping Labels and Slip")
        inner=tk.Frame(f,bg=C_BG); inner.grid(row=0,column=1)  # centered group
        # Right corner: Customer Portal upload status; click for the list + Retry now.
        self.portal_lbl=tk.Label(f,text="Portal upload: off",bg=C_BG,fg=C_SUBTLE,font=("Segoe UI",9),cursor="hand2")
        self.portal_lbl.grid(row=0,column=2,sticky="e")
        self.portal_lbl.bind("<Button-1>",lambda e: self._portal_window())
        # Two buttons, packed RIGHT to left (Save and Close first): reads Save
        # Draft · Save and Close, as before.
        _btn(inner,"✓   Save and Close Shipment",self._fin,bg=C_GREEN,fg=C_WHITE,font_size=12,padx=32,pady=12).pack(side="right",padx=6)
        # Draft = filled teal WITH a visible border (frame wrapper — reliable on
        # Windows, unlike a tk.Button highlight ring); Save and Close stays the
        # larger filled-green primary, so the two still read differently.
        draft_wrap=tk.Frame(inner,bg="#0B4F41",padx=2,pady=2)   # darker-teal border
        draft_wrap.pack(side="right",padx=6)
        _btn(draft_wrap,"Save Draft — finish later",self._save_draft,bg=C_TEAL,fg=C_WHITE,font_size=11,padx=22,pady=8).pack()
    def _set_division(self,d):
        self.division_var.set(d)
        ind=d=="Industrial"; itar=d=="MAG - ITAR"
        # Selected = filled + check glyph; unselected = enabled-looking white secondary
        # (not the disabled-looking gray it used to be). Glyph encodes selection by
        # shape too, so it's clear regardless of the green/amber color difference.
        self.btn_ind.config(text=("✓  Industrial" if ind else "Industrial"),
                            bg=C_GREEN if ind else C_WHITE, fg=C_WHITE if ind else C_DARK,
                            activebackground=C_GREEN if ind else C_WHITE)
        self.btn_itar.config(text=("✓  MAG - ITAR" if itar else "MAG - ITAR"),
                            bg=C_AMBER if itar else C_WHITE, fg=C_WHITE if itar else C_DARK,
                            activebackground=C_AMBER if itar else C_WHITE)
    def _refresh_operators(self):
        prev=self.operator_var.get()
        self.operators=load_operators(online=self._online); ns=["— Select your name —"]+self.operators
        self.op_combo["values"]=ns
        if prev in ns and not prev.startswith("—"):
            self.operator_var.set(prev)                 # keep a deliberate choice across roster refreshes
        elif DEFAULT_OPERATOR in self.operators:
            self.operator_var.set(DEFAULT_OPERATOR)     # fresh start -> default operator
        else:
            self.operator_var.set(ns[0])                # not in roster -> placeholder
    def _refresh_online_banner(self):
        """Show/hide the offline bar to match self._online (packed under the header)."""
        bar=getattr(self,"_offline_bar",None)
        if bar is None: return
        if self._online: bar.pack_forget()
        else: bar.pack(side="top",fill="x",after=self._header_frame)
    def _recheck_online(self):
        """Operator-triggered re-probe of PCC2 (e.g. after IT says it's back)."""
        was=self._online; self._online=pcc2_online()
        self.admins=load_admins(online=self._online)
        self._refresh_operators(); self._refresh_online_banner()
        if self._online and not was:
            messagebox.showinfo("Back online","PCC2 is reachable again. Work-order lookup and "
                "normal shipping are available.\n\nIf you shipped anything offline, open the "
                "Outbox and sync those folders to their WO folders.")
        elif not self._online:
            messagebox.showwarning("Still offline","PCC2 is still not reachable.")
    def _open_outbox(self):
        """Open the local Offline Outbox folder in Explorer."""
        try:
            OUTBOX_DIR.mkdir(parents=True,exist_ok=True)
            os.startfile(str(OUTBOX_DIR))
        except Exception as e:
            messagebox.showerror("Outbox",f"Could not open the Outbox folder:\n\n{OUTBOX_DIR}\n\n{e}")
    def _open_admin(self):
        d=AdminLoginDialog(self,self.admins)
        if d.result:
            if getattr(d,"_default_pw",False):
                messagebox.showwarning("Change the default password",
                    "This admin account is still using the original default "
                    "password, which is publicly known and insecure.\n\n"
                    "Open the Administrators tab and set a new password now.")
            AdminPanel(self,self.admins,self.operators,self._oas)
    def _on_admin_saved(self,a,o): self.admins=a; self.operators=o; self._refresh_operators()
    def _open_history(self):
        """Open the Shipment History window, or focus it if already open."""
        w=getattr(self,"_history_win",None)
        try:
            if w is not None and w.winfo_exists():
                w.deiconify(); w.lift(); w.focus_set(); return
        except Exception: pass
        self._history_win=HistoryViewer(self)
    def _open_receiving(self):
        """Open the Receiving window (boards back from customer testing), or focus
        it if already open. Needs the share: history + the received-boards file
        both live on \\\\PCC2, so Receiving is disabled while offline."""
        if not self._online:
            messagebox.showwarning("Offline",
                "Receiving needs the PCC2 share (the history and the received-boards "
                "file live there).\n\nUse 'Re-check PCC2' once the server is back.")
            return
        w=getattr(self,"_receiving_win",None)
        try:
            if w is not None and w.winfo_exists():
                w.deiconify(); w.lift(); w.focus_set(); return
        except Exception: pass
        self._receiving_win=ReceivingWindow(self)
    def _load_workorder_offline(self,wo,div,op):
        """Set up an urgent shipment when PCC2 is down: the WO can't be verified,
        so we confirm the typed number and create a local Outbox folder that mirrors
        a real WO folder. The rest of the app (photos, serials, COC, Packing List,
        Save and Close) then runs unchanged, writing into the Outbox instead of the
        share."""
        c=ConfirmActionDialog(self,
            title="Ship offline", heading="Start an OFFLINE shipment?",
            message=(f"PCC2 is not reachable, so the work-order folder can't be looked up "
                     f"or verified.\n\n"
                     f"Division:  {div}\n"
                     f"Work order (as typed):  {wo}\n\n"
                     f"Documents will be generated into the local Offline Outbox so you can "
                     f"PRINT them and ship the box now. Sync the Outbox folder into the real "
                     f"WO folder once PCC2 is back.\n\n"
                     f"Double-check the WO number — it cannot be verified while offline."),
            confirm_label="Start offline shipment", cancel_label="Cancel",
            heading_color=C_ORANGE, confirm_color=C_ORANGE)
        if not c.result: return
        root=create_outbox_shipment(div,wo)
        if not root:
            messagebox.showerror("Outbox error",
                f"Could not create the local Offline Outbox folder under:\n\n{OUTBOX_DIR}\n\n"
                f"Check that this PC's local disk is writable."); return
        self.pics_path=root/PICTURES_FOLDER; self.serials_path=root/SERIALS_FOLDER; self.wo_folder=root
        self.loaded_wo=wo; self.parsed_info={"customer":"","pn":"","eng":"","rev":""}
        self.pending_photos=[]; self.serials=[]; self.packing_data=None; self.coc_data=None; self.slip_data=None
        self.resumed_from_draft=None; self._offline_shipment=True
        self._started_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._started_by="" if self.operator_var.get().startswith("—") else self.operator_var.get()
        self.shipment_notes=""; self.current_draft_id=None; self._reset_dup_state()   # no history index offline
        self.lbl_div.config(text=f"{div}  ·  OFFLINE",fg=C_ORANGE); self.lbl_wo.config(text=wo,fg=C_ORANGE)
        self.lbl_cust.config(text="— fill in COC / Packing —",fg=C_SUBTLE); self.lbl_pn.config(text="—",fg=C_SUBTLE)
        self.lbl_eng.config(text="—",fg=C_SUBTLE); self.lbl_fp.config(text=str(root),fg=C_MID)
        self.lbl_st.config(text="OFFLINE shipment — add photos, scan serials, fill COC/Packing, "
                           "then Save and Close to generate printable documents in the Outbox",fg=C_ORANGE)
        self._uc(); self.current_wo.set(""); self.wo_entry.focus_set()
        self.after(250,self._sp)   # auto-open serials, same as a fresh online load
    def _load_workorder(self):
        wo=self.current_wo.get().strip(); div=self.division_var.get(); op=self.operator_var.get()
        if not wo: messagebox.showwarning("No WO","Please Scan or Type a Work Order Number first."); return
        if op.startswith("—"): messagebox.showwarning("No operator","Please Select Your Name Before Loading a Job."); return
        # Guard: loading a new WO discards the current shipment's unsaved work.
        if self.wo_folder and (self.pending_photos or self.serials or self.coc_data or self.packing_data or self.slip_data):
            warn=ConfirmActionDialog(self,
                title="Switch work order?", heading="Discard the current shipment?",
                message=(f"WO {self.loaded_wo} still has unsaved work "
                         f"({len(self.pending_photos)} photo(s), {len(self.serials)} serial(s)).\n\n"
                         f"Loading a new work order will discard it. Use Save Shipment "
                         f"first if you want to keep it."),
                confirm_label="Discard and continue", cancel_label="Cancel",
                heading_color=C_AMBER, confirm_color=C_RED)
            if not warn.result: return
        # Offline: skip the (impossible, and freeze-prone) share search entirely and
        # set up a local Outbox shipment so the urgent order can still be documented.
        if not self._online:
            self._load_workorder_offline(wo,div,op); return
        # Monday.com auto-lookup (online only): pre-fill Customer/PN/PO/Rev and get
        # Monday's canonical WO to seed the folder search (helps when an Eng# or a
        # partial WO was typed). Non-blocking with a timeout; any failure / slow /
        # rate-limit degrades to the folder-only search below with no autofill.
        self._monday_hit=None; search_term=wo; monday_hits=[]; monday_ok=False
        if MONDAY_LOOKUP_ENABLED and MONDAY_API_KEY:
            hits=self._monday_lookup_ui(wo)
            monday_hits=hits; monday_ok=getattr(self,"_monday_last_ok",False)
            if len(hits)==1:
                self._monday_hit=hits[0]
            elif len(hits)>1:
                self._monday_hit=self._pick_monday(hits)   # None => operator skipped, use folder search
            if self._monday_hit and self._monday_hit.get("wo","").strip():
                search_term=self._monday_hit["wo"].strip()
        sd=SearchingDialog(self,search_term,div); self.update()
        # Keep the "Searching..." dialog painted during the (possibly slow) share
        # walk instead of freezing to "Not Responding". SearchingDialog is modal
        # (grab_set) so pumping events here can't re-enter _lw.
        stats={}
        m=search_wo_folder(div,search_term,progress=self.update,stats=stats)
        sd.destroy()
        if not m:
            # Tell the failures apart so a real problem doesn't masquerade as a bad
            # WO number: (1) the server is down; (2) the server is up but this PC or
            # account can't open THIS division's folder (e.g. the shipping operators
            # have no MAG-ITAR access); (3) the division opened fine but the WO
            # wasn't found. A fresh socket probe separates (1) from (2).
            reachable=stats.get("reachable"); err=stats.get("error") or ""; scanned=stats.get("scanned",0)
            denied=any(s in err for s in ("PermissionError","WinError 5","Errno 13")) or "denied" in err.lower()
            if not pcc2_online():
                messagebox.showerror("Server not reachable",
                    f"Could not reach the production share for {div}:\n\n{PRODUCTION_ROOTS[div]}\n\n"
                    f"Make sure this PC is on the network, then use 'Re-check PCC2' and try again."
                    + (f"\n\nError: {err}" if err else ""))
            elif not reachable or denied:
                messagebox.showerror(f"No access to {div}",
                    f"This PC or account can't open the {div} production folder on the share:\n\n"
                    f"{PRODUCTION_ROOTS[div]}\n\n"
                    + ("You don't have access to MAG - ITAR. If you need to ship ITAR jobs, ask IT to "
                       "grant this account access to that folder — otherwise select the Industrial "
                       "division above.\n\n"
                       if div=="MAG - ITAR" else
                       "You may not have permission to this folder, or it doesn't exist here. "
                       "Confirm the Division is correct, or contact your supervisor.\n\n")
                    + (f"(Details: {err})" if err else
                       "Tip: paste the path above into File Explorer — if it won't open, it's an access issue."))
            elif self._monday_hit:
                # Monday knew the WO but no matching folder is on the share — usually
                # a wrong division or a job folder not created yet (NOT a bad number).
                messagebox.showerror("Found in Monday, not on the share",
                    f"'{wo}' matches a Monday work order (WO {self._monday_hit.get('wo','?')}"
                    + (f", {self._monday_hit['customer']}" if self._monday_hit.get('customer') else "")
                    + f") but no matching folder was found under:\n\n{PRODUCTION_ROOTS[div]}\n\n"
                    f"(Searched {stats.get('scanned',0)} job folders.)\n\n"
                    f"Confirm the Division ({div}) is correct — the job folder may live under the "
                    f"other division, or may not have been created on the share yet.")
            else:
                messagebox.showerror("Not found",
                    f"No work order or engineering number matching '{wo}' found in:\n\n"
                    f"{PRODUCTION_ROOTS[div]}\n\n(Searched {stats.get('scanned',0)} job folders.)\n\n"
                    f"Check the number, confirm the Division ({div}) is correct, or contact your supervisor.")
            return
        wf=m[0] if len(m)==1 else None
        if not wf:
            pk=WOPickerDialog(self,m)
            if pk.result is None: return
            wf=pk.result
        # Before confirming, decide whether the Monday hit really is THIS folder's
        # job. The reliable key is the Work Order folder name inside Production Data,
        # which always matches Monday's WO number; the job-folder name
        # ("Customer PN [Eng - BOM Rev - Assy Rev]") is fuzzy, so it's only a
        # secondary check. If neither the WO nor the Eng# lines up, drop the hit — no
        # PO/CPN autofill and nothing claimed in the confirm — so a fuzzy folder match
        # can never carry an unrelated job's PO.
        if self._monday_hit and not monday_hit_matches_folder(self._monday_hit, wf):
            self._monday_hit=None
        # Second chance (Dhruvil 2026-09-25: "Customer PO is not filling from
        # Monday"): Monday matches a WO number EXACTLY, so a WO typed short
        # ('16588' for the folder 16588.1) — or an Eng# whose Monday picker was
        # skipped — found nothing there although the folder search found the
        # job. First the answer already in hand (an Eng# lookup lists the
        # folder's own item), then Monday once more with the folder's own WO
        # number; either is taken only when exactly ONE item carries that WO
        # (never a guess), and the confirm below still shows its PO before
        # anything is applied. Skipped when Monday did not ANSWER the first time
        # (down / slow: don't make the operator wait out a second timeout) and
        # when the folder's number is what was typed (the same question twice).
        if self._monday_hit is None and MONDAY_LOOKUP_ENABLED and MONDAY_API_KEY and monday_ok:
            self._monday_hit=pick_monday_hit_for_folder(monday_hits, wf)
            fwo=monday_folder_wo(wf.name)
            if self._monday_hit is None and fwo and fwo.lower()!=(wo or "").strip().lower():
                self._monday_hit=pick_monday_hit_for_folder(self._monday_lookup_ui(fwo), wf)
        # Always confirm the exact job with the operator before loading — now showing
        # the Monday-matched Customer PO so they verify it before it's applied.
        dl=ConfirmDialog(self,wo,wf,div,monday=self._monday_hit)
        if not dl.result: return
        try:
            pp=wf/PICTURES_FOLDER; sp=wf/SERIALS_FOLDER
            if not ensure_folder(pp) or not ensure_folder(sp): return
            self.pics_path=pp; self.serials_path=sp; self.wo_folder=wf
            self.loaded_wo=wf.name; self.parsed_info=parse_wo_path(wf)
            self._merge_monday(self.parsed_info)   # fill gaps + add PO/CPN (folder authoritative; hit already validated above)
            self.pending_photos=[]; self.serials=[]; self.packing_data=None; self.coc_data=None; self.slip_data=None
            self.resumed_from_draft=None; self.shipment_notes=""
            # Stamp when/who started this shipment at WO load, so History's "Started"
            # is the real load time. A resumed draft restores its own start below.
            self._started_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._started_by="" if self.operator_var.get().startswith("—") else self.operator_var.get()
            self._offline_shipment=False   # an online-loaded WO is never an offline shipment
            # Start building the serial duplicate-history index for this job in the
            # background so it's ready by the time the operator scans.
            self._reset_dup_state(); self._kickoff_dup_index()
            info=self.parsed_info; root=Path(PRODUCTION_ROOTS[div])
            try: rel=wf.relative_to(root)
            except ValueError: rel=wf
            dc=C_GREEN if div=="Industrial" else C_AMBER
            self.lbl_div.config(text=div,fg=dc); self.lbl_wo.config(text=wf.name,fg=C_GREEN)
            self.lbl_cust.config(text=info["customer"],fg=C_DARK); self.lbl_pn.config(text=info["pn"],fg=C_DARK)
            self.lbl_eng.config(text=info["eng"],fg=C_DARK); self.lbl_fp.config(text=str(rel),fg=C_MID)
            # Multiple drafts can coexist per WO (one per batch). If any exist, the
            # operator CHOOSES: resume one (e.g. supervisor closing a specific
            # batch, dated rows tell them apart) or start a NEW shipment alongside.
            self.current_draft_id=None
            draft=None
            drafts=list_shipment_drafts(wf)
            if drafts:
                pk=DraftPickerDialog(self, wf.name, drafts)
                if pk.result is None:            # Cancel -> abort the load entirely
                    self._reset_form(); return
                mode, chosen = pk.result
                if mode=="resume":
                    draft=load_shipment_draft(wf, path=chosen.get("path"))
                    if draft is None:
                        messagebox.showerror("Draft unreadable",
                            "That draft file could not be read — starting a new shipment instead. "
                            "The file was left in place.")
                    else:
                        self.current_draft_id=str(draft.get("draft_id") or "")
            if draft:
                self.pending_photos=draft.get("pending_photos") or []
                self.serials=draft.get("serials") or []
                self.coc_data=draft.get("coc_data")
                self.packing_data=draft.get("packing_data")
                self.slip_data=draft.get("slip_data")     # None in drafts from before 2026-09-24
                self.shipment_notes=draft.get("notes","") or ""
                saved_at=draft.get("saved_at","earlier"); saved_by=draft.get("saved_by","")
                if draft.get("started_at"): self._started_at=str(draft.get("started_at"))
                if draft.get("started_by"): self._started_by=str(draft.get("started_by"))
                self.resumed_from_draft={"saved_at": saved_at, "saved_by": saved_by}
                tag=f" by {saved_by}" if saved_by else ""
                extra=f" ({len(drafts)-1} other draft(s) remain)" if len(drafts)>1 else ""
                self.lbl_st.config(text=f"📋  Draft resumed (saved {saved_at}{tag}){extra} — review and complete, then Save and Close",fg=C_TEAL)
            else:
                st="Ready — add photos, scan serials, fill forms"
                h=self._monday_hit
                if h:
                    pre=[]
                    if h.get("po"): pre.append(f"PO {h['po']}")
                    if monday_bom_rev(h.get("rev")): pre.append(f"BOM Rev {monday_bom_rev(h.get('rev'))}")
                    if pre: st="Ready — Monday pre-filled "+", ".join(pre)+".  Add photos, scan serials, fill forms"
                self.lbl_st.config(text=st,fg="#27500A")
                # Scanner-first: on a FRESH load the next physical action is
                # scanning serials, so open the Serial popup automatically with
                # its scan box focused. This also removes the misdirected-scan
                # hazard (a serial scanned into the WO box triggered a server
                # search + "Not found"). Draft resumes are NOT auto-opened —
                # supervisors reviewing a draft keep full control.
                self.after(250, self._sp)
            self._uc(); self.current_wo.set(""); self.wo_entry.focus_set()
        except Exception as e: messagebox.showerror("Load error",f"{type(e).__name__}: {e}")
    def _monday_lookup_ui(self, term):
        """Run monday_lookup(term) on a daemon worker thread behind a small MODAL
        'Searching Monday.com…' dialog, pumping the UI so the window never freezes.
        The dialog's grab_set means the pump can't re-enter _load_workorder.
        Returns the hit list, or [] on any failure/timeout (caller degrades to
        the folder search)."""
        import threading
        # In-flight guard: if a prior lookup's worker is still wedged (e.g. a DNS
        # stall that the socket timeout doesn't bound), don't stack another thread —
        # just degrade to the folder search for this load.
        self._monday_last_ok=False     # did Monday ANSWER (a 'no match' included)? — read by the second-chance lookup
        if getattr(self,"_monday_busy",False):
            return []
        self._monday_busy=True
        box={"done":False,"res":[]}; st={}
        def work():
            try: box["res"]=monday_lookup(term,status=st)
            except Exception: box["res"]=[]
            finally: box["done"]=True; self._monday_busy=False
        threading.Thread(target=work,daemon=True).start()
        dlg=_MondaySearchingDialog(self,term)
        start=time.monotonic()
        try:
            # Hard cap a little beyond monday_lookup's own socket timeout, on a
            # MONOTONIC clock so an NTP/clock step can't extend the wait — a wedged
            # connection can never hang the load.
            while not box["done"] and (time.monotonic()-start)<(MONDAY_TIMEOUT+3.0):
                try: self.update()
                except Exception: pass
                time.sleep(0.05)
        finally:
            dlg.destroy()
        self._monday_last_ok=bool(box["done"] and st.get("ok"))
        return box["res"] if box["done"] else []
    def _pick_monday(self, hits):
        """Choose among multiple Monday matches. Returns the chosen dict, or None
        if the operator skipped (then the normal folder search runs, no autofill)."""
        return MondayPickerDialog(self,hits).result
    def _merge_monday(self, info):
        """Fold the chosen Monday hit into parsed_info. Folder-derived
        customer/pn/eng stay AUTHORITATIVE (they reflect what physically exists
        on the share); Monday only fills those BLANKS and supplies PO + CPN,
        which the folder name never encodes. The REV is the exception
        (Dhruvil 2026-10-05: "It should always carry BOM rev. If monday has BOM
        rev column then it should always prioritize that to folder name"):
        Monday's rev IS its BOM Rev column, so a real value there replaces the
        folder name's BOM rev (parse_wo_path), which stays the fallback.
        parsed_info always gains 'po'/'cpn' keys so the COC and Packing popups
        can read them safely."""
        h=getattr(self,"_monday_hit",None) or {}
        for k in ("customer","eng"):
            cur=(info.get(k) or "").strip()
            if (not cur or cur=="N/A") and h.get(k):
                info[k]=_clean_customer_name(h[k]) if k=="customer" else h[k]
        mrev=monday_bom_rev(h.get("rev"))
        if mrev:
            info["rev"]=mrev
        # P/N: the folder name is authoritative; only when it yields none do we
        # fall back to Monday's customer P/N (still operator-editable in the COC).
        cur_pn=(info.get("pn") or "").strip()
        if (not cur_pn or cur_pn=="N/A") and h.get("cpn"):
            info["pn"]=h["cpn"]
        info["po"]=h.get("po","")
        info["cpn"]=h.get("cpn","")
    def _open_photo_popup(self):
        if not self.wo_folder: messagebox.showwarning("No WO","Load a work order first."); return
        p=PhotoPopup(self,self.pending_photos)
        if p.result is not None: self.pending_photos=p.result; self._uc()
    def _open_serial_popup(self,event=None):
        if not self.wo_folder: messagebox.showwarning("No WO","Load a work order first."); return
        dup_ctx=self._dup_ctx_for_scan()
        p=SerialPopup(self,self.serials,notes=self.shipment_notes,dup_ctx=dup_ctx,
                      job_pn=(self.parsed_info or {}).get("pn",""))
        if p.result is not None:
            self.serials=p.result
            self.shipment_notes=getattr(p,"notes_result",self.shipment_notes)
            if getattr(p,"overrides",None): self._dup_overrides.extend(p.overrides)
            self._uc()
    # ── serial duplicate-history index (scoped to the current Eng#+P/N) ───────
    def _reset_dup_state(self):
        """Clear the duplicate-history index + supervisor approvals (called on each
        WO load and on form reset). Approvals are per shipment, never carried over."""
        self._dup_index={}; self._dup_index_ready=False
        self._dup_index_key=None; self._dup_index_building=None
        self._dup_approved=set(); self._dup_overrides=[]
    def _kickoff_dup_index(self):
        """Start building the duplicate-history index for the loaded job, in the
        background, so it's ready by the time the operator scans. No-op offline (the
        share history is unreachable) or when Eng#/P/N aren't both known yet."""
        if self._offline_shipment or not self._online: return
        eng=_norm_key((self.parsed_info or {}).get("eng"))
        pn=_norm_key((self.parsed_info or {}).get("pn"))
        if eng and pn: self._ensure_dup_index(eng, pn)
    def _ensure_dup_index(self, eng_key, pn_key):
        """(Re)build the index on a daemon thread if the job changed. The UI never
        blocks; _dup_lookup returns None until the build lands (the submit sweep is
        the backstop for anything scanned meanwhile)."""
        key=(eng_key, pn_key)
        if self._dup_index_key==key and self._dup_index_ready: return
        if self._dup_index_building==key: return
        self._dup_index_building=key
        self._dup_index_key=key; self._dup_index={}; self._dup_index_ready=False
        eng=(self.parsed_info or {}).get("eng",""); pn=(self.parsed_info or {}).get("pn","")
        import threading
        def work():
            idx={}
            try: idx=build_serial_history(eng, pn)
            except Exception: idx={}
            if self._dup_index_key==key:      # ignore a stale build if the job changed
                self._dup_index=idx; self._dup_index_ready=True
            if self._dup_index_building==key: self._dup_index_building=None
        threading.Thread(target=work,daemon=True).start()
    def _dup_lookup(self, serial):
        """Return the prior occurrence of `serial` in this job's closed shipments,
        or None (also None while the index is still loading). Exact match — a
        different Rev counts as a different serial."""
        if not self._dup_index_ready: return None
        pn=(self.parsed_info or {}).get("pn","")
        return (self._dup_index or {}).get(_serial_identity(serial, pn))
    def _prompt_pn(self):
        """Confirm the P/N when a job loads without one (the dup check is P/N-scoped).
        Writes the entered P/N into parsed_info + the info card. Returns it or ''."""
        d=PNPromptDialog(self, self.loaded_wo or "", (self.parsed_info or {}).get("eng",""))
        pn=(getattr(d,"result","") or "").strip()
        if pn:
            if not isinstance(self.parsed_info, dict): self.parsed_info={}
            self.parsed_info["pn"]=pn
            try: self.lbl_pn.config(text=pn, fg=C_DARK)
            except Exception: pass
        return pn
    def _dup_ctx_for_scan(self):
        """Build the duplicate-check context passed to SerialPopup, or None to
        disable the check (offline, or no Eng#+P/N to scope on). Prompts to confirm
        the P/N first if it loaded blank."""
        if self._offline_shipment or not self._online: return None
        eng=_norm_key((self.parsed_info or {}).get("eng"))
        pn=_norm_key((self.parsed_info or {}).get("pn"))
        if not pn:
            pn=_norm_key(self._prompt_pn())
        if not eng or not pn: return None
        self._ensure_dup_index(eng, pn)
        op=self.operator_var.get(); op="" if op.startswith("—") else op
        pn_raw=(self.parsed_info or {}).get("pn","")
        return {"lookup": self._dup_lookup, "admins": self.admins, "approved": self._dup_approved,
                "identity": (lambda s, _pn=pn_raw: _serial_identity(s, _pn)),
                "eng": (self.parsed_info or {}).get("eng",""), "pn": pn_raw,
                "wo": self.loaded_wo or "", "division": self.division_var.get(), "operator": op}
    def _open_packing_popup(self,event=None):
        if not self.wo_folder: messagebox.showwarning("No WO","Load a work order first."); return
        p=PackingListPopup(self,self.loaded_wo,len(self.serials),self.packing_data,self.parsed_info,self.coc_data)
        if p.result is not None:
            self.packing_data=p.result; self._uc()
            if getattr(p,"print_requested",False):
                eng=sanitize_filename_part(self.parsed_info.get("eng") or self.loaded_wo, fallback=self.loaded_wo)
                self._preview_print("Packing List", lambda folder: write_packing_list_xlsx(
                    folder, eng, self.loaded_wo, self.packing_data, sort_serials(self.serials),
                    logo_path=str(LOGO_PATH) if LOGO_PATH else None))
    def _open_coc_popup(self,event=None):
        if not self.wo_folder: messagebox.showwarning("No WO","Load a work order first."); return
        p=COCPopup(self,self.loaded_wo,len(self.serials),self.coc_data,self.parsed_info)
        if p.result is not None:
            self.coc_data=p.result; self._uc()
            if getattr(p,"print_requested",False):
                eng=sanitize_filename_part(self.parsed_info.get("eng") or self.loaded_wo, fallback=self.loaded_wo)
                self._preview_print("COC", lambda folder: write_coc_xlsx(
                    folder, eng, self.loaded_wo, self.coc_data, sort_serials(self.serials)))
    def _open_slip_popup(self,event=None):
        """Shipping Labels and Slip (Dhruvil 2026-09-24; renamed 2026-09-25). Unlike the COC / Packing
        popups the printing happens INSIDE the popup (it stamps codes / outputs
        / printed times into the data and stays open for the other page); the
        app only keeps the returned slip_data — for drafts and the close record."""
        if not self.wo_folder: messagebox.showwarning("No WO","Load a work order first."); return
        op=self.operator_var.get(); op="" if op.startswith("—") else op
        p=SlipLabelsPopup(self,self.loaded_wo,len(self.serials),self.slip_data,self.parsed_info,
                          monday=getattr(self,"_monday_hit",None),coc=self.coc_data,packing=self.packing_data,
                          operator=op,operators=self.operators,wo_folder=self.wo_folder,
                          box_sizes=load_box_sizes(online=None if self._online else False),
                          online=None if self._online else False,division=self.division_var.get())
        if p.result is not None:
            self.slip_data=p.result; self._uc()
    def _preview_print(self, label, generate):
        """Save-and-Print: render the document into a temp folder and open it in
        Excel so the operator can review and print it (Ctrl+P). The writers set the
        print area to the written content only and fit all columns onto one page
        wide; _ensure_duplex_default() sets every printer to two-sided first. The
        FILED copy is still generated at Save and Close — this is just the print render."""
        import tempfile
        try:
            tmp = Path(tempfile.gettempdir()) / "PCC Shipping App - Print"
            tmp.mkdir(parents=True, exist_ok=True)
            for old in tmp.glob("*.xlsx"):        # best-effort tidy of prior renders
                try: old.unlink()
                except Exception: pass
            path = generate(tmp)
            if not path or not Path(path).exists():
                messagebox.showwarning("Print", f"Could not generate the {label} to print."); return
            self._ensure_duplex_default()         # set the printer to two-sided before Excel reads it
            os.startfile(str(path))               # opens in Excel; operator prints (Ctrl+P), two-sided by default
        except Exception as e:
            messagebox.showerror("Print",
                f"Could not open the {label} for printing:\n\n{type(e).__name__}: {e}\n\n"
                f"(The {label} data is still saved — you can print it after Save and Close.)")
    def _ensure_duplex_default(self):
        """Best-effort: set EVERY installed printer's default to two-sided so the
        shipping documents print double-sided no matter which printer the operator
        picks (not just the Windows default). Runs once per session; each printer is
        tried independently (printers that don't support duplex — PDF/Fax/OneNote —
        are skipped). Persistent, like setting it by hand. Uses a built-in Windows
        command (no extra libraries); any failure is silent and the manual printer
        setting remains the guaranteed fallback."""
        if getattr(self, "_duplex_set", False):
            return
        self._duplex_set = True   # attempt once per session regardless of outcome
        try:
            import subprocess
            ps = ("Get-Printer | ForEach-Object { "
                  "try { Set-PrintConfiguration -PrinterName $_.Name -DuplexingMode TwoSidedLongEdge } "
                  "catch {} }")
            subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                capture_output=True, timeout=30,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception:
            pass
    def _set_flow_state(self):
        """Guide the eye to the current step: while no WO is loaded, sections 3-6
        render dimmed (clicking them would only produce a 'Load a work order first'
        warning); their accent colors return the moment a WO loads."""
        loaded=self.wo_folder is not None
        for a,l,acc_fg,acc_bg in self._cta_banners:
            if loaded: a.config(bg=acc_bg,highlightbackground=acc_fg); l.config(bg=acc_bg,fg=acc_fg)
            else:      a.config(bg="#F0F0F0",highlightbackground=C_BORDER); l.config(bg="#F0F0F0",fg=C_SUBTLE)
    def _update_status_chips(self):
        self._set_flow_state()
        np=len(self.pending_photos); ns=len(self.serials)
        self.pcl.config(text=f"{np} photo{'s' if np!=1 else ''}",fg=C_ORANGE if np>0 else C_SUBTLE,bg=C_ORANGE_LITE if np>0 else "#F0F0F0",
                        highlightbackground=C_ORANGE if np>0 else C_BORDER)
        self.scl.config(text=f"{ns} serial{'s' if ns!=1 else ''}",fg=C_BLUE if ns>0 else C_SUBTLE,bg=C_BLUE_LITE if ns>0 else "#F0F0F0",
                        highlightbackground=C_BLUE if ns>0 else C_BORDER)
        if self.packing_data: self.plsl.config(text="✓ Filled",fg=C_TEAL,bg=C_TEAL_LITE,highlightbackground=C_TEAL)
        else: self.plsl.config(text="Not filled",fg=C_SUBTLE,bg="#F0F0F0",highlightbackground=C_BORDER)
        if self.coc_data: self.csl.config(text="✓ Filled",fg=C_PURPLE,bg=C_PURPLE_LITE,highlightbackground=C_PURPLE)
        else: self.csl.config(text="Not filled",fg=C_SUBTLE,bg="#F0F0F0",highlightbackground=C_BORDER)
        # Shipping Labels and Slip button (printer icon only, 2026-09-25): live
        # only while a WO is loaded — blue, then green once the slip or labels
        # have been sent to print for this shipment.
        b=self.slip_btn
        pr=(self.slip_data or {}).get("printed") or {}
        printed=bool(pr.get("slip") or pr.get("labels") or pr.get("pack"))
        b._base_bg=(C_GREEN if printed else C_BLUE) if self.wo_folder else "#BFC7C3"
        b.config(state="normal" if self.wo_folder else "disabled",bg=b._base_bg)
    def _save_draft(self):
        """Save the current shipment as a draft for supervisor review.
        Photos are MOVED from their source paths into the WO folder at this step
        (local PC has fast access to source files). Draft JSON then stores the
        new network paths, so supervisors on other machines can resume reliably."""
        if not self.wo_folder:
            messagebox.showwarning("No WO","Load a work order first."); return
        # Guard: drafts must contain at least photos OR serials.
        # COC/Packing data alone is not enough — drafts are for capturing shipment evidence.
        if not self.pending_photos and not self.serials:
            messagebox.showwarning(
                "Nothing to save",
                "Please add photos and serial numbers before saving the shipment.\n\n"
                "Drafts require at least photos or serial numbers — "
                "COC and Packing List info alone can't be saved as a draft.")
            return
        # Confirmation popup
        op=self.operator_var.get(); op_show=op if not op.startswith("—") else "(no operator selected)"
        n_photos=len(self.pending_photos); n_serials=len(self.serials)
        coc_status="filled" if self.coc_data else "not filled"
        pl_status="filled" if self.packing_data else "not filled"
        # Count how many of the current photos are still external (need to be moved).
        # Resolve the WO pictures root ONCE — resolving a UNC path is a network
        # round-trip, and doing it per photo made even this pre-confirm step crawl.
        _pics_root = str(Path(self.pics_path).resolve())
        n_external = sum(1 for p in self.pending_photos
                         if not str(Path(p).resolve()).startswith(_pics_root))
        photo_note = (f"\n\n{n_external} photo(s) will be moved into the WO folder."
                      if n_external else "")
        msg=(f"Save this shipment as a draft for later review?\n\n"
             f"WO: {self.loaded_wo}\n"
             f"Photos: {n_photos}  ·  Serials: {n_serials}\n"
             f"COC: {coc_status}  ·  Packing List: {pl_status}"
             f"{photo_note}\n\n"
             f"No documents will be generated yet. You can reload this WO "
             f"any time to resume from where you left off.")
        confirm=ConfirmActionDialog(self,
            title="Save Shipment", heading="Save shipment as draft?", message=msg,
            confirm_label="Save Draft", cancel_label="Cancel",
            heading_color=C_TEAL, confirm_color=C_TEAL)
        if not confirm.result: return
        # Long synchronous phase: photo moves to the share + the draft/history
        # writes can take 30s+ — show a live progress box instead of a frozen
        # window. try/finally guarantees the box always comes down.
        prog=ProgressDialog(self, "Saving draft…",
                            f"Saving shipment draft for WO {self.loaded_wo}")
        try:
            self._save_draft_work(op, n_photos, n_external, prog)
        finally:
            prog.close()
    def _save_draft_work(self, op, n_photos, n_external, prog):
        """The slow part of Save Draft (photo moves + draft + history writes),
        run under the ProgressDialog owned by _save_draft."""
        # ─── Move external photos into the WO folder ──────────────────────
        # Photos already inside self.pics_path (from a prior Save) are kept as-is.
        # External photos move into <pics_path>/<date>/<wo>_S<seq>.<ext>
        # using a "_S" prefix so saved-draft photos don't collide with close-time
        # photos (which use "{date_folder}_{seq}.{ext}").
        import traceback, re
        log = lambda m: log_shipment_event(self.wo_folder, m)
        log(f"SAVE start: wo={self.loaded_wo} photos={n_photos} (external={n_external})")
        ts=datetime.datetime.now(); date_folder=ts.strftime("%m%d")+ts.strftime("%y")
        photo_dir=self.pics_path/date_folder
        if not ensure_folder(photo_dir): photo_dir=self.pics_path
        pics_root=str(Path(self.pics_path).resolve())
        moved_photos=[]; photos_missing=[]; photos_failed=[]
        # Find next available sequence number in photo_dir for "_S" prefix
        existing_nums=[]
        try:
            for f in photo_dir.iterdir():
                m=re.match(r'^.+_S(\d+)\.', f.name)
                if m:
                    try: existing_nums.append(int(m.group(1)))
                    except ValueError: pass
        except Exception: pass
        next_seq=(max(existing_nums)+1) if existing_nums else 1

        for _pi, src in enumerate(self.pending_photos, 1):
            prog.step(f"Moving photo {_pi} of {n_photos} into the WO folder…")
            try:
                src_p=Path(src)
                # If already inside the WO Shipping Pictures area, keep as-is
                try:
                    src_resolved=str(src_p.resolve())
                    if src_resolved.startswith(pics_root):
                        moved_photos.append(str(src_p))
                        continue
                except Exception:
                    pass
                if not src_p.exists():
                    photos_missing.append(src_p.name)
                    log(f"SAVE photo missing at source: {src}")
                    continue
                ext=src_p.suffix.lower() or ".jpg"
                wo_tag=str(self.loaded_wo).replace("/","_").replace("\\","_")
                nm=f"{wo_tag}_S{next_seq:03d}{ext}"
                dest=photo_dir/nm
                # If somehow a file already exists at dest, bump sequence until unique
                while dest.exists():
                    next_seq+=1
                    nm=f"{wo_tag}_S{next_seq:03d}{ext}"
                    dest=photo_dir/nm
                shutil.move(str(src_p), str(dest))
                moved_photos.append(str(dest))
                log(f"SAVE photo moved: {src} -> {dest}")
                next_seq+=1
            except Exception as e:
                photos_failed.append(f"{Path(src).name}: {e}")
                log(f"SAVE photo move FAILED: {src}: {e}\n{traceback.format_exc()}")
        log(f"SAVE photos done: moved={len(moved_photos)} missing={len(photos_missing)} failed={len(photos_failed)}")
        # Update in-memory list so subsequent Save calls see the new paths
        self.pending_photos=list(moved_photos)
        self._uc()

        # Sort the whole shipment's serials into increasing order now that scanning
        # is done — flows into the draft file and the draft-save history record.
        self.serials = sort_serials(self.serials)
        # Persist draft with the NEW photo paths
        prog.step("Writing the draft file…")
        op_clean=op if not op.startswith("—") else ""
        # Each shipment session saves to its OWN draft file (multi-draft): a new
        # session mints its id on first save; a resumed draft keeps its own.
        if self.current_draft_id is None:
            self.current_draft_id=f"{datetime.datetime.now():%Y%m%d_%H%M%S}"
        out=save_shipment_draft(self.wo_folder, self.pending_photos, self.serials,
                                self.coc_data, self.packing_data, operator=op_clean,
                                notes=self.shipment_notes, draft_id=self.current_draft_id,
                                slip_data=self.slip_data,
                                started_at=self._started_at, started_by=self._started_by)
        if not out:
            prog.close()
            messagebox.showerror("Save failed","Could not save draft.\n\nCheck folder permissions and try again.")
            log("SAVE draft FAILED")
            return
        log(f"SAVE draft written: {out}")
        # Skip the share history write for an offline draft — the draft file itself
        # is already saved locally in the Outbox, and touching HISTORY_DIR on the
        # dead share would freeze the UI for the full SMB timeout.
        if not self._offline_shipment:
            prog.step("Recording in shipment history…")
            write_history_event("draft_save", {
                "wo": self.loaded_wo, "division": self.division_var.get(),
                "operator": op_clean, "wo_folder": str(self.wo_folder),
                "customer": self.parsed_info.get("customer", "N/A"),
                "pn": self.parsed_info.get("pn", "N/A"),
                "eng": self.parsed_info.get("eng", "N/A"),
                "serial_count": len(self.serials), "serials": list(self.serials),
                "notes": self.shipment_notes or "", "dup_overrides": list(self._dup_overrides),
                "draft_id": self.current_draft_id or "",
                "started_at": self._started_at or "", "started_by": self._started_by or "",
                "coc": self.coc_data or None, "packing": self.packing_data or None,
                "slip": self.slip_data or None,
                "photos": {"moved_count": len(moved_photos),
                           "missing": list(photos_missing), "failed": list(photos_failed)},
                "draft_path": out,
                "resumed_from_draft": self.resumed_from_draft,
            })

        # Build summary (progress comes down first so the result box stands alone)
        prog.close()
        summary_lines=[f"✓ Shipment saved as draft.",""]
        if moved_photos: summary_lines.append(f"Photos in WO folder: {len(moved_photos)}")
        if photos_missing:
            summary_lines.append(f"  ⚠ {len(photos_missing)} photo(s) missing from source (skipped):")
            for p in photos_missing[:5]:
                summary_lines.append(f"     • {p}")
            if len(photos_missing) > 5:
                summary_lines.append(f"     • ... and {len(photos_missing)-5} more")
        if photos_failed:
            summary_lines.append(f"  ⚠ {len(photos_failed)} photo(s) failed to move:")
            for p in photos_failed[:5]:
                summary_lines.append(f"     • {p}")
        summary_lines += ["",
            f"To resume: reload WO {self.loaded_wo} from the same division.",
            "Supervisor can complete and close from any network PC.",
            "", f"Saved to:", f"  {out}"]
        if photos_missing or photos_failed:
            messagebox.showwarning("Draft saved (with photo issues)", "\n".join(summary_lines))
        else:
            messagebox.showinfo("Draft saved", "\n".join(summary_lines))
        self._reset_form()
    def _reset_form(self):
        """Reset the main app state after Save or Close completes."""
        self.pics_path=self.serials_path=self.wo_folder=None
        self.pending_photos=[]; self.serials=[]
        self.packing_data=None; self.coc_data=None; self.slip_data=None
        self.loaded_wo=None; self.parsed_info={}; self.current_wo.set("")
        self.resumed_from_draft=None; self._offline_shipment=False; self._monday_hit=None
        self.shipment_notes=""; self.current_draft_id=None; self._reset_dup_state()
        self.lbl_div.config(text="—",fg=C_DARK); self.lbl_wo.config(text="—",fg=C_DARK)
        self.lbl_cust.config(text="—",fg=C_DARK); self.lbl_pn.config(text="—",fg=C_DARK)
        self.lbl_eng.config(text="—",fg=C_DARK); self.lbl_fp.config(text="—",fg=C_DARK)
        self.lbl_st.config(text="Scan a work order to begin",fg=C_SUBTLE)
        self._uc(); self.wo_entry.focus_set()
    def _write_with_retry(self, file_label, target_path, write_callable, log):
        """
        Call write_callable(). If it raises PermissionError (file open in Excel
        on another machine), show FileInUseDialog and respond to user choice:
          - Retry  -> call write_callable() again
          - Skip   -> return ("skip", None, None)
          - Cancel -> return ("cancel", None, None)
        Other exceptions return ("error", None, "<error string>") immediately.
        Success returns ("ok", <return_value>, None).
        """
        max_attempts = 10  # safety cap on retries
        for attempt in range(max_attempts):
            try:
                result = write_callable()
                if attempt > 0:
                    log(f"{file_label} WRITE succeeded on retry #{attempt}")
                return ("ok", result, None)
            except PermissionError as e:
                err_text = f"{type(e).__name__}: {e}"
                log(f"{file_label} WRITE locked (attempt {attempt+1}): {err_text}")
                dlg = FileInUseDialog(self, file_label, target_path, err_text)
                if dlg.result == "retry":
                    continue
                elif dlg.result == "skip":
                    log(f"{file_label} WRITE skipped by user")
                    return ("skip", None, None)
                else:  # cancel
                    log(f"{file_label} WRITE cancel-close by user")
                    return ("cancel", None, None)
            except Exception as e:
                err_text = f"{type(e).__name__}: {e}"
                import traceback as _tb
                log(f"{file_label} WRITE failed: {err_text}\n{_tb.format_exc()}")
                return ("error", None, err_text)
        # Hit attempt cap
        log(f"{file_label} WRITE gave up after {max_attempts} attempts")
        return ("error", None, f"Exceeded {max_attempts} retry attempts")

    def _close_shipment(self):
        if not self.pics_path: messagebox.showwarning("No WO","No work order loaded."); return
        wo=self.loaded_wo; div=self.division_var.get(); op=self.operator_var.get()
        # Close-as-DELETE: if photos and serials are both empty, the operator/supervisor
        # is choosing to discard this shipment entirely. Common scenario: supervisor
        # reviewed a draft, found bad data, cleared photos/serials, now wants to wipe.
        # COC/Packing data being filled doesn't change this — without photos/serials
        # there's no shipment to document.
        if not self.pending_photos and not self.serials:
            coc_status="filled" if self.coc_data else "not filled"
            pl_status="filled" if self.packing_data else "not filled"
            # Shipping Slip / Labels pages (Dhruvil 2026-09-24) are written at
            # PRINT time, not at close — ones already printed stay in the WO
            # folder (never touched); say so and record where they are.
            slip_out=((self.slip_data or {}).get("outputs") or {})
            slip_pages=[str(p) for p in slip_out.values() if p]
            slip_note=("\n\nThe Shipping Slip / Labels page(s) already printed stay in the WO folder:\n  "
                       +"\n  ".join(slip_pages)) if slip_pages else ""
            del_msg=(f"This shipment has no photos and no serial numbers.\n\n"
                     f"WO: {wo}\n"
                     f"COC: {coc_status}  ·  Packing List: {pl_status}\n\n"
                     f"Continuing will DELETE this shipment entirely. Any saved draft "
                     f"(including COC and Packing List data) will be discarded. "
                     f"No documents will be generated.{slip_note}\n\n"
                     f"This cannot be undone.")
            confirm=ConfirmActionDialog(self,
                title="Delete Shipment", heading="Delete this shipment?", message=del_msg,
                confirm_label="Yes, Delete", cancel_label="Cancel",
                heading_color=C_RED, confirm_color=C_RED)
            if not confirm.result: return
            # Discard: remove any saved draft sidecar, then reset the form
            ok, derr = delete_shipment_draft(self.wo_folder, draft_id=self.current_draft_id)
            if not ok:
                log_shipment_event(self.wo_folder, f"DELETE-shipment branch: draft delete failed: {derr}")
                messagebox.showwarning("Draft not removed",
                    f"Shipment was marked for deletion but the draft file could not be removed:\n\n"
                    f"  {self.wo_folder}\\{DRAFT_FILENAME}\n\n"
                    f"Reason: {derr}\n\n"
                    f"The shipment will still be reset locally. To prevent it from "
                    f"resuming on next reload, manually delete the file above.")
            # Offline: skip the share history write (would freeze on the dead
            # HISTORY_DIR); discarding an empty offline shipment needs no record.
            if not self._offline_shipment:
                write_history_event("delete", {
                    "wo": wo, "division": div,
                    "operator": op if not op.startswith("—") else "",
                    "wo_folder": str(self.wo_folder),
                    "customer": self.parsed_info.get("customer", "N/A"),
                    "pn": self.parsed_info.get("pn", "N/A"),
                    "eng": self.parsed_info.get("eng", "N/A"),
                    "deleted": {"had_coc": bool(self.coc_data), "had_packing": bool(self.packing_data),
                                "had_slip": bool(self.slip_data),
                                "slip_outputs": slip_out or None,
                                "slip_printed": (self.slip_data or {}).get("printed") or None,
                                "draft_removed": ok, "draft_error": derr},
                    "resumed_from_draft": self.resumed_from_draft,
                })
            messagebox.showinfo("Shipment deleted",
                f"Shipment for WO {wo} has been deleted.\n\nNo documents were generated.{slip_note}")
            self._reset_form()
            return
        # Require an operator before generating documents, so the placeholder
        # "— Select your name —" can never be written into the serials xlsx.
        if op.startswith("—"):
            messagebox.showwarning("No operator",
                "Please select your name before closing the shipment.")
            try: self.op_combo.focus_set()
            except Exception: pass
            return
        # Confirmation popup before destructive Close action
        n_photos=len(self.pending_photos); n_serials=len(self.serials)
        coc_status="filled" if self.coc_data else "not filled"
        pl_status="filled" if self.packing_data else "not filled"
        op_show=op if not op.startswith("—") else "(no operator selected)"
        msg=(f"Close this shipment and generate all final documents?\n\n"
             f"WO: {wo}\n"
             f"Operator: {op_show}\n"
             f"Photos: {n_photos}  ·  Serials: {n_serials}\n"
             f"COC: {coc_status}  ·  Packing List: {pl_status}\n\n"
             f"This will move photos into the WO folder and generate the "
             f"serials, COC, and Packing List documents.\n\n"
             f"This action cannot be undone.")
        # Amber header = "stop and check" (matches the app's severity language for
        # irreversible actions); the confirm button stays green (green = commit).
        confirm=ConfirmActionDialog(self,
            title="Save and Close Shipment", heading="Close this shipment?", message=msg,
            confirm_label="Yes, Save and Close", cancel_label="Cancel",
            heading_color=C_AMBER, confirm_color=C_GREEN)
        if not confirm.result: return
        # Sort the whole shipment's serials into increasing order now that scanning
        # is done — flows into the serials xlsx, COC, packing list, history, and the
        # safety re-save below. (Scanning kept them in raw scan order.)
        self.serials = sort_serials(self.serials)
        # ───── Serial-list format asked UP FRONT (before anything is written or
        # moved) so Cancel / the window-X here is a true "go back": the shipment
        # is NOT closed and nothing is generated.
        want_word=False
        if self.serials:
            ed=ExportDialog(self,wo,len(self.serials))
            if ed.result is None:   # Cancel / Escape / window-X -> abort the close
                return
            want_word=(ed.result=="both")
        # Long synchronous phase: photo moves + all document writes to the share
        # can take 30s+ — show a live progress box (per-step text) instead of a
        # frozen window. try/finally guarantees the box always comes down.
        prog=ProgressDialog(self, "Saving and closing…",
                            f"Closing shipment WO {wo} — generating documents")
        try:
            self._close_shipment_work(wo, div, op, want_word, prog)
        finally:
            prog.close()
    def _close_shipment_work(self, wo, div, op, want_word, prog):
        """The slow part of Save and Close (photo moves + serials/COC/Packing
        writes + history), run under the ProgressDialog owned by _close_shipment."""
        # ───── Diagnostic logging — recorded to <wo_folder>/.shipping_log.txt
        # so any silent failures can be traced after the fact.
        import traceback
        log = lambda msg: log_shipment_event(self.wo_folder, msg)
        log(f"CLOSE start: wo={wo} div={div} op={op} "
            f"photos={len(self.pending_photos)} serials={len(self.serials)} "
            f"coc={'yes' if self.coc_data else 'no'} pl={'yes' if self.packing_data else 'no'}")

        ts=datetime.datetime.now(); date_folder=ts.strftime("%m%d")+ts.strftime("%y")

        # ───── Photo move loop — for any photos NOT already inside the WO
        # folder (e.g. operator added new photos after a Save and went straight
        # to Close, or this is a fresh non-draft shipment). Photos already
        # moved at Save time are kept in place and counted toward the total.
        # New photos at Close go into a fresh date-based folder (timestamped
        # if same-day-second-shipment).
        sp=[]; moved_full=[]; photos_missing=[]; photos_failed=[]
        orig_photo_count=len(self.pending_photos)
        _pics_root = str(Path(self.pics_path).resolve())   # resolve the UNC root ONCE
        external_photos = [p for p in self.pending_photos
                           if not str(Path(p).resolve()).startswith(_pics_root)]
        if external_photos:
            photo_dir = unique_dir_path(self.pics_path, date_folder)
            if not ensure_folder(photo_dir):
                photo_dir = self.pics_path
            log(f"PHOTOS new dir at close: {photo_dir}")
        else:
            photo_dir = None  # no new photos to place
        if self.pending_photos:
            pics_root=str(Path(self.pics_path).resolve())
            already_in_wo=0
            seq=1
            for _pi, src in enumerate(self.pending_photos, 1):
                prog.step(f"Moving photo {_pi} of {orig_photo_count} into the WO folder…")
                try:
                    src_p=Path(src)
                    # Skip photos already inside the WO Shipping Pictures area
                    try:
                        if str(src_p.resolve()).startswith(pics_root):
                            sp.append(src_p.name); moved_full.append(str(src_p)); already_in_wo+=1
                            continue
                    except Exception:
                        pass
                    if not src_p.exists():
                        photos_missing.append(src_p.name)
                        continue
                    if photo_dir is None:
                        photos_failed.append(f"{src_p.name}: no destination folder")
                        continue
                    ext=src_p.suffix.lower() or ".jpg"
                    # Use the actual folder name (may include timestamp) as the file prefix
                    prefix = photo_dir.name
                    nm=f"{prefix}_{seq:03d}{ext}"
                    dest=photo_dir/nm
                    while dest.exists():
                        seq+=1
                        nm=f"{prefix}_{seq:03d}{ext}"
                        dest=photo_dir/nm
                    shutil.move(str(src_p),str(dest))
                    sp.append(nm); moved_full.append(str(dest)); seq+=1
                except Exception as e:
                    photos_failed.append(f"{Path(src).name}: {e}")
                    moved_full.append(str(src))
                    log(f"PHOTO move failed: {src}: {e}")
            log(f"PHOTOS: total={len(sp)} (already_in_wo={already_in_wo}, "
                f"missing={len(photos_missing)}, failed={len(photos_failed)})")

        # ───── Persist post-move state as a safety draft.
        # Photos have now physically moved into the WO folder, so the in-memory
        # list and the on-disk draft must reference their NEW paths — otherwise a
        # later cancel/failure (or a crash) would leave the draft pointing at the
        # vanished source paths and every photo would report as "missing" on
        # resume. Re-saving here makes every downstream abort path recoverable and
        # makes the "your draft is still saved" messages truthful. On a clean
        # close this draft is deleted at the end.
        op_clean=op if not op.startswith("—") else ""
        self.pending_photos=list(moved_full)
        self._uc()
        prog.step("Saving safety draft…")
        if self.current_draft_id is None:      # never-drafted close still gets its own file
            self.current_draft_id=f"{datetime.datetime.now():%Y%m%d_%H%M%S}"
        save_shipment_draft(self.wo_folder, self.pending_photos, self.serials,
                            self.coc_data, self.packing_data, operator=op_clean,
                            notes=self.shipment_notes, draft_id=self.current_draft_id,
                            slip_data=self.slip_data,
                            started_at=self._started_at, started_by=self._started_by)

        # ───── Determine engineering # for filenames (must come BEFORE writes).
        # Priority: parsed from folder path > COC entry > Packing List entry > WO# as fallback.
        eng = (self.parsed_info.get("eng") if self.parsed_info else None) or ""
        if not eng or eng == "N/A":
            eng = (self.coc_data or {}).get("job", "") or ""
        if not eng:
            eng = (self.packing_data or {}).get("eng", "") or ""
        if not eng:
            eng = wo  # final fallback — at least the file is named with the WO
        # Sanitize for filenames: strip ALL Windows-illegal characters (not just
        # slashes) so a stray ':' '?' '*' etc. in a free-text field can't abort
        # every document write. Falls back to the WO number if nothing remains.
        eng_safe = sanitize_filename_part(eng, fallback=wo)
        log(f"FILENAME ENG#: {eng_safe} (raw='{eng}')")

        # ───── Serials xlsx — always a NEW file (no append, no lock conflicts)
        serials_file=None; serials_err=None; serials_skipped=False; serials_cancelled=False
        serials_pdf=None; serials_pdf_err=None
        if self.serials:
            prog.step(f"Writing the serials Excel ({len(self.serials)} serials)…")
            fdate = fmt_file_date()
            target_preview = self.serials_path / f"{eng_safe}-{fdate} serials.xlsx"
            status, value, err = self._write_with_retry(
                "Serials", target_preview,
                lambda: write_serial_xlsx(self.serials_path, eng_safe, wo, self.serials, op),
                log)
            prog.regrab()
            if status == "ok":
                serials_file = value
                log(f"SERIALS xlsx written: {serials_file}")
                serials_pdf, serials_pdf_err = self._write_pdf("SERIALS",
                    lambda: pdf_docs.write_serials_pdf(serials_file, eng_safe, wo, self.serials, fmt_date(), fmt_time(), op), log)
            elif status == "skip":
                serials_skipped = True
            elif status == "cancel":
                serials_cancelled = True
            else:
                serials_err = err

        # If user explicitly cancelled the close at the lock dialog, abort now —
        # keep draft, keep state, let them retry later. With the new "always new
        # filename" scheme this should be extremely rare (the only way a lock
        # could occur is if the EXACT same engineering # + date combo collided).
        if serials_cancelled:
            log("CLOSE aborted by user (serials cancel)")
            prog.close()
            messagebox.showinfo("Close cancelled",
                "Shipment close was cancelled.\n\n"
                "Your draft is still saved. You can retry Save and Close once "
                "the file is no longer in use.")
            return

        # ───── Optional Word export of serials — format was chosen up front
        # (want_word); here we just write it if requested and the xlsx succeeded.
        wp=None; wp_err=None
        if self.serials and serials_file and want_word:
            prog.step("Writing the serials Word document…")
            try:
                wp=write_word_serials(self.serials_path,eng_safe,wo,self.serials)
                log(f"SERIALS word written: {wp}")
            except Exception as e:
                wp_err=f"{type(e).__name__}: {e}"
                log(f"SERIALS word FAILED: {wp_err}")

        # ───── COC xlsx — wrap with retry-on-lock + skip handling
        cf=None; coc_err=None; coc_cancelled=False; coc_skipped=False
        coc_pdf=None; coc_pdf_err=None
        if self.coc_data:
            prog.step("Writing the Certificate of Compliance…")
            try:
                cp=self.wo_folder/COC_FOLDER
                if ensure_folder(cp):
                    fd=fmt_file_date()
                    target_preview = cp / f"{eng_safe}-{fd}-COC.xlsx"
                    status, value, err = self._write_with_retry(
                        "COC", target_preview,
                        lambda: write_coc_xlsx(cp, eng_safe, wo, self.coc_data, self.serials),
                        log)
                    prog.regrab()
                    if status == "ok":
                        cf = value
                        log(f"COC written: {cf}")
                        coc_pdf, coc_pdf_err = self._write_pdf("COC",
                            lambda: pdf_docs.write_coc_pdf(cf, self.coc_data, self.serials), log)
                    elif status == "skip":
                        coc_skipped = True
                    elif status == "cancel":
                        # Cancelling here aborts the rest of the close
                        log("CLOSE aborted by user (COC cancel)")
                        prog.close()
                        messagebox.showinfo("Close cancelled",
                            "Shipment close was cancelled.\n\n"
                            "Your draft is still saved.")
                        return
                    else:
                        coc_err = err
            except Exception as e:
                coc_err=f"{type(e).__name__}: {e}"
                log(f"COC FAILED: {coc_err}\n{traceback.format_exc()}")

        # ───── Packing List xlsx — always a NEW file
        sf=None; pl_err=None; pl_cancelled=False; pl_skipped=False
        pl_pdf=None; pl_pdf_err=None
        if self.packing_data:
            prog.step("Writing the Packing List…")
            try:
                plp=self.wo_folder/PACKING_LIST_FOLDER
                if ensure_folder(plp):
                    fd=fmt_file_date()
                    target_preview = plp / f"{eng_safe}-{fd}-Packing_List.xlsx"
                    status, value, err = self._write_with_retry(
                        "Packing List", target_preview,
                        lambda: write_packing_list_xlsx(plp, eng_safe, wo, self.packing_data, self.serials, logo_path=str(LOGO_PATH)),
                        log)
                    prog.regrab()
                    if status == "ok":
                        sf = value
                        log(f"PACKING LIST written: {sf}")
                        pl_pdf, pl_pdf_err = self._write_pdf("PACKING LIST",
                            lambda: pdf_docs.write_packing_list_pdf(sf, self.packing_data, self.serials, logo_path=str(LOGO_PATH)), log)
                    elif status == "skip":
                        pl_skipped = True
                    elif status == "cancel":
                        log("CLOSE aborted by user (Packing List cancel)")
                        prog.close()
                        messagebox.showinfo("Close cancelled",
                            "Shipment close was cancelled.\n\n"
                            "Your draft is still saved.")
                        return
                    else:
                        pl_err = err
            except Exception as e:
                pl_err=f"{type(e).__name__}: {e}"
                log(f"PACKING LIST FAILED: {pl_err}\n{traceback.format_exc()}")

        # ───── Customer Portal: queue the CoC / Packing List / Serial list PDFs.
        # Industrial and MAG - ITAR jobs (Sahil 2026-11-07). enqueue_shipment() refuses
        # files outside the division's share folder and anything marked export controlled.
        # Only writes a small job file; a background thread uploads. Never blocks the close.
        if self._offline_shipment:
            portal_status, portal_msg = "offline", "Not sent (offline shipment — file it on the share first)."
        else:
            prog.step("Queuing the Customer Portal upload…")
            portal_status, portal_msg = portal_upload.enqueue_shipment(
                div, self.wo_folder, wo, eng_safe,
                coc_pdf=coc_pdf, coc_data=self.coc_data,
                packing_pdf=pl_pdf, packing_data=self.packing_data,
                serials=self.serials, serials_pdf=serials_pdf)
        log(f"PORTAL {portal_status}: {portal_msg}")

        # ───── Build summary, surfacing real errors prominently
        n_serials=len(self.serials)
        errors=[]
        if serials_err: errors.append(f"  Serials xlsx: {serials_err}")
        if coc_err:     errors.append(f"  COC: {coc_err}")
        if pl_err:      errors.append(f"  Packing List: {pl_err}")
        if wp_err:      errors.append(f"  Serials Word: {wp_err}")
        if serials_pdf_err: errors.append(f"  Serials PDF: {serials_pdf_err}")
        if coc_pdf_err:     errors.append(f"  COC PDF: {coc_pdf_err}")
        if pl_pdf_err:      errors.append(f"  Packing List PDF: {pl_pdf_err}")
        def pdf_note(xlsx, pdf):
            if not xlsx: return ""
            return "  + PDF ✓" if pdf else "  + PDF FAILED"

        # Build status label per file. Possible states:
        # ✓ written / — n/a (no data) / — skipped (user) / FAILED (unhandled error)
        def status_label(have_data, ok_value, cancelled, skipped):
            if not have_data: return "—"
            if ok_value: return "✓"
            if cancelled: return "— skipped (append/overwrite cancelled)"
            if skipped:   return "— skipped (file in use)"
            return "FAILED"

        # Critical writes succeeded? Skipped files count as user-acknowledged,
        # not failures. They will be treated like done for draft-cleanup purposes.
        critical_ok = True
        if self.serials and not serials_file and not serials_skipped: critical_ok = False
        if self.coc_data and not cf and not coc_cancelled and not coc_skipped: critical_ok = False
        if self.packing_data and not sf and not pl_cancelled and not pl_skipped: critical_ok = False

        slip_out = ((self.slip_data or {}).get("outputs") or {})   # Shipping Slip / Labels page paths (Dhruvil 2026-09-24)
        close_record = {
            "wo": wo, "division": div,
            "operator": op if not op.startswith("—") else "",
            "wo_folder": str(self.wo_folder),
            "customer": self.parsed_info.get("customer", "N/A"),
            "pn": self.parsed_info.get("pn", "N/A"),
            "eng": self.parsed_info.get("eng", "N/A"),
            "eng_safe": eng_safe,
            "serial_count": len(self.serials), "serials": list(self.serials),
            "notes": self.shipment_notes or "", "dup_overrides": list(self._dup_overrides),
            "coc": self.coc_data or None, "packing": self.packing_data or None,
            # Shipping Slip & Labels (Dhruvil 2026-09-24): the form as printed
            # (codes, box table, printed times, page paths) — the History copy
            # HistoryDetail shows and reemit_paperwork rebuilds from.
            "slip": self.slip_data or None,
            "photos": {"moved_count": len(sp), "moved": list(sp),
                       "missing": list(photos_missing), "failed": list(photos_failed),
                       "photo_dir": str(photo_dir) if photo_dir else None},
            "outputs": {"serials_xlsx": str(serials_file) if serials_file else None,
                        "serials_docx": str(wp) if wp else None,
                        "coc_xlsx": str(cf) if cf else None,
                        "packing_xlsx": str(sf) if sf else None,
                        "serials_pdf": str(serials_pdf) if serials_pdf else None,
                        "coc_pdf": str(coc_pdf) if coc_pdf else None,
                        "packing_pdf": str(pl_pdf) if pl_pdf else None,
                        "slip_html": slip_out.get("slip_html") or None,
                        "labels_html": slip_out.get("labels_html") or None,
                        "pack_html": slip_out.get("pack_html") or None},
            "file_status": {
                "serials": "ok" if serials_file else ("skipped_in_use" if serials_skipped else (("failed: " + serials_err) if serials_err else "n/a")),
                "coc": "ok" if cf else ("cancelled" if coc_cancelled else ("skipped_in_use" if coc_skipped else (("failed: " + coc_err) if coc_err else "n/a"))),
                "packing": "ok" if sf else ("cancelled" if pl_cancelled else ("skipped_in_use" if pl_skipped else (("failed: " + pl_err) if pl_err else "n/a"))),
                "word": "ok" if wp else (("failed: " + wp_err) if wp_err else "n/a"),
            },
            "critical_ok": critical_ok, "errors": list(errors),
            "portal": {"status": portal_status, "detail": portal_msg},
            "resumed_from_draft": self.resumed_from_draft,
            "draft_id": self.current_draft_id or "",
            "started_at": self._started_at or "", "started_by": self._started_by or "",
        }
        prog.step("Recording in shipment history…")
        report_update = None   # Shipping Report auto-update outcome, shown + recorded below
        if self._offline_shipment:
            # Offline: the share HISTORY_DIR is unreachable. Save the close record
            # (marked source='offline', with the offline timestamp) into the Outbox
            # as history_close.json, plus a plain-text sync README. Both travel with
            # the folder so the shipment can be filed correctly once PCC2 is back.
            hist_path=None; ts=datetime.datetime.now()
            rec={"schema_version":1,"event_type":"close","source":"offline",
                 "ts":ts.strftime("%Y-%m-%d %H:%M:%S"),"ts_epoch":time.time(),
                 "station":os.environ.get("COMPUTERNAME",""),"win_user":os.environ.get("USERNAME",""),
                 "app_version":APP_VERSION}
            rec.update(close_record)
            try: _atomic_write_json(self.wo_folder/"history_close.json", rec)
            except Exception: pass
            readme_rows=[("Serials xlsx",serials_file),("Serials Word",wp),("COC xlsx",cf),("Packing List xlsx",sf)]
            if self.slip_data:      # only when the slip popup was used — an older-style shipment's README stays as it was
                readme_rows+=[("Shipping Slip html",slip_out.get("slip_html")),("Shipping Labels html",slip_out.get("labels_html"))]
                if slip_out.get("pack_html"): readme_rows.append(("Packing Label html",slip_out.get("pack_html")))
            write_outbox_readme(self.wo_folder, div, wo, op, readme_rows)
        else:
            hist_path = write_history_event("close", close_record)
            # Keep the job's Shipping Report current after every shipment — SYNCHRONOUS
            # now and SHOWING the result (Dhruvil 2026-10-24: the silent daemon thread
            # didn't reliably run — it was killed when the app moved on). History-only
            # (~a few seconds); the close is already in History above, so the refresh
            # includes this shipment. report_update stays None for a non-test-and-return
            # job (no report line is shown then).
            prog.step("Updating the Shipping Report…")
            try:
                report_update = auto_update_shipping_report(
                    str(close_record.get("eng") or ""), str(close_record.get("pn") or ""))
            except Exception as e:
                report_update = ("failed", f"{type(e).__name__}: {e}")
            _record_report_update(hist_path, report_update)

        msg_lines = [
            f"Shipment closed{'  ✓' if critical_ok and not errors else '  ⚠'}",
            "",
            f"WO: {wo}    Div: {div}",
            f"Operator: {op}",
            "",
            f"Photos moved:    {len(sp)} of {orig_photo_count}",
        ]
        if photos_missing:
            msg_lines.append(f"  • {len(photos_missing)} photo(s) MISSING from source (not moved)")
        if photos_failed:
            msg_lines.append(f"  • {len(photos_failed)} photo(s) FAILED to move")
        msg_lines += [
            f"Serials xlsx:    {status_label(bool(self.serials), serials_file, False, serials_skipped)}{pdf_note(serials_file, serials_pdf)}",
            f"Serials Word:    {'✓' if wp else '—'}",
            f"COC:             {status_label(bool(self.coc_data), cf, coc_cancelled, coc_skipped)}{pdf_note(cf, coc_pdf)}",
            f"Packing List:    {status_label(bool(self.packing_data), sf, pl_cancelled, pl_skipped)}{pdf_note(sf, pl_pdf)}",
            f"Customer Portal: {portal_msg}",
        ]
        if report_update:
            _rs, _rd = report_update
            if _rs == "updated":
                msg_lines.append("Shipping Report: updated  ✓")
            elif _rs == "none":
                msg_lines.append(f"Shipping Report: {_rd}")
            else:   # unknown / failed
                msg_lines.append(f"Shipping Report: NOT updated — {_rd}")
        if errors:
            msg_lines += ["", "ERRORS:"] + errors
        if self._offline_shipment:
            msg_lines[0]=f"OFFLINE shipment ready to PRINT{'  ✓' if critical_ok and not errors else '  ⚠'}"
            msg_lines += ["", "⚠  This shipment was made OFFLINE. The documents are in the local Outbox:",
                          f"  {self.wo_folder}", "",
                          "1. PRINT the documents from that folder and ship the box.",
                          f"2. When PCC2 is back, sync this folder into the real WO folder for {wo}",
                          "   (the '_OFFLINE_SHIPMENT_README.txt' inside explains how; 'Open Outbox' opens it)."]
        else:
            msg_lines += ["", f"Saved to:", f"  {self.wo_folder}"]
            if hist_path is None:
                msg_lines.append("note: history record could not be written")
        msg = "\n".join(msg_lines)

        prog.close()   # done working — the result box stands alone
        if errors:
            messagebox.showwarning("Close completed with errors", msg)
        else:
            messagebox.showinfo("Done", msg)

        # ───── Draft cleanup logic
        # critical_ok=True means all writes either succeeded or were explicitly
        # skipped by the user. Either way, the draft has served its purpose —
        # delete it. The user made informed choices about what to skip.
        #
        # critical_ok=False means at least one write hit an unhandled error
        # (not a lock, not a user skip). Offer the choice: discard the draft
        # (user accepts the loss) or keep it (try again later).
        if critical_ok:
            ok, derr = delete_shipment_draft(self.wo_folder, draft_id=self.current_draft_id)
            if not ok:
                log(f"DRAFT delete failed: {derr}")
                messagebox.showwarning("Draft not removed",
                    f"Shipment closed successfully, but the draft file could not be deleted:\n\n"
                    f"  {self.wo_folder}\\{DRAFT_FILENAME}\n\n"
                    f"Reason: {derr}\n\n"
                    f"Manually delete the file above to prevent this WO from auto-resuming on next load.\n\n"
                    f"All shipment documents were saved successfully — this is only the draft sidecar.")
            else:
                log("DRAFT deleted")
        else:
            # Unhandled error path — ask the user what to do with the draft
            keep = ConfirmActionDialog(self,
                title="Keep draft for retry?",
                heading="Some files didn't get saved.",
                message=("Would you like to keep the draft so you can retry the "
                         "failed files later?\n\n"
                         "Keep Draft: leaves the draft in place so you can reload "
                         "the WO and try Close again.\n\n"
                         "Discard Draft: deletes the draft. Any successfully-written "
                         "files (COC, Packing List, etc.) stay in the WO folder — "
                         "they are not removed."),
                confirm_label="Keep Draft", cancel_label="Discard Draft",
                heading_color=C_AMBER, confirm_color=C_TEAL)
            if keep.result:
                # Re-persist the current in-memory state so the kept draft is
                # actually up to date. Do NOT reset the form if this fails —
                # the operator's photos/serials/forms must stay on screen so
                # nothing is lost.
                saved=save_shipment_draft(self.wo_folder, self.pending_photos,
                                          self.serials, self.coc_data,
                                          self.packing_data, operator=op_clean,
                                          notes=self.shipment_notes,
                                          draft_id=self.current_draft_id,
                                          slip_data=self.slip_data,
                                          started_at=self._started_at, started_by=self._started_by)
                if saved:
                    log("DRAFT kept (user choice — will retry)")
                else:
                    log("DRAFT keep FAILED to persist")
                    messagebox.showwarning("Draft not saved",
                        "Could not save the draft to the WO folder.\n\n"
                        "Your photos and serial numbers are still on screen — "
                        "do NOT close the app. Fix the network or folder issue, "
                        "then use Save Shipment again.")
                    return  # keep everything on screen; skip _reset_form
            else:
                ok, derr = delete_shipment_draft(self.wo_folder, draft_id=self.current_draft_id)
                if ok:
                    log(f"DRAFT discarded by user{' (note: ' + derr + ')' if derr else ''}")
                else:
                    log(f"DRAFT discard failed: {derr}")
                    messagebox.showwarning("Draft not removed",
                        f"Could not delete the draft file:\n\n"
                        f"  {self.wo_folder}\\{DRAFT_FILENAME}\n\n"
                        f"Reason: {derr}\n\n"
                        f"Manual delete required.")

        log("CLOSE end")
        self._reset_form()

    # ── Legacy short-name aliases ─────────────────────────────────────────
    # The methods above were renamed for readability (2026-07-22 cleanup);
    # these aliases keep every older reference, note and test script working.
    _lw=_load_workorder; _fin=_close_shipment; _uc=_update_status_chips
    _sd=_set_division; _sf=_build_footer
    _pp=_open_photo_popup; _sp=_open_serial_popup
    _cp=_open_coc_popup; _plp=_open_packing_popup
    _oa=_open_admin; _oas=_on_admin_saved

if __name__=="__main__":
    ShippingApp().mainloop()