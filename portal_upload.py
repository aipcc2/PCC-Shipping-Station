"""
Sends shipment CoC and Packing List PDFs to the PCC Customer Portal.

How it works
  1. After Save and Close writes the CoC / Packing List .xlsx, the app calls enqueue_shipment().
     That only writes a small job file on this PC, so shipping is never slowed down or blocked.
  2. A background thread (PortalUploader) picks up each job: it turns the .xlsx into a .pdf next to it
     (Excel does the conversion, so the PDF looks exactly like the printed form), then sends the PDF to
     the portal's ingest-shipping-docs function with this station's id and key.
  3. If the portal or the network is down, the job waits and retries (1 min, 5 min, 15 min, then
     hourly, for up to 7 days). Jobs the portal refuses for a reason that won't fix itself are marked
     failed and shown to the operator.

ITAR (hard rule): nothing from a MAG - ITAR job is ever sent. The guard (itar_reason) runs when the job
is queued, again before the PDF is made, and again right before upload, so even an edited queue file
can't get an ITAR job out. The portal refuses ITAR jobs as well.

Settings live on each PC (never on the network share, never in git):
  %LOCALAPPDATA%\\PCC Shipping\\portal_upload.json  ->  {"enabled", "station_id", "station_key", "url"}
Upload is OFF until an admin turns it on in Admin -> Portal upload.
"""

import datetime
import json
import os
import re
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
import uuid
from pathlib import Path

DEFAULT_URL = "https://jefedoraeqoevggxevgd.supabase.co/functions/v1/ingest-shipping-docs"
ALLOWED_DIVISION = "Industrial"
INDUSTRIAL_ROOT = r"\\PCC2\PCC-I-Shares\Production\Industrial"
ITAR_ROOT = r"\\PCC2\PCC-I-Shares\Production\MAG - ITAR"

RETRY_DELAYS = [60, 300, 900, 3600]   # seconds; after the last one, hourly
GIVE_UP_AFTER = 7 * 24 * 3600         # a job older than this is marked failed
EXPORT_TIMEOUT = 120                  # seconds Excel gets to make one PDF
HTTP_TIMEOUT = 60

DOC_COC = "CoC"
DOC_PACKING = "Packing slip"
DOC_SERIALS = "Serial list"


def local_dir():
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "PCC Shipping"


def queue_dir():
    return local_dir() / "portal_queue"


def config_path():
    return local_dir() / "portal_upload.json"


###############################################################################
# ITAR GUARD
###############################################################################

_ITAR_WORDS = [
    (re.compile(r"(^|[^a-z])itar([^a-z]|$)", re.I), "ITAR"),
    (re.compile(r"international\s+traffic\s+in\s+arms", re.I), "International Traffic in Arms"),
    (re.compile(r"(^|[^a-z])usml([^a-z]|$)", re.I), "USML"),
    (re.compile(r"export[\s-]*controlled", re.I), "export controlled"),
]


def _norm_path(p):
    return str(p).replace("/", "\\").rstrip("\\").lower()


def _under(path, root):
    p, r = _norm_path(path), _norm_path(root)
    return p == r or p.startswith(r + "\\")


def itar_reason(division, paths=(), texts=(), industrial_root=INDUSTRIAL_ROOT, itar_root=ITAR_ROOT):
    """
    Returns why a job must NOT go to the portal, or None when it may.
    A job may go only if ALL of these hold:
      - the division is exactly "Industrial"
      - every file/folder is inside the Industrial production root (and not the MAG - ITAR root)
      - no folder name, file name or text field mentions ITAR / USML / export controlled
    """
    if division != ALLOWED_DIVISION:
        return f"division is '{division}', only '{ALLOWED_DIVISION}' jobs go to the portal"
    for p in paths:
        if not p:
            continue
        if _under(p, itar_root):
            return f"file is in the MAG - ITAR area: {p}"
        if not _under(p, industrial_root):
            return f"file is outside the Industrial production folder: {p}"
        for part in re.split(r"[\\/]+", str(p)):
            for rx, label in _ITAR_WORDS:
                if rx.search(part):
                    return f"folder or file name mentions {label}: {part}"
    for t in texts:
        for rx, label in _ITAR_WORDS:
            if t and rx.search(str(t)):
                return f"shipment details mention {label}"
    return None


class ItarBlocked(Exception):
    pass


def _check_job(job, industrial_root=INDUSTRIAL_ROOT, itar_root=ITAR_ROOT):
    meta = job.get("meta") or {}
    reason = itar_reason(job.get("division"),
                         paths=[job.get("wo_folder"), job.get("xlsx_path"), job.get("pdf_path")],
                         texts=list(meta.values()),
                         industrial_root=industrial_root, itar_root=itar_root)
    if reason:
        raise ItarBlocked(reason)


###############################################################################
# SETTINGS
###############################################################################

def load_config():
    try:
        with open(config_path(), encoding="utf-8") as f:
            c = json.load(f)
    except Exception:
        c = {}
    return {
        "enabled": bool(c.get("enabled")),
        "station_id": str(c.get("station_id") or ""),
        "station_key": str(c.get("station_key") or ""),
        "url": str(c.get("url") or DEFAULT_URL),
    }


def save_config(c):
    local_dir().mkdir(parents=True, exist_ok=True)
    with open(config_path(), "w", encoding="utf-8") as f:
        json.dump(c, f, indent=2)


def config_ready(c):
    return bool(c.get("enabled") and re.fullmatch(r"[A-Za-z0-9-]{1,32}", c.get("station_id") or "")
                and len(c.get("station_key") or "") >= 24 and str(c.get("url", "")).startswith("https://"))


###############################################################################
# QUEUE
###############################################################################

def _now():
    return time.time()


def _write_job(job, qdir=None):
    qdir = Path(qdir or queue_dir())
    qdir.mkdir(parents=True, exist_ok=True)
    tmp = qdir / f"{job['id']}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(job, f, indent=2)
    os.replace(tmp, qdir / f"{job['id']}.json")


def load_jobs(qdir=None):
    qdir = Path(qdir or queue_dir())
    jobs = []
    if not qdir.exists():
        return jobs
    for p in sorted(qdir.glob("*.json")):
        try:
            with open(p, encoding="utf-8") as f:
                jobs.append(json.load(f))
        except Exception:
            pass
    return jobs


def enqueue_shipment(division, wo_folder, wo, eng, coc_xlsx=None, coc_data=None,
                     packing_xlsx=None, packing_data=None, serials=None, serials_xlsx=None,
                     send_serials=False, config=None, qdir=None,
                     industrial_root=INDUSTRIAL_ROOT, itar_root=ITAR_ROOT):
    """
    Queues the shipment's PDFs for the portal. Never raises, never blocks.
    Returns (status, message) for the Close summary:
      ("queued", "...")   jobs written; the background uploader will send them
      ("blocked", "...")  ITAR / not allowed: nothing queued
      ("off", "...")      upload not turned on for this PC
      ("none", "...")     nothing to send
      ("error", "...")    couldn't write the queue (shipping is unaffected)
    """
    try:
        coc_data = coc_data or {}
        packing_data = packing_data or {}
        files = [p for p in (coc_xlsx, packing_xlsx, serials_xlsx if send_serials else None) if p]
        reason = itar_reason(division, paths=[wo_folder] + files,
                             texts=[wo, eng, coc_data.get("po"), packing_data.get("po"), packing_data.get("packing_no")],
                             industrial_root=industrial_root, itar_root=itar_root)
        if reason:
            return ("blocked", f"Not sent to the portal ({reason}).")
        config = config if config is not None else load_config()
        if not config_ready(config):
            return ("off", "Portal upload is off on this PC.")
        if not files:
            return ("none", "No CoC or Packing List to send.")
        po = (coc_data.get("po") or packing_data.get("po") or "").strip()
        ship_date = (coc_data.get("ship_date") or packing_data.get("date") or "").strip()
        meta = {
            "work_order": str(wo or ""),
            "engineering_number": str(eng or ""),
            "po_number": po,
            "packing_list_no": (packing_data.get("packing_no") or "").strip(),
            "ship_date": ship_date,
            "serials": ", ".join(serials or []),
        }
        queued = []
        for path, doc_type in ((coc_xlsx, DOC_COC), (packing_xlsx, DOC_PACKING),
                               (serials_xlsx if send_serials else None, DOC_SERIALS)):
            if not path:
                continue
            job = {
                "id": uuid.uuid4().hex,
                "created_at": _now(),
                "division": division,
                "wo_folder": str(wo_folder),
                "xlsx_path": str(path),
                "pdf_path": str(Path(path).with_suffix(".pdf")),
                "doc_type": doc_type,
                "meta": meta,
                "attempts": 0,
                "next_try_at": _now(),
                "status": "pending",
                "last_error": "",
            }
            _write_job(job, qdir)
            queued.append({DOC_COC: "CoC", DOC_PACKING: "Packing List", DOC_SERIALS: "Serial list"}[doc_type])
        return ("queued", f"Sending to the portal: {', '.join(queued)}.")
    except Exception as e:  # shipping must never fail because of the portal
        return ("error", f"Couldn't queue the portal upload ({type(e).__name__}: {e}). Shipping is not affected.")


###############################################################################
# PDF EXPORT (Excel)
###############################################################################

_EXPORT_PS = r"""
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
try {
  $x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false; $x.AskToUpdateLinks = $false
  $wb = $x.Workbooks.Open($env:PCC_XLSX, 0, $true)
  try { $wb.ExportAsFixedFormat(0, $env:PCC_PDF, 0, $true, $false) } finally { $wb.Close($false) }
} finally {
  $x.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
"""


def export_pdf(xlsx_path, pdf_path, timeout=EXPORT_TIMEOUT):
    """Makes pdf_path from xlsx_path with Excel (through PowerShell, so no extra Python packages)."""
    if Path(pdf_path).exists() and Path(pdf_path).stat().st_size > 0:
        return pdf_path
    if not Path(xlsx_path).exists():
        raise FileNotFoundError(f"The .xlsx is gone: {xlsx_path}")
    env = dict(os.environ, PCC_XLSX=str(xlsx_path), PCC_PDF=str(pdf_path))
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", _EXPORT_PS],
                       env=env, capture_output=True, text=True, timeout=timeout, creationflags=flags)
    if r.returncode != 0 or not Path(pdf_path).exists():
        msg = (r.stderr or r.stdout or "").strip().splitlines()
        raise RuntimeError("Excel couldn't make the PDF: " + (msg[-1] if msg else f"exit code {r.returncode}"))
    return pdf_path


###############################################################################
# UPLOAD
###############################################################################

def _multipart(fields, file_field, file_name, file_bytes):
    boundary = "----PCCShipping" + uuid.uuid4().hex
    out = []
    for k, v in fields.items():
        out.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n".encode("utf-8"))
    safe_name = file_name.replace('"', "").replace("\r", "").replace("\n", "")
    out.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{file_field}\"; filename=\"{safe_name}\"\r\n"
               f"Content-Type: application/pdf\r\n\r\n".encode("utf-8"))
    out.append(file_bytes)
    out.append(f"\r\n--{boundary}--\r\n".encode("utf-8"))
    return b"".join(out), f"multipart/form-data; boundary={boundary}"


def http_send(url, headers, body, timeout=HTTP_TIMEOUT):
    """POSTs and returns (status, json-or-None). Network failures return (0, {"error": ...})."""
    req = urllib.request.Request(url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read()
            status = resp.status
    except urllib.error.HTTPError as e:
        raw = e.read() if hasattr(e, "read") else b""
        status = e.code
    except Exception as e:
        return 0, {"error": f"{type(e).__name__}: {e}"}
    try:
        return status, json.loads(raw.decode("utf-8") or "{}")
    except Exception:
        return status, {"error": raw[:200].decode("utf-8", "replace")}


def classify(status, body):
    """'done', 'retry' or 'failed' for a portal answer."""
    body = body or {}
    if 200 <= status < 300 and body.get("ok"):
        return "done"
    if status == 0 or status >= 500 or status in (408, 429):
        return "retry"
    if body.get("retry") is True:
        return "retry"
    return "failed"


def ping(config, send=http_send):
    """Checks the station id and key with the portal. Returns (ok, message)."""
    body, ctype = _multipart({"ping": "1"}, "none", "none.pdf", b"")
    status, data = send(config["url"], {"x-station-id": config["station_id"], "x-station-key": config["station_key"],
                                        "Content-Type": ctype}, body)
    if status == 200 and (data or {}).get("pong"):
        return True, f"Connected as station '{config['station_id']}'."
    data = data or {}
    if status == 404 and not data.get("error"):
        return False, "The portal's upload function isn't set up yet (HTTP 404)."
    return False, data.get("error") or data.get("message") or f"No answer from the portal (HTTP {status})."


def upload_job(job, config, send=http_send, industrial_root=INDUSTRIAL_ROOT, itar_root=ITAR_ROOT):
    """Sends one job's PDF. Returns (outcome, message, response). Raises ItarBlocked for ITAR jobs."""
    _check_job(job, industrial_root, itar_root)  # last line of defence, right before anything leaves the PC
    pdf = Path(job["pdf_path"])
    data = pdf.read_bytes()
    meta = job.get("meta") or {}
    fields = {
        "station": config["station_id"],
        "division": job["division"],
        "doc_type": job["doc_type"],
        "work_order": meta.get("work_order", ""),
        "engineering_number": meta.get("engineering_number", ""),
        "po_number": meta.get("po_number", ""),
        "packing_list_no": meta.get("packing_list_no", ""),
        "ship_date": meta.get("ship_date", ""),
        "serials": meta.get("serials", ""),
        "source_path": str(pdf),
    }
    body, ctype = _multipart(fields, "file", pdf.name, data)
    headers = {"x-station-id": config["station_id"], "x-station-key": config["station_key"], "Content-Type": ctype}
    status, resp = send(config["url"], headers, body)
    outcome = classify(status, resp)
    msg = (resp or {}).get("error") or ("Filed in the portal." if outcome == "done" else f"HTTP {status}")
    return outcome, msg, resp


def process_job(job, config, send=http_send, export=export_pdf, now=None, qdir=None,
                industrial_root=INDUSTRIAL_ROOT, itar_root=ITAR_ROOT):
    """
    Runs one pending job: guard -> PDF -> upload -> save the result. Returns the updated job.
    Never raises (except on programming errors in tests).
    """
    now = now if now is not None else _now()
    try:
        _check_job(job, industrial_root, itar_root)
        export(job["xlsx_path"], job["pdf_path"])
        _check_job(job, industrial_root, itar_root)
        outcome, msg, resp = upload_job(job, config, send, industrial_root, itar_root)
    except ItarBlocked as e:
        outcome, msg, resp = "blocked", f"Not sent: {e}", None
    except subprocess.TimeoutExpired:
        outcome, msg, resp = "retry", "Excel took too long to make the PDF.", None
    except Exception as e:
        outcome, msg, resp = "retry", f"{type(e).__name__}: {e}", None
    job["attempts"] = job.get("attempts", 0) + 1
    job["last_try_at"] = now
    job["last_error"] = "" if outcome == "done" else msg
    if outcome == "done":
        job["status"] = "done"
        job["document_id"] = (resp or {}).get("document_id")
    elif outcome in ("failed", "blocked"):
        job["status"] = outcome
    elif now - job.get("created_at", now) > GIVE_UP_AFTER:
        job["status"] = "failed"
        job["last_error"] = f"Gave up after 7 days. Last problem: {msg}"
    else:
        delay = RETRY_DELAYS[min(job["attempts"] - 1, len(RETRY_DELAYS) - 1)]
        job["status"] = "pending"
        job["next_try_at"] = now + delay
    _write_job(job, qdir)
    return job


def describe(job):
    m = job.get("meta") or {}
    kind = {DOC_COC: "CoC", DOC_PACKING: "Packing List", DOC_SERIALS: "Serial list"}.get(job.get("doc_type"), "Document")
    return f"{kind} for WO {m.get('work_order', '?')}"


def prune(qdir=None, keep_days=30):
    """Removes finished jobs older than keep_days so the queue folder stays small."""
    cutoff = _now() - keep_days * 86400
    for job in load_jobs(qdir):
        if job.get("status") in ("done", "blocked") and job.get("last_try_at", 0) < cutoff:
            try:
                (Path(qdir or queue_dir()) / f"{job['id']}.json").unlink()
            except Exception:
                pass


class PortalUploader(threading.Thread):
    """
    Background worker. Every few seconds it runs due jobs one at a time and reports to the UI through
    on_event(kind, job) where kind is 'done', 'retry', 'failed' or 'blocked'. on_event is called from
    this thread, so the UI must hand it to tkinter with a thread-safe queue (see ShippingApp).
    """

    def __init__(self, on_event, poll_seconds=5):
        super().__init__(daemon=True, name="PortalUploader")
        self.on_event = on_event
        self.poll_seconds = poll_seconds
        self._wake = threading.Event()

    def retry_now(self):
        for job in load_jobs():
            if job.get("status") == "pending":
                job["next_try_at"] = _now()
                _write_job(job)
        self._wake.set()

    def run(self):
        try:
            prune()
        except Exception:
            pass
        while True:
            try:
                config = load_config()
                if config_ready(config):
                    for job in load_jobs():
                        if job.get("status") == "pending" and job.get("next_try_at", 0) <= _now():
                            job = process_job(job, config)
                            kind = job["status"] if job["status"] != "pending" else "retry"
                            try:
                                self.on_event(kind, job)
                            except Exception:
                                pass
            except Exception:
                pass
            self._wake.wait(self.poll_seconds)
            self._wake.clear()


def summary(jobs=None):
    """Counts for the footer: waiting, failed, sent today."""
    jobs = load_jobs() if jobs is None else jobs
    today = datetime.date.today()
    waiting = sum(1 for j in jobs if j.get("status") == "pending")
    failed = sum(1 for j in jobs if j.get("status") == "failed")
    sent_today = sum(1 for j in jobs if j.get("status") == "done"
                     and datetime.date.fromtimestamp(j.get("last_try_at", 0)) == today)
    return {"waiting": waiting, "failed": failed, "sent_today": sent_today}


if __name__ == "__main__" and len(sys.argv) == 3:
    # Manual check: py portal_upload.py <file.xlsx> <file.pdf>
    print(export_pdf(sys.argv[1], sys.argv[2]))
