"""End-to-end: write receive/backfill events through the REAL write_history_event
(envelope, filename tagging), then read them back via build_receiving_state."""
import sys, tempfile, shutil, os, stat, json
from pathlib import Path
sys.path.insert(0, r"C:\Users\Dsavaliya\Claude Cowork\PCC Shipping App\Scripts")
import app_logic as A

fail = 0
def check(n, c):
    global fail
    print(("  ok  " if c else " FAIL ") + n)
    if not c: fail += 1

hist = tempfile.mkdtemp(prefix="pcc_rt_"); proj = tempfile.mkdtemp(prefix="pcc_rtp_")
_oh = A.HISTORY_DIR
try:
    A.HISTORY_DIR = hist
    pd = Path(proj) / "Job [4041]" / "Production Data"
    (pd / "16151.1").mkdir(parents=True)
    t = lambda i: f"{i:06d}-136-9624-900-A00"
    # a close through the real writer
    p1 = A.write_history_event("close", {"wo": "16151.1", "division": "Industrial",
        "operator": "Anita Patel", "eng": "4041", "pn": "136-9624-900",
        "serials": [t(i) for i in range(1, 21)], "wo_folder": str(pd / "16151.1")})
    check("close event written via real writer", bool(p1))
    # a receive through the real writer (as ReceivingWindow._save builds it)
    p2 = A.write_history_event("receive", {"wo": "136-9624-900", "division": "Industrial",
        "operator": "Lata Gandhi", "eng": "4041", "pn": "136-9624-900",
        "serial_count": 3, "serials": [t(1), t(2), t(3)],
        "not_in_shipped": [], "dup_overrides": [{"serial": t(2), "reason": "Already received", "approved_by": "admin"}],
        "notes": "UPS 1 box", "production_data": str(pd)})
    check("receive event written (P/N in filename ok)", bool(p2) and "136-9624-900" in Path(p2).name)
    # a backfill through the real writer
    p3 = A.write_history_event("backfill_shipped", {"wo": "16149.1", "division": "Industrial",
        "eng": "4041", "pn": "136-9624-900", "serials": [t(99)], "ship_date": "5/2/25",
        "production_data": str(pd)})
    check("backfill event written", bool(p3))
    st = A.build_receiving_state("4041", "136-9624-900")
    m = A.receiving_summary(st)
    check("roundtrip: shipped 21 (20 close + 1 backfill)", m["shipped"] == 21)
    check("roundtrip: received 3", m["received"] == 3)
    check("roundtrip: outstanding 18", m["outstanding"] == 18)
    check("receive ts filled by envelope", bool(st["received"][A._serial_identity(t(1), "136-9624-900")]["ts"]))
    check("pd_path recovered", st["pd_path"] == str(pd / "16151.1").rsplit("\\", 1)[0] or st["pd_path"] == str(pd))
    out = A.write_received_xlsx(st)
    check("xlsx written from roundtrip state", Path(out).exists())
    # jobs list sees it once
    jobs = A.list_testing_jobs()
    check("job listed exactly once", sum(1 for j in jobs if j["pn"] == "136-9624-900") == 1)
finally:
    A.HISTORY_DIR = _oh
    for r, _, fs in os.walk(proj):
        for f in fs:
            try: os.chmod(os.path.join(r, f), stat.S_IWRITE)
            except Exception: pass
    shutil.rmtree(hist, ignore_errors=True); shutil.rmtree(proj, ignore_errors=True)
print("ALL PASSED" if not fail else f"{fail} FAILED")
sys.exit(1 if fail else 0)
