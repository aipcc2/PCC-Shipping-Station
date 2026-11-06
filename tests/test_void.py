"""Void-record + batch-override tests: a supervisor-voided shipment stops counting
in the duplicate check AND the receiving tracker, while staying in history."""
import sys, tempfile, shutil, json, inspect
from pathlib import Path
sys.path.insert(0, r"C:\Users\Dsavaliya\Claude Cowork\PCC Shipping App\Scripts")
import app_logic as A

fail = 0
def check(n, c):
    global fail
    print(("  ok  " if c else " FAIL ") + n)
    if not c: fail += 1

hist = tempfile.mkdtemp(prefix="pcc_void_")
_oh = A.HISTORY_DIR
try:
    A.HISTORY_DIR = hist
    E, P = "4041", "136-9624-900"
    t = lambda i: f"{i:06d}-136-9624-900-A00"

    # Two real closes via the real writer (so event_ids are authentic)
    p1 = A.write_history_event("close", {"wo": "16151.1", "division": "Industrial",
        "operator": "Anita Patel", "eng": E, "pn": P,
        "serials": [t(1), t(2), t(3)], "wo_folder": ""})
    p2 = A.write_history_event("close", {"wo": "16152.1", "division": "Industrial",
        "operator": "Anita Patel", "eng": E, "pn": P,
        "serials": [t(4), t(5)], "wo_folder": ""})
    ev1, ev2 = Path(p1).stem, Path(p2).stem
    idx = A.build_serial_history(E, P)
    check("before void: all 5 serials shipped", len(idx) == 5)
    st = A.build_receiving_state(E, P)
    check("before void: tracker shipped=5", A.receiving_summary(st)["shipped"] == 5)

    # Supervisor voids the WRONG shipment (16151.1) via the real void writer
    target = {"event_id": ev1, "event_type": "close", "wo": "16151.1",
              "division": "Industrial", "eng": E, "pn": P, "ts": "x",
              "serials": [t(1), t(2), t(3)]}
    vp = A.write_void_event(target, "wrong boards shipped", "admin")
    check("void event written", bool(vp))
    vrec = json.load(open(vp, encoding="utf-8"))
    check("void carries target id + reason + approver",
          vrec["target_event_id"] == ev1 and vrec["reason"] == "wrong boards shipped"
          and vrec["approved_by"] == "admin" and vrec["event_type"] == "void")

    idx2 = A.build_serial_history(E, P)
    check("dup-check drops the voided shipment's serials (5 -> 2)",
          len(idx2) == 2 and A._serial_identity(t(1), "136-9624-900") not in idx2 and A._serial_identity(t(4), "136-9624-900") in idx2)
    st2 = A.build_receiving_state(E, P)
    check("tracker drops them too (shipped=2)", A.receiving_summary(st2)["shipped"] == 2)

    # voided RECEIVE event stops counting as received
    pr = A.write_history_event("receive", {"wo": P, "division": "Industrial",
        "operator": "Lata Gandhi", "eng": E, "pn": P, "serials": [t(4)],
        "failed": [], "production_data": ""})
    check("received before void: 1", A.receiving_summary(A.build_receiving_state(E, P))["received"] == 1)
    A.write_void_event({"event_id": Path(pr).stem, "event_type": "receive", "wo": P,
                        "eng": E, "pn": P, "ts": "x", "serials": [t(4)]}, "mis-scan", "admin")
    check("received after void: 0", A.receiving_summary(A.build_receiving_state(E, P))["received"] == 0)

    # a job whose ONLY close is voided disappears from the receiving job list
    hist2 = tempfile.mkdtemp(prefix="pcc_void2_")
    A.HISTORY_DIR = hist2
    p3 = A.write_history_event("close", {"wo": "17777.1", "eng": "9999", "pn": "555-000-900",
                                         "serials": ["x1"], "wo_folder": ""})
    check("job listed pre-void", any(j["eng"] == "9999" for j in A.list_testing_jobs()))
    A.write_void_event({"event_id": Path(p3).stem, "event_type": "close", "wo": "17777.1",
                        "eng": "9999", "pn": "555-000-900", "ts": "x", "serials": ["x1"]}, "", "admin")
    check("job gone post-void", not any(j["eng"] == "9999" for j in A.list_testing_jobs()))
    shutil.rmtree(hist2, ignore_errors=True)

    check("void without event_id refused", A.write_void_event({"wo": "x"}, "r", "a") is None)

    # ── wiring checks (Tk classes, verified by source inspection) ──
    sp = inspect.getsource(A.SerialPopup)
    check("batch-override button + toggle wired",
          "_toggle_dup_session" in sp and "allow re-ship" in sp and "_dup_session_admin" in sp)
    check("batch mode auto-approves + logs in _resolve_dup",
          "if self._dup_session_admin:" in sp and sp.index("if self._dup_session_admin:") <
          sp.index("HistoricalDuplicateDialog(self, v, occ)"))
    hd = inspect.getsource(A.HistoryDetail)
    check("HistoryDetail void button + banner + admin gate",
          "_void" in hd and "VOIDABLE" in hd and "AdminLoginDialog" in hd and "VOIDED by" in hd)
    hv = inspect.getsource(A.HistoryViewer)
    check("HistoryViewer maps + grays voided rows",
          "_voided" in hv and "(VOIDED)" in hv and '"void"' in hv.replace("'", '"'))
finally:
    A.HISTORY_DIR = _oh
    shutil.rmtree(hist, ignore_errors=True)
print("ALL PASSED" if not fail else f"{fail} FAILED")
sys.exit(1 if fail else 0)
