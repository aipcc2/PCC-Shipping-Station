"""Receiving feature test suite — heavy, many-instance coverage.
Covers: state building from history, coated matching, xlsx generation + read-only
+ lock handling, backfill parsing (real app-generated files + odd layouts + csv),
job listing/detection, WO-search exclusion, idempotency, multi-session merges,
and a 1500-board stress run. Pure logic — no Tk windows."""
import sys, os, json, tempfile, shutil, time, stat
from pathlib import Path

SCRIPTS = r"C:\Users\Dsavaliya\Claude Cowork\PCC Shipping App\Scripts"
sys.path.insert(0, SCRIPTS)
import app_logic as A

# Redirected stdout (a log file, a pipe) uses the cp1252 locale codec on this PC;
# a check NAME with a symbol outside it used to abort the whole run mid-suite
# (review round 1). Keep printing, never crash on a glyph.
try: sys.stdout.reconfigure(errors="replace")
except Exception: pass

_fail = 0
def check(name, cond):
    global _fail
    print(("  ok  " if cond else " FAIL ") + name)
    if not cond: _fail += 1

def ev(root, ym, name, rec):
    folder = Path(root) / ym[:4] / ym
    folder.mkdir(parents=True, exist_ok=True)
    (folder / f"{name}.json").write_text(json.dumps(rec), encoding="utf-8")

def close_ev(eng, pn, wo, serials, ts, wo_folder=None):
    return {"schema_version": 1, "event_type": "close", "eng": eng, "pn": pn, "wo": wo,
            "operator": "Anita Patel", "ts": ts, "serials": list(serials),
            "wo_folder": wo_folder or ""}

def recv_ev(eng, pn, serials, ts, op="Lata Gandhi", pd=""):
    return {"schema_version": 1, "event_type": "receive", "eng": eng, "pn": pn,
            "operator": op, "ts": ts, "serials": list(serials), "production_data": pd}

TPN = "136-9624-900"; CPN = "136-9624-970"; ENG = "4041"
def tser(i, rev="A00"): return f"{i:06d}-136-9624-900-{rev}"
def cser(i, rev="B00"): return f"{i:06d}-136-9624-970-{rev}"

from openpyxl import Workbook as _WB6
def _shipfile(folder, name, serials, date=None, label_beside=False,
              pn_label="161-9658-900", no_date_col=False):
    # realistic app paperwork (COC / Packing List / Shipping Serials): every file
    # carries a bare P/N value cell and a cert sentence embedding the P/N — these
    # must NEVER become phantom boards (review round 3). Module-level since
    # 2026-09-23 so the coated-only section can build -970 COC files too.
    folder.mkdir(parents=True, exist_ok=True)
    w = _WB6(); ws6 = w.active
    if label_beside and date:
        ws6.append(["Date:", date])
        for s in serials: ws6.append([s])
    elif no_date_col:
        ws6.append(["Serial Number"])
        for s in serials: ws6.append([s])
    else:
        ws6.append(["Serial Number", "Date"])
        for j, s in enumerate(serials):
            ws6.append([s, date if (j == 0 and date) else None])
    ws6.append(["P/N:", pn_label])
    ws6.append([f"This certifies that the parts listed conform to the "
                f"specifications of the {pn_label} Rev.- A00 drawings."])
    w.save(str(folder / name))

# ═══ 1. tiny helpers ═════════════════════════════════════════════════════════
print("helpers:")
check("testing_to_coated_pn 900->970", A.testing_to_coated_pn("136-9624-900") == "136-9624-970")
check("testing_to_coated_pn non-900 -> ''", A.testing_to_coated_pn("161-1061-971") == "")
check("testing_to_coated_pn blank -> ''", A.testing_to_coated_pn("") == "")
check("serial_seq strips leading zeros", A.serial_seq("000123-136-9624-900-A00") == "123")
check("serial_seq non-numeric first seg -> ''", A.serial_seq("SN12-900") == "")
check("serial_seq plain number", A.serial_seq("42") == "42")

# ═══ 2. state building from a big synthetic history ══════════════════════════
print("build_receiving_state:")
hist = tempfile.mkdtemp(prefix="pcc_rcvhist_")
proj = tempfile.mkdtemp(prefix="pcc_rcvproj_")
_orig_hist = A.HISTORY_DIR
_orig_roots = dict(A.PRODUCTION_ROOTS)
try:
    A.HISTORY_DIR = hist
    A.PRODUCTION_ROOTS["Industrial"] = proj
    # Real-ish folder tree for the job
    job = Path(proj) / "20 TestCust" / "GizmoX [4041 Rev. A]"
    pd_path = job / "Production Data"
    for wo in ("16151.1", "16152.1", "16153.1"):
        (pd_path / wo / A.SERIALS_FOLDER).mkdir(parents=True, exist_ok=True)
    # the coated (-970) WO folder exists too: since 2026-10-06 the Shipping Report homes THERE
    (pd_path / "16154.1").mkdir(parents=True, exist_ok=True)

    # Testing closes: WO 16151.1 seq 1..500 (June), 16152.1 seq 501..1000 (July),
    # 16153.1 seq 1001..1500 (August, NEWEST — its wo_folder should win pd_path)
    ev(hist, "2026-06", "a_close_16151", close_ev(ENG, TPN, "16151.1",
       [tser(i) for i in range(1, 501)], "2026-06-10 09:00:00", str(pd_path / "16151.1")))
    ev(hist, "2026-07", "a_close_16152", close_ev(ENG, TPN, "16152.1",
       [tser(i) for i in range(501, 1001)], "2026-07-12 09:00:00", str(pd_path / "16152.1")))
    ev(hist, "2026-08", "a_close_16153", close_ev(ENG, TPN, "16153.1",
       [tser(i) for i in range(1001, 1501)], "2026-08-14 09:00:00", str(pd_path / "16153.1")))
    # Coated closes: seqs 1..600 shipped coated (REV CHANGED to B00 — must still match by seq)
    ev(hist, "2026-08", "a_close_coat1", close_ev(ENG, CPN, "16154.1",
       [cser(i) for i in range(1, 601)], "2026-08-18 10:00:00", str(pd_path / "16154.1")))
    # Receives: 3 sessions over 2 months, total seqs 1..900 (session 2 re-lists 400-450 — dup across events)
    ev(hist, "2026-07", "b_recv_1", recv_ev(ENG, TPN, [tser(i) for i in range(1, 401)],
       "2026-07-20 14:00:00", "Lata Gandhi", str(pd_path)))
    ev(hist, "2026-08", "b_recv_2", recv_ev(ENG, TPN, [tser(i) for i in range(400, 651)],
       "2026-08-05 14:00:00", "Bina Patel", str(pd_path)))
    ev(hist, "2026-08", "b_recv_3", recv_ev(ENG, TPN, [tser(i) for i in range(651, 901)],
       "2026-08-19 14:00:00", "Anita Patel", str(pd_path)))
    # Noise: other job, drafts, deletes, coated of another eng
    ev(hist, "2026-07", "z_noise_close", close_ev("9999", "555-000-900", "17777.1",
       ["000001-555-000-900-A00"], "2026-07-01 08:00:00"))
    ev(hist, "2026-07", "z_noise_draft", {"event_type": "draft_save", "eng": ENG, "pn": TPN,
       "wo": "16152.1", "ts": "2026-07-11 08:00:00", "serials": [tser(9999)]})
    ev(hist, "2026-07", "z_noise_del", {"event_type": "delete", "eng": ENG, "pn": TPN,
       "wo": "16152.1", "ts": "2026-07-11 09:00:00"})

    st = A.build_receiving_state(ENG, TPN)
    m = A.receiving_summary(st)
    check("shipped = 1500", m["shipped"] == 1500)
    check("received = 900 (dups across sessions collapse)", m["received"] == 900)
    check("outstanding = 600", m["outstanding"] == 600)
    check("coated&shipped among received = 600", m["coated"] == 600)
    check("in-house awaiting coating = 300", m["in_house"] == 300)
    check("draft serial (9999) NOT in shipped universe", A._serial_identity(tser(9999), TPN) not in st["shipped"])
    check("noise job's serial not present", "000001-555-000-900-a00" not in st["shipped"])
    check("WO attribution: seq 42 -> 16151.1", st["shipped"][A._serial_identity(tser(42), TPN)]["wo"] == "16151.1")
    check("WO attribution: seq 1200 -> 16153.1", st["shipped"][A._serial_identity(tser(1200), TPN)]["wo"] == "16153.1")
    check("pd_path from NEWEST close", st["pd_path"] == str(pd_path))
    check("re-received serial keeps NEWEST occurrence (seq 400 -> Bina, Aug)",
          st["received"][A._serial_identity(tser(400), TPN)]["operator"] == "Bina Patel")
    check("coated matched by seq despite rev change (seq 42 B00 vs A00)",
          A.serial_seq(st["coated"].get("42", {}).get("serial", "")) == "42")
    # blank scope -> empty
    st0 = A.build_receiving_state("", TPN)
    check("blank eng -> empty state", not st0["shipped"] and not st0["received"])

    # ═══ 3. stress: state build + summary timing ═════════════════════════════
    print("stress:")
    t0 = time.time()
    for _ in range(3): st = A.build_receiving_state(ENG, TPN)
    dt = (time.time() - t0) / 3
    check(f"state build over 1500/900/600 avg {dt:.2f}s < 3s", dt < 3.0)

    # ═══ 4. xlsx generation ══════════════════════════════════════════════════
    print("write_received_xlsx:")
    # Dhruvil 2026-10-06: boards are no longer sent for testing - the report homes in the -970 COATED WO folder
    # for every job with a coated side (was: the newest -900 testing WO folder, 16153.1)
    check("wo_path = the COATED WO folder (16154.1) - the 2026-10-06 rule, home_side 'coated'",
          st["wo_path"] == str(pd_path / "16154.1") and st.get("home_side") == "coated")
    # plant stale copies at the legacy pd-level home AND an older WO folder —
    # real OLD-FORMAT (read-only era) workbooks: the app always saved those with
    # the READ-ONLY attribute, which is exactly what marks them as edit-free.
    fname = A.received_xlsx_path(pd_path, TPN).name
    from openpyxl import load_workbook, Workbook as _WB
    for stale_dir in (pd_path / A.RECEIVED_FOLDER, pd_path / "16151.1" / A.RECEIVED_FOLDER):
        stale_dir.mkdir(parents=True, exist_ok=True)
        _swb = _WB(); _swb.active.title = "Received"
        _swb.active.append(["old read-only-era file"]); _swb.save(str(stale_dir / fname))
        os.chmod(str(stale_dir / fname), stat.S_IREAD)
    out = A.write_received_xlsx(st)
    p = Path(out)
    check("file created INSIDE the coated WO folder (16154.1\\Shipping Reports, the 2026-10-06 home)",
          p.exists() and p.parent.name == A.RECEIVED_FOLDER
          and p.parent.parent == pd_path / "16154.1")
    check("nothing written into the old -900 home (16153.1)", not (pd_path / "16153.1" / A.RECEIVED_FOLDER / fname).exists())
    check("stale legacy pd-level copy removed",
          not (pd_path / A.RECEIVED_FOLDER / fname).exists())
    check("stale copy in an older WO folder removed",
          not (pd_path / "16151.1" / A.RECEIVED_FOLDER / fname).exists())
    check("file is WRITABLE on disk (operators edit it now)", os.access(out, os.W_OK))
    check("no overlap conflicts on a fresh build", not st.get("xlsx_conflicts"))
    wb = load_workbook(out); names = wb.sheetnames
    check("FOUR sheets in order: Received from testing / Coated & shipped / In-House / Still at customer",
          names == ["Received from testing", "Coated & shipped", "In-House", "Still at customer"])
    ws, wsc, ws2 = wb["In-House"], wb["Coated & shipped"], wb["Still at customer"]
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = next(i for i, r in enumerate(rows) if r and r[0] == "#")
    data = [r for r in rows[hdr_i + 1:] if r and r[1]]
    check("In-House sheet = in-house only (900 received - 600 coated = 300)", len(data) == 300)
    check("first in-house row is seq 601 (1-600 are coated & shipped)", data[0][1] == tser(601))
    check("coated boards NOT on the In-House sheet",
          not any(r[1] == tser(400) for r in data))
    check("In-House subtitle carries its own count", "300 board(s)" in (rows[1][0] or ""))
    r700 = next(r for r in data if r[1] == tser(700))
    check("row 700: result + operator + WO",
          r700[2] == "PASS" and r700[4] == "Anita Patel" and r700[5] == "16152.1")
    # Received-from-testing LOG: every receive occurrence, re-receives included
    wsr = wb["Received from testing"]
    rowsr = list(wsr.iter_rows(values_only=True))
    hdrr_i = next(i for i, r in enumerate(rowsr) if r and r[0] == "#")
    datar = [r for r in rowsr[hdrr_i + 1:] if r and r[1]]
    check("Received log = 901 rows (900 boards + seq 400 received TWICE)", len(datar) == 901)
    r400s = [r for r in datar if r[1] == tser(400)]
    check("seq 400 has TWO receive rows (July + Aug)", len(r400s) == 2
          and {str(r[3])[:10] for r in r400s} == {"2026-07-20", "2026-08-05"})
    check("received rows carry ship info", r400s[0][5] == "16151.1")
    # Coated & shipped LOG: rows are the COATED shipments themselves (cser form)
    rowsc = list(wsc.iter_rows(values_only=True))
    hdrc_i = next(i for i, r in enumerate(rowsc) if r and r[0] == "#")
    datac = [r for r in rowsc[hdrc_i + 1:] if r and r[1]]
    check("Coated & shipped sheet lists ALL 600 coated shipments", len(datac) == 600)
    r400c = next(r for r in datac if r[1] == cser(400))
    check("coated row: WO + ship date + receive-date link",
          r400c[2] == "16154.1" and str(r400c[3])[:10] == "2026-08-18"
          and str(r400c[4])[:10] == "2026-08-05")
    rows2 = list(ws2.iter_rows(values_only=True))
    hdr2_i = next(i for i, r in enumerate(rows2) if r and r[0] == "#")
    data2 = [r for r in rows2[hdr2_i + 1:] if r and r[1]]
    check("Still-at-customer sheet has 600 rows", len(data2) == 600)
    check("first outstanding is seq 901", data2[0][1] == tser(901))
    check("outstanding row carries WO", data2[0][2] == "16152.1")
    summary_line = rows[2][0]
    check("summary line correct", "Shipped for testing: 1500" in summary_line
          and "Still at customer: 600" in summary_line
          and "Received back: 900" in summary_line
          and "coated & shipped 600" in summary_line
          and "Coated boards on record: 600" in summary_line)
    wb.close()

    # re-save over the existing file = MERGE, not regenerate
    out2 = A.write_received_xlsx(st)
    check("re-save over existing file works (merge)", out2 == out and Path(out).exists())
    check("still writable after merge", os.access(out, os.W_OK))
    check("idempotent merge: no conflicts, no duplicate rows", not st.get("xlsx_conflicts"))
    wb = load_workbook(out)
    n2 = sum(1 for r in wb["Received from testing"].iter_rows(values_only=True)
             if r and isinstance(r[0], int))
    check("row count unchanged after idempotent merge (901)", n2 == 901)
    wb.close()
    # simulate a NEW receive folded in -> file reflects it
    st["received"][A._serial_identity(tser(901), TPN)] = {"serial": tser(901), "ts": "2026-08-21 10:00:00",
                                                 "sts": "2026-08-21 10:00:00",
                                                 "operator": "Palak Patel", "result": "pass"}
    A.write_received_xlsx(st)
    wb = load_workbook(out)
    d2 = [r for r in wb["Still at customer"].iter_rows(values_only=True) if r and r[1] == tser(901)]
    check("after new receive folded: seq 901 gone from Still-at-customer", not d2)
    wb.close()
    del st["received"][A._serial_identity(tser(901), TPN)]

    # LOCKED file -> PermissionError raised (caller shows retry dialog)
    fh = open(out, "rb")   # simulate a viewer holding the file
    try:
        try:
            A.write_received_xlsx(st); locked_ok = False
        except PermissionError: locked_ok = True
        except Exception: locked_ok = False
    finally:
        fh.close()
    check("locked file raises PermissionError for retry UX", locked_ok)
    # and it recovers after the lock clears
    check("write recovers after lock released", A.write_received_xlsx(st) == out)
    # temp files cleaned up
    stray = [f for f in p.parent.iterdir() if f.name.endswith(".tmp")]
    check("no stray .tmp files left behind", not stray)
    # no pd_path -> clear error
    try:
        A.write_received_xlsx({"eng": ENG, "pn": TPN, "shipped": {}, "received": {}, "coated": {},
                               "pd_path": None}); npe = False
    except RuntimeError: npe = True
    except Exception: npe = False
    check("missing wo_path/pd_path raises RuntimeError (not silent)", npe)

    # ═══ 5. backfill parser ══════════════════════════════════════════════════
    print("backfill parser:")
    wo_f = pd_path / "16151.1"
    # (a) REAL app-generated serials xlsx
    A.write_serial_xlsx(wo_f / A.SERIALS_FOLDER, "4041", "16151.1",
                        [tser(i) for i in range(1, 51)], "Anita Patel")
    s, d, n = A.parse_serials_from_shipping_files(wo_f, TPN)
    check("app-generated xlsx: 50 serials", len(s) == 50 and s[0] == tser(1))
    check("ship date picked up from Date column", bool(d))
    check("one file read", n == 1)
    # (b) second file with overlap -> dedupe
    A.write_serial_xlsx(wo_f / A.SERIALS_FOLDER, "4041", "16151.1",
                        [tser(i) for i in range(40, 61)], "Anita Patel")
    s, d, n = A.parse_serials_from_shipping_files(wo_f, TPN)
    check("two files, overlap deduped: 60 unique", len(s) == 60 and n == 2)
    # (c) odd manual layout (no header) -> P/N containment fallback
    from openpyxl import Workbook
    wb2 = Workbook(); w = wb2.active
    w.append(["shipment record — do not edit"])
    w.append([tser(200), "packed by J"])
    w.append(["junk", tser(201)])
    w.append(["16452.1-77"])          # other customer format, no embedded P/N -> ignored
    odd_dir = pd_path / "16152.1" / A.SERIALS_FOLDER
    wb2.save(str(odd_dir / "manual_list.xlsx"))
    s2, _, _ = A.parse_serials_from_shipping_files(pd_path / "16152.1", TPN)
    check("odd layout via P/N containment: exactly the 2 embedded serials",
          sorted(s2) == sorted([tser(200), tser(201)]))
    # (d) csv fallback file
    csv_dir = pd_path / "16153.1" / A.SERIALS_FOLDER
    (csv_dir / "old.csv").write_text(
        "Date,Time,Serial_Number,Operator,WO\n"
        f"6/1/25,9:00 AM,{tser(300)},Anita,16153.1\n"
        f"6/1/25,9:01 AM,{tser(301)},Anita,16153.1\n", encoding="utf-8")
    s3, d3, n3 = A.parse_serials_from_shipping_files(pd_path / "16153.1", TPN)
    check("csv file parsed (2 serials, date)", len(s3) == 2 and d3 == "6/1/25" and n3 == 1)
    check("missing Shipping Serials folder -> empty, no crash",
          A.parse_serials_from_shipping_files(pd_path / "nope", TPN) == ([], "", 0))

    # backfill events feed the shipped universe + are idempotent by dedupe
    ev(hist, "2026-06", "c_backfill", {"event_type": "backfill_shipped", "eng": ENG, "pn": TPN,
       "wo": "16149.1", "ts": "2026-06-01 08:00:00", "ship_date": "5/2/25",
       "serials": [tser(i) for i in range(2001, 2101)] + [tser(1)],   # tser(1) already known via close
       "production_data": str(pd_path)})
    st2 = A.build_receiving_state(ENG, TPN)
    m2 = A.receiving_summary(st2)
    check("backfill adds 100 new shipped (dup vs close ignored)", m2["shipped"] == 1600)
    check("backfilled serial has WO + ship_date", st2["shipped"][A._serial_identity(tser(2050), TPN)]["wo"] == "16149.1"
          and st2["shipped"][A._serial_identity(tser(2050), TPN)]["ts"] == "5/2/25")
    check("close beats backfill for a serial in both (newest-first walk)",
          st2["shipped"][A._serial_identity(tser(1), TPN)]["source"] == "close")

    # ═══ 5a2. pass/fail + cycles + received_date (test-and-return flow) ═══════
    print("pass/fail + cycles:")
    hist2 = tempfile.mkdtemp(prefix="pcc_pf_")
    _hh = A.HISTORY_DIR
    try:
        A.HISTORY_DIR = hist2
        E2, P2 = "5050", "222-333-900"
        s2 = lambda i, rev="A00": f"{i:06d}-222-333-900-{rev}"
        # ship 1..10 (June), receive 1..6 with 4,5,6 FAILED (July),
        # re-ship 4,5 for re-test (August, NEWER), coated close for seq 1 (Aug)
        ev(hist2, "2026-06", "a_ship1", close_ev(E2, P2, "17001.1",
           [s2(i) for i in range(1, 11)], "2026-06-01 09:00:00"))
        ev(hist2, "2026-07", "b_recv1", dict(recv_ev(E2, P2, [s2(i) for i in range(1, 7)],
           "2026-07-01 10:00:00", "Anita Patel"), failed=[s2(4), s2(5), s2(6)]))
        ev(hist2, "2026-08", "c_ship2", close_ev(E2, P2, "17002.1",
           [s2(4), s2(5)], "2026-08-01 09:00:00"))
        ev(hist2, "2026-08", "d_coat", close_ev(E2, "222-333-970", "17003.1",
           ["000001-222-333-970-B00"], "2026-08-02 09:00:00"))
        stp = A.build_receiving_state(E2, P2)
        mp = A.receiving_summary(stp)
        check("received counts only CURRENT boards (6 received - 2 re-shipped = 4)", mp["received"] == 4)
        check("re-shipped failed boards back in 'still at customer' (10-4=6)", mp["outstanding"] == 6)
        check("failed awaiting re-ship = 1 (seq 6; 4+5 re-shipped)", mp["failed"] == 1)
        check("coated & shipped = 1 (seq 1)", mp["coated"] == 1)
        check("in-house awaiting coating = 2 (seq 2,3)", mp["in_house"] == 2)
        check("_currently_received: re-shipped board is NOT received",
              not A._currently_received(stp, A._serial_identity(s2(4), P2)))
        check("_currently_received: failed-in-house board IS received",
              A._currently_received(stp, A._serial_identity(s2(6), P2)))
        # dup-check: failed board awaiting re-ship -> NO alert; passed -> alert
        idx2 = A.build_serial_history(E2, P2)
        check("dup-check SKIPS failed board awaiting re-ship (seq 6)",
              A._serial_identity(s2(6), P2) not in idx2)
        check("dup-check still alerts a PASSED received board (seq 2)",
              A._serial_identity(s2(2), P2) in idx2)
        check("dup-check alerts a re-shipped board (seq 4, newest = ship)",
              A._serial_identity(s2(4), P2) in idx2)
        check("dup-check alerts a still-at-customer board (seq 8)",
              A._serial_identity(s2(8), P2) in idx2)
        # receive AGAIN after re-ship: not 'currently received' -> allowed by checks
        ev(hist2, "2026-08", "e_recv2", dict(recv_ev(E2, P2, [s2(4)],
           "2026-08-10 10:00:00", "Bina Patel"), failed=[]))
        stp2 = A.build_receiving_state(E2, P2)
        check("second receive after re-test: board received again, now PASS",
              A._currently_received(stp2, A._serial_identity(s2(4), P2))
              and stp2["received"][A._serial_identity(s2(4), P2)]["result"] == "pass")
        # xlsx: FAIL row appears with Result=FAIL; re-shipped boards on sheet 2
        # (synthetic closes carry no wo_folder, so point the state at a real dir)
        stp2["pd_path"] = str(pd_path); stp2["wo_path"] = str(pd_path / "16153.1")
        outp = A.write_received_xlsx(stp2)
        wbp = load_workbook(outp)
        rows_p = [r for r in wbp["In-House"].iter_rows(values_only=True) if r and r[1] == s2(6)]
        check("xlsx Result column shows FAIL", rows_p and rows_p[0][2] == "FAIL")
        out2 = [r[1] for r in wbp["Still at customer"].iter_rows(values_only=True) if r and r[1]]
        check("re-shipped board (seq 5) listed as still at customer", s2(5) in out2)
        wbp.close()
        # received_date (historical import): drives display AND cycle ordering
        ev(hist2, "2026-08", "f_recv_old", dict(recv_ev(E2, P2, [s2(9)],
           "2026-08-20 10:00:00", "Lata Gandhi"), received_date="2026-05-14"))
        stp3 = A.build_receiving_state(E2, P2)
        r9 = stp3["received"][A._serial_identity(s2(9), P2)]
        check("received_date shows as the receive date", r9["ts"] == "2026-05-14")
        check("received_date sorts BEFORE the June ship -> board still at customer",
              not A._currently_received(stp3, A._serial_identity(s2(9), P2)))
        check("_parse_user_date formats", A._parse_user_date("2026-05-14") == "2026-05-14"
              and A._parse_user_date("5/14/26") == "2026-05-14"
              and A._parse_user_date("05-14-2026") == "2026-05-14"
              and A._parse_user_date("garbage") is None)
        # ── mid-job REVISION CHANGE (Dhruvil 2026-09-02, jobs 4092/4041):
        # board ships wearing -A01, comes back scanned as -B01 — SAME board.
        # (seq 77 is FRESH — a seq reused from the earlier ships would now count
        # as a second shipment under the count-based cycles, which is correct
        # behavior but not what this rev-matching test is about)
        ev(hist2, "2026-06", "g_ship_revA", close_ev(E2, P2, "18001.1",
           ["000077-222-333-900-A01"], "2026-06-05 09:00:00"))
        ev(hist2, "2026-07", "h_recv_revB", dict(recv_ev(E2, P2, ["000077-222-333-900-B02"],
           "2026-07-05 10:00:00", "Anita Patel"), failed=[]))
        strv = A.build_receiving_state(E2, P2)
        kid = A._serial_identity("000077-222-333-900-A01", P2)
        check("rev change: shipped A01 + received B02 -> ONE board, matched",
              kid in strv["shipped"] and kid in strv["received"]
              and A._currently_received(strv, kid))
        check("rev change: board NOT stuck in 'still at customer'",
              kid not in {k for k in strv["shipped"] if k not in strv["received"]})
        idxrv = A.build_serial_history(E2, P2)
        check("dup-check: scanning the B02 form finds the A01 shipment",
              A._serial_identity("000077-222-333-900-B02", P2) in idxrv)
    finally:
        A.HISTORY_DIR = _hh
        shutil.rmtree(hist2, ignore_errors=True)

    # ═══ 5a3. job 4065: ALL data imported manually (nothing in live history) ═══
    print("fully-backfilled job (4065 scenario):")
    hist3 = tempfile.mkdtemp(prefix="pcc_4065_")
    _h3 = A.HISTORY_DIR
    try:
        A.HISTORY_DIR = hist3
        E3, TP3, CP3 = "4065", "134-2201-900", "134-2201-970"
        t3 = lambda i, pn="134-2201-900", rev="A00": f"{i:06d}-{pn}-{rev}"
        # testing shipments imported via backfill (30 boards)
        ev(hist3, "2026-09", "a_bf_test", {"event_type": "backfill_shipped", "eng": E3, "pn": TP3,
           "wo": "16200.1", "ts": "2026-09-02 09:00:00", "ship_date": "3/10/26",
           "serials": [t3(i) for i in range(1, 31)], "production_data": ""})
        # received back: 12 boards, pasted with an old received_date
        ev(hist3, "2026-09", "b_recv", dict(recv_ev(E3, TP3, [t3(i) for i in range(1, 13)],
           "2026-09-02 10:00:00", "Anita Patel"), received_date="2026-05-20"))
        # COATED shipments imported via backfill under the -970 P/N (10 boards) —
        # this is the exact case that was broken: backfilled coated never counted.
        ev(hist3, "2026-09", "c_bf_coat", {"event_type": "backfill_shipped", "eng": E3, "pn": CP3,
           "wo": "16210.1", "ts": "2026-09-02 11:00:00", "ship_date": "6/1/26",
           "serials": [t3(i, "134-2201-970") for i in range(1, 11)], "production_data": ""})
        st65 = A.build_receiving_state(E3, TP3)
        m65 = A.receiving_summary(st65)
        check("4065: shipped baseline 30 (testing backfill only — coated NOT mixed in)",
              m65["shipped"] == 30)
        check("4065: received 12", m65["received"] == 12)
        check("4065: backfilled COATED now counts — coated & shipped = 10", m65["coated"] == 10)
        check("4065: in-house awaiting coating = 2 (12 received - 10 coated)", m65["in_house"] == 2)
        check("4065: still at customer = 18", m65["outstanding"] == 18)
        check("coated map carries the -970 WO", st65["coated"]["5"]["wo"] == "16210.1")
        # the -970 'job' created by the coated import must NOT be a receivable job
        jobs65 = A.list_testing_jobs()
        check("coated (-970) backfill job hidden from the Receiving dropdown",
              not any(j["pn"] == CP3 for j in jobs65)
              and any(j["pn"] == TP3 for j in jobs65))
    finally:
        A.HISTORY_DIR = _h3
        shutil.rmtree(hist3, ignore_errors=True)

    # ═══ 5a4. 4065 AS IT REALLY HAPPENED (from the live share, 2026-09-02):
    #        coated serials imported under the TESTING P/N label, plus an old
    #        close whose pn was recorded as the WHOLE folder name. Routing must
    #        follow the P/N embedded in each serial, not the event label. ═══════
    print("4065 real-data shapes (embedded-P/N routing):")
    hist4 = tempfile.mkdtemp(prefix="pcc_4065b_")
    _h4 = A.HISTORY_DIR
    try:
        A.HISTORY_DIR = hist4
        E4, TP4, CP4 = "4065", "134-2201-900", "134-2201-970"
        ts4 = lambda i, rev="A00": f"{i:06d}-1342201900-{rev}"   # testing form (no dashes in pn part)
        cs4 = lambda i, rev="A00": f"{i:06d}-1342201970-{rev}"   # coated form
        # real testing close (clean label)
        ev(hist4, "2026-08", "a_close_test", close_ev(E4, TP4, "16587.1",
           [ts4(i) for i in range(71, 121)], "2026-08-20 13:46:00"))
        # real coated close with the OLD PARSER BUG pn (= whole folder name)
        ev(hist4, "2026-08", "b_close_coat_junkpn", close_ev(
           E4, "134-2201-970 4065 Rev. A00 Rev. A00", "16588.1",
           [cs4(i) for i in range(1, 6)], "2026-08-03 14:10:00"))
        # clean coated close
        ev(hist4, "2026-08", "c_close_coat", close_ev(E4, CP4, "16588.1",
           [cs4(i) for i in range(21, 71)], "2026-08-12 09:25:00"))
        # HIS MANUAL IMPORT: coated serials, but pn field = TESTING P/N (the bug)
        ev(hist4, "2026-09", "d_bf_mislabeled", {"event_type": "backfill_shipped",
           "eng": E4, "pn": TP4, "wo": "16588.1", "ts": "2026-09-02 13:03:00",
           "ship_date": "7/27/2026", "serials": [cs4(i) for i in range(6, 11)],
           "production_data": ""})
        # testing backfill + receive of 10 (matches his description)
        ev(hist4, "2026-09", "e_bf_test", {"event_type": "backfill_shipped",
           "eng": E4, "pn": TP4, "wo": "16587.1", "ts": "2026-09-02 09:35:00",
           "ship_date": "6/29/2026", "serials": [ts4(i) for i in range(1, 11)],
           "production_data": ""})
        ev(hist4, "2026-09", "f_recv", dict(recv_ev(E4, TP4,
           [ts4(i) for i in range(1, 11)], "2026-09-02 09:36:00", "Anita Patel"), failed=[]))
        st4 = A.build_receiving_state(E4, TP4)
        m4 = A.receiving_summary(st4)
        check("4065real: coated bucket = 60 (junk-pn close 5 + clean close 50 + mislabeled import 5)",
              len(st4["coated"]) == 60)
        check("4065real: mislabeled import serials NOT in the testing shipped baseline",
              A._serial_identity(cs4(7), TP4) not in st4["shipped"]
              and not any("970" in (v["serial"].split("-")[1] if "-" in v["serial"] else "")
                          for v in st4["shipped"].values()))
        check("4065real: shipped baseline = 60 testing boards (50 close + 10 backfill)",
              m4["shipped"] == 60)
        check("4065real: received 10, coated&shipped among them counted by seq",
              m4["received"] == 10 and m4["coated"] == sum(
                  1 for v in st4["received"].values() if A.serial_seq(v["serial"]) in st4["coated"]))
        # dup-check routing: the -970 index sees the mislabeled serials; -900 doesn't
        idx900 = A.build_serial_history(E4, TP4)
        idx970 = A.build_serial_history(E4, CP4)
        check("4065real: -900 dup index has NO coated-form serials",
              A._serial_identity(cs4(7), CP4) not in idx900
              and A._serial_identity(ts4(75), TP4) in idx900)
        check("4065real: -970 dup index HAS the mislabeled-import serials",
              A._serial_identity(cs4(7), CP4) in idx970)
        check("4065real: -970 dup index HAS the junk-pn close serials",
              A._serial_identity(cs4(3), CP4) in idx970)
        # xlsx with ALL received boards already coated (his exact complaint:
        # "Received says 10 in summary but there is no serials in first sheet")
        wo4root = tempfile.mkdtemp(prefix="pcc_4065wo_")
        wo4 = Path(wo4root) / "16587.1"; wo4.mkdir()
        st4["pd_path"] = wo4root; st4["wo_path"] = str(wo4)
        out4 = A.write_received_xlsx(st4)
        wb4 = load_workbook(out4)
        rows4 = list(wb4["In-House"].iter_rows(values_only=True))
        h4i = next(i for i, r in enumerate(rows4) if r and r[0] == "#")
        serial_rows4 = [r for r in rows4[h4i + 1:] if r and isinstance(r[0], int)]
        check("4065real xlsx: In-House sheet has NO serial rows (all coated & shipped)",
              not serial_rows4)
        check("4065real xlsx: empty In-House sheet explains where the boards went",
              any(r and r[1] and "Coated & shipped" in str(r[1]) for r in rows4[h4i + 1:]))
        check("4065real xlsx: In-House subtitle says 0 board(s)",
              "0 board(s)" in (rows4[1][0] or ""))
        check("4065real xlsx: summary shows the arithmetic (received 10 = coated 10)",
              "Received back: 10" in rows4[2][0] and "coated & shipped 10" in rows4[2][0]
              and "Coated boards on record: 60" in rows4[2][0])
        done4 = [r for r in wb4["Coated & shipped"].iter_rows(values_only=True)
                 if r and isinstance(r[0], int)]
        check("4065real xlsx: Coated sheet = ALL 60 coated shipments (close+junk-pn+import)",
              len(done4) == 60 and all(r[2] == "16588.1" for r in done4))
        c7 = next(r for r in done4 if r[1] == cs4(7))
        check("4065real xlsx: received-then-coated board links its receive date",
              str(c7[4])[:10] == "2026-09-02")
        c40 = next(r for r in done4 if r[1] == cs4(40))
        check("4065real xlsx: never-tested coated board says 'not shipped for testing' (Dhruvil 2026-09-23 wording)",
              c40[4] == "not shipped for testing")
        check("4065real xlsx: mixed job's A3 counts the coated boards that never went for testing (50)",
              "Coated boards on record: 60 (50 not shipped for testing)" in rows4[2][0])
        rcv4 = [r for r in wb4["Received from testing"].iter_rows(values_only=True)
                if r and isinstance(r[0], int)]
        check("4065real xlsx: Received log = the 10 receives, no matter what shipped",
              len(rcv4) == 10)
        wb4.close()
        shutil.rmtree(wo4root, ignore_errors=True)
    finally:
        A.HISTORY_DIR = _h4
        shutil.rmtree(hist4, ignore_errors=True)

    # ═══ 5a5. ONE FILE: merge semantics (Dhruvil 2026-09-02: "If we edit and
    #      save something in that file then keep it and add other content as we
    #      scan new boards; if some data is overlapping then raise an error") ═══
    print("one-file merge semantics:")
    hist5 = tempfile.mkdtemp(prefix="pcc_merge_")
    pd5root = tempfile.mkdtemp(prefix="pcc_mergewo_")
    _h5 = A.HISTORY_DIR
    try:
        A.HISTORY_DIR = hist5
        E5, TP5, CP5 = "4100", "150-1000-900", "150-1000-970"
        t5 = lambda i, rev="A00": f"{i:06d}-1501000900-{rev}"
        c5 = lambda i, rev="A00": f"{i:06d}-1501000970-{rev}"
        pd5 = Path(pd5root)
        (pd5 / "20000.1").mkdir(); (pd5 / "20001.1").mkdir()
        ev(hist5, "2026-06", "a_close", close_ev(E5, TP5, "20000.1",
           [t5(i) for i in range(1, 6)], "2026-06-01 09:00:00", str(pd5 / "20000.1")))
        ev(hist5, "2026-07", "b_recv", dict(recv_ev(E5, TP5, [t5(i) for i in range(1, 4)],
           "2026-07-01 10:00:00", "Anita Patel", str(pd5)), failed=[]))
        ev(hist5, "2026-07", "c_recv_unshipped", dict(recv_ev(E5, TP5, [t5(9)],
           "2026-07-02 10:00:00", "Anita Patel", str(pd5)), failed=[]))
        ev(hist5, "2026-08", "d_close_coat", close_ev(E5, CP5, "20500.1",
           [c5(i) for i in range(1, 3)], "2026-08-01 11:00:00"))
        st5 = A.build_receiving_state(E5, TP5)
        out5 = A.write_received_xlsx(st5)
        wb5 = load_workbook(out5)
        wsr5 = wb5["Received from testing"]
        n0 = sum(1 for r in wsr5.iter_rows(values_only=True) if r and isinstance(r[0], int))
        check("merge: v1 receive log = 4 rows (incl. a never-shipped board)", n0 == 4)
        row9 = next(r for r in wsr5.iter_rows() if r[1].value == t5(9))
        check("merge: never-shipped receive says 'not in shipping records'",
              row9[5].value == "not in shipping records")
        # user edits: change a cell, add a manual row, add their own sheet
        row2 = next(r for r in wsr5.iter_rows() if r[1].value == t5(2))
        row2[4].value = "Manual Edit"
        wsr5.append([99, "MANUAL-001", "PASS", "2026-01-15", "Hand Entry", "", ""])
        wb5.create_sheet("My Notes").append(["remember to ask about seq 3"])
        wb5.save(out5); wb5.close()
        # new boards scanned + a backfill that now knows t5(9)'s shipment
        ev(hist5, "2026-09", "e_recv2", dict(recv_ev(E5, TP5, [t5(4), t5(5)],
           "2026-09-01 10:00:00", "Bina Patel", str(pd5)), failed=[]))
        ev(hist5, "2026-09", "f_bf", {"event_type": "backfill_shipped", "eng": E5,
           "pn": TP5, "wo": "16999.9", "ts": "2026-09-01 09:00:00",
           "ship_date": "5/1/2026", "serials": [t5(9)], "production_data": str(pd5)})
        st5 = A.build_receiving_state(E5, TP5)
        out5b = A.write_received_xlsx(st5)
        check("merge: same file path", out5b == out5)
        check("merge: no conflicts from clean append", not st5.get("xlsx_conflicts"))
        wb5 = load_workbook(out5)
        wsr5 = wb5["Received from testing"]
        rows5 = [r for r in wsr5.iter_rows(values_only=True) if r and r[1]]
        check("merge: manual row KEPT", any(r[1] == "MANUAL-001" for r in rows5))
        check("merge: edited cell KEPT",
              next(r for r in rows5 if r[1] == t5(2))[4] == "Manual Edit")
        n1 = sum(1 for r in rows5 if isinstance(r[0], int))
        check("merge: new boards APPENDED (4 + manual + 2 new = 7)", n1 == 7)
        r4 = next(r for r in rows5 if r[1] == t5(4))
        check("merge: appended numbering continues after the manual 99", r4[0] >= 100)
        check("merge: placeholder upgraded once the shipment is known",
              next(r for r in rows5 if r[1] == t5(9))[5] == "16999.9")
        check("merge: user's own sheet survives", "My Notes" in wb5.sheetnames)
        # OVERLAP: user's row now DISAGREES with history -> error raised, file kept
        wsrx = wb5["Received from testing"]
        rowx = next(r for r in wsrx.iter_rows() if r[1].value == t5(1))
        rowx[2].value = "FAIL"
        wb5.save(out5); wb5.close()
        st5 = A.build_receiving_state(E5, TP5)
        A.write_received_xlsx(st5)
        conf5 = st5.get("xlsx_conflicts") or []
        check("overlap: conflict RAISED (file says FAIL, history says PASS)",
              any(c["serial"] == t5(1) and c["field"] == "Result"
                  and c["file"] == "FAIL" and c["history"] == "PASS" for c in conf5))
        wb5 = load_workbook(out5)
        check("overlap: the file's row was KEPT as-is",
              next(r for r in wb5["Received from testing"].iter_rows(values_only=True)
                   if r and r[1] == t5(1))[2] == "FAIL")
        # undo the disagreement for the migration test
        wsrx = wb5["Received from testing"]
        next(r for r in wsrx.iter_rows() if r[1].value == t5(1))[2].value = "PASS"
        wb5.save(out5); wb5.close()
        # LOCATION MIGRATION: a newer testing WO closes -> file MOVES, content kept
        ev(hist5, "2026-10", "g_close2", close_ev(E5, TP5, "20001.1",
           [t5(i) for i in range(20, 23)], "2026-10-01 09:00:00", str(pd5 / "20001.1")))
        st5 = A.build_receiving_state(E5, TP5)
        check("migration: state points at the NEWEST WO folder",
              st5["wo_path"] == str(pd5 / "20001.1"))
        out5c = A.write_received_xlsx(st5)
        check("migration: file now lives in the new WO folder",
              Path(out5c).parent.parent == pd5 / "20001.1" and Path(out5c).exists())
        check("migration: old copy removed", not Path(out5).exists())
        wb5 = load_workbook(out5c)
        rows5 = [r for r in wb5["Received from testing"].iter_rows(values_only=True) if r and r[1]]
        check("migration: manual row + edits + notes sheet all carried over",
              any(r[1] == "MANUAL-001" for r in rows5)
              and next(r for r in rows5 if r[1] == t5(2))[4] == "Manual Edit"
              and "My Notes" in wb5.sheetnames)
        wb5.close()
        # UNREADABLE stray copy: kept on disk, never blocks the save
        junk_dir = pd5 / "19999.1" / A.RECEIVED_FOLDER
        junk_dir.mkdir(parents=True, exist_ok=True)
        junk = junk_dir / Path(out5c).name
        junk.write_text("not an xlsx", encoding="utf-8")
        out5d = A.write_received_xlsx(st5)
        check("unreadable stray copy: save still succeeds", out5d == out5c)
        check("unreadable stray copy: file KEPT (never delete what can't be merged)",
              junk.exists())

        # ── hardening (2026-09-02 adversarial review) ────────────────────
        print("data-safety hardening:")
        tgt5 = Path(out5c); fold5 = tgt5.parent
        check("no lock file left behind after a save",
              not list(fold5.glob("*.lock")))
        check("no false-positive shape detection on our own files",
              not A._xlsx_has_shapes(tgt5))
        # RENAMED 'Received from testing' tab -> save REFUSED, file untouched
        wbx = load_workbook(out5c)
        wbx[A.RCV_SHEET].title = "My Received"; wbx.save(out5c); wbx.close()
        try:
            A.write_received_xlsx(st5); renamed_ok = False
        except RuntimeError as e:
            renamed_ok = "Received from testing" in str(e)
        except Exception:
            renamed_ok = False
        check("renamed log tab: save REFUSED with instructions (not rebuilt)", renamed_ok)
        wbx = load_workbook(out5c)
        check("renamed log tab: file left untouched (manual row still there)",
              "My Received" in wbx.sheetnames
              and any(r[1] == "MANUAL-001"
                      for r in wbx["My Received"].iter_rows(values_only=True) if r and r[1]))
        wbx["My Received"].title = A.RCV_SHEET; wbx.save(out5c); wbx.close()
        # USER FORMULA in a date cell: no duplicate row, formula never overwritten
        wbx = load_workbook(out5c)
        wsx = wbx[A.RCV_SHEET]
        rowf = next(r for r in wsx.iter_rows() if r[1].value == t5(4))
        rowf[3].value = "=TODAY()"
        n_before = sum(1 for r in wsx.iter_rows(values_only=True) if r and isinstance(r[0], int))
        wbx.save(out5c); wbx.close()
        A.write_received_xlsx(st5)
        wbx = load_workbook(out5c)
        wsx = wbx[A.RCV_SHEET]
        n_after = sum(1 for r in wsx.iter_rows(values_only=True) if r and isinstance(r[0], int))
        rowf = next(r for r in wsx.iter_rows() if r[1].value == t5(4))
        check("formula date cell: NO duplicate row appended", n_after == n_before)
        check("formula date cell: user formula preserved", rowf[3].value == "=TODAY()")
        # BLANK Result cell: quietly upgraded from history, NOT a perpetual conflict
        rowb = next(r for r in wsx.iter_rows() if r[1].value == t5(5))
        rowb[2].value = None
        wbx.save(out5c); wbx.close()
        A.write_received_xlsx(st5)
        check("blank Result: refilled from history, no conflict",
              not (st5.get("xlsx_conflicts") or [])
              and next(r for r in load_workbook(out5c)[A.RCV_SHEET].iter_rows(values_only=True)
                       if r and r[1] == t5(5))[2] == "PASS")
        # MANUAL row with a BLANK date: matched by serial, date filled, no duplicate
        wbx = load_workbook(out5c); wsx = wbx[A.RCV_SHEET]
        wsx.append([200, t5(20), "PASS", "", "Hand Entry", "", ""])
        wbx.save(out5c); wbx.close()
        ev(hist5, "2026-10", "h_recv3", dict(recv_ev(E5, TP5, [t5(20)],
           "2026-10-02 10:00:00", "Anita Patel", str(pd5)), failed=[]))
        st5 = A.build_receiving_state(E5, TP5)
        A.write_received_xlsx(st5)
        rows20 = [r for r in load_workbook(out5c)[A.RCV_SHEET].iter_rows(values_only=True)
                  if r and r[1] == t5(20)]
        check("blank-date manual row: ONE row, date filled from history",
              len(rows20) == 1 and str(rows20[0][3])[:10] == "2026-10-02")
        # CLEAN new-format stray: rows carried VERBATIM, then the copy is deleted
        def _mini_copy(dirname, serial, extra_sheet=False, date="2026-03-03",
                       by="Copy Edit", res="PASS", wo=""):
            d5 = pd5 / dirname / A.RECEIVED_FOLDER
            d5.mkdir(parents=True, exist_ok=True)
            mwb = _WB()
            msr = mwb.active; msr.title = A.RCV_SHEET
            msr["A1"] = f"RECEIVED AFTER TESTING — P/N {TP5}"
            for _ in range(4): msr.append([])
            msr.append(["#", "Serial Number", "Result", "Received", "Received By",
                        "Testing WO", "Shipped"])
            msr.append([1, serial, res, date, by, wo, ""])
            msc = mwb.create_sheet(A.COAT_SHEET)
            msc["A1"] = f"RECEIVED AFTER TESTING — P/N {TP5}"
            for _ in range(4): msc.append([])
            msc.append(["#", "Serial Number", "Coated WO", "Shipped", "Received from testing"])
            if extra_sheet:
                mwb.create_sheet("Extra notes").append(["keep me"])
            p5 = d5 / tgt5.name
            mwb.save(str(p5))
            return p5
        s_clean = _mini_copy("19998.1", "COPY-ROW-1")
        A.write_received_xlsx(st5)
        got1 = [r for r in load_workbook(out5c)[A.RCV_SHEET].iter_rows(values_only=True)
                if r and r[1] == "COPY-ROW-1"]
        check("clean stray: manual row carried VERBATIM into the one file",
              len(got1) == 1 and got1[0][4] == "Copy Edit")
        check("clean stray: copy deleted after full merge", not s_clean.exists())
        # UNCLEAN stray (user added their own sheet): rows merged, file KEPT
        s_dirty = _mini_copy("19997.1", "COPY-ROW-2", extra_sheet=True)
        A.write_received_xlsx(st5)
        got2 = [r for r in load_workbook(out5c)[A.RCV_SHEET].iter_rows(values_only=True)
                if r and r[1] == "COPY-ROW-2"]
        check("unclean stray: rows still merged", len(got2) == 1)
        check("unclean stray: file KEPT (content beyond the known footprint)",
              s_dirty.exists())
        # WRITABLE old-format copy: unknown provenance -> KEPT, never deleted
        d5w = pd5 / "19996.1" / A.RECEIVED_FOLDER; d5w.mkdir(parents=True, exist_ok=True)
        owb5 = _WB(); owb5.active.title = "Received"
        owb5.active.append(["maybe edited"]); owb5.save(str(d5w / tgt5.name))
        A.write_received_xlsx(st5)
        check("WRITABLE old-format copy: KEPT (only read-only-era copies are edit-free)",
              (d5w / tgt5.name).exists())
        # LOCK: a fresh foreign lock blocks the save; a stale one is broken
        lk = fold5 / (tgt5.name + ".lock")
        lk.write_text("other-station pid 1", encoding="utf-8")
        try:
            A.write_received_xlsx(st5); lock_ok = False
        except RuntimeError as e:
            lock_ok = "another station" in str(e)
        except Exception:
            lock_ok = False
        check("fresh foreign lock: save refuses with a clear message", lock_ok)
        check("foreign lock NOT deleted by the refused save", lk.exists())
        old = time.time() - 600
        lk.write_text("crashed-station pid 1", encoding="utf-8")
        os.utime(str(lk), (old, old))
        check("stale lock: broken (atomic rename) and the save proceeds",
              A.write_received_xlsx(st5) == out5c and not lk.exists())
        check("no lock residue after the stale break", not list(fold5.glob("*.lock*")))
        # MERGED CELLS (title block + a data row): the save must not crash
        wbm = load_workbook(out5c); wsm = wbm[A.RCV_SHEET]
        wsm.merge_cells("A1:A4")
        rown = next(r for r in wsm.iter_rows() if r[1].value == t5(1))
        wsm.merge_cells(start_row=rown[0].row, start_column=6,
                        end_row=rown[0].row, end_column=7)
        wbm.save(out5c); wbm.close()
        check("merged cells in the file: save still works (no crash)",
              A.write_received_xlsx(st5) == out5c)
        # ── round-2 hardening (verification review) ───────────────────────
        print("round-2 hardening:")
        # stray edit in a NEVER-compared column (Received By): file wins, no
        # alert, and the stray is KEPT because its edit was not carried
        s_rb = _mini_copy("19995.1", t5(3), date="2026-07-01 10:00:00", by="CORRECTED NAME")
        A.write_received_xlsx(st5)
        r3 = [r for r in load_workbook(out5c)[A.RCV_SHEET].iter_rows(values_only=True)
              if r and r[1] == t5(3)]
        check("never-compared column edit: ONE row, target's value kept",
              len(r3) == 1 and r3[0][4] == "Anita Patel")
        check("never-compared column edit: no false conflict", not st5.get("xlsx_conflicts"))
        check("never-compared column edit: stray KEPT (edit was not carried)", s_rb.exists())
        s_rb.unlink()
        # write-through cache: a stray's typed value must beat SAME-save history
        ev(hist5, "2026-10", "i_recv4", dict(recv_ev(E5, TP5, [t5(30)],
           "2026-10-03 10:00:00", "Anita Patel", str(pd5)), failed=[]))
        st5 = A.build_receiving_state(E5, TP5)
        A.write_received_xlsx(st5)     # t5(30) lands with 'not in shipping records'
        s_wt = _mini_copy("19994.1", t5(30), date="2026-10-03 10:00:00",
                          by="Anita Patel", wo="WO-TYPED")
        ev(hist5, "2026-10", "j_bf30", {"event_type": "backfill_shipped", "eng": E5,
           "pn": TP5, "wo": "17000.1", "ts": "2026-10-03 11:00:00",
           "ship_date": "10/1/2026", "serials": [t5(30)], "production_data": str(pd5)})
        st5 = A.build_receiving_state(E5, TP5)
        A.write_received_xlsx(st5)
        r30 = next(r for r in load_workbook(out5c)[A.RCV_SHEET].iter_rows(values_only=True)
                   if r and r[1] == t5(30))
        check("file-typed value beats same-save history (cache write-through)",
              r30[5] == "WO-TYPED")
        check("fully-carried stray deleted (sig verified)", not s_wt.exists())
        # blank-date row arriving FROM A STRAY + dated history in the SAME save
        s_bd = _mini_copy("19993.1", t5(31), date="", by="Hand Entry")
        ev(hist5, "2026-10", "k_recv5", dict(recv_ev(E5, TP5, [t5(31)],
           "2026-10-04 10:00:00", "Bina Patel", str(pd5)), failed=[]))
        st5 = A.build_receiving_state(E5, TP5)
        A.write_received_xlsx(st5)
        r31 = [r for r in load_workbook(out5c)[A.RCV_SHEET].iter_rows(values_only=True)
               if r and r[1] == t5(31)]
        check("stray blank-date row + dated history, ONE save: single row, date filled",
              len(r31) == 1 and str(r31[0][3])[:10] == "2026-10-04")
        check("that stray fully carried -> deleted", not s_bd.exists())
        # '=' string carried from a copy must land as TEXT, never a live formula
        wbf = load_workbook(out5c); wsf = wbf[A.RCV_SHEET]
        wsf.append([300, t5(32), "PASS", "", "Hand Entry", "", ""])
        wbf.save(out5c); wbf.close()
        s_fx = _mini_copy("19992.1", t5(32), date="=B7", by="Hand Entry")
        st5 = A.build_receiving_state(E5, TP5)
        A.write_received_xlsx(st5)
        wbf = load_workbook(out5c)
        cellf = next(r for r in wbf[A.RCV_SHEET].iter_rows() if r[1].value == t5(32))[3]
        check("'=' string from a copy lands as TEXT, not a live formula",
              cellf.value == "=B7" and cellf.data_type != "f")
        wbf.close()
        # incoming blank Result (partial manual row in a copy): no false conflict
        s_ib = _mini_copy("19990.1", t5(5), date="2026-09-01 10:00:00",
                          by="Bina Patel", res="")
        A.write_received_xlsx(st5)
        check("incoming blank Result: NO false conflict", not st5.get("xlsx_conflicts"))
        check("incoming blank Result: stray fully merged and deleted", not s_ib.exists())
        # DISPUTED record vs a blank-dated manual row: both must stand
        wbf = load_workbook(out5c); wsf = wbf[A.RCV_SHEET]
        wsf.append([310, t5(33), "FAIL", "", "Hand Entry", "", ""])
        wbf.save(out5c); wbf.close()
        ev(hist5, "2026-10", "l_recv6", dict(recv_ev(E5, TP5, [t5(33)],
           "2026-10-05 10:00:00", "Bina Patel", str(pd5)), failed=[]))
        st5 = A.build_receiving_state(E5, TP5)
        A.write_received_xlsx(st5)
        r33 = [r for r in load_workbook(out5c)[A.RCV_SHEET].iter_rows(values_only=True)
               if r and r[1] == t5(33)]
        failrow = next(r for r in r33 if r[2] == "FAIL")
        check("disputed blank-date row NOT adopted: both records stand",
              len(r33) == 2 and failrow[3] in (None, ""))
        # a cell NOTE in a stray can't be carried -> the copy is KEPT
        from openpyxl.comments import Comment
        s_cm = _mini_copy("19989.1", t5(34), date="2026-10-06", by="Inspector")
        wbc = load_workbook(s_cm)
        next(r for r in wbc[A.RCV_SHEET].iter_rows() if r[1].value == t5(34))[1].comment = \
            Comment("DO NOT SHIP — solder issue", "Inspector")
        wbc.save(s_cm); wbc.close()
        A.write_received_xlsx(st5)
        check("stray with a cell NOTE: file KEPT (note can't be carried)", s_cm.exists())
        # a user note ABOVE the header row in a stray -> the copy is KEPT
        s_a5 = _mini_copy("19988.1", t5(35), date="2026-10-07", by="Anita Patel")
        wba = load_workbook(s_a5); wba[A.RCV_SHEET]["B2"] = "NOTE: boxes 3+4 in quarantine!"
        wba.save(s_a5); wba.close()
        A.write_received_xlsx(st5)
        check("stray with a note above the header: file KEPT", s_a5.exists())
        # modern THREADED comment on the ONE file: refuse (openpyxl would drop it)
        import zipfile as _zf
        with _zf.ZipFile(out5c, "a") as z:
            z.writestr("xl/threadedComments/threadedComment1.xml", "<threadedComments/>")
        try:
            A.write_received_xlsx(st5); tc_ok = False
        except RuntimeError as e:
            tc_ok = "comment" in str(e).lower()
        except Exception:
            tc_ok = False
        check("threaded comment on the ONE file: save refused with instructions", tc_ok)
        tmpz = str(out5c) + ".rebuild"
        with _zf.ZipFile(out5c) as zin, _zf.ZipFile(tmpz, "w", _zf.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                if "threadedcomment" not in it.filename.lower():
                    zout.writestr(it, zin.read(it.filename))
        os.replace(tmpz, str(out5c))
        check("threaded comment removed: saving works again",
              A.write_received_xlsx(st5) == out5c)
        # sig-verified deletion: a copy that CHANGED since harvest is kept
        pth = pd5 / "sigtest.bin"
        pth.write_text("v1", encoding="utf-8"); sg1 = A._file_sig(pth)
        time.sleep(0.02); pth.write_text("v2-changed", encoding="utf-8")
        A._delete_merged_copies([(pth, sg1)], tgt5)
        check("copy changed since harvest: NOT deleted", pth.exists())
        A._delete_merged_copies([(pth, A._file_sig(pth))], tgt5)
        check("unchanged copy: deleted", not pth.exists())
        # _date_key normalization incl. 2-digit years
        check("_date_key: datetime / 4-digit / 2-digit / M-D-YYYY all agree",
              A._date_key("2026-05-14 10:00:00") == "2026-05-14"
              and A._date_key("5/14/2026") == "2026-05-14"
              and A._date_key("5/14/26") == "2026-05-14"
              and A._date_key("05-14-2026") == "2026-05-14")
        # UI wiring (source-level, same style as the void suite)
        src5 = Path(SCRIPTS, "app_logic.py").read_text(encoding="utf-8")
        check("Receiving footer Close goes through _on_close (unsaved-scan guard)",
              '_btn(f, "Close", self._on_close' in src5)
        check("_adopt_state rejects superseded builds (box identity)",
              "box is not self._state_box" in src5)
        check("_do_save folds new receives into receive_log too",
              'state.setdefault("receive_log", []).append' in src5
              and "_fold_receive_locally(st, kept, op, rdate, ts)" in src5)     # r4: the fold is a Tk-free helper
    finally:
        A.HISTORY_DIR = _h5
        shutil.rmtree(hist5, ignore_errors=True)
        shutil.rmtree(pd5root, ignore_errors=True)

    # ═══ 5a6. 4094: COUNT-BASED CYCLES + FOLDER TRUTH (Dhruvil 2026-09-03) ═══
    print("count-based cycles + folder truth (4094):")
    hist6 = tempfile.mkdtemp(prefix="pcc_4094_")
    cust6 = tempfile.mkdtemp(prefix="pcc_4094cust_")
    _h6 = A.HISTORY_DIR
    try:
        A.HISTORY_DIR = hist6
        E6, TP6, CP6 = "4094", "161-9658-900", "161-9658-970"
        t6 = lambda i, rev="A00": f"{i:06d}-1619658900-{rev}"
        c6 = lambda i, rev="A00": f"{i:06d}-1619658970-{rev}"
        FAIL5 = [t6(i) for i in (17, 36, 44, 45, 46)]
        job900 = Path(cust6) / "161-9658-900_4094 Rev. A00 Rev. A"
        wo900 = job900 / "Production Data" / "16583.1"
        # real 4094 shape: 50 shipped Aug-24; 5 fail; re-shipped Aug-31; the
        # RECEIVE (marking the 5 FAILED) was scanned LATE on Sep-03 (no date)
        ev(hist6, "2026-08", "a_close50", close_ev(E6, TP6, "16583.1",
           [t6(i) for i in range(1, 51)], "2026-08-24 13:39:00", str(wo900)))
        ev(hist6, "2026-08", "b_reship5", close_ev(E6, TP6, "16583.1",
           FAIL5, "2026-08-31 10:04:00", str(wo900)))
        ev(hist6, "2026-09", "c_recv50", dict(recv_ev(E6, TP6,
           [t6(i) for i in range(1, 51)], "2026-09-03 11:30:00", "Anita Patel"),
           failed=list(FAIL5)))
        # anomaly board: seq 60 shipped TWICE but its one receive was a PASS
        ev(hist6, "2026-08", "d_ship60a", close_ev(E6, TP6, "16583.1",
           [t6(60)], "2026-08-24 14:00:00", str(wo900)))
        ev(hist6, "2026-08", "e_ship60b", close_ev(E6, TP6, "16583.1",
           [t6(60)], "2026-08-31 11:00:00", str(wo900)))
        ev(hist6, "2026-09", "f_recv60", dict(recv_ev(E6, TP6, [t6(60)],
           "2026-09-03 12:00:00", "Anita Patel"), failed=[]))
        # a VOIDED shipment whose folder files will still exist
        ev(hist6, "2026-08", "g_voided", dict(close_ev(E6, TP6, "16583.1",
           [t6(90), t6(91)], "2026-08-25 09:00:00", str(wo900)), event_id="ev-void-me"))
        ev(hist6, "2026-09", "h_void", {"event_type": "void", "eng": E6, "pn": TP6,
           "wo": "16583.1", "ts": "2026-09-03 14:00:00", "target_event_id": "ev-void-me"})
        st6 = A.build_receiving_state(E6, TP6)
        m6 = A.receiving_summary(st6)
        k17 = A._serial_identity(t6(17), TP6)
        check("count-based: re-shipped failed board (2 ships, 1 receive) = AT CUSTOMER",
              not A._currently_received(st6, k17))
        check("count-based: all 5 late-scanned failed boards are OUT, not in-house",
              sum(1 for s in FAIL5 if not A._currently_received(st6, A._serial_identity(s, TP6))) == 5)
        check("count-based: normal board (1 ship, 1 receive, receive newest) = in-house",
              A._currently_received(st6, A._serial_identity(t6(2), TP6)))
        check("count-based: ship counts (seq 17 = 2, seq 2 = 1)",
              st6["ship_counts"][k17] == 2
              and st6["ship_counts"][A._serial_identity(t6(2), TP6)] == 1)
        check("count-based: voided close counts ZERO ships",
              A._serial_identity(t6(90), TP6) not in st6["ship_counts"])
        check("voided shipment remembered for the folder scan",
              any(A._norm_serial(t6(90)) in v["serials"] for v in st6["voided_ships"]))
        notes6 = A.receiving_anomaly_notes(st6)
        k60 = A._serial_identity(t6(60), TP6)
        check("anomaly note: shipped 2x but never FAILED -> flagged",
              "never recorded as FAILED" in notes6.get(k60, ""))
        check("anomaly note: a genuinely failed re-ship board is NOT flagged",
              k17 not in notes6)
        # ── the job folders: 900 + TWO 970 dirs (renamed twin, like real 4094) ──
        # (_shipfile — realistic COC/PL/Serials fixture — is module-level now)
        # 900 COC: the Aug-24 batch + the Aug-31 re-ship batch (5 serials, 2nd time)
        _shipfile(wo900 / A.COC_FOLDER, "4094-COC-1.xlsx",
                  [t6(i) for i in range(1, 51)], "2026-08-24")
        _shipfile(wo900 / A.COC_FOLDER, "4094-COC-2.xlsx", FAIL5, "2026-08-31",
                  label_beside=True)
        # Shipping Serials has ONLY the first batch -> cross-check heads-up
        _shipfile(wo900 / A.SERIALS_FOLDER, "4094-serials-1.xlsx",
                  [t6(i) for i in range(1, 51)], "2026-08-24")
        # the VOIDED shipment's file still exists -> must be skipped
        _shipfile(wo900 / A.COC_FOLDER, "4094-COC-void.xlsx",
                  [t6(90), t6(91)], "2026-08-25")
        # TWO 970 job dirs; the OLD coated shipment exists ONLY as folder files
        job970a = Path(cust6) / "161-9658-970 4094 Rev. A00 Rev. A"
        job970b = Path(cust6) / "161-9658-970_4094 Rev. A00 Rev. A"
        _shipfile(job970a / "Production Data" / "16584.1" / A.COC_FOLDER,
                  "4094-COC-coat1.xlsx", [c6(i) for i in range(1, 6)], "2026-07-20",
                  pn_label=CP6)
        _shipfile(job970b / "Production Data" / "16584.1" / A.COC_FOLDER,
                  "4094-COC-coat2.xlsx", [c6(i) for i in range(20, 25)], "2026-09-02",
                  pn_label=CP6)
        scan6 = A.scan_job_folders(st6)
        check("folder scan: ok, all three job dirs found",
              scan6["ok"] and len(scan6["jobs"]) == 3)
        check("folder scan: COC counts the re-shipped 5 twice",
              scan6["ship_counts"][k17] == 2)
        check("folder scan: 'Date:' label-beside layout parsed",
              any(b["date"] == "2026-08-31" for b in scan6["batches"]))
        check("folder scan: voided shipment's file SKIPPED",
              A._serial_identity(t6(90), TP6) not in scan6["ship_counts"])
        check("folder scan: cross-check flags the 5 (COC 2x vs Shipping Serials 1x)",
              len(scan6["discrepancies"]) == 5
              and any("COC 2×" in d and "Shipping Serials 1×" in d
                      for d in scan6["discrepancies"]))
        check("folder scan: BOTH old coated shipments found (10 boards)",
              len(scan6["coated_rows"]) == 10)
        A.fold_folder_scan(st6, scan6)
        m6b = A.receiving_summary(st6)
        check("fold: coated boards on record = 10 (found on disk, zero in history)",
              m6b["coated_all"] == 10 and len(st6["coated"]) == 10)
        check("fold: received-and-coated boards move to coated bucket (seqs 1-5 + 20-24)",
              m6b["coated"] == 10)
        check("fold: idempotent (fold twice = same counts)",
              A.fold_folder_scan(st6, scan6) is st6
              and A.receiving_summary(st6)["coated_all"] == 10
              and st6["ship_counts"][k17] == 2)
        # xlsx: Note columns + Times shipped on the status sheets
        st6["pd_path"] = str(job900 / "Production Data"); st6["wo_path"] = str(wo900)
        out6 = A.write_received_xlsx(st6)
        wb6 = load_workbook(out6)
        rows6 = list(wb6["Still at customer"].iter_rows(values_only=True))
        h6i = next(i for i, r in enumerate(rows6) if r and r[0] == "#")
        check("Still-at-customer header has Times shipped + Note",
              rows6[h6i][4] == "Times shipped" and rows6[h6i][5] == "Note")
        r17 = next(r for r in rows6[h6i + 1:] if r and r[1] == t6(17))
        check("re-shipped failed board: Times shipped 2 + re-test note",
              r17[4] == 2 and "RE-TEST" in str(r17[5]))
        r60 = next(r for r in rows6[h6i + 1:] if r and r[1] == t6(60))
        check("shipped-twice-never-failed board carries the check-records note",
              "never recorded as FAILED" in str(r60[5]))
        rows6i = list(wb6["In-House"].iter_rows(values_only=True))
        h6j = next(i for i, r in enumerate(rows6i) if r and r[0] == "#")
        check("In-House header gained the Note column", rows6i[h6j][7] == "Note")
        coat6 = [r for r in wb6["Coated & shipped"].iter_rows(values_only=True)
                 if r and isinstance(r[0], int)]
        check("folder-found coated shipments land on the Coated & shipped sheet",
              len(coat6) == 10 and any(str(r[3])[:10] == "2026-07-20" for r in coat6))
        wb6.close()
        # ── round 3: phantom rejection / per-WO counts / coated one-row /
        #    unparseable files / receive shadowing ──────────────────────────
        print("round-3 hardening:")
        check("phantom P/N cells: the bare P/N never becomes a board",
              A._serial_identity(TP6, TP6) not in scan6["ship_counts"]
              and A._serial_identity(CP6, CP6)
                  not in {A._serial_identity(r["serial"], CP6)
                          for r in scan6["coated_rows"]})
        check("phantom cert sentence: no paragraph 'serials' anywhere",
              all(len(str((scan6["ship_info"].get(k) or {}).get("serial", ""))) <= 40
                  for k in scan6["ship_counts"]))
        wo2 = job900 / "Production Data" / "16590.1"
        _shipfile(wo2 / A.SERIALS_FOLDER, "reship-serials.xlsx",
                  [t6(2), t6(70)], "2026-09-05")
        _shipfile(job970a / "Production Data" / "16584.1" / A.COC_FOLDER,
                  "coat3-coc.xlsx", [c6(i) for i in range(30, 33)], "2026-09-04",
                  pn_label=CP6)
        _shipfile(job970a / "Production Data" / "16584.1" / A.SERIALS_FOLDER,
                  "coat3-serials.xlsx", [c6(i) for i in range(30, 33)],
                  pn_label=CP6, no_date_col=True)
        bad_csv = wo900 / A.COC_FOLDER / "old-export.csv"
        bad_csv.write_bytes("Serial,Date\r\n0001-x,\xe9\r\n".encode("cp1252"))
        scan7 = A.scan_job_folders(st6)
        k2 = A._serial_identity(t6(2), TP6)
        check("per-WO: COC in WO1 + Shipping-Serials-only re-ship in WO2 = 2 ships",
              scan7["ship_counts"][k2] == 2)
        c30rows = [r for r in scan7["coated_rows"]
                   if A.serial_seq(r["serial"]) in ("30", "31", "32")]
        check("coated: dated COC + dateless serials file = ONE dated row per board",
              len(c30rows) == 3 and all(r["ts"] == "2026-09-04" for r in c30rows))
        check("unparseable csv: skipped + heads-up line, scan still ok",
              scan7["ok"] and len(scan7["skipped"]) == 1
              and any("could not read" in d for d in scan7["discrepancies"]))
        st6b = A.build_receiving_state(E6, TP6)
        A.fold_folder_scan(st6b, scan7)
        check("count-based after fold: folder-only WO2 re-ship puts the board OUT",
              not A._currently_received(st6b, k2))
        k70 = A._serial_identity(t6(70), TP6)
        check("folder-only shipment carries its date into the tie-break (sts set)",
              st6b["shipped"][k70]["sts"].startswith("2026-09-05"))
        # an old paper receive imported LATER must not shadow a newer live receive
        E7, TP7 = "4077", "150-2000-900"
        t7 = lambda i: f"{i:06d}-1502000900-A00"
        ev(hist6, "2026-06", "t7_close", close_ev(E7, TP7, "19000.1", [t7(1)],
           "2026-06-01 09:00:00"))
        ev(hist6, "2026-06", "t7_recv_live", dict(recv_ev(E7, TP7, [t7(1)],
           "2026-06-15 10:00:00", "Anita Patel"), failed=[]))
        ev(hist6, "2026-09", "t7_recv_import", dict(recv_ev(E7, TP7, [t7(1)],
           "2026-09-03 09:00:00", "Anita Patel"), received_date="2026-05-10",
           failed=[t7(1)]))
        st7 = A.build_receiving_state(E7, TP7)
        k7 = A._serial_identity(t7(1), TP7)
        check("old paper receive imported later does NOT shadow the newer receive",
              st7["received"][k7]["ts"].startswith("2026-06-15")
              and st7["received"][k7]["result"] == "pass")
        m7 = A.receiving_summary(st7)
        check("that board stays in-house/passed (not in the failed bucket)",
              m7["in_house"] == 1 and m7["failed"] == 0)
        src6 = Path(SCRIPTS, "app_logic.py").read_text(encoding="utf-8")
        check("save fold also updates recv_dates",
              'state.setdefault("recv_dates", {}).setdefault(k, set()).add((rdate or ts)[:10])' in src6)
        check("cached folder scan re-folded after every fresh walk",
              src6.count("self._fold_cached_scan()") >= 2)
        # scan failure: unknown folder -> ok False, fold is a no-op
        bad = dict(st6)
        bad["pd_path"] = str(Path(cust6) / "nope" / "jobX" / "Production Data")
        sc_bad = A.scan_job_folders(bad)
        check("unreachable folders: scan says ok=False (fallback to history-only)",
              not sc_bad["ok"] and bool(sc_bad["error"]))

        # ── rev-change folder resolution (Dhruvil 2026-09-08, job 4092): the
        #    recorded wo_path is MISSPELLED / archived; the file must land in the
        #    REAL live testing folder and NEVER fabricate the stale tree ───────
        print("rev-change folder resolution:")
        revroot = tempfile.mkdtemp(prefix="pcc_4092rev_")
        custR = Path(revroot) / "377 Anduril"
        # archived OLD rev (must be ignored), live NEW rev spelled WITH a period
        (custR / "Archive" / "136-3312-900_4092 Rev A01 Rev A" / "Production Data"
         / "16585.1").mkdir(parents=True)
        live = custR / "136-3312-900_4092 Rev B01 Rev. B02" / "Production Data" / "16585.1"
        live.mkdir(parents=True)
        # a -970 coated sibling exists too (must NOT be chosen for a -900 file)
        (custR / "136-3312-970_4092 Rev B01 Rev. B02" / "Production Data"
         / "16586.1").mkdir(parents=True)
        stR = {"eng": "4092", "pn": "136-3312-900",
               "shipped": {}, "received": {}, "coated": {}, "receive_log": [],
               "coated_log": [], "ship_counts": {}, "recv_dates": {}, "voided_ships": [],
               # STALE recorded path: 'Rev B02' (no period) — does NOT exist on disk
               "wo_path": str(custR / "136-3312-900_4092 Rev B01 Rev B02"
                              / "Production Data" / "16585.1"),
               "pd_path": str(custR / "136-3312-970_4092 Rev B01 Rev. B02" / "Production Data"),
               "walk_error": False}
        resolved = A._resolve_received_wo_dir(stR)
        check("resolver skips the stale spelling and finds the REAL live folder",
              resolved == live)
        check("resolver picks the TESTING (-900) folder, not the -970 sibling",
              "900" in resolved.parent.parent.name and "970" not in resolved.parent.parent.name)
        check("resolver ignores the archived old-rev folder",
              "Archive" not in str(resolved))
        outR = A.write_received_xlsx(stR)
        check("received file lands INSIDE the real live folder",
              Path(outR).parent.parent == live)
        check("the STALE 'Rev B02' folder was NOT fabricated",
              not (custR / "136-3312-900_4092 Rev B01 Rev B02").exists())
        check("state wo_path/pd_path corrected to the real testing folder",
              stR["wo_path"] == str(live) and stR["pd_path"] == str(live.parent))
        # a genuinely unreachable/unknown recorded path with no real match -> refuse
        stX = dict(stR)
        stX["wo_path"] = str(Path(revroot) / "gone" / "j" / "Production Data" / "9.1")
        stX["pd_path"] = str(Path(revroot) / "gone" / "j" / "Production Data")
        try:
            A.write_received_xlsx(stX); refused = False
        except RuntimeError as e:
            refused = "could not be found" in str(e)
        except Exception:
            refused = False
        check("no real folder found: save REFUSES, does not fabricate", refused)
        shutil.rmtree(revroot, ignore_errors=True)
    finally:
        A.HISTORY_DIR = _h6
        shutil.rmtree(hist6, ignore_errors=True)
        shutil.rmtree(cust6, ignore_errors=True)

    # ═══ 5a7. IMPORT + CLOSE of the SAME shipment must count ONCE (Dhruvil
    #      2026-09-08: coated serials "coming twice" — a coated board recorded
    #      as both a manual backfill import AND a real close). But two CLOSES of
    #      the same serial (a real re-test re-ship) must still count twice. ═════
    print("import+close dedup (coated coming twice):")
    hist7 = tempfile.mkdtemp(prefix="pcc_impclose_")
    _h7 = A.HISTORY_DIR
    try:
        A.HISTORY_DIR = hist7
        E8, TP8, CP8 = "4200", "170-3000-900", "170-3000-970"
        t8 = lambda i: f"{i:06d}-1703000900-A00"
        c8 = lambda i: f"{i:06d}-1703000970-A00"
        # testing: 10 shipped; boards 1-3 FAIL, re-shipped (2nd close, same WO),
        #          come back PASS; all 10 eventually received
        ev(hist7, "2026-07", "a_close10", close_ev(E8, TP8, "17001.1",
           [t8(i) for i in range(1, 11)], "2026-07-01 09:00:00"))
        ev(hist7, "2026-07", "b_recv_fail", dict(recv_ev(E8, TP8,
           [t8(i) for i in range(1, 11)], "2026-07-10 10:00:00", "Anita Patel"),
           failed=[t8(1), t8(2), t8(3)]))
        ev(hist7, "2026-07", "c_reship", close_ev(E8, TP8, "17001.1",
           [t8(1), t8(2), t8(3)], "2026-07-15 09:00:00"))   # real re-ship = 2nd close
        ev(hist7, "2026-07", "d_recv_pass", dict(recv_ev(E8, TP8, [t8(1), t8(2), t8(3)],
           "2026-07-20 10:00:00", "Anita Patel"), failed=[]))
        # coated: boards 1-10 coated & shipped, recorded BOTH as a backfill IMPORT
        # (7/25) AND a real close (8/01) — the SAME shipment entered twice
        ev(hist7, "2026-08", "e_coat_import", {"event_type": "backfill_shipped",
           "eng": E8, "pn": TP8, "wo": "17002.2", "ts": "2026-08-01 09:00:00",
           "ship_date": "7/25/2026", "serials": [c8(i) for i in range(1, 11)],
           "production_data": ""})
        ev(hist7, "2026-08", "f_coat_close", close_ev(E8, CP8, "17002.1",
           [c8(i) for i in range(1, 11)], "2026-08-01 11:00:00"))
        st8 = A.build_receiving_state(E8, TP8)
        clog = st8["coated_log"]
        ids = [A._serial_identity(e["serial"], CP8) for e in clog]
        check("coated board recorded as import+close = ONE coated_log row each",
              len(clog) == 10 and len(set(ids)) == 10)
        check("coated_log prefers the real CLOSE record (WO 17002.1, not import)",
              all(e["wo"] == "17002.1" for e in clog))
        m8 = A.receiving_summary(st8)
        check("coated_all counts 10 distinct (not 20)", m8["coated_all"] == 10)
        check("re-shipped-after-fail boards still count as 2 shipments",
              st8["ship_counts"][A._serial_identity(t8(1), TP8)] == 2
              and st8["ship_counts"][A._serial_identity(t8(5), TP8)] == 1)
        # boards 1-3: shipped 2, received 2 -> back in-house/coated (not at customer)
        check("re-shipped board that came back is NOT stuck at customer",
              A._currently_received(st8, A._serial_identity(t8(1), TP8)))
        # a pure IMPORT with no close counts as one shipment (idempotent)
        ev(hist7, "2026-08", "g_coat_import2", {"event_type": "backfill_shipped",
           "eng": E8, "pn": TP8, "wo": "17002.2", "ts": "2026-08-02 09:00:00",
           "ship_date": "7/25/2026", "serials": [c8(20)], "production_data": ""})
        st8b = A.build_receiving_state(E8, TP8)
        idb = [A._serial_identity(e["serial"], CP8) for e in st8b["coated_log"]]
        check("import-only coated board (no close) still listed once",
              idb.count(A._serial_identity(c8(20), CP8)) == 1)
    finally:
        A.HISTORY_DIR = _h7
        shutil.rmtree(hist7, ignore_errors=True)

    # ═══ 5a8. in-house tested (coated-only) jobs + opt-in reports — Dhruvil 2026-09-23 ═══
    #      Testing moved in-house: a job may ship ONLY coated (-970 / -970+letter)
    #      boards. Its report is named by the REAL coated label, lives in the -970
    #      WO folder next to the COC, and is created ONLY by a manual build (the
    #      file is the opt-in switch). Fixtures follow the LIVE 4006/4007/4008
    #      History shapes read on 2026-09-23 (letter after 970, digit-only serials,
    #      8/27 "Anduril test damage" -900-labelled boards inside a -970C close).
    print("in-house tested (coated-only) jobs + opt-in reports:")
    hist9 = tempfile.mkdtemp(prefix="pcc_inhouse_")
    cust9root = tempfile.mkdtemp(prefix="pcc_inhousecust_")
    extra_roots9 = []
    _h9 = A.HISTORY_DIR
    try:
        A.HISTORY_DIR = hist9
        # ── helpers (report code only; the dup-check helpers stay letter-blind) ──
        check("_tr_pn_parts: lettered coated label",
              A._tr_pn_parts("125-1030-970C") == ("1251030", "970", "C"))
        check("_tr_pn_parts: plain testing label",
              A._tr_pn_parts("136-9624-900") == ("1369624", "900", ""))
        check("_tr_pn_parts: dash-less coated", A._tr_pn_parts("1342201970") == ("1342201", "970", ""))
        check("_tr_pn_parts: folder-name label / N/A -> empty",
              A._tr_pn_parts("134-2201-970 4065 Rev. A00 Rev. A00") == ("", "", "")
              and A._tr_pn_parts("N/A") == ("", "", "") and A._tr_pn_parts("") == ("", "", ""))
        check("_tr_testing_label: canonical -900 keeps the label's own punctuation",
              A._tr_testing_label("125-1030-970C") == "125-1030-900"
              and A._tr_testing_label("134-2201-970") == "134-2201-900"
              and A._tr_testing_label("1342201970") == "1342201900"
              and A._tr_testing_label("136-9624-900") == "136-9624-900"
              and A._tr_testing_label("N/A") == "")
        check("_pn_sibling / testing_to_coated_pn untouched (letter-blind: they feed the dup check)",
              A._pn_sibling("125-1030-970C") == "" and A._pn_sibling("125-1030-970") == "125-1030-900"
              and A.testing_to_coated_pn("125-1030-900") == "125-1030-970")
        check("r4: _coat_ident_pn — the digit-form coated P/N for a -900 key, the P/N itself otherwise (never a lettered label); _hist_has_testing prefers the walk's snapshot",
              A._coat_ident_pn("125-1030-900") == "125-1030-970" and A._coat_ident_pn("125-1030-970C") == "125-1030-970C"
              and A._coat_ident_pn("") == "" and A._coat_ident_pn(None) == ""
              and A._hist_has_testing({"hist_has_testing": False, "shipped": {"x": {"serial": "s"}}, "pn": "1"}) is False
              and A._hist_has_testing({"hist_has_testing": True, "pn": "1"}) is True
              and A._hist_has_testing({"pn": "1", "shipped": {"x": {"serial": "s"}}}) is True
              and A._hist_has_testing({"pn": "1", "shipped": {}}) is False)

        # ── the REAL 4006 shape ────────────────────────────────────────────────
        E9, TP9, CPL9, CP9 = "4006", "125-1030-900", "125-1030-970C", "125-1030-970"
        c9 = lambda i: f"{i:06d}-1251030970-C00"
        t9 = lambda i: f"{i:06d}-1251030900-C00"
        cust9 = Path(cust9root) / "377 Anduril"
        job970 = cust9 / "125-1030-970C 4006 Rev. C"
        pd970 = job970 / "Production Data"
        wo970a = pd970 / "16493.1"; wo970b = pd970 / "16493.2"
        (wo970a / A.COC_FOLDER).mkdir(parents=True); (wo970b / A.COC_FOLDER).mkdir(parents=True)
        ev(hist9, "2026-08", "a_coat_close1", close_ev(E9, CPL9, "16493.1",
           [c9(i) for i in range(851, 861)], "2026-08-20 10:00:00", str(wo970a)))
        ev(hist9, "2026-08", "b_coat_close2", dict(close_ev(E9, CPL9, "16493.2",
           [c9(i) for i in range(861, 871)] + [t9(569), t9(571)], "2026-08-27 11:00:00",
           str(wo970b)), notes="Anduril test damage"))
        # a real old-flow job in the same history (dropdown order check)
        ev(hist9, "2026-09", "z_other_close", close_ev("9998", "555-001-900", "17700.1",
           ["000001-555-001-900-A00"], "2026-09-01 08:00:00"))
        st9 = A.build_receiving_state(E9, TP9)
        check("state: coated-only job anchors the NEWEST -970 close WO (home_side coated)",
              st9["wo_path"] == str(wo970b) and st9.get("home_side") == "coated"
              and st9["pd_path"] == str(pd970))
        check("state: coated data complete (shipped {}, received {}, 22 coated_log rows, no ship_counts)",
              st9["shipped"] == {} and st9["received"] == {} and len(st9["coated_log"]) == 22
              and st9["ship_counts"] == {})
        check("state: the REAL coated label is captured", st9.get("coated_pn_label") == CPL9)
        kb = A._serial_identity(t9(569), CP9)
        rows_b = [e for e in st9["coated_log"] if e["serial"] == t9(569)]
        check("state: Case-B -900-form serial in a -970C-labelled close -> coated log, WO + Note",
              len(rows_b) == 1 and rows_b[0]["wo"] == "16493.2"
              and "scanned with a -900 testing label" in rows_b[0]["note"]
              and "Anduril test damage" in rows_b[0]["note"]
              and "569" in st9["coated"] and kb in st9.get("coat_label_mismatch", set()))
        check("state: Case-B serial NOT in shipped / ship_counts (no phantom 'Still at customer')",
              A._serial_identity(t9(569), TP9) not in st9["shipped"]
              and A._serial_identity(t9(569), TP9) not in st9["ship_counts"])
        m9 = A.receiving_summary(st9)
        check("summary: has_testing False for a coated-only job; coated_all counts every board (22)",
              m9["has_testing"] is False and m9["coated_all"] == 22 and m9["shipped"] == 0)
        check("summary: has_testing True for the 4041 fixture (section 2)",
              A.receiving_summary(st)["has_testing"] is True)
        check("old-flow (the 2026-10-06 rule: the report homes in the -970 coated WO): section-2 state home_side coated, wo_path 16154.1, clean coated label recorded",
              st.get("home_side") == "coated" and st["wo_path"] == str(pd_path / "16154.1")
              and st.get("coated_pn_label") == CPN and st.get("testing_label_seen") is True)
        check("report name: coated-only -> the REAL coated label; old-flow -> the -900 P/N",
              A._report_pn(st9) == CPL9 and A._report_pn(st) == TPN)
        check("report name spellings: -900 key + real label + digit form (rename migration)",
              A._report_pn_spellings(st9) == [TP9, CPL9, CP9])
        check("r5: coated_all counts DISTINCT boards by seq — a Case-B -900-form row and the same board's later -970 row are ONE board; a seq-less serial keys on its identity",
              A.receiving_summary({"pn": TP9, "shipped": {}, "received": {}, "coated": {},
                                   "coated_log": [{"serial": t9(569)}, {"serial": c9(569)}, {"serial": c9(1)},
                                                  {"serial": "SNX-1251030970-A"}, {"serial": "SNX-1251030970-B"}]})["coated_all"] == 3)
        # ── receive events on a coated-only job — Dhruvil 2026-09-23, decision (a) on
        #    the open Q3 question (req-r3-2): a receive flips has_testing iff at least
        #    ONE of its serials is NOT coated-form (…970…) AND NOT a Case-B board of the
        #    job (coat_label_mismatch). Supervisor-overridden scans of a coated-form or
        #    Case-B board leave the job in-house-tested: same label, tag, home, name. ──
        RW_ = A.ReceivingWindow
        def _job_entry(eng):
            js = [j for j in A.list_testing_jobs() if j["eng"] == eng]
            return js[0] if len(js) == 1 else None
        # (1) coated-form override receive (the 'Nothing to receive here' -> override -> Save path)
        ev(hist9, "2026-09", "c_recv_override", dict(recv_ev(E9, TP9, [c9(851)],
           "2026-09-05 10:00:00", "Anita Patel", ""), failed=[]))
        st9r = A.build_receiving_state(E9, TP9); j9r = _job_entry(E9)
        check("decision (a): coated-form override receive -> has_testing False, name = real label, home coated, received {} (row logged)",
              A.receiving_summary(st9r)["has_testing"] is False and st9r.get("testing_label_seen") is False
              and st9r.get("hist_has_testing") is False and A._report_pn(st9r) == CPL9
              and st9r["wo_path"] == str(wo970b) and st9r.get("home_side") == "coated"
              and st9r["received"] == {} and len(st9r["receive_log"]) == 1
              and "_recv_serials" not in st9r)
        check("decision (a): …and the dropdown entry keeps its coated-only tag (list_testing_jobs agrees with the walk)",
              j9r is not None and j9r.get("testing") is False and j9r["pn"] == TP9 and j9r.get("coated_label") == CPL9
              and RW_._job_label(j9r).endswith(A._COATED_ONLY_TAG) and RW_._coated_only(st9r) is True
              and RW_._shown_pn(st9r) == CPL9)
        # (2) Case-B board receive: a -900-form serial whose coated identity is in coat_label_mismatch
        #     (the 8/27 damage board t9(569) — the walk sees this receive BEFORE the close that files it)
        ev(hist9, "2026-09", "c_recv_caseb", dict(recv_ev(E9, TP9, [t9(569)],
           "2026-09-06 10:00:00", "Anita Patel", ""), failed=[]))
        st9c = A.build_receiving_state(E9, TP9); j9c = _job_entry(E9)
        check("decision (a): Case-B board receive -> has_testing False, name = real label, tag kept, home coated (received holds it)",
              A.receiving_summary(st9c)["has_testing"] is False and st9c.get("testing_label_seen") is False
              and A._report_pn(st9c) == CPL9 and st9c["wo_path"] == str(wo970b) and st9c.get("home_side") == "coated"
              and len(st9c["received"]) == 1 and kb in st9c.get("coat_label_mismatch", set())
              and j9c is not None and j9c.get("testing") is False and RW_._coated_only(st9c) is True)
        # (3) a GENUINE -900-form, non-Case-B board received -> a real testing shipment came back (old rule)
        ev(hist9, "2026-09", "c_recv_real", dict(recv_ev(E9, TP9, [t9(700)],
           "2026-09-07 10:00:00", "Anita Patel", ""), failed=[]))
        st9t = A.build_receiving_state(E9, TP9); j9t = _job_entry(E9)
        check("decision (a): genuine -900-form non-Case-B receive -> has_testing True, name -> -900, tag dropped (real entry)",
              A.receiving_summary(st9t)["has_testing"] is True and st9t.get("testing_label_seen") is True
              and st9t.get("hist_has_testing") is True and A._report_pn(st9t) == TP9
              and j9t is not None and "testing" not in j9t and j9t["pn"] == TP9 and RW_._coated_only(st9t) is False)
        # a mixed receive (coated-form + Case-B + one genuine -900 board in ONE event) flips too
        os.remove(Path(hist9) / "2026" / "2026-09" / "c_recv_real.json")
        ev(hist9, "2026-09", "c_recv_mixed", dict(recv_ev(E9, TP9, [c9(852), t9(571), t9(701)],
           "2026-09-07 11:00:00", "Anita Patel", ""), failed=[]))
        st9m_ = A.build_receiving_state(E9, TP9)
        check("decision (a): ONE receive holding coated-form + Case-B + a genuine -900 board flips (at least one qualifying serial)",
              A.receiving_summary(st9m_)["has_testing"] is True and A._report_pn(st9m_) == TP9
              and "testing" not in (_job_entry(E9) or {"testing": False}))
        for _n in ("c_recv_override", "c_recv_caseb", "c_recv_mixed"):
            os.remove(Path(hist9) / "2026" / "2026-09" / f"{_n}.json")
        st9z = A.build_receiving_state(E9, TP9)
        check("decision (a): with the receives removed the job is back to coated-only (fixture restored)",
              A.receiving_summary(st9z)["has_testing"] is False and A._report_pn(st9z) == CPL9
              and (_job_entry(E9) or {}).get("testing") is False)
        # the Save's LOCAL fold (the display fallback when the pre-write re-walk fails) follows
        # the walk's rule too (review round 4): a coated-form override is logged, never `received`
        stF = dict(A.build_receiving_state(E9, TP9)); nlog = len(stF["receive_log"])
        A._fold_receive_locally(stF, [{"serial": c9(851), "result": "pass"}], "Anita Patel", "", "2026-09-20 10:00:00")
        recvF1 = dict(stF["received"]); nlogF1 = len(stF["receive_log"]); hasF1 = A._has_testing(stF)
        A._fold_receive_locally(stF, [{"serial": t9(700), "result": "fail"}], "Anita Patel", "2026-09-19", "2026-09-20 10:00:00")
        k700 = A._serial_identity(t9(700), TP9)
        check("r4: _fold_receive_locally — a coated-form override stays out of received/recv_dates (logged, has_testing still False); a genuine -900 board is folded into received / recv_dates / the log with the typed date",
              recvF1 == {} and nlogF1 == nlog + 1 and hasF1 is False
              and list(stF["received"]) == [k700] and stF["received"][k700]["result"] == "fail"
              and stF["received"][k700]["ts"] == "2026-09-19" and stF["received"][k700]["sts"] == "2026-09-19 00:00:00"
              and stF["recv_dates"] == {k700: {"2026-09-19"}}
              and len(stF["receive_log"]) == nlog + 2 and stF["receive_log"][-1]["serial"] == t9(700)
              and A._has_testing(stF) is True)
        # (4) old flow untouched: the section-2 fixture (-900 closes + receives) is a testing job as before
        check("decision (a): old-flow section-2 fixture unchanged (has_testing True via the -900 label, -900 name, real entry)",
              A.receiving_summary(st)["has_testing"] is True and st.get("testing_label_seen") is True
              and A._report_pn(st) == TPN and st.get("hist_has_testing") is True and "_recv_serials" not in st)
        # a hand-built old-flow state (no walk flags, receives only) still counts as testing
        check("decision (a): _has_testing on a hand-built receive-only state (no mismatch set) -> True as before",
              A._has_testing({"pn": TPN, "received": {"x": {"serial": tser(1)}}}) is True
              and A._has_testing({"pn": TPN, "received": {}, "shipped": {}}) is False)
        # a -900 backfill beats a NEWER -970 close: a testing-side anchor always wins
        pd900 = cust9 / "125-1030-900 4006 Rev. C" / "Production Data"
        ev(hist9, "2026-07", "a_bf_test", {"event_type": "backfill_shipped", "eng": E9, "pn": TP9,
           "wo": "16400.1", "ts": "2026-07-01 09:00:00", "ship_date": "6/1/26",
           "serials": [t9(1), t9(2)], "production_data": str(pd900)})
        st9b = A.build_receiving_state(E9, TP9)
        check("state (the 2026-10-06 rule: the report homes in the -970 coated WO): a -970 close beats a -900 backfill -> home_side coated, the -970 WO 16493.2",
              st9b.get("home_side") == "coated" and Path(st9b["wo_path"]).name == "16493.2"
              and "-970" in Path(st9b["wo_path"]).parent.parent.name
              and A.receiving_summary(st9b)["has_testing"] is True and A._report_pn(st9b) == TP9)
        os.remove(Path(hist9) / "2026" / "2026-07" / "a_bf_test.json")
        # a voided coated-labelled close carrying a -900-form serial (Case B) -> its
        # coated LOG row is purge-listed; nothing of it survives in coated_log
        ev(hist9, "2026-08", "c_void_target", dict(close_ev(E9, CPL9, "16493.2",
           [t9(999), c9(999)], "2026-08-28 09:00:00", str(wo970b)), event_id="ev-caseb-void"))
        ev(hist9, "2026-09", "d_void", {"event_type": "void", "eng": E9, "pn": CPL9,
           "wo": "16493.2", "ts": "2026-09-09 10:00:00", "target_event_id": "ev-caseb-void"})
        st9v = A.build_receiving_state(E9, TP9)
        check("void: voided coated close -> Case-B + coated rows purge-listed, not in coated_log",
              (A._serial_identity(t9(999), CP9), "2026-08-28") in st9v["voided_coat_rows"]
              and (A._serial_identity(c9(999), CP9), "2026-08-28") in st9v["voided_coat_rows"]
              and not any(e["serial"] in (t9(999), c9(999)) for e in st9v["coated_log"])
              and len(st9v["coated_log"]) == 22)
        # junk-label (pre-v58 folder-name pn) -970 close still anchors; label from the clean one
        hist9j = tempfile.mkdtemp(prefix="pcc_inhousej_")
        A.HISTORY_DIR = hist9j
        ev(hist9j, "2026-08", "b_junk", close_ev(E9, "125-1030-970C 4006 Rev. C", "16493.2",
           [c9(i) for i in range(1, 4)], "2026-08-27 09:00:00", str(wo970b)))
        ev(hist9j, "2026-08", "a_clean_older", close_ev(E9, CPL9, "16493.1",
           [c9(i) for i in range(4, 6)], "2026-08-20 09:00:00", str(wo970a)))
        stj = A.build_receiving_state(E9, TP9)
        check("state: junk-label -970 close still anchors (newest close wins), 5 coated rows",
              stj["wo_path"] == str(wo970b) and stj.get("home_side") == "coated"
              and len(stj["coated_log"]) == 5)
        check("state: the coated label comes from the newest CLEAN coated close (junk ignored)",
              stj.get("coated_pn_label") == CPL9)
        shutil.rmtree(hist9j, ignore_errors=True)
        hist9k = tempfile.mkdtemp(prefix="pcc_inhousek_")
        A.HISTORY_DIR = hist9k
        ev(hist9k, "2026-08", "a_bf_coat", {"event_type": "backfill_shipped", "eng": E9, "pn": CPL9,
           "wo": "16493.1", "ts": "2026-08-20 09:00:00", "ship_date": "8/1/26",
           "serials": [c9(i) for i in range(1, 4)], "production_data": str(pd970)})
        stk = A.build_receiving_state(E9, TP9)
        check("state: -970 backfill only -> <pd>\\<wo> anchor, home_side coated, label from the import",
              stk["wo_path"] == str(pd970 / "16493.1") and stk.get("home_side") == "coated"
              and stk["pd_path"] == str(pd970) and stk.get("coated_pn_label") == CPL9)
        shutil.rmtree(hist9k, ignore_errors=True)
        A.HISTORY_DIR = hist9

        # ── list_testing_jobs ────────────────────────────────────────────────
        jobs9 = A.list_testing_jobs()
        c_ent = [j for j in jobs9 if j["eng"] == E9]
        check("list: coated-only job listed ONCE under the -900 spelling, testing False, real label carried",
              len(c_ent) == 1 and c_ent[0]["pn"] == TP9 and c_ent[0].get("testing") is False
              and c_ent[0].get("coated_label") == CPL9 and c_ent[0]["ts"].startswith("2026-08-27"))
        check("list: no entry's pn is coated-side (…970 / …970+letter)",
              all(A._tr_pn_parts(j["pn"])[1] != "970" for j in jobs9))
        check("list: coated-only entries come AFTER every real entry (dropdown order untouched)",
              jobs9[-1]["eng"] == E9 and jobs9[0]["eng"] == "9998" and "testing" not in jobs9[0])
        ev(hist9, "2026-09", "m_close900", close_ev("4300", "134-2201-900", "16600.1",
           ["000001-1342201900-A00"], "2026-09-02 09:00:00"))
        ev(hist9, "2026-09", "n_close970", close_ev("4300", "1342201970", "16601.1",
           ["000001-1342201970-A00"], "2026-09-03 09:00:00"))
        m_ent = [j for j in A.list_testing_jobs() if j["eng"] == "4300"]
        check("list: dash-less -970 spelling attaches to the hyphenated -900 job (ONE entry, real spelling, no 'testing' key)",
              len(m_ent) == 1 and m_ent[0]["pn"] == "134-2201-900" and "testing" not in m_ent[0])
        ev(hist9, "2026-09", "o_void_me", dict(close_ev("4301", "131-3722-970A", "16490.2",
           ["000538-1313722970/A-A00"], "2026-09-02 10:00:00"), event_id="ev-coat-void"))
        ev(hist9, "2026-09", "p_void", {"event_type": "void", "eng": "4301", "pn": "131-3722-970A",
           "wo": "16490.2", "ts": "2026-09-09 10:00:00", "target_event_id": "ev-coat-void"})
        check("list: a VOIDED coated-only close registers no job",
              not any(j["eng"] == "4301" for j in A.list_testing_jobs()))
        # ── review round 1 ───────────────────────────────────────────────────
        check("r1: _tr_testing_label tolerates a space/dash before the rev letter; never hands back the coated spelling",
              A._tr_testing_label("125-1030-970 C") == "125-1030-900"
              and A._tr_testing_label("125-1030-970-C") == "125-1030-900"
              and A._tr_testing_label("125-1030-970C") == "125-1030-900")
        ev(hist9, "2026-09", "u_space_letter", close_ev("4600", "125-1040-970 C", "16650.1",
           [f"{i:06d}-1251040970-C00" for i in range(1, 4)], "2026-09-04 09:00:00"))
        e46 = [j for j in A.list_testing_jobs() if j["eng"] == "4600"]
        st46 = A.build_receiving_state("4600", "125-1040-900")
        check("r1: '970 C' label -> canonical -900 key + coated-only entry; walked as COATED (no testing flag, no Case-B note)",
              len(e46) == 1 and e46[0]["pn"] == "125-1040-900" and e46[0].get("testing") is False
              and e46[0]["coated_label"] == "125-1040-970 C"
              and st46.get("testing_label_seen") is False and st46["shipped"] == {}
              and len(st46["coated_log"]) == 3 and st46.get("coated_pn_label") == "125-1040-970 C"
              and not st46.get("coat_label_mismatch") and A._report_pn(st46) == "125-1040-970 C")
        # review round 4: the substitution agrees with the parser for EVERY tail shape _tr_pn_parts
        # accepts — a period after the letter, a 3-character separator — or the job vanished from
        # list_testing_jobs (no dropdown row, no picker row, no auto-update) while the walk named its report
        check("r4: _tr_testing_label — '970C.' and '970 - C' canonicalise like '970C'",
              A._tr_testing_label("125-1030-970C.") == "125-1030-900"
              and A._tr_testing_label("125-1030-970 - C") == "125-1030-900"
              and A._tr_testing_label("125-1030-970") == "125-1030-900"
              and A._tr_testing_label("970-1030-970C") == "970-1030-900")
        ev(hist9, "2026-09", "u_dot_letter", close_ev("4777", "125-1040-970C.", "16651.1",
           [f"{i:06d}-1251040970-C00" for i in range(1, 3)], "2026-09-04 10:00:00"))
        ev(hist9, "2026-09", "u_dash_sp_letter", close_ev("4778", "125-1040-970 - C", "16652.1",
           [f"{i:06d}-1251040970-C00" for i in range(1, 3)], "2026-09-04 10:00:00"))
        e47 = [j for j in A.list_testing_jobs() if j["eng"] in ("4777", "4778")]
        st47d = A.build_receiving_state("4777", "125-1040-900")
        check("r4: '970C.' / '970 - C' labelled jobs are LISTED once each under the canonical -900 key with their real label (list and walk agree)",
              sorted((j["eng"], j["pn"], j.get("testing"), j.get("coated_label")) for j in e47)
              == [("4777", "125-1040-900", False, "125-1040-970C."), ("4778", "125-1040-900", False, "125-1040-970 - C")]
              and st47d.get("coated_pn_label") == "125-1040-970C." and A._report_pn(st47d) == "125-1040-970C."
              and len(st47d["coated_log"]) == 2 and st47d["shipped"] == {})
        for _n in ("u_dot_letter", "u_dash_sp_letter"):
            os.remove(Path(hist9) / "2026" / "2026-09" / f"{_n}.json")
        # a -900 label typed WITH a rev letter ('125-1030-900C') is a testing shipment on both sides:
        # ONE real entry (no coated-only twin) and the walk files its board in `shipped` (review round 4)
        ev(hist9, "2026-09", "u_900C", close_ev(E9, "125-1030-900C", "16400.5", [t9(700)], "2026-09-04 11:00:00"))
        e9c = [j for j in A.list_testing_jobs() if j["eng"] == E9]
        st9c_ = A.build_receiving_state(E9, TP9)
        check("r4: a '-900C'-labelled close -> ONE real entry under the -900 core key (no coated-only entry); the walk agrees (has_testing True, shipped 1, -900 name, no stop)",
              len(e9c) == 1 and "testing" not in e9c[0] and e9c[0]["pn"] == "125-1030-900C"
              and A.receiving_summary(st9c_)["has_testing"] is True and len(st9c_["shipped"]) == 1
              and A.ReceivingWindow._coated_only(st9c_) is False and A._report_pn(st9c_) == TP9
              and st9c_.get("testing_pn_labels") == ["125-1030-900C"])
        os.remove(Path(hist9) / "2026" / "2026-09" / "u_900C.json")
        # dropdown tag and window stop must agree: a junk-label (pre-v58 folder-name pn)
        # close carrying -900-FORM serials makes the job a TESTING job on both sides
        hist9r = tempfile.mkdtemp(prefix="pcc_inhouser_")
        A.HISTORY_DIR = hist9r
        ev(hist9r, "2026-08", "a_clean", close_ev(E9, CPL9, "16493.1", [c9(i) for i in range(1, 4)],
           "2026-08-20 09:00:00", str(wo970a)))
        ev(hist9r, "2026-08", "b_junk900", close_ev(E9, "125-1030-970C 4006 Rev. C", "16493.2",
           [t9(i) for i in range(11, 14)], "2026-08-21 09:00:00", str(wo970b)))
        jr = [j for j in A.list_testing_jobs() if j["eng"] == E9]
        str_ = A.build_receiving_state(E9, TP9)
        check("r1: junk-label close with -900-form serials -> a REAL dropdown entry (no tag), and the walk agrees (has_testing, -900 name, no stop)",
              len(jr) == 1 and jr[0]["pn"] == TP9 and "testing" not in jr[0] and "coated_label" not in jr[0]
              and A.receiving_summary(str_)["has_testing"] is True and A._report_pn(str_) == TP9
              and A.ReceivingWindow._coated_only(str_) is False and len(str_["shipped"]) == 3)
        check("r4 (the 2026-10-06 rule: the report homes in the -970 coated WO): testing boards on record no longer refuse the coated anchor -> home_side coated, the -970 WO 16493.1 (nothing is created until a Build)",
              str_.get("home_side") == "coated" and Path(str_["wo_path"]).name == "16493.1")
        shutil.rmtree(hist9r, ignore_errors=True)
        # BackfillDialog walks the job with the P/N auto-filled from the PICKED WO folder —
        # a -970 folder means a -970 walk: the job's own label is never 'the coated side'
        A.HISTORY_DIR = hist
        stC = A.build_receiving_state(ENG, CPN)
        A.HISTORY_DIR = hist9
        check("r1: a -970 walk keeps every -970 close serial in shipped (600, section-2 history), nothing re-routed as Case B",
              len(stC["shipped"]) == 600 and stC["coated_log"] == [] and not stC.get("coat_label_mismatch")
              and stC.get("testing_label_seen") is True and stC.get("coated_pn_labels") == [])

        # ── folder truth on a coated-only job (pd_path now known): -970 COC with
        #    2 extra seqs, real paperwork shape (bare P/N cell + cert sentence) ──
        _shipfile(wo970b / A.COC_FOLDER, "4006-COC-2.xlsx",
                  [c9(i) for i in range(861, 871)] + [c9(880), c9(881)], "2026-08-27",
                  pn_label=CPL9)
        st9f = A.build_receiving_state(E9, TP9)
        scan9 = A.scan_job_folders(st9f)
        check("folder truth: coated-only job scans ok (1 job dir), 12 coated rows, COC date, no phantoms",
              scan9["ok"] and len(scan9["jobs"]) == 1 and len(scan9["coated_rows"]) == 12
              and all(r["ts"] == "2026-08-27" for r in scan9["coated_rows"])
              and scan9["ship_counts"] == {}
              and not any(A.serial_seq(r["serial"]) == "125" for r in scan9["coated_rows"]))
        A.fold_folder_scan(st9f, scan9)
        check("folder truth: fold adds ONLY the 2 folder-only boards (coated_all 22 -> 24)",
              A.receiving_summary(st9f)["coated_all"] == 24 and st9f["shipped"] == {})

        # ── resolver: two ordered passes by home_side; sibling tree only on Rebuild ──
        r1 = A._resolve_received_wo_dir(dict(st9))
        check("resolver: only a -970 tree -> its recorded WO (fast path)", r1 == wo970b)
        stR9 = dict(st9); stR9["wo_path"] = str(pd970 / "16493.9")      # renamed WO: gone
        os.utime(str(wo970a), (time.time() - 100, time.time() - 100))   # 16493.2 is newest
        r2 = A._resolve_received_wo_dir(stR9)
        check("resolver: -970 WO renamed -> newest EXISTING -970 WO, nothing created",
              r2 == wo970b and not (pd970 / "16493.9").exists() and stR9.get("tree_switched") is False)
        job900 = cust9 / "125-1030-900 4006 Rev. C"
        wo900 = job900 / "Production Data" / "16400.1"; wo900.mkdir(parents=True)
        os.utime(str(wo900), (time.time() + 50, time.time() + 50))      # -900 touched NEWER
        r3 = A._resolve_received_wo_dir(dict(stR9))
        check("resolver: coated-only history + BOTH trees -> the -970 WO (home tree first)",
              r3 == wo970b)
        stR9t = dict(stR9); stR9t["home_side"] = "testing"
        stR9t["wo_path"] = str(wo900.parent / "16400.9")       # stale -900 spelling
        stR9t["pd_path"] = str(pd970)                          # pd deliberately INSIDE the -970 tree
        r4 = A._resolve_received_wo_dir(stR9t)
        check("resolver: old-flow (home testing), -970 touched newer -> still the -900 WO", r4 == wo900)
        custZ = Path(tempfile.mkdtemp(prefix="pcc_inhouseZ_")) / "377 Anduril"
        extra_roots9.append(custZ.parent)
        woZ900 = custZ / "125-1030-900 4006 Rev. C" / "Production Data" / "16400.1"
        woZ900.mkdir(parents=True)
        stZ = dict(st9)
        stZ["wo_path"] = str(custZ / "125-1030-970C 4006 Rev. C" / "Production Data" / "16493.2")
        stZ["pd_path"] = str(custZ / "125-1030-970C 4006 Rev. C" / "Production Data")
        check("resolver: -970 tree gone, -900 live, no allow_tree_switch -> None",
              A._resolve_received_wo_dir(dict(stZ)) is None)
        try:
            A.write_received_xlsx(dict(stZ)); refZ = False
        except RuntimeError as e:
            refZ = "could not be found" in str(e) and "-900 or -970" in str(e)
        except Exception:
            refZ = False
        check("write: refuses ('could not be found', -900 or -970 wording), nothing fabricated",
              refZ and not (custZ / "125-1030-970C 4006 Rev. C").exists()
              and not (woZ900 / A.RECEIVED_FOLDER).exists())
        stZ2 = dict(stZ); stZ2["allow_tree_switch"] = True
        rZ = A._resolve_received_wo_dir(stZ2)
        check("resolver: same with allow_tree_switch -> the -900 WO and tree_switched True",
              rZ == woZ900 and stZ2.get("tree_switched") is True)
        _orig_iterdir = A.Path.iterdir
        def _fake_iterdir(self):
            if self.name == "Production Data" and "125-1030-900" in self.parent.name:
                raise PermissionError("no access")
            return _orig_iterdir(self)
        A.Path.iterdir = _fake_iterdir
        try:
            stU = dict(stR9t); stU["allow_tree_switch"] = True
            rU = A._resolve_received_wo_dir(stU)
        finally:
            A.Path.iterdir = _orig_iterdir
        check("resolver: unreadable matched Production Data -> None (even with allow_tree_switch)",
              rU is None)

        # ── OPT-IN: the file is the switch ───────────────────────────────────
        A.auto_update_shipping_report(E9, CPL9)
        check("optin: auto_update on a job with NO report -> writes nothing, creates no folder",
              not (wo970b / A.RECEIVED_FOLDER).exists() and not (wo970a / A.RECEIVED_FOLDER).exists()
              and not list(cust9.rglob("*Shipping Report*.xlsx")))
        has0, tgt0 = A._report_exists(st9)
        check("optin: _report_exists False before the first manual build (target path known)",
              has0 is False and tgt0 == A.received_xlsx_path(wo970b, CPL9))
        out9 = A.write_received_xlsx(st9)          # what 🔁 Build report does
        check("build: coated-only report named by the REAL coated label, inside the -970 WO next to the COC",
              Path(out9) == wo970b / A.RECEIVED_FOLDER / f"{CPL9} Shipping Report.xlsx"
              and Path(out9).exists())
        check("build: only the 'Shipping Reports' leaf was created inside the WO folder",
              sorted(p.name for p in wo970b.iterdir()) == sorted([A.COC_FOLDER, A.RECEIVED_FOLDER]))
        wb9 = load_workbook(out9)
        coat9 = [r for r in wb9["Coated & shipped"].iter_rows(values_only=True)
                 if r and isinstance(r[0], int)]
        rb = next(r for r in coat9 if r[1] == t9(569))
        check("build: Coated & shipped lists 22 rows incl. the Case-B boards with their Note",
              len(coat9) == 22 and "scanned with a -900 testing label" in str(rb[5])
              and rb[2] == "16493.2")
        out_rows9 = [r for r in wb9["Still at customer"].iter_rows(values_only=True)
                     if r and isinstance(r[0], int)]
        check("build: Still at customer has NO rows (Case B never shipped for testing)", not out_rows9)
        wb9.close()
        has1, tgt1 = A._report_exists(st9)
        check("optin: _report_exists True after the build, same path", has1 is True and tgt1 == Path(out9))
        ev(hist9, "2026-09", "q_coat_close3", close_ev(E9, CPL9, "16493.2",
           [c9(i) for i in range(871, 874)], "2026-09-09 10:00:00", str(wo970b)))
        A.auto_update_shipping_report(E9, CPL9)
        n9 = sum(1 for r in load_workbook(out9)["Coated & shipped"].iter_rows(values_only=True)
                 if r and isinstance(r[0], int))
        check("optin: auto_update on a job WITH a report -> rows appended (22 -> 25), still ONE file",
              n9 == 25 and len(list(cust9.rglob("*Shipping Report*.xlsx"))) == 1)
        st9 = A.build_receiving_state(E9, TP9)
        check("build: second save appends nothing, xlsx_conflicts == []",
              A.write_received_xlsx(st9) == out9 and st9.get("xlsx_conflicts") == []
              and sum(1 for r in load_workbook(out9)["Coated & shipped"].iter_rows(values_only=True)
                      if r and isinstance(r[0], int)) == 25)
        junk9 = wo900 / A.RECEIVED_FOLDER / Path(out9).name
        junk9.parent.mkdir(parents=True, exist_ok=True)
        junk9.write_text("not an xlsx", encoding="utf-8")
        cands9 = A._stale_received_candidates(st9, Path(out9))
        check("stale: sibling-tree copies ARE candidates, under the -900 AND the coated-label names",
              any(c == junk9 for c in cands9)
              and any(c.name == f"{TP9} Shipping Report.xlsx" and "125-1030-900 4006" in str(c) for c in cands9)
              and any(c.name == A._received_legacy_fname(CPL9) for c in cands9))
        check("stale: unreadable copy in the SIBLING tree is kept and never blocks the save",
              A.write_received_xlsx(st9) == out9 and junk9.exists())
        junk9.unlink(); junk9.parent.rmdir()

        # ── migrate: the -970 tree loses its WO folders (report parked at the PD
        #    level, like an old-era file); only a manual Rebuild re-homes it ────
        wbm = load_workbook(out9)
        wbm["Coated & shipped"].append([99, "MANUAL-COAT-1", "16493.2", "2026-09-01", "", "hand note"])
        wbm.save(out9); wbm.close()
        shutil.move(str(wo970b / A.RECEIVED_FOLDER), str(pd970 / A.RECEIVED_FOLDER))
        shutil.rmtree(wo970a); shutil.rmtree(wo970b)
        parked = pd970 / A.RECEIVED_FOLDER / Path(out9).name
        st9m = A.build_receiving_state(E9, TP9)      # wo_path = the (now missing) 16493.2
        check("migrate: no allow_tree_switch -> resolver None (Save refuses), parked copy untouched",
              A._resolve_received_wo_dir(dict(st9m)) is None and parked.exists())
        A.auto_update_shipping_report(E9, CPL9)
        check("migrate: auto_update never moves a report between trees (parked, -900 WO untouched)",
              parked.exists() and not (wo900 / A.RECEIVED_FOLDER).exists())
        check("r2: _report_exists on the moved job returns the copy it FOUND (never a path under the missing WO); cached at_home True",
              A._report_exists(st9m) == (True, parked)
              and A._annotate_report_exists(dict(st9m))["_report_at_home"] is True
              and A._annotate_report_exists(dict(st9m))["_report_found_at"] == parked)
        # through the REAL Rebuild-ALL loop bodies: the picker row comes from
        # _collect_report_rows (the sibling -900 tree is live -> tickable), then the
        # ticked row is rebuilt, so the summary's 'moved' / 'with folder-truth'
        # counters are seen non-zero (review rounds 1 + 2)
        rowm = A._collect_report_rows({}, jobs=[c_ent[0]])[0]
        check("r2: picker row for the moved job — 'has report', tickable (sibling -900 WO live), path = the parked copy",
              rowm["status"] == "has report" and rowm["tickable"] is True and rowm["has"] is True
              and rowm["path"] == parked and rowm["coated_only"] is True)
        boxm = {}
        A._rebuild_report_rows([rowm], boxm)
        st9m = rowm["st"]                             # the fresh walk the loop body used
        out9m = wo900 / A.RECEIVED_FOLDER / Path(out9).name
        check("migrate: picker rebuild -> report in the -900 WO, tree_switched True, same name; box moved 1 / folders 1 / created 0",
              boxm["ok"] == 1 and boxm["moved"] == 1 and boxm["folders"] == 1 and boxm["created"] == 0
              and boxm["failed"] == [] and out9m.exists()
              and st9m.get("tree_switched") is True and st9m["wo_path"] == str(wo900)
              and "allow_tree_switch" not in st9m)
        wbm = load_workbook(out9m)
        rowsm = [r for r in wbm["Coated & shipped"].iter_rows(values_only=True) if r and r[1]]
        wbm.close()
        check("migrate: manual row + note carried, all 25 history rows present",
              any(r[1] == "MANUAL-COAT-1" and r[5] == "hand note" for r in rowsm)
              and sum(1 for r in rowsm if isinstance(r[0], int) and r[1] != "MANUAL-COAT-1") == 25)
        check("migrate: the -970 copy was removed after the verified save", not parked.exists())

        # ── archive: the -900 tree is archived, the -970 tree comes back alive.
        #    The archived copy is the merge BASE once and is NEVER modified/deleted ──
        arch = cust9 / "Archive"; arch.mkdir()
        shutil.move(str(job900), str(arch / job900.name))
        archived = (arch / job900.name / "Production Data" / "16400.1" / A.RECEIVED_FOLDER
                    / Path(out9).name)
        asig = A._file_sig(archived)
        wo970b.mkdir(parents=True)
        st9a = A.build_receiving_state(E9, TP9)      # 16493.2 exists again -> fast path
        out9a = A.write_received_xlsx(st9a)
        rows_a = [r for r in load_workbook(out9a)["Coated & shipped"].iter_rows(values_only=True)
                  if r and r[1]]
        check("archive: report re-homed in the live -970 WO with EVERY row (25 + manual) from the archived base",
              Path(out9a) == wo970b / A.RECEIVED_FOLDER / Path(out9).name
              and any(r[1] == "MANUAL-COAT-1" for r in rows_a)
              and sum(1 for r in rows_a if isinstance(r[0], int)) == 26)
        check("archive: archived copy still exists with an unchanged (mtime, size) fingerprint",
              archived.exists() and A._file_sig(archived) == asig)
        st9a = A.build_receiving_state(E9, TP9)
        A.write_received_xlsx(st9a)
        n_a = sum(1 for r in load_workbook(out9a)["Coated & shipped"].iter_rows(values_only=True)
                  if r and isinstance(r[0], int))
        check("archive: second save appends nothing (archived copy never re-read), xlsx_conflicts == []",
              n_a == 26 and st9a.get("xlsx_conflicts") == [] and archived.exists())
        check("archive: _under_archive guard", A._under_archive(archived) and not A._under_archive(out9a))

        # ── optin: a LEGACY-named file counts as an existing report (and migrates) ──
        custL = Path(tempfile.mkdtemp(prefix="pcc_inhouseL_")) / "20 TestCust"
        extra_roots9.append(custL.parent)
        woL = custL / "150-2200-900 4310 Rev. A" / "Production Data" / "16700.1"
        (woL / A.SERIALS_FOLDER).mkdir(parents=True)
        tL = lambda i: f"{i:06d}-1502200900-A00"
        ev(hist9, "2026-08", "r_close_L", close_ev("4310", "150-2200-900", "16700.1",
           [tL(i) for i in range(1, 4)], "2026-08-10 09:00:00", str(woL)))
        stL = A.build_receiving_state("4310", "150-2200-900")
        check("optin: old-flow job without any file -> _report_exists False",
              A._report_exists(stL)[0] is False and A._report_exists(stL)[1] == A.received_xlsx_path(woL, "150-2200-900"))
        legL = woL / A.RECEIVED_FOLDER_LEGACY / A._received_legacy_fname("150-2200-900")
        legL.parent.mkdir(parents=True)
        lwb = _WB(); lws = lwb.active; lws.title = A.RCV_SHEET
        lws["A1"] = "RECEIVED AFTER TESTING — P/N 150-2200-900"
        for _ in range(4): lws.append([])
        lws.append(["#", "Serial Number", "Result", "Received", "Received By", "Testing WO", "Shipped"])
        lws.append([1, "LEGACY-ROW-1", "PASS", "2026-03-03", "Copy Edit", "", ""])
        lwb.save(str(legL)); lwb.close()
        hasL, tgtL = A._report_exists(stL)
        check("optin: legacy-named file ('Received after testing\\<PN> - Received after testing.xlsx') counts as existing",
              hasL is True and tgtL == A.received_xlsx_path(woL, "150-2200-900"))
        stL2 = A._annotate_report_exists(A.build_receiving_state("4310", "150-2200-900"))
        check("r2: legacy-only copy -> cached True, _report_at_home False, and the copy itself is named (_report_found_at)",
              stL2["_report_exists"] is True and stL2["_report_at_home"] is False
              and stL2["_report_path"] == tgtL and stL2["_report_found_at"] == legL)
        A.auto_update_shipping_report("4310", "150-2200-900")
        check("optin: auto_update migrates the legacy file into 'Shipping Reports\\<PN> Shipping Report.xlsx'",
              tgtL.exists() and not legL.exists()
              and any(r[1] == "LEGACY-ROW-1" for r in load_workbook(tgtL)[A.RCV_SHEET].iter_rows(values_only=True)
                      if r and r[1]))
        A._mark_report_written(stL2, tgtL)          # what _regen_xlsx does after a successful write
        check("r2: _mark_report_written flips ALL the cached keys (at_home True, path / found = the file just written)",
              stL2["_report_exists"] is True and stL2["_report_at_home"] is True
              and stL2["_report_path"] == tgtL and stL2["_report_found_at"] == tgtL)
        _orig_exists = A.Path.exists
        def _fake_exists(self, *a, **k):
            if "Shipping Report" in self.name or "Received after testing" in self.name:
                raise PermissionError("no access")
            return _orig_exists(self, *a, **k)
        A.Path.exists = _fake_exists
        try:
            hasP, tgtP = A._report_exists(stL)
        finally:
            A.Path.exists = _orig_exists
        check("optin: _report_exists is None (could not check) under PermissionError — never True, so nothing is created/touched",
              hasP is None and tgtP == tgtL)
        # ── review round 1: tri-state — 'could not check' is NOT 'no report' ──
        nope = (Path(tempfile.gettempdir()) / "pcc_nope_r1_zz" / "377 Anduril"
                / "150-2200-900 4500 Rev. A" / "Production Data" / "16900.1")
        stN = dict(stL, eng="4500", wo_path=str(nope), pd_path=str(nope.parent))
        hasN, tgtN = A._report_exists(stN)
        A._annotate_report_exists(stN)
        check("r1: unreachable job folders -> _report_exists None (not False), cached None, target still named",
              hasN is None and tgtN == A.received_xlsx_path(nope, "150-2200-900")
              and stN["_report_exists"] is None and stN["_report_at_home"] is None)
        check("r1: the receiving-save gate answers None there (nothing written, no 'no report' claim)",
              A._receive_save_gate(stN) is None)
        ev(hist9, "2026-08", "v_nope", close_ev("4500", "150-2200-900", "16900.1", [tL(9)],
           "2026-08-11 09:00:00", str(nope)))
        rowN = A._collect_report_rows({}, jobs=[{"eng": "4500", "pn": "150-2200-900",
                                                 "ts": "2026-08-11 09:00:00"}])[0]
        check("r1: picker row for a job whose folders this PC cannot see -> \"can't check from this PC\", not tickable",
              rowN["status"] == "can't check from this PC" and rowN["tickable"] is False
              and rowN["has"] is False)
        A.auto_update_shipping_report("4500", "150-2200-900")
        check("r1: auto_update on a could-not-check job writes nothing, fabricates nothing",
              not nope.exists() and not nope.parent.parent.parent.exists())

        # ── Rebuild-ALL picker proxy (R5): several jobs, ONLY 4007 ticked -> one report ──
        job7 = cust9 / "125-1031-970D 4007 Rev. D"; wo7 = job7 / "Production Data" / "16494.2"
        (wo7 / A.COC_FOLDER).mkdir(parents=True)
        ev(hist9, "2026-08", "s_4007", close_ev("4007", "125-1031-970D", "16494.2",
           [f"{i:06d}-1251031970-D00" for i in range(362, 366)], "2026-08-24 10:00:00", str(wo7)))
        job8 = cust9 / "125-1032-970C 4008 Rev. C"; wo8 = job8 / "Production Data" / "16495.1"
        (wo8 / A.COC_FOLDER).mkdir(parents=True)
        ev(hist9, "2026-08", "t_4008", close_ev("4008", "125-1032-970C", "16495.1",
           ["000945-1251032970/C-C"], "2026-08-25 10:00:00", str(wo8)))
        jobsP = A.list_testing_jobs()
        check("list: the three real Anduril coated-only jobs (4006/4007/4008) all listed, testing False (+ the '970 C' fixture 4600)",
              {j["eng"] for j in jobsP if j.get("testing") is False} == {"4006", "4007", "4008", "4600"})
        # the REAL functions behind the picker (review round 2): _collect_report_rows ->
        # the Tk-free tick model (_picker_default_ticks / _picker_set / _picker_picked,
        # what ReportPickerDialog delegates to) -> _rebuild_report_rows on the picked rows
        before = set(cust9.rglob("*Shipping Report*.xlsx")) | set(custL.rglob("*Shipping Report*.xlsx"))
        rowsP = A._collect_report_rows({}, jobs=jobsP)
        statuses = {r["job"]["eng"]: r["status"] for r in rowsP}
        ticksP = A._picker_default_ticks(rowsP)
        check("optin: picker statuses (has report / no report / no WO folder known) per job",
              statuses["4006"] == "has report" and statuses["4310"] == "has report"
              and statuses["4007"] == "no report" and statuses["4008"] == "no report"
              and statuses["9998"] == "no WO folder known")
        check("r2: the picker opens with ONLY the jobs that already have a report ticked (has AND tickable)",
              {r["job"]["eng"] for r, t in zip(rowsP, ticksP) if t} == {"4006", "4310"})
        for i in range(len(rowsP)): A._picker_set(rowsP, ticksP, i, False)      # 'Select none'
        i7 = next(i for i, r in enumerate(rowsP) if r["job"]["eng"] == "4007")
        i98 = next(i for i, r in enumerate(rowsP) if r["job"]["eng"] == "9998")
        check("r2: an un-tickable row (no WO folder known) refuses a tick; a tickable one takes it",
              A._picker_set(rowsP, ticksP, i98, True) is False and ticksP[i98] is False
              and A._picker_set(rowsP, ticksP, i7, True) is True and ticksP[i7] is True)
        pickedP = A._picker_picked(rowsP, ticksP)
        boxP = {}
        A._rebuild_report_rows(pickedP, boxP)
        after = set(cust9.rglob("*Shipping Report*.xlsx")) | set(custL.rglob("*Shipping Report*.xlsx"))
        new_files = after - before
        check("optin: Rebuild-ALL loop body with only 4007 ticked -> exactly ONE new report, named by its label",
              [r["job"]["eng"] for r in pickedP] == ["4007"] and boxP["ok"] == 1 and boxP["created"] == 1
              and boxP["failed"] == [] and len(new_files) == 1
              and next(iter(new_files)) == wo7 / A.RECEIVED_FOLDER / "125-1031-970D Shipping Report.xlsx")
        check("optin: unticked jobs untouched (4008 has a folder but still no file/leaf)",
              not (wo8 / A.RECEIVED_FOLDER).exists())

        # ── Receiving window wiring (implementer 3): the REAL Rebuild-ALL loop bodies
        #    (_collect_report_rows / _rebuild_report_rows are what the worker threads
        #    run), the state annotation, the BackfillDialog guard, exact UI texts ──
        _orig_brs = A.build_receiving_state
        def _brs_fail(eng, pn, progress=None):
            if str(eng) == "9998": raise OSError("share hiccup")
            return _orig_brs(eng, pn, progress)
        A.build_receiving_state = _brs_fail
        try:
            boxC = {}
            rowsC = A._collect_report_rows(boxC, jobs=jobsP)
        finally:
            A.build_receiving_state = _orig_brs
        byeng = {r["job"]["eng"]: r for r in rowsC}
        check("picker rows: one per job, done flag, n/i counters filled",
              boxC.get("done") is True and boxC["n"] == len(jobsP) == len(rowsC) and boxC["i"] == len(jobsP))
        check("picker rows: statuses has report / no report / history unreadable, tickable accordingly",
              byeng["4006"]["status"] == "has report" and byeng["4006"]["has"] is True and byeng["4006"]["tickable"]
              and byeng["4007"]["status"] == "has report"
              and byeng["4008"]["status"] == "no report" and byeng["4008"]["has"] is False and byeng["4008"]["tickable"]
              and byeng["9998"]["status"] == "history unreadable" and byeng["9998"]["tickable"] is False
              and byeng["9998"]["st"] is None)
        check("picker rows: coated-only flag per job, last shipment date (yyyy-mm-dd)",
              byeng["4008"]["coated_only"] is True and byeng["4310"]["coated_only"] is False
              and byeng["4008"]["last_ts"] == "2026-08-25" and byeng["4310"]["last_ts"] == "2026-08-10")
        check("picker rows: the 'no WO folder known' status (9998 walked normally)",
              A._collect_report_rows({}, jobs=[j for j in jobsP if j["eng"] == "9998"])[0]["status"] == "no WO folder known")
        before3 = set(cust9.rglob("*Shipping Report*.xlsx")) | set(custL.rglob("*Shipping Report*.xlsx"))
        mt4006 = Path(out9).stat().st_mtime; time.sleep(0.05)
        mts_before = {p: p.stat().st_mtime for p in before3 if p != Path(out9)}
        box3 = {}
        A._rebuild_report_rows([byeng["4006"], byeng["4008"], byeng["9998"]], box3)
        after3 = set(cust9.rglob("*Shipping Report*.xlsx")) | set(custL.rglob("*Shipping Report*.xlsx"))
        check("picker rebuild: ticked 4006 (has) + 4008 (none) -> ok 2, created 1, coated-only 2, folders 2, moved 0; the unreadable row fails softly",
              box3.get("done") is True and box3["ok"] == 2 and box3["created"] == 1 and box3["coated_only"] == 2
              and box3["folders"] == 2 and box3["moved"] == 0 and len(box3["failed"]) == 1
              and "9998" not in str(after3))
        check("picker rebuild: exactly ONE new file (4008, real label, in its -970 WO); 4006 updated in place",
              after3 - before3 == {wo8 / A.RECEIVED_FOLDER / "125-1032-970C Shipping Report.xlsx"}
              and Path(out9).stat().st_mtime > mt4006)
        check("picker rebuild: allow_tree_switch is dropped from the reused states afterwards",
              "allow_tree_switch" not in byeng["4006"]["st"] and "allow_tree_switch" not in byeng["4008"]["st"])
        check("picker rebuild: unticked 4310 / 4007 files (and the archived copy) untouched — mtimes unchanged",
              len(mts_before) >= 2 and all(p.stat().st_mtime == mts_before[p] for p in mts_before))
        shutil.rmtree(wo8 / A.RECEIVED_FOLDER)      # the R3 proxy below needs 4008 WITHOUT a report
        stA = A._annotate_report_exists(A.build_receiving_state(E9, TP9))
        check("annotate: worker-thread cache == _report_exists (bool + path)",
              stA["_report_exists"] is True and stA["_report_path"] == A._report_exists(stA)[1] == Path(out9))
        stA8 = A._annotate_report_exists(A.build_receiving_state("4008", "125-1032-900"))
        check("annotate: no report -> False + the would-be target path",
              stA8["_report_exists"] is False
              and stA8["_report_path"] == wo8 / A.RECEIVED_FOLDER / "125-1032-970C Shipping Report.xlsx")
        RW = A.ReceivingWindow
        check("window: coated-only predicate follows has_testing (never Tk)",
              RW._coated_only(st9) is True and RW._coated_only(st) is False and RW._coated_only(None) is False
              and RW._coated_only(dict(st9, walk_error=True)) is False)
        check("window: dropdown labels — old-flow verbatim, coated-only under the REAL label + tag",
              RW._job_label({"eng": "4041", "pn": TPN}) == "Eng 4041  ·  P/N 136-9624-900"
              and RW._job_label(c_ent[0]) == "Eng 4006  ·  P/N 125-1030-970C  (in-house tested — coated only)")
        check("window: exact stop / save texts (class constants)",
              RW._NOTHING_MSG == ("Nothing to receive on this job — its boards are tested in-house; coated shipments "
                                  "are tracked in the job's Shipping Report once it has been built (🔁 Build report).")
              and RW._COATED_HINT == ("\n\nThis looks like a COATED (…970) label. Coated boards are not received here — "
                                      "their shipments are tracked in the job's Shipping Report once it has been built "
                                      "(🔁 Build report; the file lives in the work-order folder, next to the COC).")
              and RW._NOT_IN_RECORDS_MSG.startswith("was never shipped for testing according to the records. "
                                                     "In-house tested boards must NOT be received here")
              and RW._NOT_IN_RECORDS_MSG.endswith("(or run 'Import older shipments…' first).")
              # r7: under the opt-in model nothing is 'tracked automatically' until Build report was clicked
              and "automatically" not in RW._NOTHING_MSG + RW._COATED_HINT + RW._NOT_IN_RECORDS_MSG)
        check("backfill guard: all-sibling (-970) serials under a -900 P/N -> True; mixed / own / empty / lettered -> False",
              A._backfill_serials_look_coated(TP9, [c9(1), c9(2)]) is True
              and A._backfill_serials_look_coated(TP9, [c9(1), t9(1)]) is False
              and A._backfill_serials_look_coated(TP9, [t9(1)]) is False
              and A._backfill_serials_look_coated(TP9, []) is False
              and A._backfill_serials_look_coated(CPL9, [c9(1)]) is False
              and A._backfill_serials_look_coated(TPN, [cser(5), cser(6)]) is True)
        srcW = Path(SCRIPTS, "app_logic.py").read_text(encoding="utf-8")
        check("wiring: single-scan 'Not in shipping records' stop + pending sweep, paste keeps the reworded warning hint",
              srcW.count('self._issue(v, "Not in shipping records", self._NOT_IN_RECORDS_MSG)') == 2
              and 'elif e["status"] == "pending" and k not in st["shipped"]' in srcW
              and "of them are not in the shipping records (kept — if they are OLD" in srcW
              and "boards should not be received here):" in srcW)
        check("wiring: 'Which job?' coated sentence gated on a …970 scan (scan + paste)",
              srcW.count("re.search(r'\\d{6,7}970', _alnum(") == 2 and "+ hint, parent=self)" in srcW)
        check("wiring: coated-only stop on scan, paste (one for the list) and the pre-save sweep (+ the r4 dup-check exclusion)",
              srcW.count('self._issue(v, "Nothing to receive here", self._NOTHING_MSG)') == 2
              and srcW.count('"Nothing to receive here"') == 4 and srcW.count("self._NOTHING_MSG") == 3
              and "approve=lines" in srcW
              and '- {"Nothing to receive here", "Not in shipping records", "Coated shipment on paperwork only"})' in srcW)   # r6: the records override too; r7: the paperwork stop
        check("wiring: _do_save is UPDATE-ONLY through the gate (regen only behind it), exact popup endings",
              "gate = _receive_save_gate(cur)" in srcW
              and "        if gate:\n            out = self._regen_xlsx()" in srcW
              and srcW.count("self._regen_xlsx()") == 2        # the gated save + the manual Rebuild
              and "No Shipping Report exists for this job yet, so nothing was written to Excel — " in srcW
              and "the boards ARE saved in History. Use 🔁 Build report when you want this job's report." in srcW
              and "could not be checked from this PC (no access to its " in srcW)
        check("wiring: opt-in check runs on the worker threads (both walks), button relabel, chips report lines (tri-state)",
              srcW.count("_annotate_report_exists(build_receiving_state(") == 2
              and '"🔁 Build report" if build else "🔁 Rebuild report"' in srcW     # r4: only a true first build
              and "📄 No report for this job yet — 🔁 Build report creates it (first time only; " in srcW
              and 'f"\\n📄 Report: {st.get(\'_report_path\')}"' in srcW
              and "Can't check for this job's report from this PC" in srcW
              and "Report exists under an older name or folder — {how}: {tgt}" in srcW
              and 'how = "the next Save / after-shipment update (or 🔁 Rebuild report) brings it to"' in srcW   # r5: whichever tree
              and "same_tree" not in srcW and "Rebuild report moves it to" not in srcW
              and "In-house tested — not shipped for testing" in srcW)
        check("wiring: first-time Build confirm, allow_tree_switch only on manual paths, tree-switch message",
              "This job has no report yet. Build it now from shipment history (and the " in srcW
              and srcW.count('st["allow_tree_switch"] = True') == 2
              and "The report moved to the job's {side} work-order folder" in srcW
              and 'st.pop("allow_tree_switch", None)' in srcW)
        check("wiring: Open file uses the cached resolved path + exact 'no report' text; FileInUseDialog names the resolved file",
              'No report for this job yet — use 🔁 Build report.' in srcW
              and "No -900 or -970 work-order folder is known for this job yet" in srcW
              and "no live shipment record (a -900 close or " in srcW          # r4: the text says what is true
              and "can't be checked from this PC (no access to its work-order " in srcW
              and 'received_xlsx_path(st["wo_path"], _report_pn(st)),\n' in srcW
              and 'if st.get("wo_path") else st.get("_report_path")' not in srcW   # r3: wo_path is resolved there
              and "🔁 Rebuild file" not in srcW)
        check("wiring: Rebuild ALL is the picker (dialog class, buttons, counters), heads-up carries the coated lines",
              "class ReportPickerDialog(tk.Toplevel):" in srcW and "picked = ReportPickerDialog(self, rows).result" in srcW
              and 'f"Rebuild selected ({n})"' in srcW and '"Select with report"' in srcW
              and "created (first time)" in srcW and "in-house tested (coated only)" in srcW
              and "skipped (no -900/-970 work-order folder found)" in srcW
              and "moved to the sibling (-900/-970) work-order folder" in srcW
              and '.get("coated_headsup") or [])[:6]' in srcW
              and "{_CASE_B_NOTE} (coated WO" in srcW
              and A._CASE_B_NOTE == "scanned with a -900 testing label"
              and A._COATED_ONLY_TAG == "  (in-house tested — coated only)"
              and "pn += _COATED_ONLY_TAG" in srcW and not hasattr(A.ReportPickerDialog, "_TAG"))   # r3: one name
        check("wiring: BackfillDialog sibling-P/N warning + Import confirm",
              "_coated_scan = _backfill_serials_look_coated(pn," in srcW
              and "These serials carry the COATED (…970) P/N, not the testing P/N typed above." in srcW
              and 'messagebox.askyesno("Coated serials?",' in srcW
              and "are you sure they were " in srcW)
        # receive save on a job WITHOUT a report (R3 proxy): event written, no xlsx
        pr8 = A.write_history_event("receive", {"wo": "125-1032-900", "division": "Industrial",
              "operator": "Lata Gandhi", "eng": "4008", "pn": "125-1032-900",
              "serials": ["000945-1251032970/C-C"], "failed": [],
              "production_data": str(wo8.parent)})
        st8_ = A.build_receiving_state("4008", "125-1032-900")
        has8, _t8 = A._report_exists(st8_)
        check("optin: receive save on a job without a report -> event written, _report_exists False, no xlsx/leaf",
              pr8 and Path(pr8).exists() and len(st8_["receive_log"]) == 1 and has8 is False
              and not (wo8 / A.RECEIVED_FOLDER).exists())
        # source-text pins (engine wiring)
        src9 = Path(SCRIPTS, "app_logic.py").read_text(encoding="utf-8")
        check("wiring: auto_update is UPDATE-ONLY via _report_exists",
              "if not _report_exists(st)[0]:" in src9)
        check("wiring: the writer names the file (and the A1 header) by ONE _report_pn value",
              "rpn = _report_pn(state)\n    target = received_xlsx_path(wo_path, rpn)" in src9
              and 'f"SHIPPING REPORT — P/N {rpn}"' in src9)
        check("wiring: archived copies are never deleted (both deletion gates)",
              src9.count("_under_archive(") >= 3)
        check("wiring: legacy base search includes the archive",
              "_stale_received_candidates(state, target, include_archive=True)" in src9)

        # ── xlsx wording (plan change 6 + Q1/Q3): the coated-only 4006 report ──
        #    25 history rows (22 + the 9/9 close) + the MANUAL-COAT-1 row = 26
        st9x = A.build_receiving_state(E9, TP9)
        out9x = A.write_received_xlsx(st9x)
        check("xlsx: coated-only save lands on the SAME real-label file, zero conflicts",
              Path(out9x) == Path(out9a) and st9x.get("xlsx_conflicts") == [])
        wbx = load_workbook(out9x)
        check("xlsx: coated-only file has the 4 sheets in order and opens on 'Coated & shipped'",
              wbx.sheetnames == [A.RCV_SHEET, A.COAT_SHEET, A.IH_SHEET, A.OUT_SHEET]
              and wbx.active.title == A.COAT_SHEET)
        a1x = wbx[A.RCV_SHEET]["A1"].value; a3x = wbx[A.RCV_SHEET]["A3"].value
        check("xlsx: A1 carries the REAL coated label (Q1) on every sheet",
              a1x == f"SHIPPING REPORT — P/N {CPL9}"
              and all(wbx[n]["A1"].value == a1x for n in wbx.sheetnames))
        check("xlsx: A3 == 'In-house tested — not shipped for testing    ·    Coated boards shipped (P/N 125-1030-970C): 25'",
              a3x == f"In-house tested — not shipped for testing    ·    Coated boards shipped (P/N {CPL9}): 25"
              and all(wbx[n]["A3"].value == a3x for n in wbx.sheetnames))
        check("xlsx: RCV/COAT/IH/OUT subtitles exact (not-applicable wording)",
              wbx[A.RCV_SHEET]["A2"].value == f"Engineering # {E9}    ·    received back from customer testing — none (this job's boards are tested in-house)"
              and wbx[A.COAT_SHEET]["A2"].value == f"Engineering # {E9}    ·    EVERY coated shipment on record — 26 row(s) (boards tested in-house)"
              and wbx[A.IH_SHEET]["A2"].value == f"Engineering # {E9}    ·    IN-HOUSE — not applicable (nothing shipped for testing on this job) — 0 board(s)"
              and wbx[A.OUT_SHEET]["A2"].value == f"Engineering # {E9}    ·    Still at customer — not applicable (nothing shipped for testing on this job) — 0 board(s)")
        check("xlsx: A4 log/status notes unchanged",
              str(wbx[A.RCV_SHEET]["A4"].value).startswith("Updated ")
              and str(wbx[A.RCV_SHEET]["A4"].value).endswith("log sheet — the app only APPENDS new boards here and never changes or removes existing rows; your edits are kept. Overlapping data raises an alert.")
              and str(wbx[A.IH_SHEET]["A4"].value).endswith("STATUS VIEW — rebuilt from shipping history on every save; edits here are NOT kept (edit the log sheets instead)."))
        ihx = list(wbx[A.IH_SHEET].iter_rows(values_only=True))
        outx = list(wbx[A.OUT_SHEET].iter_rows(values_only=True))
        rcvx = [r for r in wbx[A.RCV_SHEET].iter_rows(values_only=True) if r and isinstance(r[0], int)]
        check("xlsx: IH and OUT '(none — …)' rows exact, no numbered rows; RCV has 0 data rows",
              any(r and r[1] == "(none — this job's boards are tested in-house — nothing was shipped for testing)" for r in ihx)
              and any(r and r[1] == "(none — nothing was shipped for testing on this job; boards are tested in-house)" for r in outx)
              and not any(r and isinstance(r[0], int) for r in ihx + outx) and rcvx == [])
        coatx = [r for r in wbx[A.COAT_SHEET].iter_rows(values_only=True)
                 if r and isinstance(r[0], int) and r[1] != "MANUAL-COAT-1"]
        rbx = next(r for r in coatx if r[1] == t9(569))
        check("xlsx: every COAT history row col 5 == 'not shipped for testing'; Case-B row keeps its Note",
              len(coatx) == 25 and all(r[4] == "not shipped for testing" for r in coatx)
              and "scanned with a -900 testing label" in str(rbx[5]) and "Anduril test damage" in str(rbx[5]))
        wbx.close()
        check("xlsx: no testing-side anomaly note fires for a coated-only job (Case B included)",
              A.receiving_anomaly_notes(st9x) == {})
        st9x = A.build_receiving_state(E9, TP9)
        A.write_received_xlsx(st9x)
        check("xlsx: second save appends nothing, xlsx_conflicts == []",
              st9x.get("xlsx_conflicts") == []
              and sum(1 for r in load_workbook(out9x)[A.COAT_SHEET].iter_rows(values_only=True)
                      if r and isinstance(r[0], int)) == 26)
        # a pre-change file cell ('sent without testing') upgrades on the next save
        wbp = load_workbook(out9x); wsp = wbp[A.COAT_SHEET]
        rp = next(r for r in range(7, wsp.max_row + 1) if wsp.cell(row=r, column=2).value == c9(851))
        wsp.cell(row=rp, column=5, value="sent without testing"); wbp.save(out9x); wbp.close()
        st9x = A.build_receiving_state(E9, TP9)
        A.write_received_xlsx(st9x)
        wbp = load_workbook(out9x)
        check("xlsx: pre-planted 'sent without testing' cell upgrades to 'not shipped for testing' with zero conflicts",
              wbp[A.COAT_SHEET].cell(row=rp, column=5).value == "not shipped for testing"
              and st9x.get("xlsx_conflicts") == [])
        wbp.close()

        # ── folder truth, coated side (plan change 9 + Case B on disk): the REAL
        #    8/27 COC shape — coated serials + the two -900-labelled damage boards —
        #    a duplicate COC for one board and a Packing List that lists it once ──
        _shipfile(wo970b / A.COC_FOLDER, "4006-COC-2.xlsx",
                  [c9(i) for i in range(861, 871)] + [t9(569), t9(571), c9(880), c9(881)],
                  "2026-08-27", pn_label=CPL9)
        _shipfile(wo970b / A.COC_FOLDER, "4006-COC-2b.xlsx", [c9(861)], "2026-08-27", pn_label=CPL9)
        _shipfile(wo970b / A.PACKING_LIST_FOLDER, "4006-PL-2.xlsx",
                  [c9(i) for i in range(861, 871)] + [t9(569), t9(571), c9(880), c9(881)],
                  "2026-08-27", pn_label=CPL9)
        st9g = A.build_receiving_state(E9, TP9)
        scan9g = A.scan_job_folders(st9g)
        rows_g = {r["serial"]: r for r in scan9g["coated_rows"]}
        r880 = rows_g.get(c9(880)) or {}
        check("folder truth: Case-B serial in the COATED tree's COC -> coated row with the Note, never a testing ship",
              scan9g["ok"] and scan9g["ship_counts"] == {} and scan9g["ship_info"] == {}
              and (rows_g.get(t9(569)) or {}).get("note") == "scanned with a -900 testing label"
              and (rows_g.get(t9(571)) or {}).get("wo") == "16493.2"
              and "note" not in r880 and r880.get("ts") == "2026-08-27")
        check("folder truth: 2 COC vs 1 PL for one coated serial -> ONE '(WO …, coated)' heads-up line, one row",
              scan9g["discrepancies"] == [f"{c9(861)} (WO 16493.2, coated):  COC 2× · Packing List 1×"]
              and sum(1 for r in scan9g["coated_rows"] if r["serial"] == c9(861)) == 1)
        check("folder truth: 'coated board shipped 2×' heads-up under scan['coated_headsup'] (popup text only)",
              scan9g["coated_headsup"] == [f"{c9(861)}:  coated board shipped 2× — check records"])
        A.fold_folder_scan(st9g, scan9g)
        m9g = A.receiving_summary(st9g)
        cb_rows = [e for e in st9g["coated_log"] if e["serial"] == t9(569)]
        check("folder truth: fold keeps has_testing False, real-label name, coated_all 25 -> 27, Case-B row NOT duplicated",
              st9g["shipped"] == {} and m9g["has_testing"] is False and A._report_pn(st9g) == CPL9
              and m9g["coated_all"] == 27 and len(cb_rows) == 1 and "Anduril test damage" in cb_rows[0]["note"]
              and st9g["folder_scan"]["coated_headsup"] == scan9g["coated_headsup"]
              and A.receiving_anomaly_notes(st9g) == {})
        out9g = A.write_received_xlsx(st9g)
        wbg = load_workbook(out9g)
        check("folder truth: report after fold stays the SAME file, opens on Coated & shipped, A3 ': 27', no Still-at-customer rows",
              Path(out9g) == Path(out9a) and wbg.active.title == A.COAT_SHEET
              and wbg[A.COAT_SHEET]["A3"].value.endswith(f"Coated boards shipped (P/N {CPL9}): 27")
              and not any(r and isinstance(r[0], int) for r in wbg[A.OUT_SHEET].iter_rows(values_only=True))
              and any(r and r[1] == c9(880) and r[4] == "not shipped for testing"
                      for r in wbg[A.COAT_SHEET].iter_rows(values_only=True)))
        wbg.close()
        # a hand-typed Shipping Serials list (NO P/N cell) under the COATED WO listing the
        # 8/27 -900-form damage boards: the label-less fallback ('under the -970 tree')
        # routes them as Case B — never as testing shipments (review round 2). Removed
        # again afterwards: the pinned disagreement lines above count files per slot.
        _shipfile(wo970b / A.SERIALS_FOLDER, "hand-list-970.xlsx", [t9(569), t9(571)], "2026-08-27", pn_label="")
        st9s = A.build_receiving_state(E9, TP9)
        scan9s = A.scan_job_folders(st9s)
        rows_s = {r["serial"]: r for r in scan9s["coated_rows"]}
        A.fold_folder_scan(st9s, scan9s)
        check("r2: label-less -900 list under the -970 WO -> Case-B rows with the Note, ship_counts {}, has_testing still False after the fold",
              scan9s["ok"] and scan9s["ship_counts"] == {}
              and (rows_s.get(t9(569)) or {}).get("note") == A._CASE_B_NOTE
              and (rows_s.get(t9(571)) or {}).get("note") == A._CASE_B_NOTE
              and A.receiving_summary(st9s)["has_testing"] is False and st9s["shipped"] == {})
        shutil.rmtree(wo970b / A.SERIALS_FOLDER)

        # ── mixed job (has_testing True): honest '(k not shipped for testing)' suffix ──
        histM = tempfile.mkdtemp(prefix="pcc_mixed_")
        custM = Path(tempfile.mkdtemp(prefix="pcc_mixedcust_")) / "20 TestCust"
        extra_roots9.append(custM.parent)
        woM900 = custM / "150-3300-900 4400 Rev. A" / "Production Data" / "16800.1"
        woM970 = custM / "150-3300-970 4400 Rev. A" / "Production Data" / "16801.1"
        woM900.mkdir(parents=True); woM970.mkdir(parents=True)
        tM = lambda i: f"{i:06d}-1503300900-A00"; cM = lambda i: f"{i:06d}-1503300970-A00"
        A.HISTORY_DIR = histM
        try:
            ev(histM, "2026-08", "a_test", close_ev("4400", "150-3300-900", "16800.1",
               [tM(i) for i in range(1, 4)], "2026-08-01 09:00:00", str(woM900)))
            ev(histM, "2026-08", "b_coat", close_ev("4400", "150-3300-970", "16801.1",
               [cM(i) for i in range(1, 6)], "2026-08-15 09:00:00", str(woM970)))
            stM = A.build_receiving_state("4400", "150-3300-900")
            outM = A.write_received_xlsx(stM)
        finally:
            A.HISTORY_DIR = hist9
        wbM = load_workbook(outM)
        coatM = [r for r in wbM[A.COAT_SHEET].iter_rows(values_only=True) if r and isinstance(r[0], int)]
        check("mixed (the 2026-10-06 rule: the report homes in the -970 coated WO): report in the -970 WO 16801.1, still the -900 name/A1, opens on 'Received from testing'",
              Path(outM).name == "150-3300-900 Shipping Report.xlsx" and Path(outM).parent.name == A.RECEIVED_FOLDER
              and Path(outM).parent.parent.name == "16801.1" and "-970" in Path(outM).parent.parent.parent.parent.name
              and not (woM900 / A.RECEIVED_FOLDER).exists()
              and wbM[A.RCV_SHEET]["A1"].value == "SHIPPING REPORT — P/N 150-3300-900"
              and wbM.active.title == A.RCV_SHEET)
        check("mixed: A3 == today's line + ' (2 not shipped for testing)' after 'Coated boards on record: 5'",
              wbM[A.RCV_SHEET]["A3"].value ==
              "Shipped for testing: 3    ·    Still at customer: 3    ·    Received back: 0  =  in-house 0 "
              "(awaiting coating 0, failed 0)  +  coated & shipped 0    ·    Coated boards on record: 5 (2 not shipped for testing)")
        check("mixed: COAT col 5 = 'no receive record' for tested seqs, 'not shipped for testing' for the rest",
              [r[4] for r in coatM] == ["no receive record"] * 3 + ["not shipped for testing"] * 2)
        check("mixed: testing-side subtitles/rows exactly as today (OUT 3 rows, IH none-row)",
              wbM[A.OUT_SHEET]["A2"].value == "Engineering # 4400    ·    shipped for testing, NOT yet received back — 3 board(s)"
              and wbM[A.IH_SHEET]["A2"].value == "Engineering # 4400    ·    IN-HOUSE — received back, not yet coated & shipped — 0 board(s)"
              and sum(1 for r in wbM[A.OUT_SHEET].iter_rows(values_only=True) if r and isinstance(r[0], int)) == 3
              and any(r and r[1] == "(none — no boards are in-house right now)" for r in wbM[A.IH_SHEET].iter_rows(values_only=True)))
        wbM.close()
        # ── review round 2 on the mixed job: the NEW placeholder literal upgrades to a
        #    later receive date (and is styled as a placeholder), and a folder-side
        #    Case-B row on a job WITH testing history routes through coat_label_mismatch ──
        A.HISTORY_DIR = histM
        try:
            ev(histM, "2026-09", "c_recv4", recv_ev("4400", "150-3300-900", [tM(4)], "2026-09-20 10:00:00"))
            stM2 = A.build_receiving_state("4400", "150-3300-900")
            A.write_received_xlsx(stM2)
            wbM2 = load_workbook(outM); wsM2 = wbM2[A.COAT_SHEET]
            rM4 = next(r for r in range(7, wsM2.max_row + 1) if wsM2.cell(row=r, column=2).value == cM(4))
            rM5 = next(r for r in range(7, wsM2.max_row + 1) if wsM2.cell(row=r, column=2).value == cM(5))
            check("r2: a 'not shipped for testing' cell upgrades to the later receive timestamp with zero conflicts; a fresh placeholder cell is italic grey",
                  wsM2.cell(row=rM4, column=5).value == "2026-09-20 10:00:00" and stM2.get("xlsx_conflicts") == []
                  and wsM2.cell(row=rM5, column=5).value == "not shipped for testing"
                  and wsM2.cell(row=rM5, column=5).font.italic is True)
            wbM2.close()
            # the in-progress -900-serialized batch ships under the -970 label (Case B) on
            # this MIXED job, and its COC under the -970 WO lists the same -900-form serial
            ev(histM, "2026-09", "d_caseb", close_ev("4400", "150-3300-970", "16801.1", [tM(6)],
               "2026-09-21 09:00:00", str(woM970)))
            _shipfile(woM970 / A.COC_FOLDER, "coc-970.xlsx", [cM(i) for i in range(1, 6)] + [tM(6)],
                      "2026-08-15", pn_label="150-3300-970")
            stM3 = A.build_receiving_state("4400", "150-3300-900")
            scM3 = A.scan_job_folders(stM3)
            rowsM3 = {r["serial"]: r for r in scM3["coated_rows"]}
            kM6 = A._serial_identity(tM(6), "150-3300-900")           # testing-side identity (shipped / ship_counts)
            kM6c = A._serial_identity(tM(6), "150-3300-970")          # coated identity (coat_label_mismatch)
            check("r2: mixed job — History files tM(6) as Case B and the folder scan follows (coat_label_mismatch), never a testing count",
                  kM6c in stM3.get("coat_label_mismatch", set()) and stM3.get("hist_has_testing") is True
                  and scM3["ok"] and kM6 not in scM3["ship_counts"] and scM3["ship_counts"] == {}
                  and (rowsM3.get(tM(6)) or {}).get("note") == A._CASE_B_NOTE)
            A.fold_folder_scan(stM3, scM3)
            check("r2: after the fold the mixed job still has 3 boards out (Case B never lands on Still at customer)",
                  A.receiving_summary(stM3)["outstanding"] == 3 and A.receiving_summary(stM3)["shipped"] == 3
                  and kM6 not in stM3["shipped"])
        finally:
            A.HISTORY_DIR = hist9
        shutil.rmtree(histM, ignore_errors=True)

        # ── old flow byte-identical (plan change 12): the section-2 state written
        #    with the NEW code equals the 2026-09-16 output — pre-change literals
        #    copied from the BEFORE snapshot's _write_head / subtitle / row code ──
        tmpO = Path(tempfile.mkdtemp(prefix="pcc_oldflow_")); extra_roots9.append(tmpO)
        pdO = tmpO / "GizmoX [4041 Rev. A]" / "Production Data"; woO = pdO / "16153.1"; woO.mkdir(parents=True)
        stO = dict(st); stO["pd_path"] = str(pdO); stO["wo_path"] = str(woO)
        outO = A.write_received_xlsx(stO)
        wbO = load_workbook(outO)
        EXP_A1 = "SHIPPING REPORT — P/N 136-9624-900"
        EXP_A3 = ("Shipped for testing: 1500    ·    Still at customer: 600    ·    Received back: 900  =  in-house 300 "
                  "(awaiting coating 300, failed 0)  +  coated & shipped 600    ·    Coated boards on record: 600")
        EXP_A2 = {A.RCV_SHEET: "Engineering # 4041    ·    EVERY board received back from testing — 901 row(s)",
                  A.COAT_SHEET: "Engineering # 4041    ·    EVERY coated shipment on record (testing or not) — 600 row(s)",
                  A.IH_SHEET: "Engineering # 4041    ·    IN-HOUSE — received back, not yet coated & shipped — 300 board(s)",
                  A.OUT_SHEET: "Engineering # 4041    ·    shipped for testing, NOT yet received back — 600 board(s)"}
        EXP_A4_LOG = ("log sheet — the app only APPENDS new boards here and never changes or removes existing rows; "
                      "your edits are kept. Overlapping data raises an alert.")
        EXP_A4_ST = "STATUS VIEW — rebuilt from shipping history on every save; edits here are NOT kept (edit the log sheets instead)."
        check("old-flow: A1 / A3 verbatim on all four sheets (no suffix: k = 0), opens on 'Received from testing'",
              all(wbO[n]["A1"].value == EXP_A1 and wbO[n]["A3"].value == EXP_A3 for n in wbO.sheetnames)
              and wbO.active.title == A.RCV_SHEET and stO.get("xlsx_conflicts") == [])
        check("old-flow: A2 subtitles verbatim on all four sheets",
              all(wbO[n]["A2"].value == EXP_A2[n] for n in EXP_A2))
        check("old-flow: A4 notes verbatim (log vs status)",
              all(str(wbO[n]["A4"].value).startswith("Updated 20") and str(wbO[n]["A4"].value).endswith(EXP_A4_LOG)
                  for n in (A.RCV_SHEET, A.COAT_SHEET))
              and all(str(wbO[n]["A4"].value).endswith(EXP_A4_ST) for n in (A.IH_SHEET, A.OUT_SHEET)))
        hdrO = lambda ws_: next(i for i, r in enumerate(ws_.iter_rows(values_only=True)) if r and r[0] == "#")
        rowsO = {n: list(wbO[n].iter_rows(values_only=True)) for n in wbO.sheetnames}
        check("old-flow: header rows verbatim (RCV_COLS / COAT_COLS / IH = RCV_COLS / OUT)",
              list(rowsO[A.RCV_SHEET][5]) == A.RCV_COLS and list(rowsO[A.COAT_SHEET][5]) == A.COAT_COLS
              and list(rowsO[A.IH_SHEET][5]) == A.RCV_COLS
              and list(rowsO[A.OUT_SHEET][5]) == ["#", "Serial Number", "Testing WO", "Shipped", "Times shipped", "Note"])
        dO = {n: [r for r in rowsO[n][hdrO(wbO[n]) + 1:] if r and isinstance(r[0], int)] for n in wbO.sheetnames}
        r400O = next(r for r in dO[A.COAT_SHEET] if r[1] == cser(400))
        r700O = next(r for r in dO[A.IH_SHEET] if r[1] == tser(700))
        check("old-flow: row counts 901 / 600 / 300 / 600 and sample row values verbatim",
              [len(dO[n]) for n in (A.RCV_SHEET, A.COAT_SHEET, A.IH_SHEET, A.OUT_SHEET)] == [901, 600, 300, 600]
              and r400O[2] == "16154.1" and r400O[3] == "2026-08-18 10:00:00" and r400O[4] == "2026-08-05 14:00:00"
              and (r400O[5] or "") == ""
              and r700O[2] == "PASS" and r700O[4] == "Anita Patel" and r700O[5] == "16152.1" and r700O[6] == "2026-07-12 09:00:00"
              and dO[A.OUT_SHEET][0][1:5] == (tser(901), "16152.1", "2026-07-12 09:00:00", 1)
              and (dO[A.OUT_SHEET][0][5] or "") == ""
              and not any(r and isinstance(r[1], str) and r[1].startswith("(none") for n in wbO.sheetnames for r in rowsO[n]))
        wbO.close()

        # ══ review round 1 ═══════════════════════════════════════════════════
        _r1_files = lambda: sorted(p for p in cust9.rglob("125-1030-9*Shipping Report*.xlsx")
                                   if not A._under_archive(p))
        _rows_of = lambda p, sheet: [r for r in load_workbook(p)[sheet].iter_rows(values_only=True) if r and r[1]]
        # ── a rev-letter change on the coated label MIGRATES the one file (4006 was
        #    built under 970C; the next WO closes under 970D — the letter IS the rev) ──
        wo970c = pd970 / "16493.3"; (wo970c / A.COC_FOLDER).mkdir(parents=True)
        cur9 = Path(out9a)
        n_before = len(_rows_of(cur9, A.COAT_SHEET))
        check("r1: starting point — exactly ONE live 4006 report, the 970C-named one",
              _r1_files() == [cur9])
        ev(hist9, "2026-09", "w_coat_970D", close_ev(E9, "125-1030-970D", "16493.3",
           [f"{i:06d}-1251030970-D00" for i in range(901, 904)], "2026-09-12 10:00:00", str(wo970c)))
        st9d = A.build_receiving_state(E9, TP9)
        check("r1: both letters remembered (newest first); every spelling feeds the candidate names",
              st9d.get("coated_pn_label") == "125-1030-970D"
              and st9d.get("coated_pn_labels") == ["125-1030-970D", CPL9]
              and A._report_pn_spellings(st9d) == [TP9, "125-1030-970D", CPL9, CP9]
              and A._report_pn(st9d) == "125-1030-970D")
        hasD, tgtD = A._report_exists(st9d)
        check("r1: the 970C file still counts as THE report after the 970D close (opt-in kept, target = new name)",
              hasD is True and tgtD == wo970c / A.RECEIVED_FOLDER / "125-1030-970D Shipping Report.xlsx")
        A.auto_update_shipping_report(E9, "125-1030-970D")
        check("r1: the after-close update MIGRATED it: ONE file, newest label, newest WO, +3 rows, manual row carried, old copy gone",
              _r1_files() == [tgtD] and not cur9.exists()
              and len(_rows_of(tgtD, A.COAT_SHEET)) == n_before + 3
              and any(r[1] == "MANUAL-COAT-1" for r in _rows_of(tgtD, A.COAT_SHEET))
              and any(r[1] == "000901-1251030970-D00" for r in _rows_of(tgtD, A.COAT_SHEET))
              and load_workbook(tgtD)[A.COAT_SHEET]["A1"].value == "SHIPPING REPORT — P/N 125-1030-970D")

        # ── a close landing while the Rebuild-ALL picker is open is in the report it writes ──
        rows7 = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == "4007"])
        ev(hist9, "2026-09", "z_4007_late", close_ev("4007", "125-1031-970D", "16494.2",
           ["000099-1251031970-D00"], "2026-09-14 10:00:00", str(wo7)))
        box7 = {}; A._rebuild_report_rows(rows7, box7)
        f7 = wo7 / A.RECEIVED_FOLDER / "125-1031-970D Shipping Report.xlsx"
        check("r1: picker rebuild re-walks each ticked job — the late close's board is in the file it wrote",
              box7["ok"] == 1 and f7.exists()
              and any(r[1] == "000099-1251031970-D00" for r in _rows_of(f7, A.COAT_SHEET)))

        # ── the receiving-save gate (the ONE decision _do_save makes about Excel) ──
        st8g = A.build_receiving_state("4008", "125-1032-900")
        g0 = A._receive_save_gate(st8g)
        check("r1: save gate on a job without a report -> False, nothing appears (no leaf, no xlsx)",
              g0 is False and st8g["_report_exists"] is False and not (wo8 / A.RECEIVED_FOLDER).exists())
        out8 = A.write_received_xlsx(A.build_receiving_state("4008", "125-1032-900"))     # 🔁 Build report
        st8g = A.build_receiving_state("4008", "125-1032-900")
        check("r1: save gate after the manual build -> True with that file's path; the receive row is on the RCV sheet",
              A._receive_save_gate(st8g) is True and st8g["_report_path"] == Path(out8)
              and st8g["_report_at_home"] is True
              and any(r[1] == "000945-1251032970/C-C" for r in _rows_of(out8, A.RCV_SHEET)))

        # ── an ARCHIVED copy alone does not opt a job in (R6: a deleted report stays deleted) ──
        custA = Path(tempfile.mkdtemp(prefix="pcc_inhouseA_")) / "20 TestCust"; extra_roots9.append(custA.parent)
        jobA_old = custA / "161-9674-900 4095 Rev. A"; woA_old = jobA_old / "Production Data" / "16999.1"
        woA_old.mkdir(parents=True)
        tA = lambda i: f"{i:06d}-1619674900-A00"
        ev(hist9, "2026-08", "x_4095", close_ev("4095", "161-9674-900", "16999.1", [tA(i) for i in range(1, 4)],
           "2026-08-12 09:00:00", str(woA_old)))
        outA = A.write_received_xlsx(A.build_receiving_state("4095", "161-9674-900"))   # built once, by hand
        (custA / "Archive").mkdir()
        shutil.move(str(jobA_old), str(custA / "Archive" / jobA_old.name))               # old rev archived
        archA = custA / "Archive" / jobA_old.name / "Production Data" / "16999.1" / A.RECEIVED_FOLDER / Path(outA).name
        jobA_new = custA / "161-9674-900 4095 Rev. B"; woA_new = jobA_new / "Production Data" / "17001.1"
        woA_new.mkdir(parents=True)
        ev(hist9, "2026-09", "y_4095b", close_ev("4095", "161-9674-900", "17001.1", [tA(4)],
           "2026-09-13 09:00:00", str(woA_new)))
        stA1 = A.build_receiving_state("4095", "161-9674-900")
        hasA, tgtA = A._report_exists(stA1)
        A.auto_update_shipping_report("4095", "161-9674-900")
        check("r1: archive-only copy -> _report_exists False (opted out), auto_update writes nothing, archived copy untouched",
              hasA is False and tgtA == A.received_xlsx_path(woA_new, "161-9674-900")
              and not (woA_new / A.RECEIVED_FOLDER).exists() and archA.exists())
        rowA = A._collect_report_rows({}, jobs=[{"eng": "4095", "pn": "161-9674-900", "ts": "2026-09-13"}])[0]
        # review round 4: the answer stays False (not opted in), but every surface says
        # 'archived' rather than 'no report' — the copy is NAMED as `found`
        check("r1/r4: its picker row reads 'report archived — rebuild re-homes it' (live WO exists), tickable, has False, NOT pre-ticked",
              rowA["status"] == "report archived — rebuild re-homes it" and rowA["tickable"] is True
              and rowA["has"] is False and rowA["found"] == archA
              and A._picker_default_ticks([rowA]) == [False])
        stA1x = A._annotate_report_exists(dict(stA1))
        check("r4: _report_lookup names the Archive-only copy as `found` with a False answer; the cache carries it; save gate still False",
              A._report_lookup(stA1)[0] is False and A._report_lookup(stA1)[2] == archA
              and stA1x["_report_exists"] is False and stA1x["_report_found_at"] == archA
              and A._receive_save_gate(stA1x) is False)
        boxA = {}; A._rebuild_report_rows([rowA], boxA)
        outA2 = tgtA
        check("r1/r4: manual Build (picker row) re-creates it in the live WO (4 boards still at customer), archive untouched — NOT counted as 'created (first time)'",
              boxA["ok"] == 1 and boxA["created"] == 0 and boxA["moved"] == 0 and boxA["failed"] == []
              and outA2.exists() and archA.exists()
              and sum(1 for r in _rows_of(outA2, A.OUT_SHEET) if isinstance(r[0], int)) == 4)
        shutil.rmtree(jobA_new)                       # the whole live tree gone (finished job)
        rowA2 = A._collect_report_rows({}, jobs=[{"eng": "4095", "pn": "161-9674-900", "ts": "2026-09-13"}])[0]
        check("r1: no live -900/-970 tree at all -> picker row not tickable (a rebuild could only fail), not pre-ticked",
              rowA2["status"] == "report found — no live -900/-970 work-order folder" and rowA2["tickable"] is False
              and rowA2["has"] is False and rowA2["found"] == archA)
        # a copy parked at the Production Data level of a tree with NO live WO folder (review round 2)
        (jobA_new / "Production Data" / A.RECEIVED_FOLDER).mkdir(parents=True)
        shutil.copy(str(archA), str(jobA_new / "Production Data" / A.RECEIVED_FOLDER / archA.name))
        rowA3 = A._collect_report_rows({}, jobs=[{"eng": "4095", "pn": "161-9674-900", "ts": "2026-09-13"}])[0]
        check("r2: a copy found but no live -900/-970 WO anywhere -> 'report found — no live -900/-970 work-order folder', has True, not tickable",
              rowA3["status"] == "report found — no live -900/-970 work-order folder"
              and rowA3["tickable"] is False and rowA3["has"] is True)
        # review round 6 (test-r6-4): a job with NO copy anywhere and no live WO in either tree must read
        # 'no -900/-970 work-order folder found' — not 'report found — …' (there is nothing to find)
        histN = tempfile.mkdtemp(prefix="pcc_r6nolive_")
        jobN = custA / "161-9676-900 4097 Rev. A"; woN = jobN / "Production Data" / "16998.1"; woN.mkdir(parents=True)
        A.HISTORY_DIR = histN
        try:
            ev(histN, "2026-08", "x_4097", close_ev("4097", "161-9676-900", "16998.1", ["000001-1619676900-A00"],
               "2026-08-12 09:00:00", str(woN)))
            shutil.move(str(jobN), str(custA / "Archive" / jobN.name))                  # finished job: whole tree archived
            rowN = A._collect_report_rows({}, jobs=[{"eng": "4097", "pn": "161-9676-900", "ts": "2026-08-12"}])[0]
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histN, ignore_errors=True)
            shutil.rmtree(custA / "Archive" / jobN.name, ignore_errors=True)
        check("r6: no copy anywhere + both trees archived -> picker 'no -900/-970 work-order folder found', has False, found None, not tickable",
              rowN["status"] == "no -900/-970 work-order folder found" and rowN["has"] is False
              and rowN["found"] is None and rowN["tickable"] is False)

        # ── -900-form paperwork under a coated-only job's -900 tree (the in-progress
        #    -900-serialized batch ships as -970: its COC's P/N cell is the coated label) ──
        job900d = cust9 / "125-1030-900 4006 Rev. D"; wo900d = job900d / "Production Data" / "16401.1"
        _shipfile(wo900d / A.COC_FOLDER, "4006-COC-900batch.xlsx",
                  [f"{i:06d}-1251030900-D00" for i in range(921, 924)], "2026-09-15", pn_label="125-1030-970D")
        check("r1: _file_pn_label_side reads the file's P/N cell only (never a serial, never the cert sentence)",
              A._file_pn_label_side([["P/N:", "125-1030-970D"]], "1251030") == "970"
              and A._file_pn_label_side([["136-9624-900"]], "1369624") == "900"
              and A._file_pn_label_side([["000945-1251032970/C-C"], [t9(569)]], "1251030") == ""
              and A._file_pn_label_side([["This certifies that the parts conform to the 125-1030-970C Rev.- A00 drawings."]], "1251030") == ""
              and A._file_pn_label_side([["125-1030-970D"]], "1369624") == "")
        st9e = A.build_receiving_state(E9, TP9)
        scan9e = A.scan_job_folders(st9e)
        rows_e = {r["serial"]: r for r in scan9e["coated_rows"]}
        check("r1: that COC under the -900 tree -> coated rows carrying the Case-B Note, NO testing count",
              scan9e["ok"] and scan9e["ship_counts"] == {} and scan9e["ship_info"] == {}
              and all(rows_e.get(f"{i:06d}-1251030900-D00", {}).get("note") == A._CASE_B_NOTE for i in range(921, 924)))
        A.fold_folder_scan(st9e, scan9e)
        check("r1: fold keeps has_testing False and the coated name",
              A.receiving_summary(st9e)["has_testing"] is False and A._report_pn(st9e) == "125-1030-970D")
        # the folder-only Case-B rows carry their Note INTO coated_log and onto the sheet
        # (History has no row for 921-923: only fold_folder_scan can supply the tag — review round 3)
        b921 = {f"{i:06d}-1251030900-D00" for i in range(921, 924)}
        check("r3: folder-only Case-B rows keep the Note in coated_log after the fold",
              sum(1 for e in st9e["coated_log"] if e["serial"] in b921) == 3
              and all(e.get("note") == A._CASE_B_NOTE for e in st9e["coated_log"] if e["serial"] in b921))
        outE0 = A.write_received_xlsx(st9e)
        r921 = next((r for r in _rows_of(outE0, A.COAT_SHEET) if r[1] == "000921-1251030900-D00"), None)
        check("r3: …and that Note lands in the Coated & shipped sheet's Note cell (col 5 = 'not shipped for testing')",
              Path(outE0) == tgtD and r921 is not None and r921[5] == A._CASE_B_NOTE
              and r921[4] == "not shipped for testing" and r921[2] == "16401.1")
        # review round 4: the app's OWN Shipping-Serials file (Date/Time/Serial_Number/Operator/WO —
        # NO P/N cell, written by the close beside that COC) must not ALSO count the same boards as
        # testing shipments (one board on 'Coated & shipped' AND 'Still at customer', Q2 stop defeated)
        (wo900d / A.SERIALS_FOLDER).mkdir(parents=True, exist_ok=True)
        ssE = A.write_serial_xlsx(wo900d / A.SERIALS_FOLDER, E9, "16401.1",
                                  [f"{i:06d}-1251030900-D00" for i in range(921, 924)], "Anita Patel")
        st9ss = A.build_receiving_state(E9, TP9)
        scan9ss = A.scan_job_folders(st9ss)
        A.fold_folder_scan(st9ss, scan9ss)
        check("r4: the real Shipping-Serials shape of the SAME shipment under the -900 tree -> Case B like its COC: ship_counts {} / ship_info {}, shipped {} and has_testing False after the fold, 3 Case-B rows once",
              ssE and Path(ssE).exists() and Path(ssE).suffix == ".xlsx"
              and scan9ss["ok"] and scan9ss["files"] >= 2 and scan9ss["ship_counts"] == {} and scan9ss["ship_info"] == {}
              and st9ss["shipped"] == {} and A.receiving_summary(st9ss)["has_testing"] is False
              and A.ReceivingWindow._coated_only(st9ss) is True
              and sum(1 for e in st9ss["coated_log"] if e["serial"] in b921) == 3)
        _shipfile(wo900d / A.SERIALS_FOLDER, "hand-list.xlsx",
                  [f"{i:06d}-1251030900-D00" for i in range(931, 934)], "2026-09-16", pn_label="")
        st9h = A.build_receiving_state(E9, TP9)
        A.fold_folder_scan(st9h, A.scan_job_folders(st9h))
        check("r1: a label-less -900 list there IS a testing shipment (folder truth, on the sheets) — but the report NAME is a history fact: unchanged",
              A.receiving_summary(st9h)["has_testing"] is True and len(st9h["shipped"]) == 3
              and st9h.get("hist_has_testing") is False and A._report_pn(st9h) == "125-1030-970D"
              and A._report_exists(st9h) == (True, tgtD))
        outE = A.write_received_xlsx(st9h)                    # 🔁 Rebuild with folder truth
        A.auto_update_shipping_report(E9, "125-1030-970D")    # the next after-close update
        check("r1: Rebuild and the next after-close update land on the SAME file — no rename ping-pong",
              Path(outE) == tgtD and _r1_files() == [tgtD])
        # review round 6 (test-r6-2): a -900-LABELLED file on a coated-only job (a testing shipment made before
        # History / outside the app) IS a testing shipment — the folder-side Case B needs the file's OWN P/N cell to
        # be the coated label (or no cell, under the -970 tree); a '900' cell never re-routes
        _shipfile(wo900d / A.COC_FOLDER, "4006-old-900-COC.xlsx", [t9(1), t9(2), t9(3)], "2026-05-01", pn_label="125-1030-900")
        st9o = A.build_receiving_state(E9, TP9)
        sc9o = A.scan_job_folders(st9o)
        A.fold_folder_scan(st9o, sc9o)
        m9o, m9h = A.receiving_summary(st9o), A.receiving_summary(st9h)
        k9o = lambda i: A._serial_identity(t9(i), TP9)
        check("r6: a -900-labelled COC under the -900 tree of a coated-only job stays a TESTING shipment: 3 ship_counts, no Case-B rows for them, shipped / outstanding +3 over the label-less list alone, has_testing True, name still the coated label",
              sc9o["ok"] and all(sc9o["ship_counts"].get(k9o(i)) == 1 for i in (1, 2, 3))
              and not any(r["serial"] in {t9(1), t9(2), t9(3)} for r in sc9o["coated_rows"])
              and all(k9o(i) in st9o["shipped"] for i in (1, 2, 3))
              and (m9o["shipped"], m9o["outstanding"]) == (m9h["shipped"] + 3, m9h["outstanding"] + 3)
              and m9o["has_testing"] is True and A._report_pn(st9o) == "125-1030-970D")
        (wo900d / A.COC_FOLDER / "4006-old-900-COC.xlsx").unlink()      # (the later per-slot pins count files)

        # ── an old-flow job's folders route exactly as before (the embedded P/N decides) ──
        histO2 = tempfile.mkdtemp(prefix="pcc_oldflow2_")
        custO2 = Path(tempfile.mkdtemp(prefix="pcc_oldflow2c_")) / "20 TestCust"; extra_roots9.append(custO2.parent)
        woO900 = custO2 / "136-9700-900 4700 Rev. A" / "Production Data" / "17000.1"
        woO970 = custO2 / "136-9700-970 4700 Rev. A" / "Production Data" / "17001.1"
        woO900.mkdir(parents=True)
        tO = lambda i: f"{i:06d}-1369700900-A00"; cO = lambda i: f"{i:06d}-1369700970-A00"
        _shipfile(woO970 / A.SERIALS_FOLDER, "returned-list.xlsx", [tO(2), tO(3)], "2026-09-06", pn_label="136-9700-900")
        _shipfile(woO970 / A.COC_FOLDER, "coc-970.xlsx", [cO(1), tO(4)], "2026-09-07", pn_label="136-9700-970")
        _shipfile(woO970 / A.PACKING_LIST_FOLDER, "pl-nolabel.xlsx", [tO(5)], "2026-09-08", pn_label="")
        A.HISTORY_DIR = histO2
        try:
            ev(histO2, "2026-08", "a_test", close_ev("4700", "136-9700-900", "17000.1",
               [tO(i) for i in range(1, 6)], "2026-08-01 09:00:00", str(woO900)))
            stO2 = A.build_receiving_state("4700", "136-9700-900")
            scO2 = A.scan_job_folders(stO2)
            A.fold_folder_scan(stO2, scO2)
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histO2, ignore_errors=True)
        kO = lambda i: A._serial_identity(tO(i), "136-9700-900")
        check("r1: old-flow job — -900-form serials under the -970 tree stay TESTING shipments whatever the file's label; no Case-B rows",
              scO2["ok"] and stO2.get("hist_has_testing") is True
              and all(scO2["ship_counts"].get(kO(i)) == 1 for i in (2, 3, 4, 5))
              and [r["serial"] for r in scO2["coated_rows"]] == [cO(1)]
              and not any(r.get("note") for r in scO2["coated_rows"])
              and [e["serial"] for e in stO2["coated_log"]] == [cO(1)]
              and A.receiving_summary(stO2)["outstanding"] == 5 and A._report_pn(stO2) == "136-9700-900")

        # ══ review round 2 ═══════════════════════════════════════════════════
        # ── a prefix-related Eng# sharing the P/N (the live '…4016LF Rev. A02' folder) is
        #    NOT this job: its folders and its report must never be matched, merged or deleted ──
        check("r2: _job_dir_is — Eng# as a whole token in the live folder shapes; a longer Eng# never matches a shorter one",
              A._job_dir_is("GizmoX [4041 Rev. A]", "4041") and A._job_dir_is("125-1030-970C 4006 Rev. C", "4006")
              and A._job_dir_is("122-9538-970 4016LF Rev. A02", "4016LF") and A._job_dir_is("122-9538-970 4016 Rev. A02", "4016")
              and not A._job_dir_is("122-9538-970 4016LF Rev. A02", "4016")
              and not A._job_dir_is("122-9538-970 4016 Rev. A02", "4016LF")
              and not A._job_dir_is("125-1030-970C 4006 Rev. C", "") and not A._job_dir_is("125-1030-970C 4006 Rev. C", "N/A"))
        histE = tempfile.mkdtemp(prefix="pcc_engsub_")
        custE = Path(tempfile.mkdtemp(prefix="pcc_engsubc_")) / "377 Anduril"; extra_roots9.append(custE.parent)
        woE = custE / "122-9538-970 4016 Rev. A02" / "Production Data" / "16800.1"
        woE_old = custE / "122-9538-970 4016 Rev. A02" / "Production Data" / "16799.1"
        woLF = custE / "122-9538-970 4016LF Rev. A02" / "Production Data" / "16850.1"
        (woE / A.COC_FOLDER).mkdir(parents=True); woE_old.mkdir(parents=True); (woLF / A.COC_FOLDER).mkdir(parents=True)
        eS = lambda i: f"{i:06d}-1229538970-A02"
        A.HISTORY_DIR = histE
        try:
            ev(histE, "2026-08", "a_4016", close_ev("4016", "122-9538-970", "16800.1", [eS(i) for i in range(1, 4)],
               "2026-08-05 09:00:00", str(woE)))
            ev(histE, "2026-08", "b_4016LF", close_ev("4016LF", "122-9538-970", "16850.1", [eS(i) for i in range(35, 40)],
               "2026-08-06 09:00:00", str(woLF)))
            outE16 = A.write_received_xlsx(A.build_receiving_state("4016", "122-9538-900"))       # both built by hand
            outLF = A.write_received_xlsx(A.build_receiving_state("4016LF", "122-9538-900"))
            sigLF = A._file_sig(outLF)
            candsE = A._stale_received_candidates(A.build_receiving_state("4016", "122-9538-900"), Path(outE16))
            check("r2: 4016's candidate scan never enters 4016LF's tree; scan_job_folders lists only 4016's own job folder",
                  Path(outLF).name == Path(outE16).name and not any("4016LF" in str(c) for c in candsE)
                  and A.scan_job_folders(A.build_receiving_state("4016", "122-9538-900"))["jobs"] == [str(woE.parent.parent)])
            ev(histE, "2026-09", "c_4016_more", close_ev("4016", "122-9538-970", "16800.1", [eS(4)],
               "2026-09-01 09:00:00", str(woE)))
            A.auto_update_shipping_report("4016", "122-9538-970")
            rowsE = _rows_of(outE16, A.COAT_SHEET)
            check("r2: after 4016's after-close update, 4016LF's report still exists UNCHANGED and none of its serials are on 4016's sheet",
                  Path(outLF).exists() and A._file_sig(outLF) == sigLF
                  and any(r[1] == eS(4) for r in rowsE) and not any(r[1] == eS(35) for r in rowsE)
                  and A._report_exists(A.build_receiving_state("4016LF", "122-9538-900"))[0] is True)
            # second, independent guard on the delete path: a same-named copy in OUR tree whose
            # own head line names the OTHER job is kept, never merged, never our base
            plantE = woE_old / A.RECEIVED_FOLDER / Path(outLF).name
            plantE.parent.mkdir(parents=True); shutil.copy(str(outLF), str(plantE)); psig = A._file_sig(plantE)
            check("r2: _report_copy_eng / _extract_received_rows / _copy_names_other_job read the copy's 'Engineering #' line",
                  A._report_copy_eng(load_workbook(plantE)) == "4016LF"
                  and A._extract_received_rows(plantE, "4016") is None
                  and A._extract_received_rows(plantE, "4016LF") is not None
                  and A._extract_received_rows(plantE) is not None
                  and A._copy_names_other_job(plantE, "4016") is True and A._copy_names_other_job(plantE, "4016LF") is False)
            stE2 = A.build_receiving_state("4016", "122-9538-900")
            A.write_received_xlsx(stE2)
            rowsE = _rows_of(outE16, A.COAT_SHEET)
            check("r2: a save with that copy in our own tree leaves it on disk untouched and merges nothing from it",
                  plantE.exists() and A._file_sig(plantE) == psig and not any(r[1] == eS(35) for r in rowsE)
                  and stE2.get("xlsx_conflicts") == [])
            Path(outE16).unlink()
            check("r2: with our own report gone, the other job's copy is not our merge BASE either",
                  A._find_legacy_received_file(A.build_receiving_state("4016", "122-9538-900"), Path(outE16)) is None)
            # …and it does not opt OUR job in either: the R6 delete sticks, the daemon writes nothing
            ev(histE, "2026-09", "d_4016_after", close_ev("4016", "122-9538-970", "16800.1", [eS(5)],
               "2026-09-02 09:00:00", str(woE)))
            hasE3 = A._report_exists(A.build_receiving_state("4016", "122-9538-900"))
            A.auto_update_shipping_report("4016", "122-9538-970")
            check("r3: another job's same-named copy in our tree is NOT this job's report — _report_exists False, the after-close update re-creates nothing, the copy stays",
                  hasE3[0] is False and not Path(outE16).exists() and plantE.exists() and A._file_sig(plantE) == psig
                  and A._find_legacy_received_file(A.build_receiving_state("4016", "122-9538-900"), Path(outE16)) is None)
            # review round 5 (datasafety-r5-5): the other job's report dropped AT OUR TARGET PATH is not ours either —
            # the switch does not turn True on the file name alone, the write refuses to adopt it (its rows would have
            # been merged among ours and its 'Engineering #' line rewritten to ours), the daemon writes nothing
            shutil.copy(str(plantE), str(outE16)); tsigE = A._file_sig(outE16)
            stE5 = A._annotate_report_exists(A.build_receiving_state("4016", "122-9538-900"))
            try:
                A.write_received_xlsx(dict(stE5)); msgE5 = ""
            except RuntimeError as e:
                msgE5 = str(e)
            A.auto_update_shipping_report("4016", "122-9538-970")
            check("r5: 4016LF's report planted at 4016's own target -> _report_exists False (never True), save gate False, the write REFUSES naming the other job, the daemon writes nothing, the file byte-identical, _wo_holds_report False for that WO",
                  stE5["_report_exists"] is False and A._receive_save_gate(stE5) is False
                  and "names another job (Engineering # 4016LF)" in msgE5 and A._file_sig(outE16) == tsigE
                  and A._wo_holds_report(woE, stE5) is False and A._report_copy_eng(load_workbook(outE16)) == "4016LF"
                  and A._report_copy_eng_at(outE16) == "4016LF" and A._report_copy_eng_at(woE / "nope.xlsx") == "")
            Path(outE16).unlink()
            # a stale cached answer after a FAILED pre-save re-walk (optin-1)
            stE4 = A._annotate_report_exists(A.build_receiving_state("4016LF", "122-9538-900"))
            Path(outLF).unlink()                                        # opt-out (R6) after the walk
            g_stale = A._receive_save_gate(stE4)
            A._forget_report_cache(stE4)
            check("r2: a stale cached True would re-create a deleted report; _forget_report_cache makes the gate re-check disk -> False",
                  g_stale is True and A._receive_save_gate(stE4) is False and stE4["_report_exists"] is False)
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histE, ignore_errors=True)

        # ── one job, two spellings of its -900 label: ONE entry (newest spelling), ONE picker
        #    row, ONE report — the walk keys either spelling and a file named by the other
        #    spelling is this job's report too (review rounds 2 + 3: two rows used to make two
        #    files, of which only one kept updating) ──
        histS = tempfile.mkdtemp(prefix="pcc_twin_")
        custS = Path(tempfile.mkdtemp(prefix="pcc_twinc_")) / "377 Anduril"; extra_roots9.append(custS.parent)
        woS1 = custS / "134-2201-900 4065 Rev. A" / "Production Data" / "16300.1"
        woS2 = custS / "134-2201-900 4065 Rev. A" / "Production Data" / "16400.2"
        woS1.mkdir(parents=True); woS2.mkdir(parents=True)
        tS = lambda i: f"{i:06d}-1342201900-A00"
        A.HISTORY_DIR = histS
        try:
            ev(histS, "2026-08", "a_dashed", close_ev("4065", "134-2201-900", "16300.1", [tS(i) for i in range(1, 7)],
               "2026-08-10 09:00:00", str(woS1)))
            ev(histS, "2026-09", "b_flat", dict(close_ev("4065", "1342201900", "16400.2", [tS(i) for i in range(7, 10)],
               "2026-09-10 09:00:00", str(woS2)), event_id="ev-flat-1"))
            entS = [(j["eng"], j["pn"]) for j in A.list_testing_jobs() if j["eng"] == "4065"]
            stS = A.build_receiving_state("4065", "134-2201-900")
            stS2 = A.build_receiving_state("4065", "1342201900")
            check("r3: spelling-twin job — ONE dropdown entry (newest spelling); a walk under EITHER spelling sees all 9 boards, anchors the newest -900 WO and remembers the other spelling",
                  entS == [("4065", "1342201900")]
                  and len(stS["shipped"]) == 9 and stS["wo_path"] == str(woS2) and stS.get("testing_label_seen") is True
                  and stS.get("testing_pn_labels") == ["1342201900"] and A._report_pn_spellings(stS)[:2] == ["134-2201-900", "1342201900"]
                  and len(stS2["shipped"]) == 9 and stS2["wo_path"] == str(woS2) and stS2.get("testing_pn_labels") == ["134-2201-900"])
            rowsS = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == "4065"])
            check("r3: the Rebuild-ALL picker shows ONE row for it ('no report', tickable)",
                  len(rowsS) == 1 and rowsS[0]["status"] == "no report" and rowsS[0]["tickable"] is True)
            outS = A.write_received_xlsx(stS)                    # built by hand under the DASHED spelling
            nS = sum(1 for r in _rows_of(outS, A.OUT_SHEET) if isinstance(r[0], int))
            ev(histS, "2026-09", "c_flat2", dict(close_ev("4065", "1342201900", "16400.2", [tS(i) for i in range(10, 13)],
               "2026-09-12 09:00:00", str(woS2)), event_id="ev-flat-2"))
            A.auto_update_shipping_report("4065", "1342201900")
            fS = woS2 / A.RECEIVED_FOLDER / "1342201900 Shipping Report.xlsx"
            check("r2/r3: the after-close update lands on the ONE existing report — folded into the entry's spelling (+3 still-at-customer rows), no second file",
                  Path(outS) == woS2 / A.RECEIVED_FOLDER / "134-2201-900 Shipping Report.xlsx" and nS == 9
                  and fS.exists() and not Path(outS).exists()
                  and sum(1 for r in _rows_of(fS, A.OUT_SHEET) if isinstance(r[0], int)) == 12
                  and sorted(p.name for p in custS.rglob("*Shipping Report*.xlsx")) == [fS.name])
            # a second file under the OTHER spelling, left behind by the old two-row picker, folds in too
            twinS = woS1 / A.RECEIVED_FOLDER / "134-2201-900 Shipping Report.xlsx"
            twinS.parent.mkdir(parents=True); shutil.copy(str(fS), str(twinS))
            ev(histS, "2026-09", "d_flat3", dict(close_ev("4065", "1342201900", "16400.2", [tS(i) for i in range(13, 15)],
               "2026-09-13 09:00:00", str(woS2)), event_id="ev-flat-3"))
            A.auto_update_shipping_report("4065", "1342201900")
            check("r3: a leftover twin file under the other spelling is folded into the one report (+2 rows) and removed",
                  fS.exists() and not twinS.exists()
                  and sum(1 for r in _rows_of(fS, A.OUT_SHEET) if isinstance(r[0], int)) == 14
                  and sorted(p.name for p in custS.rglob("*Shipping Report*.xlsx")) == [fS.name])
            # review round 5 (datasafety-r5-2): the supervisor voids every dashless close (the typo) in 🕘 History
            # and re-closes the boards dashed — the flat spelling now survives ONLY on voided events, yet the file
            # the daemon renamed to that spelling is still THE report: it must fold into ONE dashed file, never be
            # orphaned beside a fresh one
            for i_, eid in enumerate(("ev-flat-1", "ev-flat-2", "ev-flat-3"), 1):
                ev(histS, "2026-09", f"e_void_{i_}", {"event_type": "void", "eng": "4065", "pn": "1342201900", "wo": "16400.2",
                   "ts": "2026-09-14 09:00:00", "target_event_id": eid})
            ev(histS, "2026-09", "f_dashed_again", close_ev("4065", "134-2201-900", "16400.2", [tS(i) for i in range(7, 15)],
               "2026-09-14 10:00:00", str(woS2)))
            entS3 = [(j["eng"], j["pn"]) for j in A.list_testing_jobs() if j["eng"] == "4065"]
            stS3 = A._annotate_report_exists(A.build_receiving_state("4065", "134-2201-900"))
            A.auto_update_shipping_report("4065", "134-2201-900")
            fS3 = woS2 / A.RECEIVED_FOLDER / "134-2201-900 Shipping Report.xlsx"
            check("r5: a -900 spelling that survives only in VOIDED closes still names this job's file — testing_pn_labels keeps it, _report_exists True (found = the flat-named file), the update folds it into ONE dashed file (14 rows), the flat name is gone",
                  entS3 == [("4065", "134-2201-900")] and stS3.get("testing_pn_labels") == ["1342201900"]
                  and "1342201900" in A._report_pn_spellings(stS3)
                  and stS3["_report_exists"] is True and stS3["_report_found_at"] == fS
                  and fS3.exists() and not fS.exists()
                  and sum(1 for r in _rows_of(fS3, A.OUT_SHEET) if isinstance(r[0], int)) == 14
                  and sorted(p.name for p in custS.rglob("*Shipping Report*.xlsx")) == [fS3.name])
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histS, ignore_errors=True)

        # ── review round 5 (coated-r5-2): a -900 label hand-typed WITH a rev letter ('134-2201-900A') must not
        #    become the job key while a plain spelling exists — under a lettered key the walk's coated key was ''
        #    and the 4065 import shape's -970-form serials counted as testing shipments again (report renamed,
        #    Still-at-customer +10); and even the lettered-key walk itself now derives its coated key from the digits ──
        histK = tempfile.mkdtemp(prefix="pcc_letter900_")
        custK = Path(tempfile.mkdtemp(prefix="pcc_letter900c_")) / "377 Anduril"; extra_roots9.append(custK.parent)
        woK = custK / "134-2201-900 4065 Rev. A00" / "Production Data" / "16587.1"; (woK / A.COC_FOLDER).mkdir(parents=True)
        tK = lambda i: f"{i:06d}-1342201900-A00"; cK = lambda i: f"{i:06d}-1342201970-A00"
        A.HISTORY_DIR = histK
        try:
            ev(histK, "2026-08", "a_close_test", close_ev("4065", "134-2201-900", "16587.1", [tK(i) for i in range(1, 31)],
               "2026-08-10 09:00:00", str(woK)))
            ev(histK, "2026-08", "b_bf_coat_mislabel", {"event_type": "backfill_shipped", "eng": "4065", "pn": "134-2201-900",
               "wo": "16588.1", "ts": "2026-08-12 09:00:00", "ship_date": "7/1/26", "serials": [cK(i) for i in range(1, 11)],
               "production_data": str(woK.parent)})
            ev(histK, "2026-08", "c_recv", recv_ev("4065", "134-2201-900", [tK(i) for i in range(1, 6)], "2026-08-20 10:00:00",
               pd=str(woK.parent)))
            outK = Path(A.write_received_xlsx(A.build_receiving_state("4065", "134-2201-900")))     # opted in by hand
            ev(histK, "2026-09", "d_close_letter", close_ev("4065", "134-2201-900A", "16587.1", [tK(i) for i in range(31, 36)],
               "2026-09-15 09:00:00", str(woK)))
            entK = [j["pn"] for j in A.list_testing_jobs() if j["eng"] == "4065"]
            A.auto_update_shipping_report("4065", "134-2201-900A")                    # the close flow's daemon call
            stK = A.build_receiving_state("4065", "134-2201-900"); mK = A.receiving_summary(stK)
            stKl = A.build_receiving_state("4065", "134-2201-900A"); mKl = A.receiving_summary(stKl)
            check("r5: a '-900A' close -> the entry keeps the PLAIN spelling, the walk remembers the lettered one; shipped 35 / outstanding 30 / coated_all 10, name unchanged, ONE file (+5 Still-at-customer rows); the lettered-key walk routes the 10 coated boards to coated_log too",
                  entK == ["134-2201-900"] and stK.get("testing_pn_labels") == ["134-2201-900A"]
                  and (mK["shipped"], mK["outstanding"], mK["coated_all"]) == (35, 30, 10) and A._report_pn(stK) == "134-2201-900"
                  and sorted(p.name for p in custK.rglob("*Shipping Report*.xlsx")) == [outK.name]
                  and sum(1 for r in _rows_of(outK, A.OUT_SHEET) if isinstance(r[0], int)) == 30
                  and (mKl["shipped"], mKl["coated_all"]) == (35, 10) and len(stKl["coated_log"]) == 10
                  and A._coat_ident_pn("134-2201-900A") == "1342201970")
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histK, ignore_errors=True)

        # ── the Save-path refusal says what is true: the home tree has no live WO and a
        #    🔁 Rebuild report on THIS PC may re-home it — not "go to another PC" ──
        try:
            A.write_received_xlsx(dict(stZ)); msgZ = ""
        except RuntimeError as e:
            msgZ = str(e)
        stU2 = dict(stZ)
        stU2["wo_path"] = str(Path(tempfile.gettempdir()) / "pcc_nope_r2_zz" / "377 Anduril"
                              / "125-1030-970C 4006 Rev. C" / "Production Data" / "16493.2")
        stU2["pd_path"] = str(Path(stU2["wo_path"]).parent)
        try:
            A.write_received_xlsx(stU2); msgU = ""
        except RuntimeError as e:
            msgU = str(e)
        check("r2: refusal texts — Save with a live sibling WO: 'could not be found' + 'Rebuild report on this PC' / sibling; folders unreachable: today's 'PC that can open' text",
              "could not be found" in msgZ and "-900 or -970" in msgZ and "Rebuild report on this PC" in msgZ
              and "sibling" in msgZ and "PC that can open" not in msgZ
              and "could not be found" in msgU and "PC that can open" in msgU and "sibling" not in msgU)

        # ── R6 opt-out vs an older copy the merge cannot remove (user sheet in it): today's
        #    R1 rule keeps the job opted in — the chips / Save popup now NAME that copy ──
        woL2 = woL.parent / "16701.1"; woL2.mkdir()
        ev(hist9, "2026-09", "r_close_L2", close_ev("4310", "150-2200-900", "16701.1", [tL(i) for i in range(4, 6)],
           "2026-09-15 09:00:00", str(woL2)))
        A.auto_update_shipping_report("4310", "150-2200-900")
        tgtL2 = A.received_xlsx_path(woL2, "150-2200-900")
        check("r2: the report followed the newest WO (16700.1 -> 16701.1); the clean older copy was removed",
              tgtL2.exists() and not tgtL.exists())
        strayL = woL / A.RECEIVED_FOLDER / tgtL2.name
        strayL.parent.mkdir(parents=True, exist_ok=True); shutil.copy(str(tgtL2), str(strayL))
        wbS = load_workbook(strayL); wbS.create_sheet("My notes")["A1"] = "keep me"; wbS.save(strayL); wbS.close()
        stL3 = A.build_receiving_state("4310", "150-2200-900")
        A.write_received_xlsx(stL3)
        check("r2: an older copy with a user-added sheet is KEPT by the save (not provably deletable)", strayL.exists())
        tgtL2.unlink()                                                   # the operator opts the job out (R6)
        hasL3, pathL3, foundL3 = A._report_lookup(stL3)
        stL3 = A._annotate_report_exists(A.build_receiving_state("4310", "150-2200-900"))
        check("r2: after the opt-out delete the kept copy still answers True — _report_lookup and the cache NAME it (chips: 'delete it too')",
              hasL3 is True and pathL3 == tgtL2 and foundL3 == strayL
              and stL3["_report_exists"] is True and stL3["_report_at_home"] is False and stL3["_report_found_at"] == strayL)
        A.auto_update_shipping_report("4310", "150-2200-900")
        check("r2: (today's R1 rule, flagged for Dhruvil) the next after-close update re-homes that copy WHOLE as the report (user sheet kept) — the opt-out did not stick",
              tgtL2.exists() and not strayL.exists() and "My notes" in load_workbook(tgtL2).sheetnames)

        # ── review round 4 (datasafety-r4-2): a PARTIAL candidate enumeration (a share blip on the
        #    Production Data listing) never becomes a definite 'no report' — the switch answers None,
        #    so no first-time Build is offered that would have FORKED the report ──
        woL3 = woL.parent / "16702.1"; woL3.mkdir()
        ev(hist9, "2026-09", "r_close_L3", close_ev("4310", "150-2200-900", "16702.1", [tL(6)],
           "2026-09-16 09:00:00", str(woL3)))
        stL4 = A.build_receiving_state("4310", "150-2200-900")      # target = 16702.1 (missing); the copy sits in 16701.1
        okL4 = A._report_lookup(stL4)
        def _blip_iterdir(self):
            if self.name == "Production Data" and "4310" in self.parent.name:
                raise OSError(64, "The specified network name is no longer available")
            return _orig_iterdir(self)
        A.Path.iterdir = _blip_iterdir
        try:
            blipL4 = A._report_lookup(stL4)
            gateL4 = A._receive_save_gate(A.build_receiving_state("4310", "150-2200-900"))
            rowL4 = A._collect_report_rows({}, jobs=[{"eng": "4310", "pn": "150-2200-900", "ts": "2026-09-16"}])[0]
            # review round 5 (optin-r5-1 / datasafety-r5-1): the human Rebuild's write under the SAME blip —
            # the home WO resolves (fast path), the target is missing, the enumeration that would have found
            # the 16701.1 copy is partial: it must REFUSE, not build a fresh file next to the real one
            stL4w = dict(A.build_receiving_state("4310", "150-2200-900"), allow_tree_switch=True)
            try:
                A.write_received_xlsx(stL4w); msgL4 = ""
            except RuntimeError as e:
                msgL4 = str(e)
        finally:
            A.Path.iterdir = _orig_iterdir
        check("r4: normal lookup True (copy in the previous WO); under a listing blip -> None (never False), save gate None, picker \"can't check from this PC\" (not tickable)",
              okL4[0] is True and okL4[2] == tgtL2
              and blipL4[0] is None and gateL4 is None
              and rowL4["status"] == "can't check from this PC" and rowL4["tickable"] is False)
        check("r5: under that blip the Rebuild write (allow_tree_switch) REFUSES ('could not confirm'): no leaf in 16702.1, the 16701.1 copy untouched — no fork; the Rebuild button refuses a None answer up front (pin)",
              "could not confirm" in msgL4 and not (woL3 / A.RECEIVED_FOLDER).exists() and tgtL2.exists()
              and sorted(p.name for p in custL.rglob("*Shipping Report*.xlsx")) == [tgtL2.name]
              and 'if st.get("_report_exists") is None:\n            messagebox.showwarning("Can\'t check",' in src9)
        A.auto_update_shipping_report("4310", "150-2200-900")
        tgtL3 = A.received_xlsx_path(woL3, "150-2200-900")
        check("r4: once the blip clears the after-close update migrates the ONE file to 16702.1 (no fork)",
              tgtL3.exists() and not tgtL2.exists()
              and sorted(p.name for p in custL.rglob("*Shipping Report*.xlsx")) == [tgtL3.name])

        # ── review round 5 (test-adequacy-r5-1): an UNLISTABLE Archive folder (or an archived job folder's
        #    Production Data) is not a partial LIVE enumeration — an archived copy never decides opt-in (R6),
        #    so the answer stays definite (True with a live copy elsewhere, False with none), the picker row
        #    stays tickable and the save gate says False, not "can't check" ──
        (custL / "Archive" / "150-2200-900 4310 Rev. 0" / "Production Data" / "16000.1").mkdir(parents=True)
        woL4 = woL.parent / "16703.1"; woL4.mkdir()
        ev(hist9, "2026-09", "r_close_L4", close_ev("4310", "150-2200-900", "16703.1", [tL(7)],
           "2026-09-17 09:00:00", str(woL4)))
        tgtL4 = A.received_xlsx_path(woL4, "150-2200-900")
        def _deny_arch(self):
            if self.name == "Archive" or (self.name == "Production Data" and "Rev. 0" in self.parent.name):
                raise PermissionError(13, "Access is denied")
            return _orig_iterdir(self)
        A.Path.iterdir = _deny_arch
        try:
            stL5 = dict(A.build_receiving_state("4310", "150-2200-900"))       # target = 16703.1 (missing); copy in 16702.1
            lookL5 = A._report_lookup(stL5)
            A._stale_received_candidates(stL5, tgtL4, include_archive=True)    # (the flags land on the state passed in)
            bakL5 = tgtL3.with_suffix(".bak"); tgtL3.rename(bakL5)              # …and with NO live copy anywhere
            try:
                lookL6 = A._report_lookup(dict(A.build_receiving_state("4310", "150-2200-900")))
                rowL6 = A._collect_report_rows({}, jobs=[{"eng": "4310", "pn": "150-2200-900", "ts": "2026-09-17"}])[0]
                gateL6 = A._receive_save_gate(A.build_receiving_state("4310", "150-2200-900"))
            finally:
                bakL5.rename(tgtL3)
        finally:
            A.Path.iterdir = _orig_iterdir
        check("r5: Archive not listable -> _arch_partial (never _cand_partial): the 16702.1 copy still answers True; with no live copy a definite False (not None), picker 'no report' tickable, save gate False",
              lookL5 == (True, tgtL4, tgtL3) and stL5.get("_arch_partial") is True and stL5.get("_cand_partial") is False
              and lookL6 == (False, tgtL4, None)
              and rowL6["status"] == "no report" and rowL6["tickable"] is True and gateL6 is False)
        A.auto_update_shipping_report("4310", "150-2200-900")
        check("r5: the after-close update then migrates the ONE file to 16703.1",
              tgtL4.exists() and not tgtL3.exists()
              and sorted(p.name for p in custL.rglob("*Shipping Report*.xlsx")) == [tgtL4.name])

        # ── the 970C spelling survives ONLY in a voided close (rev bump re-closed under 970D):
        #    the file built under 970C is still THE report and migrates to the 970D name ──
        histV = tempfile.mkdtemp(prefix="pcc_voidlbl_")
        custV = Path(tempfile.mkdtemp(prefix="pcc_voidlblc_")) / "377 Anduril"; extra_roots9.append(custV.parent)
        woV2 = custV / "125-1030-970C 4606 Rev. C" / "Production Data" / "16493.2"
        woV3 = custV / "125-1030-970C 4606 Rev. C" / "Production Data" / "16493.3"
        woV2.mkdir(parents=True); woV3.mkdir(parents=True)
        A.HISTORY_DIR = histV
        try:
            ev(histV, "2026-08", "a_970C", dict(close_ev("4606", CPL9, "16493.2", [c9(i) for i in range(1, 4)],
               "2026-08-20 09:00:00", str(woV2)), event_id="ev-970c"))
            outV = A.write_received_xlsx(A.build_receiving_state("4606", TP9))            # built under 970C
            ev(histV, "2026-09", "b_void", {"event_type": "void", "eng": "4606", "pn": CPL9, "wo": "16493.2",
               "ts": "2026-09-01 10:00:00", "target_event_id": "ev-970c"})
            ev(histV, "2026-09", "c_970D", close_ev("4606", "125-1030-970D", "16493.3",
               [f"{i:06d}-1251030970-D00" for i in range(1, 4)], "2026-09-02 09:00:00", str(woV3)))
            stV = A.build_receiving_state("4606", TP9)
            hasV, tgtV = A._report_exists(stV)
            A.auto_update_shipping_report("4606", "125-1030-970D")
            check("r2: a label seen only in a VOIDED close still feeds the spellings -> the 970C file counts, and the update migrates it to ONE 970D file",
                  Path(outV).name == f"{CPL9} Shipping Report.xlsx" and CPL9 in A._report_pn_spellings(stV)
                  and hasV is True and tgtV == woV3 / A.RECEIVED_FOLDER / "125-1030-970D Shipping Report.xlsx"
                  and tgtV.exists() and not Path(outV).exists()
                  and sorted(p.name for p in custV.rglob("*Shipping Report*.xlsx")) == [tgtV.name])
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histV, ignore_errors=True)

        # ── receive-only old-flow job (its -900 shipment predates History) + a -970 close:
        #    NO coated anchor — wo_path None as before ('No job folder … Import older shipments') ──
        histR = tempfile.mkdtemp(prefix="pcc_rcvonly_")
        A.HISTORY_DIR = histR
        try:
            ev(histR, "2026-09", "a_recv", recv_ev("4110", "150-1000-900", ["000001-1501000900-A00", "000002-1501000900-A00"],
               "2026-09-01 09:00:00", pd=str(cust9 / "150-1000-900 4110 Rev. A" / "Production Data")))
            ev(histR, "2026-09", "b_coat", close_ev("4110", "150-1000-970", "16700.1", ["000001-1501000970-A00"],
               "2026-09-05 09:00:00", str(cust9 / "150-1000-970 4110 Rev. A" / "Production Data" / "16700.1")))
            stRo = A.build_receiving_state("4110", "150-1000-900")
            check("r2 (the 2026-10-06 rule: the report homes in the -970 coated WO): receive-only job + a -970 close -> home_side coated, the -970 WO 16700.1 (the -900-named report files there on a Build)",
                  stRo.get("home_side") == "coated" and Path(stRo["wo_path"]).name == "16700.1" and len(stRo["received"]) == 2
                  and A.receiving_summary(stRo)["has_testing"] is True and A._report_pn(stRo) == "150-1000-900")
            # …while a supervisor-overridden receive of a COATED-form board keeps an in-house-tested job's coated home
            # (and, since decision (a), its in-house-tested status: has_testing stays False)
            ev(histR, "2026-09", "c_coat_only", close_ev("4111", CPL9, "16493.2", [c9(1), c9(2)], "2026-09-05 09:00:00", str(wo970b)))
            ev(histR, "2026-09", "d_override", recv_ev("4111", TP9, [c9(1)], "2026-09-06 09:00:00"))
            stOv = A.build_receiving_state("4111", TP9)
            check("r2: coated-only job + an overridden coated-form receive -> still home_side coated (received stays empty; has_testing False — decision (a))",
                  stOv["wo_path"] == str(wo970b) and stOv.get("home_side") == "coated" and stOv["received"] == {}
                  and A.receiving_summary(stOv)["has_testing"] is False and A._report_pn(stOv) == CPL9)
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histR, ignore_errors=True)

        # ── review round 4 (tk-ui-r4-1): the opt-in answer follows the writers' precondition — a
        #    state with NO wo_path (its only -900 record voided AFTER a receive was saved with
        #    production_data) is (False, None): chips / Save / 📄 Open file / picker / daemon agree ──
        histW = tempfile.mkdtemp(prefix="pcc_r4void_")
        custW = Path(tempfile.mkdtemp(prefix="pcc_r4voidc_")) / "20 TestCust"; extra_roots9.append(custW.parent)
        woW = custW / "150-1000-900 4110 Rev. A" / "Production Data" / "16600.1"; woW.mkdir(parents=True)
        A.HISTORY_DIR = histW
        try:
            ev(histW, "2026-08", "a_bf", {"event_type": "backfill_shipped", "eng": "4110", "pn": "150-1000-900",
               "wo": "16600.1", "ts": "2026-08-01 09:00:00", "ship_date": "7/1/26", "event_id": "ev-bf-4110",
               "serials": ["000001-1501000900-A00", "000002-1501000900-A00"], "production_data": str(woW.parent)})
            outW = Path(A.write_received_xlsx(A.build_receiving_state("4110", "150-1000-900")))   # 🔁 Build report
            ev(histW, "2026-09", "b_recv", recv_ev("4110", "150-1000-900", ["000001-1501000900-A00"],
               "2026-09-01 09:00:00", "Lata Gandhi", str(woW.parent)))
            ev(histW, "2026-09", "c_void", {"event_type": "void", "eng": "4110", "pn": "150-1000-900",
               "wo": "16600.1", "ts": "2026-09-02 10:00:00", "target_event_id": "ev-bf-4110"})
            stW = A._annotate_report_exists(A.build_receiving_state("4110", "150-1000-900"))
            rowW = A._collect_report_rows({}, jobs=[{"eng": "4110", "pn": "150-1000-900", "ts": "2026-09-01"}])[0]
            check("r4: wo_path None (voided anchor) + pd_path set + the file on disk -> _report_exists (False, None), cache False / no path, save gate False, picker 'no WO folder known' — no surface claims a report it cannot open or update",
                  stW["wo_path"] is None and stW.get("pd_path") == str(woW.parent) and outW.exists()
                  and A.receiving_summary(stW)["has_testing"] is True
                  and A._report_exists(stW) == (False, None)
                  and stW["_report_exists"] is False and stW["_report_path"] is None and stW["_report_at_home"] is None
                  and stW["_report_found_at"] is None and A._receive_save_gate(stW) is False
                  and rowW["status"] == "no WO folder known" and rowW["tickable"] is False)
            check("r5: the picker's 'Last shipment' is BLANK when nothing shipped is on record (never the receive / entry ts)",
                  rowW["last_ts"] == "" and rowW["job"]["ts"].startswith("2026-09-01") and stW["shipped"] == {})
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histW, ignore_errors=True)

        # ── review round 4 (test-adequacy-r4-2): an OLD-FLOW entry (a -900-labelled close) is never
        #    dropped by the decision-(a) receive rule, whatever its receives look like ──
        histG = tempfile.mkdtemp(prefix="pcc_r4guard_")
        A.HISTORY_DIR = histG
        try:
            tG = lambda i: f"{i:06d}-1504400900-A00"; cG = lambda i: f"{i:06d}-1504400970-A00"
            ev(histG, "2026-08", "a_close900", close_ev("7001", "150-4400-900", "16700.1", [tG(i) for i in range(1, 4)],
               "2026-08-01 09:00:00"))
            ev(histG, "2026-09", "b_recv_coatform", recv_ev("7001", "150-4400-900", [cG(1)], "2026-09-01 09:00:00"))
            eG = [j for j in A.list_testing_jobs() if j["eng"] == "7001"]
            stG = A.build_receiving_state("7001", "150-4400-900")
            check("r4: -900 close + a receive of ONLY coated-form serials -> the REAL entry stays (one entry, no 'testing' key), has_testing True, no coated-only stop",
                  len(eG) == 1 and "testing" not in eG[0] and eG[0]["pn"] == "150-4400-900"
                  and A.receiving_summary(stG)["has_testing"] is True and A.ReceivingWindow._coated_only(stG) is False)
            # the Case-B variant: the job's own -900-form board shipped under the coated label, then received
            ev(histG, "2026-08", "c_close970_caseb", close_ev("7001", "150-4400-970", "16701.1", [cG(2), tG(9)],
               "2026-08-15 09:00:00"))
            os.remove(Path(histG) / "2026" / "2026-09" / "b_recv_coatform.json")
            ev(histG, "2026-09", "d_recv_caseb", recv_ev("7001", "150-4400-900", [tG(9)], "2026-09-02 09:00:00"))
            eG2 = [j for j in A.list_testing_jobs() if j["eng"] == "7001"]
            stG2 = A.build_receiving_state("7001", "150-4400-900")
            check("r4: -900 close + a receive of ONLY its Case-B board -> still ONE real entry (no coated-only twin), has_testing True, -900 name",
                  len(eG2) == 1 and "testing" not in eG2[0] and eG2[0]["pn"] == "150-4400-900"
                  and A.receiving_summary(stG2)["has_testing"] is True and A._report_pn(stG2) == "150-4400-900"
                  and A._serial_identity(tG(9), "150-4400-970") in stG2.get("coat_label_mismatch", set()))
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histG, ignore_errors=True)

        # ── a 6-digit-core P/N: the junk-label -900-form serials still make the job a REAL
        #    (testing) entry — the core is no longer guessed from the serial ──
        histC6 = tempfile.mkdtemp(prefix="pcc_core6_")
        A.HISTORY_DIR = histC6
        try:
            ev(histC6, "2026-08", "a_clean", close_ev("4700", "192-146-970C", "16900.1", ["000020-192146970-A00"], "2026-08-20 09:00:00"))
            ev(histC6, "2026-08", "b_junk900", close_ev("4700", "192-146-970C 4700 Rev. A", "16900.2", ["000021-192146900-A00"], "2026-08-21 09:00:00"))
            j47 = [j for j in A.list_testing_jobs() if j["eng"] == "4700"]
            st47 = A.build_receiving_state("4700", "192-146-900")
            check("r2: 6-digit core — dropdown entry is REAL (no tag) and the walk agrees (has_testing True, -900 name, no stop)",
                  len(j47) == 1 and j47[0]["pn"] == "192-146-900" and "testing" not in j47[0]
                  and A.receiving_summary(st47)["has_testing"] is True and A._report_pn(st47) == "192-146-900"
                  and A.ReceivingWindow._coated_only(st47) is False and len(st47["shipped"]) == 1)
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histC6, ignore_errors=True)

        # ── a coated-LABELLED import (⏪ on the -970 WO) carrying the in-progress -900-serialized
        #    batch: Case B for backfills too — tag, stop, report name and header all agree ──
        histB = tempfile.mkdtemp(prefix="pcc_bfcaseb_")
        A.HISTORY_DIR = histB
        try:
            bS = lambda i: f"{i:06d}-1251031900-D00"
            ev(histB, "2026-08", "a_clean", close_ev("4007", "125-1031-970D", "16494.2",
               [f"{i:06d}-1251031970-D00" for i in range(362, 366)], "2026-08-24 10:00:00", str(wo7)))
            ev(histB, "2026-08", "b_bf_caseb", {"event_type": "backfill_shipped", "eng": "4007", "pn": "125-1031-970D",
               "wo": "16494.2", "ts": "2026-08-25 09:00:00", "ship_date": "8/1/26", "event_id": "ev-bf-caseb",
               "serials": [bS(12), bS(34), bS(35)], "production_data": str(wo7.parent)})
            j7b = [j for j in A.list_testing_jobs() if j["eng"] == "4007"]
            st7b = A.build_receiving_state("4007", "125-1031-900")
            cb7 = [e for e in st7b["coated_log"] if e["serial"] == bS(12)]
            check("r2: coated-labelled BACKFILL with -900-form serials -> coated-only entry AND Case-B coated rows (Note), shipped {}, has_testing False, coated name, stop fires",
                  len(j7b) == 1 and j7b[0].get("testing") is False and j7b[0]["coated_label"] == "125-1031-970D"
                  and st7b["shipped"] == {} and len(cb7) == 1 and cb7[0]["note"] == A._CASE_B_NOTE and cb7[0]["wo"] == "16494.2"
                  and A._serial_identity(bS(12), "125-1031-970") in st7b.get("coat_label_mismatch", set())
                  and A.receiving_summary(st7b)["has_testing"] is False and A.receiving_summary(st7b)["coated_all"] == 7
                  and A._report_pn(st7b) == "125-1031-970D" and A.ReceivingWindow._coated_only(st7b) is True)
            check("r2: BackfillDialog guard fires for a coated label typed with -900-form serials (letter-tolerant); the -900-typed direction unchanged",
                  A._backfill_serials_look_coated("125-1031-970D", [bS(12), bS(34)]) is True
                  and A._backfill_serials_look_coated("125-1031-970D", [bS(12), "000362-1251031970-D00"]) is False
                  and A._backfill_serials_look_coated("125-1031-900", ["000362-1251031970-D00"]) is True
                  and A._backfill_serials_look_coated(TP9, [c9(1), c9(2)]) is True
                  and A._backfill_serials_look_coated("125-1031-970D 4007 Rev. D", [bS(12)]) is False)
            # review round 5 (test-adequacy-r5-3): a supervisor-overridden receive of a BACKFILL-origin Case-B board
            # keeps the coated-only entry (list side and walk agree — decision (a)); a genuine -900 board flips both
            ev(histB, "2026-09", "b2_recv_caseb", recv_ev("4007", "125-1031-900", [bS(12)], "2026-09-01 09:00:00"))
            j7c = [j for j in A.list_testing_jobs() if j["eng"] == "4007"]; st7c = A.build_receiving_state("4007", "125-1031-900")
            ev(histB, "2026-09", "b3_recv_real", recv_ev("4007", "125-1031-900", [bS(700)], "2026-09-01 09:30:00"))
            j7d = [j for j in A.list_testing_jobs() if j["eng"] == "4007"]; st7d = A.build_receiving_state("4007", "125-1031-900")
            for _n in ("b2_recv_caseb", "b3_recv_real"):
                os.remove(Path(histB) / "2026" / "2026-09" / f"{_n}.json")
            check("r5: receive of a backfill-origin Case-B board -> entry still coated-only (tag, coated_label), walk has_testing False, coated name, received 1; a genuine -900 board then flips BOTH sides (real entry, -900 name)",
                  len(j7c) == 1 and j7c[0].get("testing") is False and j7c[0]["coated_label"] == "125-1031-970D"
                  and A.ReceivingWindow._job_label(j7c[0]).endswith(A._COATED_ONLY_TAG)
                  and A.receiving_summary(st7c)["has_testing"] is False and A.ReceivingWindow._coated_only(st7c) is True
                  and A._report_pn(st7c) == "125-1031-970D" and len(st7c["received"]) == 1
                  and len(j7d) == 1 and "testing" not in j7d[0] and j7d[0]["pn"] == "125-1031-900"
                  and A.receiving_summary(st7d)["has_testing"] is True and A._report_pn(st7d) == "125-1031-900")
            ev(histB, "2026-09", "c_void_bf", {"event_type": "void", "eng": "4007", "pn": "125-1031-970D", "wo": "16494.2",
               "ts": "2026-09-01 10:00:00", "target_event_id": "ev-bf-caseb"})
            st7v = A.build_receiving_state("4007", "125-1031-900")
            check("r2: voiding that import purge-lists its Case-B coated LOG rows (same rule as the live route)",
                  (A._serial_identity(bS(12), "125-1031-970"), A._date_key("8/1/26")) in st7v["voided_coat_rows"]
                  and not any(e["serial"] == bS(12) for e in st7v["coated_log"]) and len(st7v["coated_log"]) == 4)
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histB, ignore_errors=True)

        # ── the picker's tick model, Tk-free, plus the window helpers / source pins ──
        rowsX = [{"job": {"eng": str(i)}, "has": h, "tickable": t}
                 for i, (h, t) in enumerate([(True, True), (True, False), (False, True), (False, True), (False, False)])]
        tX = A._picker_default_ticks(rowsX)
        check("r2: picker model — default ticks = has AND tickable; un-tickable rows refuse; picked = ticked AND tickable",
              tX == [True, False, False, False, False]
              and A._picker_set(rowsX, tX, 1, True) is False and tX[1] is False
              and A._picker_set(rowsX, tX, 4, True) is False and tX[4] is False
              and A._picker_set(rowsX, tX, 9, True) is False
              and A._picker_set(rowsX, tX, 2, True) is True
              and [r["job"]["eng"] for r in A._picker_picked(rowsX, tX)] == ["0", "2"]
              and A._picker_picked(rowsX, [True] * 5) == [rowsX[0], rowsX[2], rowsX[3]]
              and A._picker_set(rowsX, tX, 0, False) is True and [r["job"]["eng"] for r in A._picker_picked(rowsX, tX)] == ["2"])
        check("r2: window — _shown_pn is the real coated label for an in-house-tested job, the -900 key otherwise; _job_display_pn",
              A.ReceivingWindow._shown_pn(st9) == CPL9 and A.ReceivingWindow._shown_pn(st) == TPN
              and A._job_display_pn(c_ent[0]) == CPL9 and A._job_display_pn({"eng": "4041", "pn": TPN}) == TPN)
        src2 = Path(SCRIPTS, "app_logic.py").read_text(encoding="utf-8")
        srcPk = src2[src2.index("class ReportPickerDialog(tk.Toplevel):"):src2.index("class BackfillDialog(tk.Toplevel):")]
        check("r2: wiring — the dialog delegates to the model; no <Return> binding on its Treeview (a scan's Enter can't flip a tick)",
              "self._ticked = _picker_default_ticks(self._rows)" in srcPk
              and "if not _picker_set(self._rows, self._ticked, i, val): return" in srcPk
              and "picked = _picker_picked(self._rows, self._ticked)" in srcPk
              and '.bind("<Return>"' not in srcPk and 'self._tv.bind("<space>"' in srcPk)
        check("r2: wiring — _regen_xlsx marks the cache written; a failed pre-save re-walk forgets it; chips + Save popup name the copy; Rebuild refreshes chips before asking",
              "_mark_report_written(st, out)" in src2
              and "fresh_ok = self._rebuild_state_sync(" in src2 and "        if not fresh_ok:\n" in src2
              and "_forget_report_cache(cur)" in src2
              and "the copy that keeps this job opted in: {found}" in src2
              and "Delete that copy too if this job should stop getting a report." in src2
              and "self._refresh_chips()      # the fresh state is on screen even if we stop below" in src2)
        check("r2: wiring — job-entry display P/N in ONE helper; the Save-path refusal branches; BackfillDialog texts follow the typed side",
              src2.count("_job_display_pn(j)") >= 5
              and "j.get('coated_label') or j.get('pn')" not in src2 and 'j.get("coated_label") or j.get("pn")' not in src2
              and 'if not state.get("allow_tree_switch") and not state.get("_resolve_unknown"):' in src2
              and "recorded as COATED shipments (this P/N is the coated side)" in src2
              and 'messagebox.askyesno("Testing-form serials?",' in src2
              and "label=_job_display_pn(j)" in src2 and "P/N {self._shown_pn(st)}" in src2)

        # ══ review round 3 ═══════════════════════════════════════════════════
        _nrows = lambda p, sheet: sum(1 for r in _rows_of(p, sheet) if isinstance(r[0], int))
        # ── a coated-only job WITH a report + a supervisor-overridden receive of one of its
        #    -900-labelled Case-B boards (the 8/27 damage boards): the report keeps its home ──
        hist3 = tempfile.mkdtemp(prefix="pcc_r3caseb_")
        cust3 = Path(tempfile.mkdtemp(prefix="pcc_r3casebc_")) / "377 Anduril"; extra_roots9.append(cust3.parent)
        wo3 = cust3 / "125-1030-970C 4006 Rev. C" / "Production Data" / "16493.2"; (wo3 / A.COC_FOLDER).mkdir(parents=True)
        wo3n = cust3 / "125-1030-970C 4111 Rev. C" / "Production Data" / "16493.2"; (wo3n / A.COC_FOLDER).mkdir(parents=True)
        A.HISTORY_DIR = hist3
        try:
            ev(hist3, "2026-08", "a_coat", dict(close_ev(E9, CPL9, "16493.2",
               [c9(i) for i in range(861, 871)] + [t9(569), t9(571)], "2026-08-27 11:00:00", str(wo3)),
               notes="Anduril test damage"))
            out3 = A.write_received_xlsx(A.build_receiving_state(E9, TP9))          # 🔁 Build report
            # the receive event exactly as _do_save writes it after 'Nothing to receive here'
            # -> Supervisor override -> Save (pn = the entry's -900 key, production_data = the -970 PD)
            ev(hist3, "2026-09", "b_override", dict(recv_ev(E9, TP9, [t9(569)], "2026-09-06 09:00:00",
               "Anita Patel", str(wo3.parent)), failed=[],
               dup_overrides=[{"serial": t9(569), "reason": "Nothing to receive here"}]))
            st3r = A.build_receiving_state(E9, TP9)
            st3a = A._annotate_report_exists(dict(st3r))
            check("r3: coated-only job WITH a report + an overridden -900-form Case-B receive -> keeps its coated home, _report_exists True, save gate True",
                  st3r["wo_path"] == str(wo3) and st3r.get("home_side") == "coated" and len(st3r["received"]) == 1
                  and A.receiving_summary(st3r)["has_testing"] is False       # decision (a): a Case-B board is not testing proof
                  and A._report_pn(st3r) == CPL9
                  and st3a["_report_exists"] is True and A._receive_save_gate(st3a) is True)
            rows3 = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == E9])
            check("r3: its Rebuild-ALL row stays 'has report' / tickable (not 'no WO folder known')",
                  len(rows3) == 1 and rows3[0]["status"] == "has report" and rows3[0]["tickable"] is True)
            ev(hist3, "2026-09", "c_coat2", close_ev(E9, CPL9, "16493.2", [c9(i) for i in range(871, 874)],
               "2026-09-09 10:00:00", str(wo3)))
            A.auto_update_shipping_report(E9, CPL9)
            files3 = sorted(cust3.rglob("*Shipping Report*.xlsx"))
            # (decision (a), req-r3-2: the overridden Case-B receive is NOT testing activity —
            #  the file KEEPS the real coated label; the daemon never renames it to -900)
            check("r3: the next coated close's after-close update appends to that ONE report in the -970 WO (12 -> 15 COAT rows)",
                  len(files3) == 1 and files3[0].parent.parent == wo3 and _nrows(files3[0], A.COAT_SHEET) == 15)
            check("decision (a): that file keeps the real-label NAME after the Case-B override (no -900 rename by the daemon)",
                  files3[0] == Path(out3) and files3[0].name == f"{CPL9} Shipping Report.xlsx")
            wb3 = load_workbook(files3[0])
            check("decision (a): its A1 is the real label, it opens on 'Coated & shipped', A3 says in-house tested, and the RCV subtitle owns its 1 overridden row (never 'none' above a row)",
                  wb3[A.RCV_SHEET]["A1"].value == f"SHIPPING REPORT — P/N {CPL9}" and wb3.active.title == A.COAT_SHEET
                  and str(wb3[A.RCV_SHEET]["A3"].value).startswith("In-house tested — not shipped for testing")
                  and _nrows(files3[0], A.RCV_SHEET) == 1
                  and wb3[A.RCV_SHEET]["A2"].value == f"Engineering # {E9}    ·    received back from customer testing — 1 row(s), supervisor-overridden scans (this job's boards are tested in-house)"
                  and wb3[A.IH_SHEET]["A2"].value == f"Engineering # {E9}    ·    IN-HOUSE — not applicable (nothing shipped for testing on this job) — 0 board(s)")
            wb3.close()
            r569 = next((r for r in _rows_of(files3[0], A.RCV_SHEET) if r[1] == t9(569)), None)
            check("r6: the RCV row of the overridden Case-B receive points at its coated shipment ('16493.2 (coated, -900 label)' + the 8/27 date) instead of 'not in shipping records'",
                  r569 is not None and r569[5] == "16493.2 (coated, -900 label)" and str(r569[6]).startswith("2026-08-27"))
            box3r = {}; A._rebuild_report_rows(rows3, box3r)
            check("r3: a manual Rebuild of that row writes (no 'no work-order folder known' refusal)",
                  box3r["ok"] == 1 and box3r["failed"] == [] and box3r["moved"] == 0
                  and len(sorted(cust3.rglob("*Shipping Report*.xlsx"))) == 1)
            # THE FILE IS THE SWITCH: even a -900-form board that is NOT a Case-B board of this job
            # (never closed anywhere) keeps the home while the report exists in the coated WO…
            ev(hist3, "2026-09", "d_override2", dict(recv_ev(E9, TP9, [t9(700)], "2026-09-10 09:00:00",
               "Anita Patel", str(wo3.parent)), failed=[]))
            st3f = A.build_receiving_state(E9, TP9)
            check("r3: 'the file is the switch' — a non-Case-B -900-form override receive still keeps the coated home while the report exists there",
                  st3f["wo_path"] == str(wo3) and st3f.get("home_side") == "coated" and len(st3f["received"]) == 2
                  and A._report_exists(st3f)[0] is True)
            # review round 4 (datasafety-r4-1): that switch must find the report the way the WRITE
            # path does — right after a NEW coated WO's first close the file still sits in the
            # PREVIOUS WO (it follows on the next write); the raw newest-close string alone said
            # 'no report here' and orphaned it (wo_path None: nothing updated, Rebuild 'No job folder')
            wo3b = wo3.parent / "16493.3"; (wo3b / A.COC_FOLDER).mkdir(parents=True)
            ev(hist3, "2026-09", "h_coat3", close_ev(E9, CPL9, "16493.3", [c9(i) for i in range(881, 884)],
               "2026-09-11 10:00:00", str(wo3b)))
            st3g = A.build_receiving_state(E9, TP9)
            rows3g = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == E9])
            A.auto_update_shipping_report(E9, CPL9)
            files3g = sorted(cust3.rglob("*Shipping Report*.xlsx"))
            # (the genuine t9(700) receive flipped has_testing — decision (a) — so the ONE file now
            #  carries the -900 name; the 970C-named copy is one of its spellings and folds in)
            check("r4: after the override receive a NEW coated WO's first close keeps the coated home (wo_path = the new WO), the picker row stays 'has report', and the after-close update MIGRATES the one file there (15 -> 18 COAT rows)",
                  st3g["wo_path"] == str(wo3b) and st3g.get("home_side") == "coated"
                  and len(rows3g) == 1 and rows3g[0]["status"] == "has report" and rows3g[0]["tickable"] is True
                  and files3g == [wo3b / A.RECEIVED_FOLDER / f"{TP9} Shipping Report.xlsx"]
                  and _nrows(files3g[0], A.COAT_SHEET) == 18)
            # …and a REV RENAME of the job folder ('Rev. C' -> 'Rev. D', the recorded string is stale)
            # does not orphan it either: the home stays coated and the resolver finds the renamed tree
            job3d = cust3 / "125-1030-970C 4006 Rev. D"
            shutil.move(str(wo3.parent.parent), str(job3d))
            wo3d = job3d / "Production Data" / "16493.3"
            st3h = A.build_receiving_state(E9, TP9)
            r3h = A._resolve_received_wo_dir(dict(st3h))
            ev(hist3, "2026-09", "i_coat4", close_ev(E9, CPL9, "16493.3", [c9(884)], "2026-09-12 10:00:00", str(wo3d)))
            A.auto_update_shipping_report(E9, CPL9)
            f3h = wo3d / A.RECEIVED_FOLDER / f"{TP9} Shipping Report.xlsx"
            check("r4: a rev-renamed coated job folder keeps the coated home too (wo_path known, resolver -> the renamed WO) and the next update lands on the SAME file (19 rows, one live file)",
                  st3h["wo_path"] == str(wo3b) and st3h.get("home_side") == "coated"
                  and r3h == wo3d and A._report_exists(st3h)[0] is True
                  and sorted(cust3.rglob("*Shipping Report*.xlsx")) == [f3h] and _nrows(f3h, A.COAT_SHEET) == 19)
            # review round 5 (datasafety-r5-3): a rev-TWIN coated job folder ('125-1030-970B 4006 Rev. E' — a
            # second LIVE -970 job dir, the 4094 shape) takes the next batch: the 'file is the switch' probe
            # looked in that ONE job dir only, found nothing, and the job lost its home for good (daemon
            # skipped, picker 'no WO folder known', Build refused — self-perpetuating)
            job3t = cust3 / "125-1030-970B 4006 Rev. E"; wo3t = job3t / "Production Data" / "17100.1"
            (wo3t / A.COC_FOLDER).mkdir(parents=True)
            ev(hist3, "2026-09", "j_twin", close_ev(E9, "125-1030-970B", "17100.1", [c9(890), c9(891)],
               "2026-09-14 10:00:00", str(wo3t)))
            st3t = A.build_receiving_state(E9, TP9)
            rows3t = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == E9])
            A.auto_update_shipping_report(E9, "125-1030-970B")
            f3t = wo3t / A.RECEIVED_FOLDER / f"{TP9} Shipping Report.xlsx"
            check("r5: first close in a live rev-TWIN coated job dir -> wo_path = the twin's WO (home coated), picker 'has report' / tickable, the update MIGRATES the one file there (19 -> 21 rows)",
                  st3t["wo_path"] == str(wo3t) and st3t.get("home_side") == "coated"
                  and len(rows3t) == 1 and rows3t[0]["status"] == "has report" and rows3t[0]["tickable"] is True
                  and sorted(cust3.rglob("*Shipping Report*.xlsx")) == [f3t] and _nrows(f3t, A.COAT_SHEET) == 21)
            # …and a report a human parked in the -900 tree (an idle -900 WO) is found from there too
            wo3k = cust3 / "125-1030-900 4006 Rev. C" / "Production Data" / "16400.1"
            (wo3k / A.RECEIVED_FOLDER).mkdir(parents=True)
            shutil.move(str(f3t), str(wo3k / A.RECEIVED_FOLDER / f3t.name)); f3t.parent.rmdir()
            wo3u = job3t / "Production Data" / "17100.2"; (wo3u / A.COC_FOLDER).mkdir(parents=True)
            ev(hist3, "2026-09", "k_twin2", close_ev(E9, "125-1030-970B", "17100.2", [c9(892)], "2026-09-15 10:00:00", str(wo3u)))
            st3k = A.build_receiving_state(E9, TP9)
            A.auto_update_shipping_report(E9, "125-1030-970B")
            f3u = wo3u / A.RECEIVED_FOLDER / f"{TP9} Shipping Report.xlsx"
            check("r5: the report parked in the -900 tree keeps the coated home for a NEW coated WO's first close, and the update brings the one file there (22 rows)",
                  st3k["wo_path"] == str(wo3u) and st3k.get("home_side") == "coated"
                  and sorted(cust3.rglob("*Shipping Report*.xlsx")) == [f3u] and _nrows(f3u, A.COAT_SHEET) == 22)
            # review round 6 (test-r6-1): the probe's 'a matched folder was unlistable' clause — the report parked in
            # the idle -900 WO again, and THAT job folder's Production Data cannot be listed from this PC: 'can't
            # check', never 'no WO folder known' (a checkable walk keeps the home through the file — see above)
            (wo3k / A.RECEIVED_FOLDER).mkdir(parents=True, exist_ok=True)
            shutil.move(str(f3u), str(wo3k / A.RECEIVED_FOLDER / f3u.name))
            def _deny_900_pd(self):
                if self.name == "Production Data" and "125-1030-900 4006" in self.parent.name:
                    raise PermissionError(13, "Access is denied")
                return _orig_iterdir(self)
            A.Path.iterdir = _deny_900_pd
            try:
                st3m = A._annotate_report_exists(A.build_receiving_state(E9, TP9))
                rows3m = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == E9])
                gate3m = A._receive_save_gate(dict(st3m))
            finally:
                A.Path.iterdir = _orig_iterdir
            check("r6: the -900 job folder holding the parked report is UNLISTABLE -> wo_path None + _home_unknown, _report_lookup (None, None, None), cache None, gate None, picker \"can't check from this PC\" (not 'no WO folder known')",
                  st3m["wo_path"] is None and st3m.get("_home_unknown") is True and A._report_lookup(st3m) == (None, None, None)
                  and st3m["_report_exists"] is None and gate3m is None
                  and len(rows3m) == 1 and rows3m[0]["status"] == "can't check from this PC" and rows3m[0]["tickable"] is False)
            shutil.move(str(wo3k / A.RECEIVED_FOLDER / f3u.name), str(f3u)); (wo3k / A.RECEIVED_FOLDER).rmdir()   # back where the update put it
            # review round 5 (optin-r5-2): the same job walked where the customer folder is UNREACHABLE — the probe
            # could not CHECK, which is not 'the folders hold no report': wo_path stays None (no writer places a
            # file) but the switch answers None and every consumer says "can't check", never 'no WO folder known'
            hidden3 = cust3.parent / "377 Anduril (hidden)"
            os.rename(str(cust3), str(hidden3))
            try:
                st3u = A._annotate_report_exists(A.build_receiving_state(E9, TP9))
                rows3u = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == E9])
                gate3u = A._receive_save_gate(st3u)
            finally:
                os.rename(str(hidden3), str(cust3))
            check("r5: customer folder unreachable -> wo_path None but _home_unknown, _report_lookup (None, None, None), cache None, save gate None, picker \"can't check from this PC\" (not 'no WO folder known')",
                  st3u["wo_path"] is None and st3u.get("_home_unknown") is True
                  and A._report_lookup(st3u) == (None, None, None) and st3u["_report_exists"] is None and gate3u is None
                  and len(rows3u) == 1 and rows3u[0]["status"] == "can't check from this PC" and rows3u[0]["tickable"] is False
                  and "_home_unknown" not in st3k)     # (a checkable walk never carries the flag)
            # …while the same receive on a coated-only job WITHOUT a report still refuses the coated
            # anchor (round 2's receive-only rule: a -900 board received = a -900 shipment somewhere)
            ev(hist3, "2026-09", "e_4111", dict(close_ev("4111", CPL9, "16493.2", [c9(1), c9(2), t9(800)],
               "2026-09-05 09:00:00", str(wo3n)), notes="Anduril test damage"))
            # first its own Case-B board comes back (overridden): History alone keeps the coated
            # home — no report needed, so the picker can still build the job's FIRST report
            ev(hist3, "2026-09", "f_4111_rcv_caseb", dict(recv_ev("4111", TP9, [t9(800)], "2026-09-06 08:00:00",
               "Anita Patel", str(wo3n.parent)), failed=[]))
            st3c = A.build_receiving_state("4111", TP9)
            rows3c = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == "4111"])
            check("r3: coated-only job with NO report + an overridden receive of its own Case-B board -> coated home kept from History alone (picker row 'no report', tickable)",
                  st3c["wo_path"] == str(wo3n) and st3c.get("home_side") == "coated" and len(st3c["received"]) == 1
                  and A._report_exists(st3c) == (False, A.received_xlsx_path(wo3n, CPL9))   # decision (a): still the real-label name
                  and A.receiving_summary(st3c)["has_testing"] is False
                  and len(rows3c) == 1 and rows3c[0]["status"] == "no report" and rows3c[0]["coated_only"] is True)
            ev(hist3, "2026-09", "g_4111_rcv", dict(recv_ev("4111", TP9, [t9(700)], "2026-09-06 09:00:00",
               "Anita Patel", str(wo3n.parent)), failed=[]))
            st3n = A.build_receiving_state("4111", TP9)
            check("r3 (the 2026-10-06 rule: the report homes in the -970 coated WO): with NO report anywhere a -900-form receive homes in the -970 WO, and nothing is created (opt-in: no Build yet)",
                  st3n.get("home_side") == "coated" and Path(st3n["wo_path"]).name == "16493.2"
                  and not (wo3n / A.RECEIVED_FOLDER).exists())
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(hist3, ignore_errors=True)

        # ── spelling variant of the walked job's OWN coated label (BackfillDialog's -970 walk):
        #    '125-1030-970' closes stay in `shipped` of a '125-1030-970C' walk (idempotent import) ──
        hist3v = tempfile.mkdtemp(prefix="pcc_r3own_")
        A.HISTORY_DIR = hist3v
        try:
            ev(hist3v, "2026-08", "a_970C", close_ev(E9, CPL9, "16493.1", [c9(1), c9(2)], "2026-08-20 09:00:00", str(wo970a)))
            ev(hist3v, "2026-08", "b_970", close_ev(E9, CP9, "16493.2", [c9(3), c9(4)], "2026-08-27 09:00:00", str(wo970b)))
            stOwn = A.build_receiving_state(E9, CPL9)
            stOwn9 = A.build_receiving_state(E9, TP9)
            check("r3: a -970C walk keeps the same job's plain '-970' close in shipped (4, nothing re-routed as Case B); the -900 walk still files both as coated",
                  len(stOwn["shipped"]) == 4 and stOwn["coated_log"] == [] and not stOwn.get("coat_label_mismatch")
                  and stOwn.get("testing_pn_labels") == [CP9]
                  and stOwn9["shipped"] == {} and len(stOwn9["coated_log"]) == 4 and stOwn9.get("coated_pn_label") == CP9)
            # the dropdown's coated label follows the newest CLOSE, like coated_pn_label —
            # a NEWER import typed '125-1030-970 C' must not relabel the entry
            ev(hist3v, "2026-09", "c_bf_space", {"event_type": "backfill_shipped", "eng": E9, "pn": "125-1030-970 C",
               "wo": "16493.1", "ts": "2026-09-01 09:00:00", "ship_date": "8/1/26", "serials": [c9(5)],
               "production_data": str(pd970)})
            ent3v = [j for j in A.list_testing_jobs() if j["eng"] == E9]
            check("r3: list_testing_jobs — coated_label = the newest CLOSE's spelling even when a newer import spells it differently (ts = newest event, no private key)",
                  len(ent3v) == 1 and ent3v[0]["coated_label"] == CP9 and ent3v[0]["ts"].startswith("2026-09-01")
                  and "_from_close" not in ent3v[0] and ent3v[0].get("testing") is False)
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(hist3v, ignore_errors=True)

        # ── merge BASE ranking: a LIVE previous-WO copy (user sheet, OLDER mtime) beats an
        #    archived old-rev snapshot that someone opened in Excel (NEWER mtime) — the 4092 shape ──
        hist3b = tempfile.mkdtemp(prefix="pcc_r3arch_")
        cust3b = Path(tempfile.mkdtemp(prefix="pcc_r3archc_")) / "377 Anduril"; extra_roots9.append(cust3b.parent)
        pdB = cust3b / "136-3312-900_4092 Rev B01 Rev. B02" / "Production Data"
        (pdB / "16586.2").mkdir(parents=True); (pdB / "16600.1").mkdir(parents=True)
        tB = lambda i: f"{i:06d}-136-3312-900-B02"
        A.HISTORY_DIR = hist3b
        try:
            ev(hist3b, "2026-07", "a_close_1", close_ev("4092", "136-3312-900", "16586.2", [tB(i) for i in range(1, 6)],
               "2026-07-12 09:00:00", str(pdB / "16586.2")))
            ev(hist3b, "2026-07", "c_recv", recv_ev("4092", "136-3312-900", [tB(i) for i in range(1, 4)],
               "2026-07-20 14:00:00", "Lata Gandhi", str(pdB)))
            liveB = Path(A.write_received_xlsx(A.build_receiving_state("4092", "136-3312-900")))
            wbB = load_workbook(liveB); wbB.create_sheet("My notes")["A1"] = "keep me"; wbB.save(liveB); wbB.close()
            archB = (cust3b / "Archive" / "136-3312-900_4092 Rev A01 Rev A" / "Production Data" / "16500.1"
                     / A.RECEIVED_FOLDER / liveB.name)
            archB.parent.mkdir(parents=True); shutil.copy2(str(liveB), str(archB))
            awb = load_workbook(archB); awb.remove(awb["My notes"])
            awb[A.RCV_SHEET].append([99, "STALE-ROW-FROM-OLD-REV", "PASS", "2026-02-02", "Old Copy", "", ""])
            awb.save(archB); awb.close()
            nowB = time.time(); os.utime(str(liveB), (nowB - 3600, nowB - 3600)); os.utime(str(archB), (nowB, nowB))
            asigB = A._file_sig(archB)                 # fingerprint AFTER the mtime bump
            ev(hist3b, "2026-08", "a_close_2", close_ev("4092", "136-3312-900", "16600.1", [tB(i) for i in range(6, 11)],
               "2026-08-14 09:00:00", str(pdB / "16600.1")))
            stB = A.build_receiving_state("4092", "136-3312-900")
            check("r3: _find_legacy_received_file ranks the LIVE copy above the newer-by-mtime archived one",
                  A._find_legacy_received_file(stB, A.received_xlsx_path(pdB / "16600.1", "136-3312-900")) == liveB)
            outB = Path(A.write_received_xlsx(stB))                      # the after-close update's write
            wbB2 = load_workbook(outB)
            rowsB = [r[1] for r in wbB2[A.RCV_SHEET].iter_rows(values_only=True) if r and r[1]]
            liveB_files = sorted(p for p in cust3b.rglob("*Shipping Report*.xlsx") if not A._under_archive(p))
            check("r3: new-WO report carries the user sheet, has NO archived stale row, the previous-WO copy is gone — exactly ONE live file; archive untouched",
                  outB.parent.parent == pdB / "16600.1" and "My notes" in wbB2.sheetnames
                  and "STALE-ROW-FROM-OLD-REV" not in rowsB and not liveB.exists()
                  and liveB_files == [outB] and archB.exists() and A._file_sig(archB) == asigB)
            wbB2.close()
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(hist3b, ignore_errors=True)

        # ── after a manual Rebuild re-homed the report into the SIBLING tree, the everyday
        #    updates (after-close daemon, receiving Save) keep reaching it — both directions ──
        hist3p = tempfile.mkdtemp(prefix="pcc_r3post_")
        cust3p = Path(tempfile.mkdtemp(prefix="pcc_r3postc_")) / "20 TestCust"; extra_roots9.append(cust3p.parent)
        jobP900 = cust3p / "150-3300-900 5001 Rev. A"; woP900 = jobP900 / "Production Data" / "17001.1"
        jobP970 = cust3p / "150-3300-970 5001 Rev. A"; woP970 = jobP970 / "Production Data" / "17002.1"
        (woP900 / A.COC_FOLDER).mkdir(parents=True); (woP970 / A.COC_FOLDER).mkdir(parents=True)
        tP = lambda i: f"{i:06d}-1503300900-A00"; cP = lambda i: f"{i:06d}-1503300970-A00"
        A.HISTORY_DIR = hist3p
        try:
            ev(hist3p, "2026-06", "a_close_900", close_ev("5001", "150-3300-900", "17001.1", [tP(i) for i in range(1, 4)],
               "2026-06-10 09:00:00", str(woP900)))
            ev(hist3p, "2026-08", "b_close_970", close_ev("5001", "150-3300-970", "17002.1", [cP(1), cP(2)],
               "2026-08-18 10:00:00", str(woP970)))
            outP = Path(A.write_received_xlsx(A.build_receiving_state("5001", "150-3300-900")))    # old-flow: -900 WO
            (cust3p / "Archive").mkdir(); shutil.move(str(jobP900), str(cust3p / "Archive" / jobP900.name))
            ev(hist3p, "2026-09", "c_close_970b", close_ev("5001", "150-3300-970", "17002.1", [cP(3)], "2026-09-01 10:00:00", str(woP970)))
            A.auto_update_shipping_report("5001", "150-3300-970")
            check("r3 (the 2026-10-06 rule: the report homes in the -970 coated WO): the first write already files the report in the -970 WO 17002.1; nothing in the -900 tree",
                  outP.parent.parent == woP970 and not (woP900 / A.RECEIVED_FOLDER).exists())
            rowP = A._collect_report_rows({}, jobs=[{"eng": "5001", "pn": "150-3300-900", "ts": "2026-09-01"}])[0]
            boxP = {}; A._rebuild_report_rows([rowP], boxP)
            movedP = woP970 / A.RECEIVED_FOLDER / "150-3300-900 Shipping Report.xlsx"
            check("r3 (the 2026-10-06 rule: the report homes in the -970 coated WO): Rebuild ALL finds it already home in the -970 WO (ok 1, moved 0, 3 COAT rows)",
                  boxP["ok"] == 1 and boxP["moved"] == 0 and movedP.exists() and _nrows(movedP, A.COAT_SHEET) == 3)
            ev(hist3p, "2026-09", "d_close_970c", close_ev("5001", "150-3300-970", "17002.1", [cP(4), cP(5)], "2026-09-10 10:00:00", str(woP970)))
            stP = A.build_receiving_state("5001", "150-3300-900")
            rP = A._resolve_received_wo_dir(dict(stP))
            A.auto_update_shipping_report("5001", "150-3300-970")
            check("r3: the after-close update then reaches the moved report (resolver -> the -970 WO with NO tree switch; 3 -> 5 COAT rows)",
                  rP == woP970 and _nrows(movedP, A.COAT_SHEET) == 5)
            stP2 = A._annotate_report_exists(A.build_receiving_state("5001", "150-3300-900"))
            outP2 = A.write_received_xlsx(stP2)                          # what a receiving Save does (no allow_tree_switch)
            liveP = sorted(p for p in cust3p.rglob("*Shipping Report*.xlsx") if not A._under_archive(p))
            check("r3: a receiving Save's gate is True and its write succeeds on the moved file — tree_switched False, still exactly ONE live report",
                  A._receive_save_gate(stP2) is True and stP2["_report_at_home"] is True and Path(outP2) == movedP
                  and stP2.get("tree_switched") is False and liveP == [movedP])
            # review round 6 (test-r6-3): a copy parked at the SIBLING tree's Production Data level (the legacy-era
            # location) with a live WO below it -> that tree's newest WO, for the daemon / Save too (no flag)
            pdcopyP = jobP970 / "Production Data" / A.RECEIVED_FOLDER / movedP.name
            pdcopyP.parent.mkdir(); shutil.move(str(movedP), str(pdcopyP)); movedP.parent.rmdir()
            stP6 = A.build_receiving_state("5001", "150-3300-900"); stP6r = dict(stP6)
            rP6 = A._resolve_received_wo_dir(stP6r)
            lookP6 = A._report_lookup(stP6)
            A.auto_update_shipping_report("5001", "150-3300-970")                          # the after-close daemon
            liveP6 = sorted(p for p in cust3p.rglob("*Shipping Report*.xlsx") if not A._under_archive(p))
            check("r6: a PD-level copy in the sibling -970 tree -> the resolver answers that tree's live WO 17002.1 (tree_switched False), the switch says True, the daemon migrates the ONE live file down into it (5 rows), the PD-level copy is gone",
                  rP6 == woP970 and stP6r.get("tree_switched") is False and lookP6[0] is True
                  and liveP6 == [movedP] and _nrows(movedP, A.COAT_SHEET) == 5 and not pdcopyP.exists())
            # review round 4 (test-adequacy-r4-1): a SECOND live sibling WO must not make the report
            # ping-pong between the two — the everyday update stays on the WO that HOLDS the file
            woP970b = jobP970 / "Production Data" / "17003.1"; (woP970b / A.COC_FOLDER).mkdir(parents=True)
            ev(hist3p, "2026-09", "e_close_970d", close_ev("5001", "150-3300-970", "17003.1", [cP(6), cP(7)],
               "2026-09-12 10:00:00", str(woP970b)))
            A.auto_update_shipping_report("5001", "150-3300-970")                        # after-close daemon
            stP3 = A._annotate_report_exists(A.build_receiving_state("5001", "150-3300-900"))
            outP3 = A.write_received_xlsx(stP3)                                          # a receiving Save
            A.auto_update_shipping_report("5001", "150-3300-970")                        # daemon again
            liveP3 = sorted(p for p in cust3p.rglob("*Shipping Report*.xlsx") if not A._under_archive(p))
            check("r4: with a second live -970 WO the report stays IN PLACE across daemon / Save / daemon (no ping-pong): one live file in 17002.1, 7 COAT rows, cache at_home True, the Save wrote that same file",
                  liveP3 == [movedP] and _nrows(movedP, A.COAT_SHEET) == 7 and Path(outP3) == movedP
                  and stP3["_report_at_home"] is True and stP3["_report_path"] == movedP
                  and not (woP970b / A.RECEIVED_FOLDER).exists())
            # …and a -970 REV-TWIN job folder ('150-3300-970B 5001 Rev. B', newer) does not hide it:
            # the opt-in answer and the write path agree (fresh-r4-3)
            jobP970B = cust3p / "150-3300-970B 5001 Rev. B"; woP970B = jobP970B / "Production Data" / "17010.1"
            (woP970B / A.COC_FOLDER).mkdir(parents=True)
            os.utime(str(woP970B), (time.time() + 60, time.time() + 60))
            ev(hist3p, "2026-09", "f_close_970B", close_ev("5001", "150-3300-970B", "17010.1", [cP(8), cP(9)],
               "2026-09-13 10:00:00", str(woP970B)))
            stP4 = A._annotate_report_exists(A.build_receiving_state("5001", "150-3300-900"))
            gP4 = A._receive_save_gate(stP4)
            stP4r = dict(stP4); rP4 = A._resolve_received_wo_dir(stP4r)
            A.auto_update_shipping_report("5001", "150-3300-970B")
            outP4 = A.write_received_xlsx(A.build_receiving_state("5001", "150-3300-900"))
            liveP4 = sorted(p for p in cust3p.rglob("*Shipping Report*.xlsx") if not A._under_archive(p))
            check("r4: a newer -970 rev-twin job folder — resolver (no switch) -> the holder WO in the OTHER twin, daemon +2 rows (9), save gate True == the Save write succeeds on that file, one live file",
                  rP4 == movedP.parent.parent and stP4r.get("tree_switched") is False
                  and gP4 is True and stP4["_report_at_home"] is True
                  and _nrows(movedP, A.COAT_SHEET) == 9 and Path(outP4) == movedP and liveP4 == [movedP])
            # review round 5 (datasafety-r5-4 / fresh-r5-1): a HUMAN Rebuild (allow_tree_switch) with two live -970
            # WOs (+ the newer rev twin) — the report stays in the WO that HOLDS it; before, the newest WO by mtime
            # won and the emptied leaf's rmdir made the file ping-pong on every Rebuild (popup: 'the previous folder
            # no longer exists' about a live folder, Rebuild ALL counting it 'moved' every run)
            outP5s = []; swP5 = []
            for _ in range(2):
                stP5 = A.build_receiving_state("5001", "150-3300-900"); stP5["allow_tree_switch"] = True
                outP5s.append(Path(A.write_received_xlsx(stP5))); swP5.append(stP5.get("tree_switched"))
            rowP5 = A._collect_report_rows({}, jobs=[{"eng": "5001", "pn": "150-3300-900", "ts": "2026-09-13"}])[0]
            boxP5 = {}; A._rebuild_report_rows([rowP5], boxP5)
            liveP5 = sorted(p for p in cust3p.rglob("*Shipping Report*.xlsx") if not A._under_archive(p))
            check("r5: two manual Rebuilds + a ticked Rebuild-ALL row leave the report in the holder WO 17002.1 — no ping-pong, tree_switched False (nothing moved), moved 0, no leaf in the newer WOs, one live file, 9 rows",
                  outP5s == [movedP, movedP] and swP5 == [False, False]
                  and boxP5["ok"] == 1 and boxP5["moved"] == 0 and boxP5["failed"] == [] and liveP5 == [movedP]
                  and _nrows(movedP, A.COAT_SHEET) == 9
                  and not (woP970b / A.RECEIVED_FOLDER).exists() and not (woP970B / A.RECEIVED_FOLDER).exists())
            # review round 5 (datasafety-r5-1): the -900 tree comes back (archive undone) with its live WO -> the
            # home pass resolves 17001.1 and the target there is missing; on a PC that cannot list the -970 job
            # folder the human Rebuild's write must REFUSE (the copy may sit there), never build a fresh file —
            # blip cleared, the same write migrates the ONE file home
            shutil.move(str(cust3p / "Archive" / jobP900.name), str(jobP900))
            # (the archived tree carried the frozen first-build copy back with it — an archived base is
            #  never deleted; the finding's shape is a live -900 WO WITHOUT a report, so drop that copy)
            shutil.rmtree(str(woP900 / A.RECEIVED_FOLDER))
            def _blip_970(self):
                if self.name == "Production Data" and "150-3300-970 5001" in self.parent.name:
                    raise OSError(64, "The specified network name is no longer available")
                return _orig_iterdir(self)
            A.Path.iterdir = _blip_970
            try:
                stP7 = A._annotate_report_exists(A.build_receiving_state("5001", "150-3300-900"))
                rP7 = A._resolve_received_wo_dir(dict(stP7, allow_tree_switch=True))
                stP7w = dict(A.build_receiving_state("5001", "150-3300-900"), allow_tree_switch=True)
                try:
                    A.write_received_xlsx(stP7w); msgP7 = ""
                except RuntimeError as e:
                    msgP7 = str(e)
            finally:
                A.Path.iterdir = _orig_iterdir
            check("r5: sibling-folder blip with a live HOME WO -> cache None while the resolver answers the home WO; the Rebuild write REFUSES ('could not confirm'), no leaf in 17001.1, the -970 copy untouched (no fork)",
                  stP7["_report_exists"] is None and rP7 == woP900
                  and "could not confirm" in msgP7 and not (woP900 / A.RECEIVED_FOLDER).exists()
                  and movedP.exists() and sorted(p for p in cust3p.rglob("*Shipping Report*.xlsx")) == [movedP])
            stP8 = dict(A.build_receiving_state("5001", "150-3300-900"), allow_tree_switch=True)
            outP8 = Path(A.write_received_xlsx(stP8))
            homeP = woP900 / A.RECEIVED_FOLDER / "150-3300-900 Shipping Report.xlsx"
            check("r5: blip cleared -> the same write brings the ONE file home to 17001.1 (9 COAT rows, tree_switched False, the -970 copy gone)",
                  outP8 == homeP and _nrows(homeP, A.COAT_SHEET) == 9 and not movedP.exists()
                  and stP8.get("tree_switched") is False
                  and sorted(p for p in cust3p.rglob("*Shipping Report*.xlsx")) == [homeP])
            # the coated-only direction: report parked in the -900 WO after a Rebuild, -970 tree gone
            jobQ970 = cust3p / "125-1030-970C 5002 Rev. C"; woQ970 = jobQ970 / "Production Data" / "16493.2"
            woQ900 = cust3p / "125-1030-900 5002 Rev. C" / "Production Data" / "16400.1"
            (woQ970 / A.COC_FOLDER).mkdir(parents=True); woQ900.mkdir(parents=True)
            ev(hist3p, "2026-08", "q_coat", close_ev("5002", CPL9, "16493.2", [c9(1), c9(2)], "2026-08-20 09:00:00", str(woQ970)))
            outQ = Path(A.write_received_xlsx(A.build_receiving_state("5002", TP9)))
            # the -970 tree loses its WO folders (report parked at the PD level); the -900 WO is idle
            shutil.move(str(woQ970 / A.RECEIVED_FOLDER), str(woQ970.parent / A.RECEIVED_FOLDER)); shutil.rmtree(woQ970)
            rowQ = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == "5002"])[0]
            boxQ = {}; A._rebuild_report_rows([rowQ], boxQ)
            movedQ = woQ900 / A.RECEIVED_FOLDER / outQ.name
            ev(hist3p, "2026-09", "q_coat2", close_ev("5002", CPL9, "16493.2", [c9(3)], "2026-09-10 09:00:00", str(woQ970)))
            A.auto_update_shipping_report("5002", CPL9)
            check("r3: coated-only direction — report moved to the -900 WO, the next coated close's update lands on it (2 -> 3 rows), one file",
                  boxQ["ok"] == 1 and boxQ["moved"] == 1 and movedQ.exists() and _nrows(movedQ, A.COAT_SHEET) == 3
                  and sorted(cust3p.rglob(f"{CPL9} Shipping Report*.xlsx")) == [movedQ])
            # review round 5 (test-adequacy-r5-2): the OTHER two partial-enumeration branches — the customer
            # folder itself unlistable, a sibling job dir raising on is_dir — with the only copy in the sibling
            # tree and a live home WO (fast path): 'could not check', never a definite 'no report'
            woQn = jobQ970 / "Production Data" / "16493.3"; (woQn / A.COC_FOLDER).mkdir(parents=True)
            ev(hist3p, "2026-09", "q_coat3", close_ev("5002", CPL9, "16493.3", [c9(4)], "2026-09-11 09:00:00", str(woQn)))
            def _cust_blip(self):
                if self.name == "20 TestCust":
                    raise OSError(64, "The specified network name is no longer available")
                return _orig_iterdir(self)
            _orig_isdir = A.Path.is_dir
            def _sib_isdir(self):
                if self.name == "125-1030-900 5002 Rev. C":
                    raise OSError(64, "The specified network name is no longer available")
                return _orig_isdir(self)
            resQ = []
            for attr_, fn_, orig_ in (("iterdir", _cust_blip, _orig_iterdir), ("is_dir", _sib_isdir, _orig_isdir)):
                setattr(A.Path, attr_, fn_)
                try:
                    stQb = dict(A.build_receiving_state("5002", TP9))
                    resQ.append((A._report_lookup(stQb)[0], A._receive_save_gate(dict(stQb)),
                                 A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == "5002"])[0]["status"]))
                finally:
                    setattr(A.Path, attr_, orig_)
            check("r5: customer folder unlistable / sibling job dir raising on is_dir -> None, save gate None, picker \"can't check from this PC\" (the two remaining except branches of the enumeration)",
                  resQ == [(None, None, "can't check from this PC")] * 2)
            os.remove(Path(hist3p) / "2026" / "2026-09" / "q_coat3.json"); shutil.rmtree(woQn)
            # a sibling WO that holds NO report is still refused without allow_tree_switch (nothing moves)
            stQz = dict(A.build_receiving_state("5002", TP9)); movedQ.unlink()
            check("r3: with the moved file gone the sibling WO is refused again (no report there to update; no fabrication)",
                  A._resolve_received_wo_dir(stQz) is None and stQz.get("tree_switched") is False
                  and stQz.get("_resolve_unknown") is False)
            # a FIRST build whose home tree has no live WO lands in the idle sibling WO: 'created',
            # never 'moved' — nothing existed anywhere to move (review round 4)
            woR900 = cust3p / "125-1030-900 5003 Rev. C" / "Production Data" / "16410.1"; woR900.mkdir(parents=True)
            ev(hist3p, "2026-08", "r_coat", close_ev("5003", CPL9, "16499.2", [c9(1), c9(2)], "2026-08-20 09:00:00",
               str(cust3p / "125-1030-970C 5003 Rev. C" / "Production Data" / "16499.2")))   # -970 tree never on disk
            rowR = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == "5003"])[0]
            boxR = {}; A._rebuild_report_rows([rowR], boxR)
            check("r4: a FIRST build into the idle sibling WO (no copy anywhere) is counted 'created' (1), NOT 'moved' (0); its row read 'no report' with nothing found",
                  rowR["status"] == "no report" and rowR["found"] is None and rowR["tickable"] is True
                  and boxR["ok"] == 1 and boxR["created"] == 1 and boxR["moved"] == 0 and boxR["failed"] == []
                  and (woR900 / A.RECEIVED_FOLDER / f"{CPL9} Shipping Report.xlsx").exists()
                  and rowR["st"].get("tree_switched") is True)
            # review round 4 (optin-r4-2 / quality-r4-1): a SIBLING job folder that exists but cannot be
            # listed from this PC -> 'could not check' (None), never 'no report' — as the home pass does
            def _sib_unreadable(self):
                if self.name == "Production Data" and "125-1030-900 5002" in self.parent.name:
                    raise PermissionError("no access")
                return _orig_iterdir(self)
            A.Path.iterdir = _sib_unreadable
            try:
                stQu = dict(A.build_receiving_state("5002", TP9))
                rQu = A._resolve_received_wo_dir(stQu)
                hasQu = A._report_exists(stQu)[0]
                rowQu = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == "5002"])[0]
                try:
                    A.write_received_xlsx(dict(A.build_receiving_state("5002", TP9))); msgQu = ""
                except RuntimeError as e:
                    msgQu = str(e)
            finally:
                A.Path.iterdir = _orig_iterdir
            check("r4: unlistable SIBLING job folder -> resolver None with _resolve_unknown True, _report_exists None, picker \"can't check from this PC\", the refusal says 'PC that can open' (not 'no live work-order folder')",
                  rQu is None and stQu.get("_resolve_unknown") is True and hasQu is None
                  and rowQu["status"] == "can't check from this PC" and rowQu["tickable"] is False
                  and "PC that can open" in msgQu and "sibling" not in msgQu
                  and not (woQ900 / A.RECEIVED_FOLDER / outQ.name).exists())
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(hist3p, ignore_errors=True)

        # ── Build/Rebuild race: a close landing DURING the folder scan is in the file the
        #    picker writes (the walk now follows the scan) ──
        rows7b = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == "4007"])
        _orig_scan = A.scan_job_folders
        def _scan_then_close(state, progress=None):
            res = _orig_scan(state, progress)
            ev(hist9, "2026-09", "z_4007_midscan", close_ev("4007", "125-1031-970D", "16494.2",
               ["000098-1251031970-D00"], "2026-09-16 10:00:00", str(wo7)))
            return res
        A.scan_job_folders = _scan_then_close
        try:
            box7b = {}; A._rebuild_report_rows(rows7b, box7b)
        finally:
            A.scan_job_folders = _orig_scan
        check("r3: picker rebuild — a close landing during the folder scan is in the file it writes (walk after the scan, scan folded in)",
              box7b["ok"] == 1 and box7b["folders"] == 1 and box7b["failed"] == []
              and any(r[1] == "000098-1251031970-D00" for r in _rows_of(f7, A.COAT_SHEET)))
        # a refused write names WHY in the Rebuild-ALL summary (first line of the refusal)
        _orig_write = A.write_received_xlsx
        def _refuse(state):
            raise RuntimeError("the file contains 2 drawn object(s) (text boxes / shapes) that this app cannot "
                               "round-trip — saving would silently destroy it.\nClose/repair the file first:\n"
                               + str(f7))
        A.write_received_xlsx = _refuse
        try:
            box7c = {}; A._rebuild_report_rows(rows7b, box7c)
        finally:
            A.write_received_xlsx = _orig_write
        check("r3: Rebuild-ALL failure line = '<label>: RuntimeError: <first line of the refusal>' (not the bare type name)",
              box7c["ok"] == 0 and len(box7c["failed"]) == 1
              and box7c["failed"][0].startswith("125-1031-970D: RuntimeError: the file contains 2 drawn object(s)")
              and "\n" not in box7c["failed"][0] and len(box7c["failed"][0]) <= 200)
        # an old-flow job whose boards ALL came back (has_testing True, 0 at customer):
        # the OUT sheet writes no numbered row AND no '(none — …in-house)' row (BEFORE behaviour)
        hist3d = tempfile.mkdtemp(prefix="pcc_r3done_")
        cust3d = Path(tempfile.mkdtemp(prefix="pcc_r3donec_")) / "20 TestCust"; extra_roots9.append(cust3d.parent)
        woD = cust3d / "150-2200-900 4320 Rev. A" / "Production Data" / "16750.1"; woD.mkdir(parents=True)
        A.HISTORY_DIR = hist3d
        try:
            ev(hist3d, "2026-08", "a_close", close_ev("4320", "150-2200-900", "16750.1", [tL(i) for i in range(1, 4)],
               "2026-08-10 09:00:00", str(woD)))
            ev(hist3d, "2026-09", "b_recv_all", recv_ev("4320", "150-2200-900", [tL(i) for i in range(1, 4)],
               "2026-09-01 10:00:00", "Lata Gandhi", str(woD.parent)))
            stD = A.build_receiving_state("4320", "150-2200-900")
            outD = A.write_received_xlsx(stD)
            wbD = load_workbook(outD); outD_rows = list(wbD[A.OUT_SHEET].iter_rows(values_only=True))
            check("r3: finished test-and-return job (0 still at customer) -> OUT sheet: 0 numbered rows, NO '(none' row, subtitle '— 0 board(s)'; IH lists the 3",
                  A.receiving_summary(stD)["has_testing"] is True and A.receiving_summary(stD)["outstanding"] == 0
                  and not any(r and isinstance(r[0], int) for r in outD_rows)
                  and not any(r and isinstance(r[1], str) and r[1].startswith("(none") for r in outD_rows)
                  and wbD[A.OUT_SHEET]["A2"].value == "Engineering # 4320    ·    shipped for testing, NOT yet received back — 0 board(s)"
                  and sum(1 for r in wbD[A.IH_SHEET].iter_rows(values_only=True) if r and isinstance(r[0], int)) == 3)
            wbD.close()
            # the picker's 'Last shipment' column is the newest SHIPMENT, not the entry's newest event
            # (here a receive — review round 4)
            rowD = A._collect_report_rows({}, jobs=[j for j in A.list_testing_jobs() if j["eng"] == "4320"])[0]
            check("r4: picker 'Last shipment' = the close date (2026-08-10), not the newer receive the entry's ts carries",
                  rowD["job"]["ts"].startswith("2026-09-01") and rowD["last_ts"] == "2026-08-10"
                  and A._last_shipment_date(stD) == "2026-08-10" and A._last_shipment_date({"shipped": {}, "coated_log": []}) == ""
                  and A._last_shipment_date({"shipped": {"k": {"ts": "6/1/26"}}, "coated_log": [{"ts": "2026-05-02 09:00:00"}]}) == "2026-06-01")
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(hist3d, ignore_errors=True)

        # ── the window's scan path, headless: a walk_error state is NOT the records ──
        import types as _types
        def _fake_win(state):
            f = _types.SimpleNamespace()
            class _V:
                def __init__(s): s.v = ""
                def get(s): return s.v
                def set(s, v): s.v = v
            class _W:
                def __getattr__(s, n): return lambda *a, **k: None
            f._var = _V(); f._entry = _W(); f._lb = _W(); f._state = state
            f._approved = {}; f._session = []; f._mode = "pass"; f._issues = []
            f._jobs = [{"eng": state["eng"], "pn": state["pn"]}]; f._jobs_ready = True
            f._NOT_IN_RECORDS_MSG = RW._NOT_IN_RECORDS_MSG; f._NOTHING_MSG = RW._NOTHING_MSG
            f._COATED_HINT = RW._COATED_HINT; f._coated_only = RW._coated_only
            f._PAPERWORK_MSG = RW._PAPERWORK_MSG; f._shown_pn = lambda s: RW._shown_pn(s)
            f._current_job = lambda: {"eng": state["eng"], "pn": state["pn"]}
            f._sid = lambda s: RW._sid(f, s); f._carries_pn = lambda s: RW._carries_pn(f, s)
            f._dup_checks_due = lambda k: RW._dup_checks_due(f, k)
            f._pn_check_due = lambda k: RW._pn_check_due(f, k)
            f._ref = lambda: None
            f._issue = lambda serial, heading, message, approve=None: (f._issues.append(heading), False)[1]
            return f
        fbW = _fake_win({"eng": E9, "pn": TP9, "shipped": {}, "received": {}, "coated": {},
                         "pd_path": None, "wo_path": None, "walk_error": True})     # _start_state_build's fallback
        fbW._var.set(t9(1)); RW._add(fbW)
        goodW = _fake_win(dict(st9, walk_error=False))                            # a good walk: t9(1) never shipped
        goodW._approved = {}; goodW._var.set(t9(1)); RW._add(goodW)
        # (st9 is coated-only -> its stop comes first; a mixed state shows the records stop itself)
        mixW = _fake_win({"eng": "4400", "pn": "150-3300-900", "shipped": {A._serial_identity(tM(1), "150-3300-900"): {"ts": "x"}},
                          "received": {}, "coated": {}, "recv_dates": {}, "pd_path": None, "wo_path": None,
                          "walk_error": False, "testing_label_seen": True})
        mixW._var.set(tM(9)); RW._add(mixW)
        check("r3: scan into a walk_error state -> NO 'Not in shipping records' stop, the scan is kept as 'pending' (the Save sweep re-checks it after the re-pick)",
              fbW._issues == [] and [e["status"] for e in fbW._session] == ["pending"]
              and goodW._issues == ["Nothing to receive here"] and goodW._session == []
              and mixW._issues == ["Not in shipping records"] and mixW._session == [])
        # review round 4 (tk-ui-r4-2): overriding the coated-only stop must NOT wave through a board
        # that is already in-house — the duplicate stops are asked independently of that override.
        # Fixture: the 4006 shape with its Case-B board t9(569) currently received (decision (a):
        # the job is still coated-only), a supervisor who overrides every stop.
        kb9 = A._serial_identity(t9(569), TP9)
        dupW = _fake_win(dict(st9, walk_error=False,
                              received={kb9: {"serial": t9(569), "ts": "2026-09-10 09:00:00", "sts": "2026-09-10 09:00:00",
                                              "operator": "Anita Patel", "result": "pass", "note": ""}},
                              recv_dates={kb9: {"2026-09-10"}}))
        def _approve(serial, heading, message, approve=None):
            dupW._issues.append(heading)
            for s_ in (approve if approve is not None else [serial]):
                dupW._approved.setdefault(dupW._sid(s_), set()).add(heading)    # what the real _issue does
            return True
        dupW._issue = _approve
        check("r4: fixture — that Case-B board is currently received on a job that is still coated-only",
              RW._coated_only(dupW._state) is True and A._currently_received(dupW._state, kb9) is True)
        dupW._var.set(t9(569)); RW._add(dupW)                       # the received board again
        dupW._var.set(c9(900)); RW._add(dupW)                       # a fresh coated board (override)
        dupW._var.set(c9(900)); RW._add(dupW)                       # …and the same one twice
        dupW._var.set(c9(901)); RW._add(dupW)                       # another fresh one, scanned once
        check("r4: coated-only override does not clear the duplicate stops — 'Already received' still fires for the received board, 'Scanned twice' for a same-session repeat; a duplicate override still clears them as before",
              dupW._issues == ["Nothing to receive here", "Already received", "Nothing to receive here", "Scanned twice",
                               "Nothing to receive here"]
              and [e["serial"] for e in dupW._session] == [t9(569), c9(900), c9(900), c9(901)]
              and RW._dup_checks_due(dupW, kb9) is False                 # an 'Already received' override clears them (as before)
              and RW._dup_checks_due(dupW, dupW._sid(c9(900))) is False  # so does a 'Scanned twice' one
              and RW._dup_checks_due(dupW, dupW._sid(c9(901))) is True   # a coated-only override alone does not
              and RW._dup_checks_due(dupW, "never-scanned") is True)
        # review round 5 (coated-r5-1): the coated-only override answers ITS question only — a scan carrying
        # NEITHER P/N (a WO / PO barcode from the paperwork) still gets 'Doesn't match this job' (unchecked it
        # reached History and, being neither coated-form nor Case B, flipped the job to testing); the job's own
        # -970 boards and a -900-form one pass the P/N check; a refused second stop drops the scan; old flow unchanged
        pnW = _fake_win(dict(st9, walk_error=False))
        def _approve_all(serial, heading, message, approve=None):
            pnW._issues.append(heading); pnW._approved.setdefault(pnW._sid(serial), set()).add(heading); return True
        pnW._issue = _approve_all
        pnW._var.set("16493.2"); RW._add(pnW)            # WO barcode mis-scan: both stops, supervisor overrides both
        pnW._var.set(c9(902)); RW._add(pnW)              # the job's own coated board: the coated-only stop only
        pnW._var.set(t9(700)); RW._add(pnW)              # a -900-form board: the coated-only stop only
        refW = _fake_win(dict(st9, walk_error=False))
        def _first_only(serial, heading, message, approve=None):
            refW._issues.append(heading)
            if heading != "Nothing to receive here": return False
            refW._approved.setdefault(refW._sid(serial), set()).add(heading); return True
        refW._issue = _first_only
        refW._var.set("PO-778812"); RW._add(refW)
        mix2W = _fake_win(dict(mixW._state)); mix2W._var.set("ZZ-BADSCAN-01"); RW._add(mix2W)
        check("r5: coated-only job — a P/N-less scan gets 'Nothing to receive here' AND 'Doesn't match this job'; -970 / -900-form boards only the first; a refused P/N stop drops the scan; old-flow P/N stop as before",
              pnW._issues == ["Nothing to receive here", "Doesn't match this job", "Nothing to receive here", "Nothing to receive here"]
              and [e["serial"] for e in pnW._session] == ["16493.2", c9(902), t9(700)]
              and refW._issues == ["Nothing to receive here", "Doesn't match this job"] and refW._session == []
              and RW._pn_check_due(pnW, pnW._sid(c9(902))) is True and RW._pn_check_due(pnW, pnW._sid("16493.2")) is False
              and mix2W._issues == ["Doesn't match this job"] and mix2W._session == [])
        src3 = Path(SCRIPTS, "app_logic.py").read_text(encoding="utf-8")
        check("r3: wiring — the records stop is gated on a trustworthy walk; a walk_error scan stays pending",
              'elif not st.get("walk_error") and k not in st["shipped"] and not st["received"].get(k):' in src3
              and 'status = ("pending" if st is None or st.get("walk_error")' in src3)
        check("r4: wiring — per-reason approvals: the duplicate gates ask _dup_checks_due in _add (2), the paste (1) and the Save sweep (2) (+ _pn_check_due's own body); the P/N gate asks _pn_check_due in _add, the paste and the sweep (r5); approvals reset as a dict; the fold is the Tk-free helper",
              src3.count("self._dup_checks_due(k)") == 6 and src3.count("self._pn_check_due(k)") == 3
              and src3.count("self._approved = {}") == 3
              and 'self._approved.setdefault(k, set()).add(heading)' in src3
              and "self._approved = set()" not in src3
              and "_fold_receive_locally(st, kept, op, rdate, ts)" in src3)
        check("r4: wiring — Build? confirm and the tree-switch sentence are gated on a prior copy (found0); the no-wo_path chips/Save texts; _report_lookup's wo_path gate",
              'if st.get("wo_path") and st.get("_report_exists") is False and not found0:' in src3
              and "The report was placed in the job's {side} work-order folder" in src3
              and "No -900/-970 work-order folder is known for this job yet — close a shipment" in src3
              and "No -900/-970 work-order folder is known for this job (no live shipment record" in src3
              and "Report exists only under Archive — 🔁 Rebuild report re-homes it" in src3
              and 'if not state.get("wo_path"):\n        if state.get("_home_unknown"):\n            return None, None, None\n        return False, None, None' in src3   # r5: 'could not check' beats 'no WO'
              and "report archived — rebuild re-homes it" in src3)
        check("r3: wiring — Q2: _adopt_jobs lists EVERY entry (no plan-change-8 filter); the tagged 4006 label is in the dropdown list",
              'self._jobs = box.get("jobs") or []; self._jobs_ready = True' in src3
              and 'if j.get("testing", True)]' not in src3
              and "Eng 4006  ·  P/N 125-1030-970C  (in-house tested — coated only)" in [RW._job_label(j) for j in jobsP])
        check("r3: wiring — Rebuild ALL hands the PICKED rows (not all rows) to the worker; picker count line wraps",
              "threading.Thread(target=_rebuild_report_rows, args=(picked, box2), daemon=True).start()" in src3
              and "args=(rows, box2)" not in src3
              and 'wraplength=860, justify="left")' in src3)
        check("r3: wiring — 🔁 Rebuild walks AGAIN after the folder scan (and the Build? question) before writing",
              src3.count('self._rebuild_state_sync("Rebuilding from history…")') == 2
              and "if not self._rebuild_state_sync(\"Rebuilding from history…\"):\n            if scan and scan.get(\"ok\"):\n                fold_folder_scan(self._state, scan)" in src3)
        check("r3: wiring — chips: the coated-only line itself (not the docstring / A3) is pinned; Build/Rebuild button reset while loading + on a failed walk; loading line keeps the shown label",
              "P/N {lbl}     In-house tested — not shipped for testing" in src3
              and "Coated boards shipped (P/N {lbl}): {m['coated_all']}" in src3
              and src3.count('self._rebuild_btn.config(text="🔁 Rebuild report")') == 2
              and src3.count("label=lbl)") == 3
              and src3.count("lbl = self._shown_pn(self._state) if self._state else _job_display_pn(job)") == 2   # r5: never the -900 key
              and "lbl=rw._shown_pn(rw._state) if rw._state else _job_display_pn(job)" in src3
              and "if self._state else None" not in src3 and "if rw._state else None" not in src3
              and src3.count("label=_job_display_pn(j))") == 3)       # r4: the two auto-detect scans + the History void refresh too
        check("r3: wiring — 📄 Open file refuses a walk_error state with 'History incomplete' (never 'no work-order folder')",
              '"This job\'s history could not be read — re-pick the job to retry, then open "' in src3
              and src3.index('"This job\'s history could not be read — re-pick the job to retry, then open "')
                  < src3.index("No -900 or -970 work-order folder is known for this job yet"))
        check("r3: wiring — BackfillDialog: _typed_coated initialised/reset like _coated_scan; confirm header + Imported popup follow the typed side",
              src3.count("self._typed_coated = False") == 2 and 'getattr(self, "_typed_coated"' not in src3
              and "{'Coated' if typed_coated else 'Testing'} P/N {pn}" in src3
              and '("COATED shipments (this P/N is the coated side)." if typed_coated' in src3
              and 'else "the shipped-for-testing baseline.")' in src3)
        check("r3: wiring — one refusal head shared by both write_received_xlsx refusals; _find_legacy ranks live over archived; _report_lookup skips other jobs' copies",
              src3.count("the job's work-order folder could not be found on the share — the ") == 1
              and "key = (0 if _under_archive(c) else 1, c.stat().st_mtime)" in src3
              and "if _copy_names_other_job(c, state.get(\"eng\")):\n                # another job's" in src3
              and "held = _coated_tree_holds_report(cw or cbf, state)\n        if held:" in src3     # r4/r5: resolved like the write path, tri-state
              and 'elif held is None:\n            state["_home_unknown"] = True' in src3
              and "held, sib_unreadable = _sibling_wo_holding_report(cust, ek, second, state)" in src3   # r4/r5: the WO that HOLDS it, for every caller
              and src3.index("held, sib_unreadable = _sibling_wo_holding_report(cust, ek, second, state)")
                  < src3.index("best, unreadable = _pick_report_wo(cust, ek, second, want_wo)")        # r5: the newest-WO pick only after the holder search
              and src3.count("_pick_report_wo(cust, ek, second, want_wo)") == 1
              and 'if st.get("_report_exists") is None:\n            messagebox.showwarning("Can\'t check",' in src3   # r5: Rebuild refuses a None answer up front
              and "if legacy is None and state.get(\"_cand_partial\"):" in src3                         # r5: a first-time build refuses on a partial enumeration
              and "return best if _wo_holds_report(best, state, whole_tree=True) else None" not in src3)
        check("r3: wiring — list_testing_jobs: the coated label shown follows the newest CLOSE over an import's spelling",
              'coat_jobs[ck]["coated_label"] = label; coat_jobs[ck]["_from_close"] = True' in src3
              and not any("_from_close" in j for j in A.list_testing_jobs()))

        # ══ review round 6 ═══════════════════════════════════════════════════
        src6 = Path(SCRIPTS, "app_logic.py").read_text(encoding="utf-8")
        # ── datasafety-r6-1: the HOME tree's pass ranked live WOs by mtime whenever the recorded WO FOLDER was dead
        #    (renamed / archived WO folder — not a job-folder rev rename, which keeps the WO name): the merge's rmdir of
        #    the emptied leaf made the WO the file just LEFT the newest, so every daemon run / receiving Save moved the
        #    one file into the other WO (round 5 fixed the same ping-pong for the sibling tree only). The holder wins. ──
        hist6 = tempfile.mkdtemp(prefix="pcc_r6pp_")
        cust6 = Path(tempfile.mkdtemp(prefix="pcc_r6ppc_")) / "20 TestCust"; extra_roots9.append(cust6.parent)
        _files6 = lambda eng_: sorted(p for p in cust6.rglob("*Shipping Report*.xlsx") if f" {eng_} " in str(p.relative_to(cust6)))
        A.HISTORY_DIR = hist6
        try:
            outs6 = {}
            for eng6, pn6, jd6, ser6 in (("4044", "136-9630-900", "136-9630-900 4044 Rev. A",
                                          lambda a, b: [f"{i:06d}-1369630900-A00" for i in range(a, b)]),
                                         ("4009", "125-1033-970C", "125-1033-970C 4009 Rev. C",
                                          lambda a, b: [f"{i:06d}-1251033970-C00" for i in range(a, b)])):
                pd6 = cust6 / jd6 / "Production Data"
                woA6, woB6, woC6 = pd6 / "18001.1", pd6 / "18002.1", pd6 / "18003.1"
                (woA6 / A.COC_FOLDER).mkdir(parents=True); (woB6 / A.COC_FOLDER).mkdir(parents=True)
                (pd6 / "18003.1x").mkdir()                             # the newest close's WO folder, renamed since
                ev(hist6, "2026-09", f"a_{eng6}", close_ev(eng6, pn6, "18001.1", ser6(1, 6), "2026-09-01 10:00:00", str(woA6)))
                j6 = next(j for j in A.list_testing_jobs() if j["eng"] == eng6)
                st6 = A.build_receiving_state(eng6, j6["pn"]); st6["allow_tree_switch"] = True
                out6 = Path(A.write_received_xlsx(st6)); outs6[eng6] = out6          # 🔁 Build report -> 18001.1
                ev(hist6, "2026-09", f"b_{eng6}", close_ev(eng6, pn6, "18003.1", ser6(6, 9), "2026-09-10 10:00:00", str(woC6)))
                for p6, secs in ((pd6 / "18003.1x", -300), (woA6, -200), (woB6, -100)):   # 18002.1 = newest by mtime
                    t6 = time.time() + secs; os.utime(str(p6), (t6, t6))
                paths6 = []; sw6 = []
                for _ in range(3):
                    st6 = A.build_receiving_state(eng6, j6["pn"]); st6.pop("allow_tree_switch", None)   # daemon / Save shape
                    paths6.append(Path(A.write_received_xlsx(st6))); sw6.append(st6.get("tree_switched"))
                st6a = A._annotate_report_exists(A.build_receiving_state(eng6, j6["pn"]))
                check(f"r6: {eng6} — recorded WO folder dead + two live WOs + the report in the OLDER one: three daemon/Save-shaped writes land on the SAME file (the holder), tree_switched False, one file, no leaf in the newer WO, cache at_home True",
                      out6.parent.parent == woA6 and paths6 == [out6] * 3 and sw6 == [False] * 3
                      and _files6(eng6) == [out6] and not (woB6 / A.RECEIVED_FOLDER).exists()
                      and st6a["_report_exists"] is True and st6a["_report_at_home"] is True and st6a["_report_path"] == out6)
            # ── fresh-r6-1: a user edit to the app-owned A2 head line is NOT another job's file ──
            out44 = outs6["4044"]
            wb44 = load_workbook(out44); wb44[A.RCV_SHEET]["A2"] = "Engineering # 4044 — checked by JD"; wb44.save(str(out44)); wb44.close()
            eng44 = A._report_copy_eng_at(out44)
            st44 = A._annotate_report_exists(A.build_receiving_state("4044", "136-9630-900"))
            gate44 = A._receive_save_gate(st44)
            row44 = A._collect_report_rows({}, jobs=[{"eng": "4044", "pn": "136-9630-900", "ts": "2026-09-10"}])[0]
            out44b = Path(A.write_received_xlsx(dict(st44)))
            wb44 = load_workbook(out44b); a2_44 = wb44[A.RCV_SHEET]["A2"].value; wb44.close()
            check("r6: a user edit to the A2 head line ('Engineering # 4044 — checked by JD') keeps the file THIS job's report: Eng# read '4044', _report_exists True, save gate True, picker 'has report', the write updates the SAME file and rewrites the line, no second file",
                  eng44 == "4044" and st44["_report_exists"] is True and gate44 is True and row44["status"] == "has report"
                  and out44b == out44 and str(a2_44).startswith("Engineering # 4044    ·    ") and _files6("4044") == [out44])
            check("r6: the head-line Eng# is ONE whole token, shape-checked — 4016 vs 4016LF still differ; text after the Eng# is ignored; a token that is no Eng# (or no head line) -> '' (decided on other evidence); one parse for both gates",
                  A._eng_from_head_line("Engineering # 4016LF    ·    received back") == "4016LF"
                  and A._eng_from_head_line("Engineering # 4016    ·    x") == "4016"
                  and A._eng_from_head_line("Engineering # 4041 (old rev)    ·    x") == "4041"
                  and A._eng_from_head_line("Engineering # 4041 - received back from testing") == "4041"
                  and A._eng_from_head_line("Engineering # 4041—checked") == ""
                  and A._eng_from_head_line("Received from testing — 4041") == "" and A._eng_from_head_line(None) == ""
                  and "    return _eng_from_head_line(_report_head_line_at(path))\n" in src6     # r7: one parse (_head_eng_tokens) behind both readers
                  and "    return _eng_from_head_line(_report_head_line(wb))\n" in src6
                  and src6.count(r"re.match(r'\s*Engineering\s*#\s*") == 1)
            # ── datasafety-r6-2 / tk-ui-r6-1: another job's file at OUR exact target while our own copy sits in the
            #    previous WO (migrate pending): the switch names OUR copy, never the foreign file; at_home from the
            #    lookup's evidence; the blocker recorded for the chips; every writer still refuses ──
            woD6 = cust6 / "136-9630-900 4044 Rev. A" / "Production Data" / "18004.1"; (woD6 / A.COC_FOLDER).mkdir(parents=True)
            ev(hist6, "2026-09", "c_4044", close_ev("4044", "136-9630-900", "18004.1", [f"{i:06d}-1369630900-A00" for i in range(9, 11)],
               "2026-09-12 10:00:00", str(woD6)))
            tgt44 = A.received_xlsx_path(woD6, "136-9630-900")
            stM6 = A._annotate_report_exists(A.build_receiving_state("4044", "136-9630-900"))      # migrate pending
            tgt44.parent.mkdir(parents=True); shutil.copy(str(outs6["4009"]), str(tgt44)); sigF6 = A._file_sig(tgt44)
            stF6 = A._annotate_report_exists(A.build_receiving_state("4044", "136-9630-900"))
            rowF6 = A._collect_report_rows({}, jobs=[{"eng": "4044", "pn": "136-9630-900", "ts": "2026-09-12"}])[0]
            try:
                A.write_received_xlsx(dict(stF6)); msgF6 = ""
            except RuntimeError as e:
                msgF6 = str(e)
            A.auto_update_shipping_report("4044", "136-9630-900")
            check("r6: migrate pending (before the plant): the answer rests on the 18001.1 copy, path = the 18004.1 target, at_home False, no blocker",
                  stM6["_report_exists"] is True and stM6["_report_at_home"] is False and stM6["_report_path"] == tgt44
                  and stM6["_report_found_at"] == out44 and "_report_blocked_by" not in stM6)
            check("r6: another job's file (Eng# 4009) AT our target + our own copy in the previous WO -> path = OUR copy (its head line names 4044), at_home True, blocker (target, '4009') cached, the Save write refuses naming 4009, the daemon writes nothing, the foreign file byte-identical, picker 'has report' with our copy as its path",
                  stF6["_report_exists"] is True and stF6["_report_path"] == out44 and stF6["_report_found_at"] == out44
                  and A._report_copy_eng_at(stF6["_report_path"]) == "4044" and stF6["_report_at_home"] is True
                  and stF6.get("_report_blocked_by") == (tgt44, "4009")
                  and "names another job (Engineering # 4009)" in msgF6 and A._file_sig(tgt44) == sigF6 and out44.exists()
                  and rowF6["status"] == "has report" and rowF6["path"] == out44)
            stU6 = {"eng": "4044", "pn": "136-9630-900", "wo_path": str(woD6)}
            _orig_lookup6 = A._report_lookup
            A._report_lookup = lambda s: (True, tgt44, out44)      # an EXISTING path that is not the copy the answer rests on
            try: A._annotate_report_exists(stU6)
            finally: A._report_lookup = _orig_lookup6
            check("r6: _report_at_home is derived from the lookup's evidence (found == path), never from a stat of the path; the chips name the blocker; the blocker is a cache key (forgotten / cleared with the others)",
                  stU6["_report_at_home"] is False and stU6["_report_path"] == tgt44
                  and "report location: {bp} — move it away; every update is refused until then." in src6
              and "sits at this job's " in src6 and 'blk = st.get("_report_blocked_by")' in src6
                  and "_report_blocked_by" in A._REPORT_CACHE_KEYS
                  and 'state.pop("_report_blocked_by", None)   # a write succeeded' in src6)
            tgt44.unlink(); tgt44.parent.rmdir()
            A.auto_update_shipping_report("4044", "136-9630-900")
            check("r6: the foreign file moved away -> the next after-close update migrates OUR copy to the target (one file, in 18004.1)",
                  _files6("4044") == [tgt44] and not out44.exists())
            # ══ review round 7 (on the same fixture) ═════════════════════════════════════════════
            # ── test-adequacy-r7-3: a head line whose FIRST token is no Eng# shape, AT the target: undecidable ->
            #    the file at the target is this job's report (the switch stays True, the daemon keeps appending) ──
            wb44 = load_workbook(tgt44); wb44[A.RCV_SHEET]["A2"] = "Engineering # — 4044 checked by JD"; wb44.save(str(tgt44)); wb44.close()
            eng44b = A._report_copy_eng_at(tgt44)
            st44e = A._annotate_report_exists(A.build_receiving_state("4044", "136-9630-900"))
            row44e = A._collect_report_rows({}, jobs=[{"eng": "4044", "pn": "136-9630-900", "ts": "2026-09-12"}])[0]
            n44e = _nrows(tgt44, A.OUT_SHEET)
            ev(hist6, "2026-09", "d_4044", close_ev("4044", "136-9630-900", "18004.1", [f"{i:06d}-1369630900-A00" for i in range(11, 14)],
               "2026-09-13 10:00:00", str(woD6)))
            A.auto_update_shipping_report("4044", "136-9630-900")
            check("r7: A2 edited to 'Engineering # — 4044 checked by JD' (first token no Eng# shape) AT the target: Eng# reads '', _report_exists True, at_home True, no blocker, save gate True, picker 'has report', the daemon appends to the SAME file (+3 rows)",
                  eng44b == "" and st44e["_report_exists"] is True and st44e["_report_at_home"] is True
                  and "_report_blocked_by" not in st44e and A._receive_save_gate(st44e) is True and row44e["status"] == "has report"
                  and _files6("4044") == [tgt44] and _nrows(tgt44, A.OUT_SHEET) == n44e + 3)
            # ── optin-r7-1 / tk-ui-r7-1 / tk-ui-r7-2: another job's file at OUR target and NO copy of our own ──
            hold44 = tgt44.parent / "hold.bak"; tgt44.rename(hold44)          # our own report out of the way
            shutil.copy(str(outs6["4009"]), str(tgt44)); sigB6 = A._file_sig(tgt44)
            stB6 = A._annotate_report_exists(A.build_receiving_state("4044", "136-9630-900"))
            rowB6 = A._collect_report_rows({}, jobs=[{"eng": "4044", "pn": "136-9630-900", "ts": "2026-09-13"}])[0]
            class _Cfg7:
                def __init__(s): s.kw = {}
                def config(s, **k): s.kw.update(k)
            chW7 = _types.SimpleNamespace(_state=stB6, _chips=_Cfg7(), _rebuild_btn=_Cfg7())
            RW._refresh_chips(chW7)
            chips7 = str(chW7._chips.kw.get("text", ""))
            opened7 = []; warns7 = []
            _origSF7, _origMB7 = A.os.startfile, A.messagebox
            A.os.startfile = lambda p: opened7.append(p)
            A.messagebox = _types.SimpleNamespace(showwarning=lambda *a, **k: warns7.append(("w", a[1] if len(a) > 1 else str(k.get("message")))),
                                                  showinfo=lambda *a, **k: warns7.append(("i", a[1] if len(a) > 1 else str(k.get("message")))),
                                                  showerror=lambda *a, **k: warns7.append(("e", a[1] if len(a) > 1 else str(k.get("message")))))
            try:
                RW._open_file(_types.SimpleNamespace(_state=stB6))
            finally:
                A.os.startfile, A.messagebox = _origSF7, _origMB7
            A.auto_update_shipping_report("4044", "136-9630-900")
            check("r7: another job's file (Eng# 4009) at our target, no own copy: _report_exists False with path None (never the foreign file), blocker (target, '4009') cached, save gate False, picker row 'another job's file at its report location' (not tickable), the daemon writes nothing, the foreign file byte-identical",
                  stB6["_report_exists"] is False and stB6["_report_path"] is None and stB6["_report_found_at"] is None
                  and stB6.get("_report_blocked_by") == (tgt44, "4009") and A._receive_save_gate(stB6) is False
                  and rowB6["status"] == "another job's file at its report location" and rowB6["tickable"] is False
                  and A._file_sig(tgt44) == sigB6)
            check("r7: chips for that state name the blocker and never promise 'Build report creates it'; 📄 Open file refuses with one warning naming Eng# 4009 and never calls os.startfile",
                  "sits at its report location" in chips7 and "Engineering # 4009" in chips7 and "Build report creates it" not in chips7
                  and "No -900/-970 work-order folder" not in chips7
                  and opened7 == [] and len(warns7) == 1 and warns7[0][0] == "w" and "Engineering # 4009" in warns7[0][1]
                  and str(tgt44) in warns7[0][1])
            tgt44.unlink(); hold44.rename(tgt44)                               # our own report back in place
            stR6 = A._annotate_report_exists(A.build_receiving_state("4044", "136-9630-900"))
            check("r7: the foreign file gone, our copy back: the switch is True again at home, no blocker",
                  stR6["_report_exists"] is True and stR6["_report_path"] == tgt44 and "_report_blocked_by" not in stR6)
            # ── test-adequacy-r7-2: the recorded WO by NAME wins under a RENAMED job folder (the fast path's twin),
            #    even though another WO of the home tree holds the report: 4009's newest close names 18003.1, which
            #    now exists again; the job folder goes 'Rev. C' -> 'Rev. D' (the recorded string is stale) ──
            pdC9 = cust6 / "125-1033-970C 4009 Rev. C" / "Production Data"
            (pdC9 / "18003.1" / A.COC_FOLDER).mkdir(parents=True)             # the recorded WO, live again by name
            jdD9 = cust6 / "125-1033-970C 4009 Rev. D"; pdC9.parent.rename(jdD9)
            st09 = A.build_receiving_state("4009", "125-1033-900"); st09.pop("allow_tree_switch", None)
            r09 = A._resolve_received_wo_dir(dict(st09))
            out09 = Path(A.write_received_xlsx(st09))                         # daemon / Save shape
            check("r7: job folder renamed ('Rev. C' -> 'Rev. D'), recorded WO 18003.1 live by name while 18001.1 HOLDS the report: the resolver picks the recorded WO under the renamed folder, a daemon-shaped write migrates the one file there, tree_switched False, one file",
                  r09 == jdD9 / "Production Data" / "18003.1" and out09 == A.received_xlsx_path(jdD9 / "Production Data" / "18003.1", "125-1033-970C")
                  and st09.get("tree_switched") is False and _files6("4009") == [out09]
                  and not (jdD9 / "Production Data" / "18001.1" / A.RECEIVED_FOLDER).exists())
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(hist6, ignore_errors=True)

        # ── oldflow-r6-1: overriding the 'Not in shipping records' stop answers only ITS question — the same-session
        #    'Scanned twice' stop stays live for that board (BEFORE this stop existed such a board had no stop at all, so
        #    its duplicate checks were live; one records override must not put the board into the receive event twice) ──
        recW = _fake_win(dict(mixW._state))
        def _approve_records(serial, heading, message, approve=None):
            recW._issues.append(heading)
            if heading != "Not in shipping records": return False
            recW._approved.setdefault(recW._sid(serial), set()).add(heading); return True
        recW._issue = _approve_records
        recW._var.set(tM(9)); RW._add(recW)          # not on record -> records stop, the supervisor overrides
        recW._var.set(tM(9)); RW._add(recW)          # scanner double-fire: 'Scanned twice' (refused -> dropped)
        check("r6: old-flow scan — after a 'Not in shipping records' override the same board scanned again still gets 'Scanned twice'; the session holds it once; the dup / P/N checks stay due after a records-only override",
              recW._issues == ["Not in shipping records", "Scanned twice"]
              and [e["serial"] for e in recW._session] == [tM(9)]
              and RW._dup_checks_due(recW, recW._sid(tM(9))) is True and RW._pn_check_due(recW, recW._sid(tM(9))) is True)
        # …and the Save sweep (two 'pending' scans of that board taken before the history landed), driven headless with
        # the confirm dialog stubbed to Cancel so nothing is written — the confirm texts (req-r6-1) are pinned here too
        cap6 = {}
        class _FakeConfirm6:
            def __init__(self, *a, **k): cap6["msg"] = k.get("message"); self.result = False
        _origC6, _origMB6 = A.ConfirmActionDialog, A.messagebox
        A.ConfirmActionDialog = _FakeConfirm6
        A.messagebox = _types.SimpleNamespace(showinfo=lambda *a, **k: None, showwarning=lambda *a, **k: None,
                                              showerror=lambda *a, **k: None)
        def _save_win(state, session, issue_fn):
            w = _fake_win(state)
            w._session = session; w._sess_overrides = []; w._wait_state = lambda: True
            w._rdate = type(w._var)(); w._notes = w._entry
            opv = type(w._var)(); opv.set("Anita Patel"); w._app = _types.SimpleNamespace(operator_var=opv)
            w._shown_pn = lambda s: RW._shown_pn(s)
            w._issue = lambda serial, heading, message, approve=None: issue_fn(w, serial, heading)
            return w
        try:
            def _rec_only(w, serial, heading):
                w._issues.append(heading)
                if heading != "Not in shipping records": return False
                w._approved.setdefault(w._sid(serial), set()).add(heading); return True
            swW = _save_win(dict(mixW._state), [{"serial": tM(9), "status": "pending", "result": "pass"},
                                                {"serial": tM(9), "status": "pending", "result": "pass"}], _rec_only)
            RW._do_save(swW); msg_old = str(cap6.get("msg") or "")
            coW = _save_win(dict(st9, walk_error=False), [{"serial": c9(905), "status": "not_shipped", "result": "pass"}],
                            lambda w, s, h: (w._issues.append(h), False)[1])
            coW._approved = {coW._sid(c9(905)): {"Nothing to receive here"}}      # the scan's own override
            RW._do_save(coW); msg_co = str(cap6.get("msg") or "")
        finally:
            A.ConfirmActionDialog, A.messagebox = _origC6, _origMB6
        check("r6: Save sweep — two 'pending' scans of one off-record board: the records override, then 'Scanned twice' (refused -> dropped); the confirm reads 'Record 1 board(s) as RECEIVED back from testing?' (old-flow text byte-identical)",
              swW._issues == ["Not in shipping records", "Scanned twice"] and [e["serial"] for e in swW._session] == [tM(9)]
              and msg_old.startswith("Record 1 board(s) as RECEIVED back from testing?\n(1 PASSED  ·  0 FAILED)\n\n"
                                     "Eng # 4400  ·  P/N 150-3300-900\nOperator: Anita Patel"))
        check("r6: the Save confirm on an in-house-tested job (override path) says the boards were NOT shipped for testing and names the coated label",
              coW._issues == [] and [e["serial"] for e in coW._session] == [c9(905)]
              and msg_co.startswith("Record 1 board(s) as RECEIVED on this in-house-tested job\n"
                                    "(supervisor override — these boards were not shipped for testing)?\n"
                                    "(1 PASSED  ·  0 FAILED)\n\nEng # 4006  ·  P/N 125-1030-970C"))
        # ── test-r6-5: the paste path's gates by BEHAVIOUR (polarity), not a call count ──
        shown6 = []
        A.messagebox = _types.SimpleNamespace(showinfo=lambda *a, **k: shown6.append(a[1] if len(a) > 1 else str(k.get("message"))),
                                              showwarning=lambda *a, **k: None, showerror=lambda *a, **k: None)
        try:
            pmW = _fake_win(dict(mixW._state)); pmW._wait_state = lambda: True
            RW._add_many(pmW, ["ZZ-BADSCAN-01", tM(1), tM(1)])
            pcW = _fake_win(dict(dupW._state)); pcW._wait_state = lambda: True          # coated-only, t9(569) currently received
            def _approve_paste(serial, heading, message, approve=None):
                pcW._issues.append(heading)
                for s_ in (approve if approve is not None else [serial]):
                    pcW._approved.setdefault(pcW._sid(s_), set()).add(heading)
                return True
            pcW._issue = _approve_paste
            RW._add_many(pcW, [c9(902), c9(903), "16493.2", t9(569)])
        finally:
            A.messagebox = _origMB6
        check("r6: paste — old flow: the garbage line is skipped as 'don't carry this job's P/N', the on-record board added once (its repeat 'already in this session'); coated-only: ONE 'Nothing to receive here' stop for the list, the job's -970 boards added, the WO barcode skipped as wrong P/N, the currently-received Case-B board skipped as already received",
              [e["serial"] for e in pmW._session] == [tM(1)] and len(shown6) == 2
              and "1 skipped — don't carry this job's P/N (150-3300-900)" in shown6[0] and "1 skipped — already in this session's list" in shown6[0]
              and pcW._issues == ["Nothing to receive here"] and [e["serial"] for e in pcW._session] == [c9(902), c9(903)]
              and "1 skipped — don't carry this job's P/N (125-1030-970C)" in shown6[1]     # r7: the label the operator sees (req-r7-1)
              and "1 skipped — already received before" in shown6[1])
        # ── coated-r6-1 / quality-r6-1: ONE derivation of the serial embed keys everywhere. A job whose only -900
        #    spelling carries a rev letter ('134-2201-900A' — the shape round 4 made a listed entry) must be the same
        #    job for the walk, the resolver, the candidate scan, the folder scan, the local fold and the P/N check ──
        check("r6: _tr_embed_keys — digit-only keys for plain, lettered and -970 labels; a label with no core keeps _alnum",
              A._tr_embed_keys("136-9624-900") == ("1369624900", "1369624970")
              and A._tr_embed_keys("134-2201-900A") == ("1342201900", "1342201970")
              and A._tr_embed_keys("125-1030-900C") == ("1251030900", "1251030970")
              and A._tr_embed_keys("125-1030-970C") == ("1251030970", "")
              and A._tr_embed_keys("555-000-900") == ("555000900", "555000970")
              and A._tr_embed_keys("N/A") == (A._alnum("N/A"), "") and A._tr_embed_keys("") == ("", "") and A._tr_embed_keys(None) == ("", ""))
        histL = tempfile.mkdtemp(prefix="pcc_r6letter_")
        custL6 = Path(tempfile.mkdtemp(prefix="pcc_r6letterc_")) / "377 Anduril"; extra_roots9.append(custL6.parent)
        woL9 = custL6 / "134-2201-900A 4321 Rev. A" / "Production Data" / "18101.1"
        woL7 = custL6 / "134-2201-970 4321 Rev. A" / "Production Data" / "18102.1"
        (woL9 / A.COC_FOLDER).mkdir(parents=True); (woL7 / A.COC_FOLDER).mkdir(parents=True)
        tLA = lambda i: f"{i:06d}-1342201900-A00"; tLB = lambda i: f"{i:06d}-1342201900-B00"; cL6 = lambda i: f"{i:06d}-1342201970-A00"
        _shipfile(woL9 / A.COC_FOLDER, "coc-900A.xlsx", [tLA(1), tLA(2), tLB(3)], "2026-08-27", pn_label="134-2201-900A")
        _shipfile(woL7 / A.COC_FOLDER, "coc-970.xlsx", [cL6(1), cL6(2), cL6(3), tLB(3)], "2026-08-28", pn_label="134-2201-970")  # the rev-B board re-shipped
        A.HISTORY_DIR = histL
        try:
            ev(histL, "2026-08", "a_900A", close_ev("4321", "134-2201-900A", "18101.1", [tLA(1), tLA(2), tLB(3)], "2026-08-27 09:00:00", str(woL9)))
            ev(histL, "2026-08", "b_970", close_ev("4321", "134-2201-970", "18102.1", [cL6(1), cL6(2), cL6(3)], "2026-08-28 09:00:00", str(woL7)))
            entL = [j["pn"] for j in A.list_testing_jobs() if j["eng"] == "4321"]
            stL = A.build_receiving_state("4321", "134-2201-900A")
            nsL = _types.SimpleNamespace(_state=stL, _coated_only=RW._coated_only)
            scL = A.scan_job_folders(stL)
            candL = A._stale_received_candidates(dict(stL), A.received_xlsx_path(str(woL9), A._report_pn(stL)))
            stLf = dict(stL, received={}, receive_log=[], recv_dates={})
            A._fold_receive_locally(stLf, [{"serial": cL6(1), "result": "pass"}], "op", "", "2026-09-23 10:00:00")
            shutil.rmtree(woL9.parent.parent)                        # the -900 tree gone: only the -970 tree is live
            stLr = dict(stL, allow_tree_switch=True); rL = A._resolve_received_wo_dir(stLr)
            kB3 = A._serial_identity(tLB(3), stL["pn"])
            check("r6: an only-lettered -900 key ('134-2201-900A') is the SAME job everywhere: entry keeps that spelling; _carries_pn True for its rev-B board (it IS in shipped), False for a WO barcode; the folder scan opens BOTH trees, counts the re-shipped rev-B board 2x and lists the 3 coated rows; the -970 tree contributes candidates; a coated-form fold stays out of `received`; the resolver reaches the -970 WO with allow_tree_switch",
                  entL == ["134-2201-900A"]
                  and kB3 in stL["shipped"] and RW._carries_pn(nsL, tLB(3)) is True and RW._carries_pn(nsL, "18101.1") is False
                  and scL["ok"] and sorted(Path(p).name for p in scL["jobs"]) == [woL9.parent.parent.name, woL7.parent.parent.name]
                  and scL["ship_counts"].get(kB3) == 2
                  and [r["serial"] for r in scL["coated_rows"]] == [cL6(1), cL6(2), cL6(3)]
                  and any(str(woL7.parent.parent) in str(c) for c in candL)
                  and stLf["received"] == {} and len(stLf["receive_log"]) == 1
                  and rL == woL7 and stLr.get("tree_switched") is True)
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histL, ignore_errors=True)
        check("r6: wiring — ONE embed-key derivation (_tr_embed_keys: definition + the walk, resolver, candidate scan, local fold, folder scan, P/N check); the letter-blind form survives only inside the helper; the home pass is holder-first when the recorded WO is dead; the records override is exempt from the dup-check waiver; RCV Case-B cross-reference; BackfillDialog resets before its early return",
              src6.count("_tr_embed_keys(") == 8 and src6.count("_alnum(testing_to_coated_pn(") == 2   # r7: + _coated_tree_holds_report
              and "held, h_unreadable = _sibling_wo_holding_report(cust, ek, first, state)" in src6
              and src6.index("held, h_unreadable = _sibling_wo_holding_report(cust, ek, first, state)")
                  < src6.index("held, sib_unreadable = _sibling_wo_holding_report(cust, ek, second, state)")
              and '- {"Nothing to receive here", "Not in shipping records", "Coated shipment on paperwork only"})' in src6
              and "(coated, -900 label)" in src6
              and src6.count("self._coated_scan = False; self._typed_coated = False") == 1
              and src6.index("self._coated_scan = False; self._typed_coated = False")
                  < src6.index('"Fill in the Testing P/N, then find the job folder."'))

        # ══ review round 7 ═══════════════════════════════════════════════════
        import re as _re7
        # ── datasafety-r7-1 / oldflow-r7-1 / test-adequacy-r7-1: an Eng# with SEVERAL tokens — the folder parser's
        #    own output for a bracket-style job name — owns its report. Round 6 read one token off the head line
        #    ('4044') and compared it with '4044 LF': the job's OWN report was 'another job's file' for good. ──
        hist7 = tempfile.mkdtemp(prefix="pcc_r7eng_")
        cust7 = Path(tempfile.mkdtemp(prefix="pcc_r7engc_")) / "20 TestCust"; extra_roots9.append(cust7.parent)
        pd7 = cust7 / "136-9630-900 GizmoX [4044 LF Rev. A]" / "Production Data"
        wo7a, wo7b = pd7 / "18001.1", pd7 / "18002.1"
        (wo7a / A.COC_FOLDER).mkdir(parents=True); (wo7b / A.COC_FOLDER).mkdir(parents=True)
        eng7 = A.parse_wo_path(wo7a)["eng"]                                    # what a close records for this folder
        s7 = lambda a, b: [f"{i:06d}-1369630900-A00" for i in range(a, b)]
        A.HISTORY_DIR = hist7
        try:
            ev(hist7, "2026-09", "a_lf", close_ev(eng7, "136-9630-900", "18001.1", s7(1, 6), "2026-09-01 10:00:00", str(wo7a)))
            j7 = next(j for j in A.list_testing_jobs() if j["eng"] == eng7)
            st7 = A.build_receiving_state(eng7, j7["pn"]); st7["allow_tree_switch"] = True
            out7 = Path(A.write_received_xlsx(st7))                               # 🔁 Build report
            a2_7 = A._report_head_line_at(out7)
            st7a = A._annotate_report_exists(A.build_receiving_state(eng7, j7["pn"]))
            row7 = A._collect_report_rows({}, jobs=[j7])[0]
            try:
                out7b = Path(A.write_received_xlsx(dict(st7a)))                   # daemon / Save shape: an UPDATE
            except RuntimeError as e:
                out7b = str(e)[:120]                     # (round 6 refused here: 'names another job (Engineering # 4044)')
            ext7 = A._extract_received_rows(out7, eng7)
            ev(hist7, "2026-09", "b_lf", close_ev(eng7, "136-9630-900", "18002.1", s7(6, 9), "2026-09-10 10:00:00", str(wo7b)))
            A.auto_update_shipping_report(eng7, "136-9630-900")
            files7 = sorted(cust7.rglob("*Shipping Report*.xlsx"))
            check("r7: a MULTI-TOKEN Eng# ('4044 LF', parsed from '… [4044 LF Rev. A]') owns its report: the head line names it, _copy_names_other_job False, _report_exists True at home with no blocker, save gate True, picker 'has report', a daemon-shaped write updates the SAME file, the copy harvests, the next WO's update MIGRATES the one file (no fork)",
                  eng7 == "4044 LF" and str(a2_7).startswith("Engineering # 4044 LF    ·    ")
                  and A._head_names_eng(a2_7, eng7) is True and A._copy_names_other_job(out7, eng7) is False
                  and st7a["_report_exists"] is True and st7a["_report_at_home"] is True and "_report_blocked_by" not in st7a
                  and A._receive_save_gate(st7a) is True and row7["status"] == "has report"
                  and out7b == out7 and ext7 is not None
                  and files7 == [A.received_xlsx_path(wo7b, "136-9630-900")] and _nrows(files7[0], A.OUT_SHEET) == 8)
            check("r7: _head_names_eng — whole TOKENS, the job's Eng# tokens must LEAD the segment, tri-state: multi-token ours (with / without a user note), one-token file is not the two-token job's, 4016 / 4016LF differ both ways, a user note after the Eng# is ignored, a non-Eng# first token / no head line / blank Eng# -> None",
                  A._head_names_eng("Engineering # 4044 LF    ·    x", "4044 LF") is True
                  and A._head_names_eng("Engineering # 4044 LF — checked by JD", "4044 LF") is True
                  and A._head_names_eng("Engineering # 4044    ·    x", "4044 LF") is False
                  and A._head_names_eng("Engineering # 4016LF    ·    x", "4016") is False
                  and A._head_names_eng("Engineering # 4016    ·    x", "4016LF") is False
                  and A._head_names_eng("Engineering # 4044 — checked by JD", "4044") is True
                  and A._head_names_eng("Engineering # 4041 4042    ·    x", "4041 4042") is True
                  and A._head_names_eng("Engineering # 4041 4042    ·    x", "4043") is False
                  and A._head_names_eng("Engineering # — 4044 checked by JD", "4044") is None
                  and A._head_names_eng("Received from testing — 4044", "4044") is None
                  and A._head_names_eng(None, "4044") is None and A._head_names_eng("Engineering # 4044", "") is None
                  and A._eng_from_head_line("Engineering # 4044 LF    ·    x") == "4044")
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(hist7, ignore_errors=True)
        # ── oldflow-r7-2: the -900 entry's spelling (= the report's file name) follows the newest CLOSE; a hand-typed
        #    import or a legacy receive under the other spelling must not rename an existing old-flow report ──
        histT = tempfile.mkdtemp(prefix="pcc_r7twin_")
        custT = Path(tempfile.mkdtemp(prefix="pcc_r7twinc_")) / "377 Anduril"; extra_roots9.append(custT.parent)
        woT = custT / "134-2201-900 4065 Rev. A" / "Production Data" / "16300.1"; (woT / A.COC_FOLDER).mkdir(parents=True)
        tT = lambda i: f"{i:06d}-1342201900-A00"
        A.HISTORY_DIR = histT
        try:
            ev(histT, "2026-08", "a_dashed", close_ev("4065", "134-2201-900", "16300.1", [tT(i) for i in range(1, 7)], "2026-08-10 09:00:00", str(woT)))
            stT = A.build_receiving_state("4065", "134-2201-900"); stT["allow_tree_switch"] = True
            outT = Path(A.write_received_xlsx(stT))                               # built under the DASHED spelling
            ev(histT, "2026-09", "b_flat_import", {"event_type": "backfill_shipped", "eng": "4065", "pn": "1342201900", "wo": "16300.1",
               "ts": "2026-09-01 09:00:00", "ship_date": "8/20/26", "serials": [tT(7), tT(8)], "production_data": str(woT.parent)})
            entT1 = [j["pn"] for j in A.list_testing_jobs() if j["eng"] == "4065"]
            ev(histT, "2026-09", "c_flat_recv", recv_ev("4065", "1342201900", [tT(1)], "2026-09-02 09:00:00"))
            entT2 = [j["pn"] for j in A.list_testing_jobs() if j["eng"] == "4065"]
            A.auto_update_shipping_report("4065", "1342201900")
            filesT = sorted(custT.rglob("*Shipping Report*.xlsx"))
            ev(histT, "2026-09", "d_flat_close", close_ev("4065", "1342201900", "16300.1", [tT(9)], "2026-09-03 09:00:00", str(woT)))
            entT3 = [j["pn"] for j in A.list_testing_jobs() if j["eng"] == "4065"]
            check("r7: dashed closes + a NEWER flat-spelled import and a newer flat receive -> the entry stays '134-2201-900' and the update lands on the dashed file (no rename: 6 + 2 shipped, 1 back = 7 out); a newer flat CLOSE then wins (the WO folder's spelling)",
                  entT1 == ["134-2201-900"] and entT2 == ["134-2201-900"] and filesT == [outT] and _nrows(outT, A.OUT_SHEET) == 7
                  and entT3 == ["1342201900"] and not any("_from_close" in j for j in A.list_testing_jobs()))
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(histT, ignore_errors=True)
        # ── test-adequacy-r7-4: coated-r5-1's 'old flow byte-identical' — a COATED-form serial of the job's own core
        #    on a TESTING job is the one input that tells the _coated_only gate apart ──
        mix3W = _fake_win(dict(mixW._state)); mix3W._var.set("000002-1503300970-A00"); RW._add(mix3W)
        shown7 = []
        A.messagebox = _types.SimpleNamespace(showinfo=lambda *a, **k: shown7.append(a[1] if len(a) > 1 else str(k.get("message"))),
                                              showwarning=lambda *a, **k: None, showerror=lambda *a, **k: None)
        try:
            pm7W = _fake_win(dict(mixW._state)); pm7W._wait_state = lambda: True
            RW._add_many(pm7W, [tM(1), "000002-1503300970-A00"])
        finally:
            A.messagebox = _origMB6
        check("r7: old flow — a coated-form serial of the job's own core scanned into a TESTING job gets 'Doesn't match this job' (session unchanged); pasted, it is skipped as \"don't carry this job's P/N (150-3300-900)\" — _carries_pn accepts the coated digits on coated-only jobs ONLY",
              mix3W._issues == ["Doesn't match this job"] and mix3W._session == []
              and RW._carries_pn(mix3W, "000002-1503300970-A00") is False and RW._carries_pn(mix3W, tM(2)) is True
              and [e["serial"] for e in pm7W._session] == [tM(1)] and len(shown7) == 1
              and "1 skipped — don't carry this job's P/N (150-3300-900)" in shown7[0])
        # ── req-r7-1: the P/N stop names the label the operator sees everywhere else on an in-house-tested job ──
        capW = _fake_win(dict(st9, walk_error=False)); msgs7 = []
        def _cap7(serial, heading, message, approve=None):
            msgs7.append((heading, message)); capW._approved.setdefault(capW._sid(serial), set()).add(heading); return True
        capW._issue = _cap7
        capW._var.set("16493.2"); RW._add(capW)
        check("r7: the 'Doesn't match this job' body on an in-house-tested job names the coated label (125-1030-970C), never the invented -900 key",
              [h for h, _m in msgs7] == ["Nothing to receive here", "Doesn't match this job"]
              and f"this job's P/N ({CPL9})" in msgs7[1][1] and TP9 not in msgs7[1][1])
        # ── coated-r7-2 + fresh-r7-1 on a mixed job: A3's '(k not shipped for testing)' counts BOARDS (by seq) like
        #    'Coated boards on record'; the saved file carries exactly ONE selected tab, the active one ──
        hist7m = tempfile.mkdtemp(prefix="pcc_r7knt_")
        cust7m = Path(tempfile.mkdtemp(prefix="pcc_r7kntc_")) / "20 TestCust"; extra_roots9.append(cust7m.parent)
        wo7m9 = cust7m / "150-3300-900 4400 Rev. A" / "Production Data" / "16800.1"; wo7m9.mkdir(parents=True)
        A.HISTORY_DIR = hist7m
        try:
            ev(hist7m, "2026-08", "a_test", close_ev("4400", "150-3300-900", "16800.1", [tM(1), tM(2), tM(3)], "2026-08-01 09:00:00", str(wo7m9)))
            ev(hist7m, "2026-08", "b_recv", recv_ev("4400", "150-3300-900", [tM(1), tM(2), tM(3)], "2026-08-05 09:00:00"))
            ev(hist7m, "2026-08", "c_coat", close_ev("4400", "150-3300-970", "17001.1", [cM(1), cM(2), cM(3), tM(6)], "2026-08-15 09:00:00"))
            ev(hist7m, "2026-09", "d_coat2", close_ev("4400", "150-3300-970", "17002.1", [cM(6)], "2026-09-01 09:00:00"))   # board 6 re-shipped coated
            st7m = A.build_receiving_state("4400", "150-3300-900")
            out7m = Path(A.write_received_xlsx(st7m))
            wb7m = load_workbook(out7m); a3_7m = wb7m[A.RCV_SHEET]["A3"].value; wb7m.close()
            check("r7: A3 on the mixed job — 'Coated boards on record: 4 (1 not shipped for testing)': the Case-B board and its -970 re-ship are ONE untested board, the same basis as the count it qualifies",
                  A.receiving_summary(st7m)["coated_all"] == 4 and str(a3_7m).endswith("Coated boards on record: 4 (1 not shipped for testing)"))
            # a user saved the report with 'Still at customer' active (Excel writes tabSelected on that sheet); the
            # next daemon update must not leave two sheets selected ('[Group]' in Excel: a note typed on one sheet
            # lands on the other's LOG row too)
            wb7m = load_workbook(out7m); wb7m.active = 3; wb7m.worksheets[3].sheet_view.tabSelected = True; wb7m.save(str(out7m)); wb7m.close()
            A.write_received_xlsx(A.build_receiving_state("4400", "150-3300-900"))
            import zipfile as _zip7
            with _zip7.ZipFile(str(out7m)) as z7:
                wbxml7 = z7.read("xl/workbook.xml").decode("utf-8")
                sel7 = sorted(n for n in z7.namelist() if n.startswith("xl/worksheets/sheet") and 'tabSelected="1"' in z7.read(n).decode("utf-8"))
            mAT7 = _re7.search(r'activeTab="(\d+)"', wbxml7)
            act7 = int(mAT7.group(1)) if mAT7 else 0
            check("r7: after an update the saved workbook has exactly ONE tabSelected sheet and it is the activeTab sheet ('Received from testing' on a mixed job) — never a grouped selection",
                  act7 == 0 and sel7 == [f"xl/worksheets/sheet{act7 + 1}.xml"])
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(hist7m, ignore_errors=True)
        # ── coated-r7-1 (+ quality-r7-1): a Case-B shipment that exists ONLY as coated paperwork in the job folder
        #    (a real -900-form damage shipment never closed in-app). The fold registers its identity, so the folded
        #    state's rules agree with the sheet it writes, and the Save sweep warns before an override receive of
        #    such a board (the daemon's history-only walk cannot know it is Case B). Then the same tree with an
        #    access-denied stat inside the 'Shipping Reports' leaf: 'could not check', never 'does not hold'. ──
        hist7c = tempfile.mkdtemp(prefix="pcc_r7fcb_")
        cust7c = Path(tempfile.mkdtemp(prefix="pcc_r7fcbc_")) / "377 Anduril"; extra_roots9.append(cust7c.parent)
        wo7c1 = cust7c / "125-1040-970C 4700 Rev. C" / "Production Data" / "16493.1"; (wo7c1 / A.COC_FOLDER).mkdir(parents=True)
        wo7c2 = cust7c / "125-1040-970C 4700 Rev. C" / "Production Data" / "16493.2"; (wo7c2 / A.COC_FOLDER).mkdir(parents=True)
        c7 = lambda i: f"{i:06d}-1251040970-C00"; t7 = lambda i: f"{i:06d}-1251040900-C00"
        _shipfile(wo7c2 / A.COC_FOLDER, "coc-damage.xlsx", [t7(569), t7(571)], "2026-08-27", pn_label="125-1040-970C")   # folder-ONLY Case B
        A.HISTORY_DIR = hist7c
        try:
            ev(hist7c, "2026-08", "a_coat", close_ev("4700", "125-1040-970C", "16493.1", [c7(i) for i in range(851, 856)], "2026-08-20 10:00:00", str(wo7c1)))
            j7c = next(j for j in A.list_testing_jobs() if j["eng"] == "4700")
            st7c = A.build_receiving_state("4700", j7c["pn"]); st7c["allow_tree_switch"] = True
            out7c = Path(A.write_received_xlsx(st7c))                             # 🔁 Build report (history only)
            sc7c = A.scan_job_folders(st7c)
            A.fold_folder_scan(st7c, sc7c)
            kb7 = A._serial_identity(t7(569), A._coat_ident_pn(j7c["pn"]))
            stf7 = dict(st7c, received={}, receive_log=[], recv_dates={})
            A._fold_receive_locally(stf7, [{"serial": t7(569), "result": "pass"}], "op", "", "2026-09-23 10:00:00")
            st7cx = dict(st7c, testing_label_seen=True, hist_has_testing=True)   # the shape after a flip: the scan must still file them as Case B
            sc7cx = A.scan_job_folders(st7cx)
            check("r7: folder-ONLY Case-B boards join coat_label_mismatch on the fold (and are remembered as folder-found): the folded state stays in-house tested after a local receive of one, and a later scan on a testing-history state keeps them out of ship_counts",
                  sc7c["ok"] and sc7c["ship_counts"] == {} and sorted(r["serial"] for r in sc7c["coated_rows"]) == [t7(569), t7(571)]
                  and kb7 in st7c["coat_label_mismatch"] and kb7 in st7c["_caseb_folder"]
                  and A._has_testing(st7c) is False and RW._coated_only(st7c) is True
                  and len(stf7["received"]) == 1 and A._has_testing(stf7) is False
                  and sc7cx["ship_counts"] == {} and len(sc7cx["coated_rows"]) == 2)
            _origC7, _origMB7b = A.ConfirmActionDialog, A.messagebox
            class _FakeConfirm7:
                def __init__(self, *a, **k): self.result = False
            A.ConfirmActionDialog = _FakeConfirm7
            A.messagebox = _types.SimpleNamespace(showinfo=lambda *a, **k: None, showwarning=lambda *a, **k: None, showerror=lambda *a, **k: None)
            try:
                pwW = _save_win(dict(st7c), [{"serial": t7(569), "status": "not_shipped", "result": "pass"}],
                                lambda w, s, h: (w._issues.append(h), False)[1])
                pwW._approved = {pwW._sid(t7(569)): {"Nothing to receive here"}}     # the scan's own override
                RW._do_save(pwW)
                pw2W = _save_win(dict(st7c), [{"serial": c7(851), "status": "not_shipped", "result": "pass"}],
                                 lambda w, s, h: (w._issues.append(h), False)[1])
                pw2W._approved = {pw2W._sid(c7(851)): {"Nothing to receive here"}}
                RW._do_save(pw2W)
            finally:
                A.ConfirmActionDialog, A.messagebox = _origC7, _origMB7b
            check("r7: the Save sweep stops an overridden receive of a folder-only Case-B board ('Coated shipment on paperwork only'; refused -> dropped); a coated board of the job gets no such stop",
                  pwW._issues == ["Coated shipment on paperwork only"] and pwW._session == []
                  and pw2W._issues == [] and [e["serial"] for e in pw2W._session] == [c7(851)]
                  and RW._dup_checks_due(_types.SimpleNamespace(_approved={"k": {"Coated shipment on paperwork only"}}), "k") is True)
            # quality-r7-1: a GENUINE -900-form receive (not Case B) makes finalize reach the 'file is the switch' layer
            ev(hist7c, "2026-09", "b_genuine", recv_ev("4700", j7c["pn"], [t7(700)], "2026-09-10 09:00:00", "Anita Patel", str(wo7c1.parent)))
            st7g = A.build_receiving_state("4700", j7c["pn"])                    # readable: the report keeps the coated home
            _origIF7 = Path.is_file
            def _deny_leaf7(self):
                if self.parent.name in (A.RECEIVED_FOLDER, A.RECEIVED_FOLDER_LEGACY):
                    raise PermissionError(13, "Access is denied")
                return _origIF7(self)
            Path.is_file = _deny_leaf7
            try:
                h7 = A._wo_holds_report(wo7c1, st7g, whole_tree=True)
                c7h = A._coated_tree_holds_report(str(wo7c1), st7g)
                st7u = A._annotate_report_exists(A.build_receiving_state("4700", j7c["pn"]))
                row7u = A._collect_report_rows({}, jobs=[j7c])[0]
            finally:
                Path.is_file = _origIF7
            check("r7: an access-denied stat inside the 'Shipping Reports' leaf is 'could not check', never 'does not hold': _wo_holds_report None, _coated_tree_holds_report None, the walk sets _home_unknown (wo_path None), _report_exists None, picker \"can't check from this PC\"; readable, the file keeps the coated home",
                  st7g["wo_path"] == str(wo7c1) and st7g.get("home_side") == "coated" and A._wo_holds_report(wo7c1, st7g, whole_tree=True) is True
                  and h7 is None and c7h is None
                  and st7u["wo_path"] is None and st7u.get("_home_unknown") is True and st7u["_report_exists"] is None
                  and row7u["status"] == "can't check from this PC")
        finally:
            A.HISTORY_DIR = hist9
            shutil.rmtree(hist7c, ignore_errors=True)
        src7 = Path(SCRIPTS, "app_logic.py").read_text(encoding="utf-8")
        check("r7: wiring — blocked-no-copy handled before 'Build report?' / in the Save tail / in the picker; one 'brings it to' sentence for chips + Open file; coated-only Save popup line; first build says 'Built'; tri-state _wo_holds_report; lookup never hands out the foreign path; _coated_tree_holds_report on _tr_embed_keys; close-wins spelling; one spelling for the discarded rev letter",
              'if st.get("_report_exists") is False and blk0:' in src7
              and src7.index('if st.get("_report_exists") is False and blk0:') < src7.index('messagebox.askyesno("Build report?",')
              and 'elif cur.get("_report_blocked_by"):' in src7
              and src7.count("the next Save / after-shipment update (or 🔁 Rebuild report) brings it to") == 2
              and 'status="another job\'s file at its report location"' in src7
              and '"Recorded as a supervisor-overridden receive on this in-house-tested job "' in src7
              and 'messagebox.showinfo("Built" if first else "Rebuilt"' in src7
              and "return None if raised else False" in src7
              and "return False, (None if blocked else target), archived" in src7
              and 'pk_al, ck_al = _tr_embed_keys(state.get("pn"))\n            for side_al in (ck_al, pk_al):' in src7
              and 'j["pn"] = spelled; j["_from_close"] = True' in src7
              and not any(_re7.search(r"\b(_l|_letter|_ol|_l_)\b", ln) for ln in src7.splitlines() if "_tr_pn_parts(" in ln))
    finally:
        A.HISTORY_DIR = _h9
        shutil.rmtree(hist9, ignore_errors=True)
        for _root in [cust9root] + extra_roots9:
            for r_, _, fs in os.walk(_root):
                for f in fs:
                    try: os.chmod(os.path.join(r_, f), stat.S_IWRITE)
                    except Exception: pass
            shutil.rmtree(_root, ignore_errors=True)
    # the section-2/4 file EXISTS (opted in): auto_update from a -970 close updates
    # THAT file — same path write_received_xlsx returned, no second file anywhere
    mt_before = Path(out).stat().st_mtime
    time.sleep(0.05)
    A.auto_update_shipping_report(ENG, CPN)
    check("auto_update on the 4041 fixture updates the SAME file write_received_xlsx returned; no second 4041 file",
          Path(out).exists() and Path(out).stat().st_mtime > mt_before
          and len(list(job.rglob("136-9624-9*Shipping Report*.xlsx"))) == 1)

    # ═══ 5b. review fixes: walk_error + backfill in the shipping dup-check ═══
    print("review fixes:")
    # shipping dup-check now sees backfilled serials as 'already shipped'
    idx = A.build_serial_history(ENG, TPN)
    check("dup-check finds a backfilled serial", A._serial_identity(tser(2050), TPN) in idx)
    check("dup-check backfill hit shows WO + ship_date",
          idx[A._serial_identity(tser(2050), TPN)]["wo"] == "16149.1"
          and idx[A._serial_identity(tser(2050), TPN)]["ts"] == "5/2/25")
    check("dup-check still finds close serials", A._serial_identity(tser(42), TPN) in idx)
    # a broken history root -> walk_error flagged, file refuses to regenerate
    _h2 = A.HISTORY_DIR
    badf = Path(proj) / "not_a_dir.txt"; badf.write_text("x", encoding="utf-8")
    A.HISTORY_DIR = str(badf)
    stbad = A.build_receiving_state(ENG, TPN)
    check("broken walk sets walk_error", stbad.get("walk_error") is True)
    stbad["pd_path"] = str(pd_path)
    try:
        A.write_received_xlsx(stbad); refused = False
    except RuntimeError: refused = True
    except Exception: refused = False
    check("xlsx write REFUSES a walk_error state (file untouched)", refused)
    A.HISTORY_DIR = _h2
    check("missing history root is NOT an error (fresh install)",
          A.build_receiving_state(ENG, TPN) is not None)
    # tmp naming carries the station name (PID collisions between stations)
    src_tmp = Path(SCRIPTS, "app_logic.py").read_text(encoding="utf-8")
    check("tmp name carries station id + pid (source pin)",
          'f"~{station}_{os.getpid()}_{target.name}.tmp"' in src_tmp)

    # ═══ 6. job listing + auto-detect ════════════════════════════════════════
    print("job listing / detection:")
    jobs = A.list_testing_jobs()
    keys = {(j["eng"], j["pn"]) for j in jobs}
    check("testing job listed", (ENG, TPN) in keys)
    check("-900 noise job listed too", ("9999", "555-000-900") in keys)
    check("coated (970) job NOT listed", all(j["pn"] != CPN for j in jobs))
    j = A.detect_job_for_serial(tser(77), jobs)
    check("auto-detect from scan", j and j["pn"] == TPN)
    check("no match for other-format scan", A.detect_job_for_serial("16452.1-77", jobs) is None)
    long_jobs = [{"eng": "1", "pn": "136-9624"}, {"eng": "2", "pn": "136-9624-900"}]
    check("longest P/N wins", A.detect_job_for_serial(tser(5), long_jobs)["eng"] == "2")
    check("short P/N (<6 alnum) never grabs", A.detect_job_for_serial("000001-12345-A00",
          [{"eng": "x", "pn": "12345"}]) is None)

    # ═══ 7. WO search ignores the Received folder ════════════════════════════
    print("WO search exclusion:")
    (pd_path / A.RECEIVED_FOLDER).mkdir(exist_ok=True)
    hits = A.search_wo_folder("Industrial", "4041")     # eng search lists all subdirs
    names = {h.name for h in hits}
    check("eng search finds the WOs", "16151.1" in names)
    check("'Shipping Reports' never offered as a WO", A.RECEIVED_FOLDER not in names)
    check("WO-number search unaffected", [h.name for h in A.search_wo_folder("Industrial", "16152.1")] == ["16152.1"])
finally:
    A.HISTORY_DIR = _orig_hist
    A.PRODUCTION_ROOTS.clear(); A.PRODUCTION_ROOTS.update(_orig_roots)
    # read-only file blocks rmtree on Windows — clear attrs first
    for r, _, fs in os.walk(proj):
        for f in fs:
            try: os.chmod(os.path.join(r, f), stat.S_IWRITE)
            except Exception: pass
    shutil.rmtree(hist, ignore_errors=True); shutil.rmtree(proj, ignore_errors=True)

print()
print("ALL TESTS PASSED" if _fail == 0 else f"{_fail} TEST(S) FAILED")
sys.exit(1 if _fail else 0)
