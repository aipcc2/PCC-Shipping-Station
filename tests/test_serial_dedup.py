"""Focused tests for the 2026-08-17 serial dedup + outlier + notes work.
Run:  python test_serial_dedup.py   (from anywhere; it adds Scripts to sys.path)
Exercises pure logic only — no Tk window is created."""
import sys, os, json, tempfile, datetime, shutil
from pathlib import Path

SCRIPTS = r"C:\Users\Dsavaliya\Claude Cowork\PCC Shipping App\Scripts"
sys.path.insert(0, SCRIPTS)
import app_logic as A

_fail = 0
def check(name, cond):
    global _fail
    print(("  ok  " if cond else " FAIL ") + name)
    if not cond: _fail += 1

# ── helpers ──────────────────────────────────────────────────────────────────
def write_close(root, ym, name, eng, pn, serials, event_type="close"):
    folder = Path(root) / ym[:4] / ym
    folder.mkdir(parents=True, exist_ok=True)
    rec = {"schema_version": 1, "event_type": event_type, "eng": eng, "pn": pn,
           "wo": name, "operator": "Anita Patel",
           "ts": "2026-07-30 10:00:00", "event_id": f"{ym}_{name}_{event_type}",
           "serials": list(serials)}
    (folder / f"{ym}_{name}_{event_type}.json").write_text(json.dumps(rec), encoding="utf-8")

# ── 1. build_serial_history scoping (Eng# + P/N) ─────────────────────────────
print("build_serial_history:")
tmp = tempfile.mkdtemp(prefix="pcc_hist_")
_orig_hist = A.HISTORY_DIR
try:
    A.HISTORY_DIR = tmp
    # Same Eng#, TWO P/Ns, SAME serial numbers 1..3 — must NOT merge.
    write_close(tmp, "2026-06", "16151.1", "4041", "192-146-900", ["1", "2", "3"])
    write_close(tmp, "2026-07", "16152.1", "4041", "192-146-900", ["4", "5", "6"])
    write_close(tmp, "2026-07", "16154.1", "4041", "192-146-970", ["1", "2", "3"])
    # A draft with serials must be IGNORED (only 'close' counts).
    write_close(tmp, "2026-07", "16199.9", "4041", "192-146-900", ["999"], event_type="draft_save")

    idx900 = A.build_serial_history("4041", "192-146-900")
    check("testing P/N sees its 6 serials across 2 WOs", set(idx900) == {"1","2","3","4","5","6"})
    check("serial '1' in testing maps to a testing WO (16151.1), NOT the coated one",
          idx900.get("1", {}).get("wo") == "16151.1")
    check("draft-only serial 999 is excluded", "999" not in idx900)

    idx970 = A.build_serial_history("4041", "192-146-970")
    check("coated P/N sees only its own 1..3", set(idx970) == {"1","2","3"})
    check("serial '1' in coated maps to 16154.1 (P/N scoping keeps them separate)",
          idx970.get("1", {}).get("wo") == "16154.1")

    check("blank P/N -> no scope -> empty index", A.build_serial_history("4041", "") == {})
    check("blank Eng -> no scope -> empty index", A.build_serial_history("", "192-146-900") == {})
    check("unknown job -> empty", A.build_serial_history("9999", "000-000-000") == {})
    # case/space-insensitive matching
    idxci = A.build_serial_history(" 4041 ", "192-146-900")
    check("eng/pn matching is trim-insensitive", set(idxci) == {"1","2","3","4","5","6"})
finally:
    A.HISTORY_DIR = _orig_hist
    shutil.rmtree(tmp, ignore_errors=True)

# ── 1b. REV-AGNOSTIC identity (Dhruvil 2026-09-02: jobs 4092/4041 changed rev
#        mid-job — same board wears -A01 in old shipments, -B01 later; identity
#        is Eng scope + P/N + seq, REV IGNORED. Supersedes the 08-19 exact rule.)
print("rev-agnostic identity:")
si = A._serial_identity
check("rev stripped when job P/N embedded",
      si("000633-1629645-900-A01", "1629645-900") == "000633-1629645-900")
check("A01 and B02 of the same board -> SAME identity",
      si("000633-1629645-900-A01", "1629645-900") == si("000633-1629645-900-B02", "1629645-900"))
check("4092-style P/N with dashes works",
      si("000123-136-3312-900-A01", "136-3312-900") == si("000123-136-3312-900-B01", "136-3312-900"))
check("different seq stays a different board",
      si("000633-1629645-900-A01", "1629645-900") != si("000634-1629645-900-A01", "1629645-900"))
check("<WO>-<seq> format never collapsed (no P/N inside)",
      si("16158.1-001", "1629645-900") == "16158.1-001")
check("no/short P/N -> exact string", si("000001-1921682900-AOO", "") == "000001-1921682900-aoo")
tmp3 = tempfile.mkdtemp(prefix="pcc_rev_")
try:
    A.HISTORY_DIR = tmp3
    write_close(tmp3, "2026-07", "16151.1", "4041", "192-168-2900",
                ["000001-1921682900-AOO", "000002-1921682900-AOO"])
    idxr = A.build_serial_history("4041", "192-168-2900")
    check("same serial incl. Rev -> hit",
          idxr.get(si("000001-1921682900-AOO", "192-168-2900")) is not None)
    check("SAME board, DIFFERENT Rev -> ALSO a hit (rev ignored)",
          idxr.get(si("000001-1921682900-B00", "192-168-2900")) is not None)
    check("genuinely new seq -> not a hit",
          idxr.get(si("000009-1921682900-AOO", "192-168-2900")) is None)
    check("index keys are rev-stripped identities",
          set(idxr) == {"000001-1921682900", "000002-1921682900"})
finally:
    A.HISTORY_DIR = _orig_hist
    shutil.rmtree(tmp3, ignore_errors=True)

# ── 2. detect_serial_outliers ────────────────────────────────────────────────
print("detect_serial_outliers:")

def flagged_set(serials, **kw):
    return {s for s, _ in A.detect_serial_outliers(serials, **kw)}

# specific customer: <seq>-<P/N>-<Rev>
clean_c = [f"{i:06d}-1921682900-AOO" for i in range(1, 26)]
check("clean seq-PN-Rev batch: no outliers", A.detect_serial_outliers(clean_c) == [])

wrong_pn = list(clean_c); wrong_pn[7] = "000008-1921682970-AOO"   # a coated board in a bare batch
check("wrong P/N in the middle is flagged", "000008-1921682970-AOO" in flagged_set(wrong_pn))

wild_seq = list(clean_c) + ["000999-1921682900-AOO"]             # a high but plausible board #
check("a high sequence number is NOT flagged (gaps are normal)",
      "000999-1921682900-AOO" not in flagged_set(wild_seq))

garbage = list(clean_c) + ["8675309"]                            # random junk scan
check("random junk scan is flagged", "8675309" in flagged_set(garbage))

# other customers: <WO>-<seq>
clean_w = [f"16158.1-{i:03d}" for i in range(1, 26)]
check("clean WO-seq batch: no outliers", A.detect_serial_outliers(clean_w) == [])
bad_w = list(clean_w) + ["16152.111"]                            # the user's example
check("the 16152.111 example is flagged", "16152.111" in flagged_set(bad_w))
wrong_wo = list(clean_w); wrong_wo[3] = "16159.1-004"            # a board from a different WO
check("wrong WO prefix is flagged", "16159.1-004" in flagged_set(wrong_wo))
# 1-/2-/3-digit sequence numbers in ONE WO must NOT flag (Dhruvil 2026-08-19 false positive)
dhruvil_case = [f"16452.1-{n}" for n in (42, 1, 2, 3, 4, 5, 6, 45, 15)]
check("Dhruvil's mixed 1-/2-digit seq: NO flags", A.detect_serial_outliers(dhruvil_case) == [])
big_seq = [f"16452.1-{n}" for n in range(1, 121)]                # 1..120 spans 1/2/3 digits
check("1..120 seq (1/2/3-digit) in one WO: NO flags", A.detect_serial_outliers(big_seq) == [])
# a high but plausible seq among low ones is NOT flagged — gaps are normal (Dhruvil 2026-08-19)
check("45 among 1-5 (same WO) is NOT flagged",
      A.detect_serial_outliers([f"16452.1-{n}" for n in (1, 2, 3, 4, 5, 45)]) == [])
check("1,2,3 then 45 (big gap) is NOT flagged",
      A.detect_serial_outliers([f"16452.1-{n}" for n in (1, 2, 3, 45)]) == [])

# conservative: don't over-flag
check("tiny batch (<5) is left alone", A.detect_serial_outliers(["1","2","999"]) == [])
check("big clean numeric batch 1..200: no outliers",
      A.detect_serial_outliers([str(i) for i in range(1, 201)]) == [])
# a legit internal gap (two sub-lots in one shipment) must NOT trip the range check
gap = [f"16158.1-{i:03d}" for i in list(range(1, 11)) + list(range(51, 61))]
check("legit internal gap is not flagged", A.detect_serial_outliers(gap) == [])
# genuinely mixed batch (different segment counts -> no dominant shape) -> no guessing
mixed = ["1-2", "1-2-3", "a-b-c-d", "xx", "y-z"]
check("no dominant shape -> no flags", A.detect_serial_outliers(mixed) == [])

# ── 2b. P/N-anchored odd-serial check (Dhruvil's cases; works for 2-scan batches) ─
print("odd-serial check — P/N-anchored:")
PNJ = "192-146-900"
def odd(serials, pn=PNJ):
    return {s for s, _ in A.detect_serial_outliers(serials, pn=pn)}
check("case1: same seq+P/N, A00 vs B00 -> one flagged",
      len(odd(["000171-192-146-900-A00", "000171-192-146-900-B00"])) >= 1)
check("case2: A00 vs 000 -> one flagged",
      len(odd(["000171-192-146-900-A00", "000171-192-146-900-000"])) >= 1)
check("case3: malformed 000000 (no P/N) flagged",
      "000000" in odd(["000171-192-146-900-A00", "000000"]))
check("wrong P/N (-970 in a -900 batch) flagged",
      "000006-192-146-970-A00" in odd([f"{i:06d}-192-146-900-A00" for i in range(1, 6)] + ["000006-192-146-970-A00"]))
check("lone wrong revision B00 in an A00 batch flagged",
      "000025-192-146-900-B00" in odd([f"{i:06d}-192-146-900-A00" for i in range(1, 25)] + ["000025-192-146-900-B00"]))
check("a high sequence number is NOT flagged (no numeric-gap check)",
      "999999-192-146-900-A00" not in odd([f"{i:06d}-192-146-900-A00" for i in range(1, 20)] + ["999999-192-146-900-A00"]))
check("CLEAN dashed-P/N batch (all A00): no flags",
      A.detect_serial_outliers([f"{i:06d}-192-146-900-A00" for i in range(1, 26)], pn=PNJ) == [])
check("CLEAN 2-scan same-rev batch: no flags",
      A.detect_serial_outliers(["000001-192-146-900-A00", "000002-192-146-900-A00"], pn=PNJ) == [])

# ── 2d. Archives folders are ignored in the WO search ───────────────────────
print("archives exclusion in WO search:")
tmpwo = tempfile.mkdtemp(prefix="pcc_wo_")
_orig_roots = dict(A.PRODUCTION_ROOTS)
try:
    A.PRODUCTION_ROOTS["Industrial"] = tmpwo
    live = Path(tmpwo) / "379 Liber8Now" / "GSR10 [4033 Rev. 1.1.11]" / "Production Data" / "16534.1"
    arch = Path(tmpwo) / "379 Liber8Now" / "Archives" / "GSR10 [4033 Rev. 1.1.9]" / "Production Data" / "16475.1"
    live.mkdir(parents=True); arch.mkdir(parents=True)
    check("live WO 16534.1 is found",
          [p.name for p in A.search_wo_folder("Industrial", "16534.1")] == ["16534.1"])
    check("archived WO 16475.1 is IGNORED",
          A.search_wo_folder("Industrial", "16475.1") == [])
    check("eng-# search returns only the live WO (archived one skipped)",
          [p.name for p in A.search_wo_folder("Industrial", "4033")] == ["16534.1"])
    # additional excluded folder names: Archived, Old
    (Path(tmpwo) / "Cust2" / "Archived" / "JobX [5000]" / "Production Data" / "17001.1").mkdir(parents=True)
    (Path(tmpwo) / "Cust3" / "Old" / "JobY [5001]" / "Production Data" / "17002.1").mkdir(parents=True)
    check("'Archived' folder is ignored", A.search_wo_folder("Industrial", "17001.1") == [])
    check("'Old' folder is ignored", A.search_wo_folder("Industrial", "17002.1") == [])
    # a name that merely CONTAINS 'old' as a substring is NOT excluded (segment match)
    (Path(tmpwo) / "Boldera [5002]" / "Production Data" / "17003.1").mkdir(parents=True)
    check("substring 'old' (Boldera) is NOT excluded",
          [p.name for p in A.search_wo_folder("Industrial", "17003.1")] == ["17003.1"])
finally:
    A.PRODUCTION_ROOTS.clear(); A.PRODUCTION_ROOTS.update(_orig_roots)
    shutil.rmtree(tmpwo, ignore_errors=True)

# ── 3. notes round-trip through the draft ────────────────────────────────────
print("notes round-trip:")
tmp2 = tempfile.mkdtemp(prefix="pcc_draft_")
try:
    note = "Ship to Building C dock 4 — call ext 210 first"
    out = A.save_shipment_draft(tmp2, [], ["1", "2"], None, None,
                                operator="Anita Patel", notes=note)
    check("draft saved", bool(out))
    d = A.load_shipment_draft(tmp2)
    check("notes survive save/load", (d or {}).get("notes") == note)
    check("serials survive save/load", (d or {}).get("serials") == ["1", "2"])
    # backward-compat: a draft written WITHOUT notes loads as empty, not a crash
    A.save_shipment_draft(tmp2, [], ["3"], None, None, operator="X")
    d2 = A.load_shipment_draft(tmp2)
    check("no-notes draft -> empty string", (d2 or {}).get("notes", "") == "")
finally:
    shutil.rmtree(tmp2, ignore_errors=True)

# ── 4. sanity: existing parsers still importable/working ─────────────────────
print("regression sanity (parsers unchanged):")
pn, inside = A.split_job_name("161-1061-970 4081 Rev. B00")
eng, rev = A.eng_from_inside(inside), A.rev_from_inside(inside)
check("split_job_name delimiter-less format still parses", (pn, eng, rev) == ("161-1061-970", "4081", "B00"))
check("sort_serials natural order", A.sort_serials(["SN10","SN2","SN1"]) == ["SN1","SN2","SN10"])
check("_norm_serial trims+casefolds", A._norm_serial("  AbC ") == "abc")

# (Rev behavior is covered in section 1b above: matching is EXACT, so a different
#  Rev is a different serial and is NOT treated as a duplicate.)

# ── 6. COC assembly-process options (Change 2) ───────────────────────────────
print("COC assembly-process options:")
check("AO list is exactly the 4 requested options (No PCB Assembly removed)",
      A.COCPopup.AO == ["Lead free", "Leaded", "Special", "N/A"])

print()
print(("ALL TESTS PASSED" if _fail == 0 else f"{_fail} TEST(S) FAILED"))
sys.exit(1 if _fail else 0)
