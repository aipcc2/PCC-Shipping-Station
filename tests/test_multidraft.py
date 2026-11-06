"""Multi-draft feature tests: coexisting drafts per WO, targeted save/load/delete,
legacy single-draft compatibility, listing order."""
import sys, tempfile, shutil, json, time
from pathlib import Path
sys.path.insert(0, r"C:\Users\Dsavaliya\Claude Cowork\PCC Shipping App\Scripts")
import app_logic as A

fail = 0
def check(n, c):
    global fail
    print(("  ok  " if c else " FAIL ") + n)
    if not c: fail += 1

wo = tempfile.mkdtemp(prefix="pcc_md_")
try:
    # legacy draft (no id) + two new-style drafts
    p0 = A.save_shipment_draft(wo, [], ["L1", "L2"], None, None, operator="Old Op")           # legacy
    time.sleep(0.02)
    p1 = A.save_shipment_draft(wo, [], [f"A{i}" for i in range(50)], None, None,
                               operator="Anita Patel", notes="batch 1", draft_id="20260821_100000")
    time.sleep(0.02)
    p2 = A.save_shipment_draft(wo, [], [f"B{i}" for i in range(50)], None, None,
                               operator="Lata Gandhi", notes="batch 2", draft_id="20260821_110000")
    check("three distinct draft files", len({p0, p1, p2}) == 3 and all((p0, p1, p2)))
    ds = A.list_shipment_drafts(wo)
    check("list sees all three", len(ds) == 3)
    check("newest first", ds[0]["draft_id"] == "20260821_110000" or ds[0]["saved_at"] >= ds[-1]["saved_at"])
    check("rows carry date + saver + counts",
          all(d["saved_at"] and "serial_count" in d for d in ds)
          and any(d["saved_by"] == "Anita Patel" and d["serial_count"] == 50 for d in ds))
    # targeted load
    d1 = A.load_shipment_draft(wo, path=next(d["path"] for d in ds if d["draft_id"] == "20260821_100000"))
    check("targeted load gets batch 1", d1["serials"][0] == "A0" and d1["notes"] == "batch 1")
    dleg = A.load_shipment_draft(wo)     # default = legacy file
    check("legacy default load still works", dleg and dleg["serials"] == ["L1", "L2"])
    check("legacy draft_id is blank", (A.list_shipment_drafts(wo)[-1]["draft_id"] == "" or
                                        any(d["draft_id"] == "" for d in ds)))
    # re-save same id overwrites (no new file)
    A.save_shipment_draft(wo, [], [f"A{i}" for i in range(60)], None, None,
                          operator="Anita Patel", draft_id="20260821_100000")
    check("re-save same id keeps 3 files", len(A.list_shipment_drafts(wo)) == 3)
    check("re-save updated content",
          next(d for d in A.list_shipment_drafts(wo) if d["draft_id"] == "20260821_100000")["serial_count"] == 60)
    # targeted delete removes ONLY that draft
    ok, err = A.delete_shipment_draft(wo, draft_id="20260821_100000")
    left = A.list_shipment_drafts(wo)
    check("targeted delete ok", ok and err is None)
    check("only that draft gone; other 2 remain",
          len(left) == 2 and all(d["draft_id"] != "20260821_100000" for d in left))
    # legacy delete via blank id
    ok2, _ = A.delete_shipment_draft(wo, draft_id="")
    check("legacy delete via blank id", ok2 and len(A.list_shipment_drafts(wo)) == 1)
    ok3, _ = A.delete_shipment_draft(wo, draft_id="20260821_110000")
    check("all drafts cleanly removable", ok3 and A.list_shipment_drafts(wo) == [])
    # deleting a nonexistent draft is fine
    ok4, _ = A.delete_shipment_draft(wo, draft_id="nope")
    check("delete of missing draft is a no-op success", ok4)
    # draft filename sanitization (weird id can't break the path)
    p5 = A.save_shipment_draft(wo, [], ["X"], None, None, draft_id='we?ird:id*')
    check("weird draft id sanitized into a valid file", p5 and Path(p5).exists())
finally:
    shutil.rmtree(wo, ignore_errors=True)
print("ALL PASSED" if not fail else f"{fail} FAILED")
sys.exit(1 if fail else 0)
