import sys, json
from pathlib import Path
from collections import defaultdict, Counter
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "Scripts"))   # the bundle's Scripts folder
import app_logic as A
H = Path(A.HISTORY_DIR); recs = []
for ym in ("2026-09", "2026-10"):
    for f in sorted((H / "2026" / ym).glob("*.json")):
        try: r = json.load(open(f, encoding="utf-8")); r["_f"] = f.name; recs.append(r)
        except Exception: pass
print("events Sep-Oct:", Counter(r.get("event_type") for r in recs))
keys = Counter(k for r in recs if r.get("event_type") in ("draft_save", "close") for k in r)
print("keys on draft/close:", sorted(keys))
for wo in ("16588.1", "16582.1"):
    print(f"\n=== WO {wo}")
    for r in sorted((r for r in recs if r.get("wo") == wo), key=lambda r: r.get("ts", "")):
        print(f"  {r.get('ts')} {r.get('event_type'):10s} op={r.get('operator')!s:14s} station={r.get('station')} draft_id={r.get('draft_id')!s:18s} serials={r.get('serial_count', len(r.get('serials') or []))} resumed={r.get('resumed_from_draft')} saved_by={r.get('saved_by')}")
# how many draft_save events per draft_id
dd = Counter(r.get("draft_id") for r in recs if r.get("event_type") == "draft_save")
print("\ndraft_save events per draft_id (top):", dd.most_common(6))
closes_with_did = sum(1 for r in recs if r.get("event_type") == "close" and r.get("draft_id"))
print("closes:", sum(1 for r in recs if r.get('event_type')=='close'), "with draft_id:", closes_with_did)
