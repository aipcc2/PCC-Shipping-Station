# READ-ONLY: last 3 weeks' closes of test-and-return jobs, and what auto_update_shipping_report would do — WITHOUT writing.
import sys, json, re, time, glob, os
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "Scripts"))   # the bundle's Scripts folder
import app_logic as A
H = Path(A.HISTORY_DIR)
evs = []
for ym in ("2026-09", "2026-10"):
    for f in (H / "2026" / ym).glob("*_close_*.json"):
        try: r = json.load(open(f, encoding="utf-8"))
        except Exception: continue
        if r.get("ts", "") >= "2026-10-01": evs.append((r.get("ts"), r.get("eng"), r.get("pn"), r.get("wo"), r.get("station"), r.get("app_version"), r.get("voided")))
evs.sort()
print("closes since 10/1:", len(evs))
seen = set()
for ts, eng, pn, wo, stn, ver, vd in evs:
    tr = bool(re.search(r'(\d{6,7})(?:900|970)', A._alnum(pn or "")))
    print(f"  {ts} eng={eng} pn={pn} wo={wo} station={stn} v={ver} test-and-return={tr}")
    if tr: seen.add((eng, pn))
print("MAG-ITAR reachable from this PC:", Path(A.PRODUCTION_ROOTS["MAG - ITAR"]).is_dir() if True else None)
for eng, pn in sorted(seen):
    t0 = time.time()
    ek = A._norm_key(eng); core = re.search(r'(\d{6,7})(?:900|970)', A._alnum(pn)).group(1)
    jobs = [j for j in A.list_testing_jobs() if A._norm_key(j.get("eng")) == ek and core in A._alnum(j.get("pn"))]
    for j in jobs:
        st = A.build_receiving_state(j["eng"], j["pn"])
        has = A._report_exists(st) if st and st.get("wo_path") else ("no wo_path", None)
        print(f"  JOB {eng} {pn} -> entry pn={j['pn']} walk_error={st.get('walk_error')} home={st.get('home_side')} unknown={st.get('_home_unknown')} wo_path={str(st.get('wo_path'))[-60:]} report_exists={has[0]} ({time.time()-t0:.1f}s)")
    if not jobs: print(f"  JOB {eng} {pn} -> NOT in list_testing_jobs (update skipped)")
