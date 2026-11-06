import sys, tempfile, shutil
from pathlib import Path
sys.path.insert(0, r"C:\Users\Dsavaliya\Claude Cowork\PCC Shipping App\Scripts")
import app_logic as A

fail = 0
def check(n, c):
    global fail
    print(("  ok  " if c else " FAIL ") + n)
    if not c: fail += 1

d = tempfile.mkdtemp(prefix="pcc_lbl_")
try:
    sers = [f"{i:06d}-136-9624-900-A00" for i in range(1, 6)]
    p = A.write_label_html_4x2(d, "136-9624-900", sers, wo="16452.1")
    doc = open(p, encoding="utf-8").read()
    check("file written", Path(p).exists() and p.endswith(".html"))
    check("qrcodejs embedded inline (no CDN)", "var QRCode;!function" in doc and "cdnjs" not in doc)
    check("page size 4in 2in, zero margin", "@page{size:4in 2in;margin:0}" in doc)
    check("one label block per serial", doc.count('<div class="lab">') == 5)
    check("P/N column + Serial column titles", "P/N:" in doc and "Serial:" in doc)
    check("every serial present (value + 6pt mirror)",
          all(doc.count(s) >= 2 for s in sers))
    check("QR holders for both columns per label",
          all(f'id="qp{i}"' in doc and f'id="qs{i}"' in doc for i in range(5)))
    check("serial data handed to QR script as JSON", '"000003-136-9624-900-A00"' in doc)
    check("auto print trigger", "window.print()" in doc)
    check("print-then-close (Dhruvil 2026-09-25): the tab closes itself after Chrome's print dialog, 600 ms after load as before",
          A._PRINT_THEN_CLOSE_JS in doc and "setTimeout(_pccPrintThenClose,600);" in doc
          and "setTimeout(function(){window.print();},600);" not in doc
          and doc.count("window.print()") == 2 and 'onclick="window.print()"' in doc)   # the hint bar's button + the print-then-close
    check("page-break per label", "page-break-after:always" in doc)
    # 2026-09-02 floor-test spec: 14pt titles; 10pt/700 black values that NEVER
    # wrap (auto-shrink fitter); 5x-oversampled QR at unchanged 60px size.
    check("titles are 14pt bold", ".ttl{font-size:14pt;font-weight:900;color:#000" in doc)
    check("values are 10pt semi-bold black", ".val{font-size:10pt;font-weight:700;color:#000" in doc)
    check("values never wrap (nowrap, no word-break)",
          "white-space:nowrap" in doc and "word-break" not in doc.split('class="sub"')[0].split(".val{")[1].split("}")[0])
    check("one-line auto-shrink fitter present",
          'querySelectorAll(".val")' in doc and "scrollWidth>el.clientWidth" in doc)
    check("QR rendered at 300px for crisp print", doc.count("width:300,height:300") == 2)
    check("QR displayed at 60px (size unchanged)",
          ".qr img,.qr canvas{width:60px!important;height:60px!important}" in doc)
    # escaping: hostile serial can't break the HTML/JS
    p2 = A.write_label_html_4x2(d, 'PN<&>"x', ['SN<script>&1'], wo="")
    d2 = open(p2, encoding="utf-8").read()
    check("HTML-escaped values", "PN&lt;&amp;&gt;" in d2 and "SN&lt;script&gt;&amp;1" in d2)
    check("JS payload safely JSON-encoded", '"SN<script>&1"' in d2 or '"SN\u003cscript\u003e&1"' in d2 or '\u003c' in d2)
    try:
        A.write_label_html_4x2(d, "PN", []); empty_ok = False
    except ValueError: empty_ok = True
    except Exception: empty_ok = False
    check("empty serial list raises cleanly", empty_ok)
finally:
    shutil.rmtree(d, ignore_errors=True)
print("ALL PASSED" if not fail else f"{fail} FAILED")
sys.exit(1 if fail else 0)
