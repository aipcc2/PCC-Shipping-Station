"""Shipping Slip + Shipping Labels (Apex-identical HTML printing) — generator /
fidelity half (Dhruvil 2026-09-24). Temp dirs only; no Tk; no share access.

FIDELITY: the app's generated #lbl-ship / #lbl-shipl markup is parsed with
html.parser and its element sequence (tag, id, class, inline style, input type,
nesting depth, the remaining attributes such as colspan/name/value, and every
STATIC text node — field labels, headers, 'Total:', 'Approved By:', 'Box'/'of')
must equal Apex 3.11's own markup (frozen copy in tests/fixtures) after the
documented substitutions:
  - the .box-gate and .label-nav blocks, <option>s and <datalist> are not part
    of the printed label and are omitted;
  - text/date <input>s and <select>s become static <span>s (or a <div> for the
    block-level #ship-via and the box-table cells) with the SAME class names
    and inline styles; checkboxes/radios stay <input>s;
  - the .label gets 'active'; #ship-form/#shipl-form inline display:none ->
    display:flex (what Apex's unlock sets);
  - box rows come from Apex's buildBoxTable row template (checked separately);
  - the ONE deliberate difference (Dhruvil 2026-09-25): each label barcode row
    carries the wrappers that centre its barcode under its value (an invisible
    copy of the row's title + a box around the <svg>); they are undone for the
    comparison and pinned on their own (section 5b).
Every CSS rule Apex applies to these labels must appear VERBATIM in the page.
Run: python test_slip_labels.py   (or > out.txt on a cp1252 console)
"""
import sys, os, re, shutil, datetime, json
from pathlib import Path
from html.parser import HTMLParser
SCRIPTS = r"C:\Users\Dsavaliya\Claude Cowork\PCC Shipping App\Scripts"
sys.path.insert(0, SCRIPTS)
import app_logic as A

fail = 0
def check(n, c):
    global fail
    print(("  ok  " if c else " FAIL ") + n)
    if not c: fail += 1

HERE = Path(__file__).resolve().parent
APEX_CANDIDATES = [
    HERE / "fixtures" / "PCC_Apex_Charlie_3.11.html",
    Path(os.environ.get("TEMP", "")) / "claude" / "C--Users-Dsavaliya-Claude-Cowork-PCC-Shipping-App"
        / "c9de82cb-c3b7-4550-8ccd-576abbe81514" / "scratchpad" / "slip" / "PCC_Apex_Charlie_3.11.html",
]
APEX = next((p for p in APEX_CANDIDATES if p.is_file()), None)
check("frozen Apex source available (tests/fixtures/PCC_Apex_Charlie_3.11.html)", APEX is not None)
apex_src = open(APEX, encoding="utf-8").read() if APEX else ""

# ── html.parser element-sequence extractor ───────────────────────────────────
VOID = {"input", "br", "img", "hr", "meta", "link", "col", "area", "base", "embed", "source", "track", "wbr"}

class Seq(HTMLParser):
    """Element sequence of the subtree rooted at id=root_id. `drop(tag, attrs)`
    removes an element and its subtree; `leaf(tag, attrs)` keeps the element
    but skips its children. Records per element (tag, id, class, style, type,
    parent tag, depth below the root, the OTHER attributes sorted) in .seq and
    the element's OWN text (its direct text nodes) in .own."""
    def __init__(self, root_id, drop=None, leaf=None):
        super().__init__(convert_charrefs=True)
        self.root_id = root_id
        self.drop = drop or (lambda t, a: False)
        self.leaf = leaf or (lambda t, a: False)
        self.stack = []            # (tag, id, index into seq or None)
        self.seq = []; self.own = []
        self.in_root = False; self.root_depth = None; self.skip_depth = None; self.done = False
    def _rec(self, tag, a, depth):
        parent = self.stack[-1][0] if self.stack else ""
        rest = tuple(sorted((k, v or "") for k, v in a.items() if k not in ("id", "class", "style", "type")))
        self.seq.append((tag, a.get("id", ""), a.get("class", ""), a.get("style", ""), a.get("type", ""), parent,
                         depth - self.root_depth, rest))
        self.own.append("")
        return len(self.seq) - 1
    def handle_starttag(self, tag, attrs):
        if self.done:
            return
        a = dict(attrs); depth = len(self.stack); idx = None
        if not self.in_root and a.get("id") == self.root_id:
            self.in_root = True; self.root_depth = depth
        if self.in_root and self.skip_depth is None:
            if self.drop(tag, a) and a.get("id") != self.root_id:
                self.skip_depth = depth
            else:
                idx = self._rec(tag, a, depth)
                if self.leaf(tag, a):
                    self.skip_depth = depth
        if tag not in VOID:
            self.stack.append((tag, a.get("id"), idx))
    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in VOID and self.stack and self.stack[-1][0] == tag:
            self.stack.pop()
            self._after_pop()
    def _after_pop(self):
        depth = len(self.stack)
        if self.skip_depth is not None and depth <= self.skip_depth:
            self.skip_depth = None
        if self.in_root and self.root_depth is not None and depth <= self.root_depth:
            self.done = True; self.in_root = False
    def handle_endtag(self, tag):
        if tag in VOID or self.done:
            return
        if any(t == tag for t, _, _ in self.stack):
            while self.stack:
                t, _, _ = self.stack.pop()
                if t == tag: break
            self._after_pop()
    def handle_data(self, data):
        if self.in_root and self.skip_depth is None and not self.done and self.stack and self.stack[-1][2] is not None:
            self.own[self.stack[-1][2]] += data

def seq_of(html, root_id, drop=None, leaf=None):
    p = Seq(root_id, drop, leaf); p.feed(html); p.close()
    return p

def is_field(tag, typ):
    return tag == "select" or (tag == "input" and typ in ("text", "date"))

DIV_IDS = {"ship-via"}   # block-level control -> <div>; everything else in a flex row -> <span>
VALUE_IDS = ("ship-", "shipl-", "bc-", "qr-")   # ids whose text is the shipment's VALUE (empty in Apex's static markup)

def norm(s): return " ".join(str(s).split())

def project(p, apex):
    """Comparable tuples (tag, id, class, style, type, depth, attrs, text) for
    one parsed subtree. Apex side: the documented substitutions applied. Both
    sides: no text and no attributes for the substituted controls (Apex's
    inputs/selects vs the app's spans/divs) and for the value-bearing ids, no
    attributes on the run-time-decorated <svg>; 'checked' and on* handlers
    ignored (the app freezes its checkboxes/radios with onclick). Everything
    else — the static label text, colspan, name/value, nesting depth — must
    match Apex exactly."""
    out = []
    for (tag, i, cls, style, typ, parent, depth, rest), text in zip(p.seq, p.own):
        sub = is_field(tag, typ) or (not apex and tag in ("span", "div") and parent == "td")
        if apex and is_field(tag, typ):
            tag = "div" if (i in DIV_IDS or parent == "td") else "span"; typ = ""
        if apex and i in ("lbl-ship", "lbl-shipl") and cls == "label":
            cls = "label active"
        if apex and i in ("ship-form", "shipl-form"):
            style = style.replace("display:none;", "display:flex;")
        valued = i.startswith(VALUE_IDS)
        attrs = () if (sub or valued or tag == "svg") else tuple((k, v) for k, v in rest if k != "checked" and not k.startswith("on"))
        own = "" if (sub or valued) else norm(text)
        out.append((tag, i, cls, "" if tag == "svg" else style, typ, depth, attrs, own))
    return out

def first_diff(a, b):
    for k, (x, y) in enumerate(zip(a, b)):
        if x != y:
            return f"#{k}: apex={x} app={y}"
    return f"length apex={len(a)} app={len(b)}"

apex_drop = lambda t, a: ("box-gate" in a.get("class", "") or "label-nav" in a.get("class", "")
                          or t in ("option", "datalist"))
app_leaf = lambda t, a: a.get("id") == "ship-box-tbody" or t == "svg"

# ── sample shipment ──────────────────────────────────────────────────────────
def sample():
    d = A.slip_data_template()
    d.update(dict(customer="377 Anduril", cpn="125-1030-970/C", po="PO-778", wo="16452.1", eng="4092",
                  rev="C", qty="15", date="2026-09-24", prepared_by="Dhruvil Savaliya", notes="2 boxes",
                  boxes=[{"size": "10x10x5", "qty": "10", "weight": "1.25"},
                         {"size": "", "qty": "5", "weight": "2"}],
                  customer_parts="Yes", fair=True, si=False, ppap=False, cofc=True,
                  ship_via="UPS GND", label_notes="fragile", slip_code="4092-16452.1-DS-SLIP-20260924-1234",
                  labels_code="4092-16452.1-XX-LABEL-20260924-4321"))
    return d

slip_doc = A.build_slip_html(sample())
labels_doc = A.build_labels_html(sample())
label_pages = labels_doc.split('<div class="sheet">')[1:]      # one chunk per 4x6 page (box)
# undo the 2026-09-25 barcode-alignment wrappers (docstring) for the Apex element-sequence comparison
BC_WRAP = re.compile(r'<span class="bc-sp" aria-hidden="true"><span class="fi-lbl shipl-lbl">([^<]*)</span></span>'
                     r'<span class="bc-box" style="min-width:min\((\d+)px,100%\)">(<svg.*?</svg>)</span>', re.S)
def undo_bc_align(html):
    return BC_WRAP.sub(lambda m: m.group(3), html).replace('<div class="barcode-row bc-under">', '<div class="barcode-row">')
labels_cmp = undo_bc_align(labels_doc)
label_pages_cmp = labels_cmp.split('<div class="sheet">')[1:]

# ══ 1. FIDELITY — element sequence vs the frozen Apex markup ═════════════════
if APEX:
    ax = seq_of(apex_src, "lbl-ship", drop=apex_drop)
    ap = seq_of(slip_doc, "lbl-ship", leaf=app_leaf)
    exp, got = project(ax, True), project(ap, False)
    check("slip: Apex #lbl-ship element sequence (tag/id/class/style/depth/attrs/static text) reproduced", exp == got)
    if exp != got: print("      first difference:", first_diff(exp, got))
    check("slip: the static texts really are compared (labels, headers, Total:, Approved By:)",
          {t for *_, t in exp} >= {"Printed Circuits Corp.", "SHIPPING SLIP", "Customer P/N", "Weight (lbs)", "Total:", "Approved By:", "C of C", "Boxes"})
    check("slip: nesting + colspan compared (tfoot Total cell colspan=2 at depth 7, code display inside .hdr-right at depth 3)",
          any(t == "td" and ("colspan", "2") in a and d == 7 for t, i, c, s, y, d, a, x in got)
          and next(d for t, i, *_r, d, a, x in got if i == "ship-code-display") == 3)
    check("slip: Apex sequence is non-trivial (>60 elements)", len(exp) > 60)
    check("slip: nothing from the gate / options leaked in",
          not any("box-gate" in s[2] or s[0] in ("option", "select", "datalist") for s in got))

    ax2 = seq_of(apex_src, "lbl-shipl", drop=apex_drop)
    ap2 = seq_of(labels_cmp, "lbl-shipl", leaf=app_leaf)
    exp2, got2 = project(ax2, True), project(ap2, False)
    check("labels: Apex #lbl-shipl element sequence (incl. depth/attrs/static text) reproduced (first box)", exp2 == got2)
    if exp2 != got2: print("      first difference:", first_diff(exp2, got2))
    # the Seq parser stops after the first #lbl-shipl root — pages 2..N are compared one by one
    per_page = [project(seq_of(pg, "lbl-shipl", leaf=app_leaf), False) for pg in label_pages_cmp]
    check("labels: EVERY page reproduces the Apex #lbl-shipl sequence (2 boxes = 2 pages, each 'label active')",
          len(per_page) == 2 and all(pg == exp2 for pg in per_page))
    for k, pg in enumerate(per_page, 1):
        if pg != exp2: print(f"      page {k} first difference:", first_diff(exp2, pg))
    check("labels: 'Box' / 'of' / 'Weight (lbs):' / 'SHIPPING LABEL' texts compared",
          {t for *_, t in exp2} >= {"Box", "of", "Weight (lbs):", "SHIPPING LABEL", "Work Order/Lot"})
    check("labels: nav bar / gate / datalist not printed",
          not any("label-nav" in s[2] or "box-gate" in s[2] or s[0] == "datalist" for s in got2))
    check("labels: the seven barcode <svg> ids in Apex order",
          [s[1] for s in got2 if s[0] == "svg"] == ["bc-cpn", "bc-po", "bc-wo", "bc-eng", "bc-rev", "bc-date", "bc-qty"])

    # box rows: Apex buildBoxTable's row template (from its JS) vs the app's rows
    m = re.search(r"tr\.innerHTML='(.*?)';tbody\.appendChild", apex_src)
    check("Apex buildBoxTable row template found", bool(m))
    if m:
        tpl = re.sub(r"<option value=\"\">--</option>'\+sizes\.map\(.*?\)\.join\(''\)\+'", "", m.group(1))
        # the template is a <tr>'s innerHTML -> expected rows = <tr> + the template's cells, i/N per row
        row_exp = []
        for i in (1, 2):
            tp = seq_of("<tbody id=\"T\"><tr>" + tpl.replace("'+i+'/'+count+'", f"{i}/2") + "</tr></tbody>", "T", drop=lambda t, a: t == "option")
            row_exp += project(tp, True)[1:]                       # [1:] drops the wrapper tbody
        rows = project(seq_of(slip_doc, "ship-box-tbody", leaf=lambda t, a: t == "svg"), False)[1:]
        check("slip: box rows follow Apex's buildBoxTable template (2 rows, i/N text, depth, attrs)", rows == row_exp)
        if rows != row_exp: print("      first difference:", first_diff(row_exp, rows))
    check("slip: box rows read 1/2 and 2/2", re.search(r'col-box"[^>]*>1/2<', slip_doc) and re.search(r'>2/2<', slip_doc))
    check("slip: an empty box size prints '--' like Apex's blank option", slip_doc.count(">--</div>") == 1)

# ══ 2. CSS — every rule Apex applies, verbatim ═══════════════════════════════
def apex_rules():
    """selector -> [verbatim 'selector{decls}' strings] from the Apex file, plus
    the flattened CSS they came from (for rule positions)."""
    css = "\n".join(re.findall(r"<style[^>]*>(.*?)</style>", apex_src, re.S))
    css = re.sub(r"/\*.*?\*/", lambda m: " " * len(m.group(0)), css, flags=re.S)
    rules = {}
    for m in re.finditer(r"([^{}]+)\{([^{}]*)\}", css):
        sel = m.group(1).strip()
        rules.setdefault(sel, []).append(sel + "{" + m.group(2) + "}")
    return rules, css

if APEX:
    R, apex_css = apex_rules()
    styles = "\n".join(re.findall(r"<style[^>]*>(.*?)</style>", slip_doc, re.S))
    lstyles = "\n".join(re.findall(r"<style[^>]*>(.*?)</style>", labels_doc, re.S))
    REQUIRED = [  # (selector, only rules containing this text; None = all of them)
        ("*", None), (".label", None), (".label.active", None), (".lbl-hdr-split", None), (".hdr-co", None),
        (".hdr-right", None), (".hdr-title", None), (".label-code", None), (".ship-hdr .hdr-co", None),
        (".shipl-hdr .hdr-co", None), (".job-info", None), (".fi-row", None), (".fi-row:last-child", None),
        (".fi-lbl", None), (".ship-lbl", None), (".shipl-lbl", None), (".fi-input", None),
        (".fi-row.qty .fi-input", None), (".fi-select", None), (".fi-date", None), (".footer-qr", None),
        (".footer-qr canvas,.footer-qr img", None), (".opt", None), (".opt input[type=checkbox]", None), (".opt-ship input", None),
        (".box-table", None), (".box-table th", None), (".box-table td", None),
        (".box-table input,.box-table select", None), (".col-box", None), (".col-size", None),
        (".col-qty", None), (".col-wt", None), (".box-table tfoot td", None), (".ship-section-lbl", None),
        (".ship-approved-footer", None), (".barcode-row", None), (".barcode-row svg", None),
        ("@page", None), (".box-gate", "display:none"), ("html,body", None), ("body", "display:block"),
        (".label:not(.active)", None), ("#pack-code-display,#shipl-code-display", None), (":root", None),
    ]
    missing = []
    n_rules = 0
    for sel, only in REQUIRED:
        rs = [r for r in R.get(sel, []) if (only is None or only in r)]
        if not rs:
            missing.append(f"{sel} (not found in Apex)"); continue
        for r in rs:
            n_rules += 1
            if r not in styles: missing.append(f"slip: {r[:70]}")
            if r not in lstyles: missing.append(f"labels: {r[:70]}")
    check(f"every listed Apex CSS rule present VERBATIM in both pages ({n_rules} rules)", not missing)
    for x in missing[:8]: print("      missing:", x)
    # ORDER matters for the cascade: Customer / Customer P/N / PO / WO / Engineering carry
    # class="fi-select fi-input" and get .fi-select's 12px only because it follows .fi-input
    # in Apex -> the required rules must sit in the generated Apex block in Apex's own order
    req = sorted({(apex_css.find(r), r) for sel, only in REQUIRED for r in R.get(sel, []) if only is None or only in r})
    pos = [styles.find(r) for _, r in req]; lpos = [lstyles.find(r) for _, r in req]
    check("Apex rules appear in the page in Apex's relative order (cascade identical), both pages",
          all(p >= 0 for p in pos + lpos) and all(a <= b for a, b in zip(pos, pos[1:])) and all(a <= b for a, b in zip(lpos, lpos[1:])))
    # the FIRST <style> block (the copied Apex CSS) may hold NOTHING BUT Apex rules, as an ordered subsequence of
    # Apex's own CSS — an override appended inside it (`.fi-lbl{min-width:60px;}`) passed "present verbatim" +
    # "in order" while changing what prints
    for name, doc in (("slip", slip_doc), ("labels", labels_doc)):
        first = re.sub(r"/\*.*?\*/", "", re.findall(r"<style[^>]*>(.*?)</style>", doc, re.S)[0], flags=re.S)
        block_rules = [m.group(1).strip() + "{" + m.group(2) + "}" for m in re.finditer(r"([^{}]+)\{([^{}]*)\}", first)]
        p = 0; alien = []
        for r in block_rules:
            q = apex_css.find(r, p)
            if q < 0: alien.append(r)
            else: p = q
        check(f"{name}: the Apex <style> block holds ONLY Apex rules ({len(block_rules)}), each verbatim and in Apex's file order — nothing app-side hides inside it",
              len(block_rules) > 50 and not alien)
        for r in alien[:3]: print("      not an Apex rule (or out of order):", r[:90])
    # the app-side <style> blocks after the Apex block: only the page chrome (body / .bar / .sheet /
    # @page) plus EXACTLY the three documented rules for the control stand-ins — nothing else may
    # touch an Apex class or id (an override there changes what prints while every Apex rule
    # is still "verbatim")
    SUBSTITUTE_RULES = [
        "span.fi-input,span.fi-select,span.fi-date,#ship-via,.box-table td>div,#shipl-weight{white-space:nowrap;overflow:hidden;}",
        ".box-table td>div,#shipl-weight{color:#000;}",
        '#shipl-weight:empty::before,#ship-via:empty::before{content:"\\00a0";}']
    later = "\n".join(re.findall(r"<style[^>]*>(.*?)</style>", slip_doc, re.S)[1:])
    llater = "\n".join(re.findall(r"<style[^>]*>(.*?)</style>", labels_doc, re.S)[1:])
    check("stand-ins clip on one line like Apex's inputs/selects (nowrap + overflow:hidden) and the two UA-coloured controls print black — the three documented app-side rules, both pages",
          all(r in later and r in llater for r in SUBSTITUTE_RULES))
    # final-fidelity-1: a blank <input>/<select> keeps its control height, a blank span/div has no
    # line box — the third rule gives the EMPTY stand-in a &nbsp; line box (prints nothing). For
    # :empty to match, the stand-in must be emitted with nothing at all between its tags.
    _empty_rule = SUBSTITUTE_RULES[2]
    _later_of = lambda doc: "\n".join(re.findall(r"<style[^>]*>(.*?)</style>", doc, re.S)[1:])
    _blank_w = A.build_labels_html(dict(sample(), boxes=[{"size": "", "qty": "3", "weight": ""}]))
    _blank_v = A.build_slip_html(dict(sample(), ship_via=""))
    check("blank weight: the label page carries the :empty::before nbsp rule (app-side block, not Apex's) and #shipl-weight is emitted truly empty (<span ...></span>) so :empty matches",
          _empty_rule in _later_of(_blank_w)
          and _empty_rule not in re.findall(r"<style[^>]*>(.*?)</style>", _blank_w, re.S)[0]
          and re.search(r'<span id="shipl-weight" style="[^"]*"></span>', _blank_w) is not None)
    check("blank Ship Via: the slip page carries the same rule and #ship-via is emitted truly empty (<div ...></div>) so :empty matches",
          _empty_rule in _later_of(_blank_v)
          and _empty_rule not in re.findall(r"<style[^>]*>(.*?)</style>", _blank_v, re.S)[0]
          and re.search(r'<div class="fi-select" id="ship-via" style="[^"]*"></div>', _blank_v) is not None)
    check("the rule is an :empty pseudo rule only — filled weight / Ship Via still print their value (no nbsp injected into the markup)",
          re.search(r'<span id="shipl-weight" style="[^"]*">1\.25</span>', labels_doc) is not None
          and re.search(r'<div class="fi-select" id="ship-via" style="[^"]*">UPS GND</div>', slip_doc) is not None
          and "&nbsp;" not in _blank_w and "&nbsp;" not in _blank_v and "\u00a0" not in _blank_w and "\u00a0" not in _blank_v)
    # the barcode-alignment rules (Dhruvil 2026-09-25): pinned verbatim, app-side only, the only rules naming
    # .barcode-row / .bc-sp / .bc-box outside Apex's own block
    ALIGN_RULES = [".barcode-row.bc-under{padding-left:0;padding-right:0;justify-content:flex-start;}",
                   ".bc-sp{flex:0 1 auto;min-width:0;overflow:hidden;visibility:hidden;display:flex;}",
                   ".bc-sp>.fi-lbl{box-sizing:content-box;padding-right:4px;}",
                   ".bc-box{flex:1 1 auto;display:flex;justify-content:center;padding:0 8px;}"]
    _first_l = re.findall(r"<style[^>]*>(.*?)</style>", labels_doc, re.S)[0]
    check("barcodes under their values: the four alignment rules sit VERBATIM in the app-side block of the labels page, none inside Apex's block, and no other app-side rule names .barcode-row / .bc-sp / .bc-box",
          all(r in llater and r not in _first_l for r in ALIGN_RULES)
          and [m.group(1).strip() for m in re.finditer(r"([^{}]+)\{([^{}]*)\}", re.sub(r"/\*.*?\*/", "", llater, flags=re.S))
               if re.search(r"barcode-row|bc-sp|bc-box", m.group(1))] == [r.split("{")[0] for r in ALIGN_RULES])
    rest_css = re.sub(r"/\*.*?\*/", "", later, flags=re.S)
    for r in SUBSTITUTE_RULES + ALIGN_RULES: rest_css = rest_css.replace(r, "")
    sels = [m.group(1).strip() for m in re.finditer(r"([^{}]+)\{([^{}]*)\}", rest_css)]
    bad = [s for s in sels if not re.fullmatch(r"body|\.bar( \w+)?|\.sheet(:last-of-type)?|@page", s)]
    check("no other app-side rule reaches an Apex class/id (selectors limited to body / .bar* / .sheet* / @page)", bool(sels) and not bad)
    for s in bad[:5]: print("      offending selector:", s)
    # the app-side PRINT GEOMETRY is load-bearing (the .sheet wrapper is what puts each label on its own 4x6 page):
    # its rules are pinned verbatim, and they are the ONLY .sheet rules on the page
    PRINT_GEOMETRY = [".sheet{position:relative;width:max-content;margin:16px auto;}",
                      ".sheet{width:4in;height:6in;margin:0;overflow:hidden;page-break-after:always;break-after:page;}",
                      ".sheet:last-of-type{page-break-after:auto;break-after:auto;}"]
    later_rules = [m.group(1).strip() + "{" + m.group(2) + "}" for m in re.finditer(r"([^{}]+)\{([^{}]*)\}", re.sub(r"/\*.*?\*/", "", later, flags=re.S))]
    llater_rules = [m.group(1).strip() + "{" + m.group(2) + "}" for m in re.finditer(r"([^{}]+)\{([^{}]*)\}", re.sub(r"/\*.*?\*/", "", llater, flags=re.S))]
    check("app-side print geometry pinned verbatim: the screen .sheet wrapper, then inside @media print the 4in x 6in / margin:0 / page-break rules — and no other .sheet rule, both pages",
          [r for r in later_rules if r.startswith(".sheet")] == PRINT_GEOMETRY and [r for r in llater_rules if r.startswith(".sheet")] == PRINT_GEOMETRY
          and later.index("@media print{") < later.index(PRINT_GEOMETRY[1]) < later.index(PRINT_GEOMETRY[2]))
    APEX_PAGE_RULES = {"@page{margin:0;}", "@page{size:4in 6in;margin:0}"}
    nc_slip = re.sub(r"/\*.*?\*/", "", slip_doc, flags=re.S); nc_labels = re.sub(r"/\*.*?\*/", "", labels_doc, flags=re.S)   # Apex's own comments mention @page
    check("the ONLY @page rules on either page are Apex's two (a later @page{margin:…} wins the cascade and clips the label)",
          set(re.findall(r"@page\{[^}]*\}", nc_slip)) == APEX_PAGE_RULES and set(re.findall(r"@page\{[^}]*\}", nc_labels)) == APEX_PAGE_RULES
          and nc_slip.count("@page") == 2 and nc_labels.count("@page") == 2)
    check("body colour = Apex's --ap-text (#142b1f: what the i/N cells and totals inherit on Apex's page)", "body{background:#4a5b50;color:#142b1f;" in later)
    # the 4x6 size rules that name #lbl-ship / #lbl-shipl (screen + 3.9 print block)
    size_rules = [r for sel, rs in R.items() if "#lbl-ship.active" in sel and "#lbl-shipl.active" in sel for r in rs]
    check("4x6 print-size rules naming #lbl-ship/#lbl-shipl present verbatim (>=2)",
          len(size_rules) >= 2 and all(r in styles and r in lstyles for r in size_rules))
    noprint = [r for sel, rs in R.items() if ".no-print" in sel for r in rs]
    check("print block hides .no-print (the hint bar) - verbatim Apex selector line",
          noprint and all(r in styles for r in noprint))
    check("@media print block present", "@media print{" in styles and "@media print{" in lstyles)
    check(".label-code theme override + :root var present (prints #8a958c on the slip)",
          ".label-code{color:var(--ap-text-3)!important;}" in styles and "--ap-text-3:#8a958c;" in styles)
    check("print rule hides the LABEL's code display (Apex prints no code on labels)",
          "#pack-code-display,#shipl-code-display{display:none!important;}" in lstyles)
    check("Roboto font stack rule verbatim (Apex line 14 + theme line)",
          styles.count("font-family:'Roboto','Helvetica Neue',Arial,sans-serif!important") == 2)
check("page rule @page{size:4in 6in;margin:0} (Apex setPrintSize)",
      "@page{size:4in 6in;margin:0}" in slip_doc and "@page{size:4in 6in;margin:0}" in labels_doc)
check("Apex CSS block is the first <style> (app additions come after it)",
      slip_doc.index(".label{width:4in;height:6in;") < slip_doc.index(".sheet{"))

# ══ 3. Page chrome, QR, auto-print ═══════════════════════════════════════════
check("qrcode.js embedded inline, no CDN script", "var QRCode;!function" in slip_doc and "cdnjs" not in slip_doc)
check("slip footer QR: text WO+'-1', 38x38, level M into #qr-ship",
      'var WO="16452.1"' in slip_doc and "text:WO+'-1',width:38,height:38" in slip_doc
      and "correctLevel:QRCode.CorrectLevel.M" in slip_doc and "getElementById('qr-ship')" in slip_doc)
check("labels page has no footer QR (Apex's label has none)", "new QRCode(" not in labels_doc)
check("no WO -> no QR call runs (guard like Apex's refreshFooterQR)", "if(!WO)return;" in slip_doc)
check("printer hint names the Zebra ZP 450 CTP, 4 x 6, margins none",
      "Zebra ZP 450 CTP" in slip_doc and "4 &times; 6 in" in slip_doc and "Margins to None" in labels_doc)
check("hint bar is class no-print (hidden by Apex's print block)", '<div class="no-print bar">' in slip_doc)
check("Print button + auto print", 'onclick="window.print()"' in slip_doc and "window.print();" in slip_doc)
check("no network: SLIP_FONT_LINK is off, no Google Fonts <link>, nothing on either page fetches a URL",
      A.SLIP_FONT_LINK is False and "fonts.googleapis" not in slip_doc and "fonts.googleapis" not in labels_doc
      and not re.search(r'(src|href)=["\']https?://', slip_doc + labels_doc))
check("auto print: the 4x2 page's plain timer — no document.fonts wait when there is no web font to wait for",
      "setTimeout(function(){window.print();},600);" in slip_doc and "document.fonts" not in slip_doc and "document.fonts" not in labels_doc)
_ofl = A.SLIP_FONT_LINK
try:
    A.SLIP_FONT_LINK = True; _on = A.build_slip_html(sample())
    check("the one-line switch: SLIP_FONT_LINK=True puts Apex's Roboto <link> AFTER the embedded script and waits (bounded) for the font",
          "fonts.googleapis.com/css2?family=Roboto" in _on and _on.index("</script>") < _on.index("fonts.googleapis")
          and "document.fonts.ready" in _on and "setTimeout(_goPrint,2500)" in _on)
finally:
    A.SLIP_FONT_LINK = _ofl
check("pages are self-contained UTF-8 HTML with a title", slip_doc.startswith("<!DOCTYPE html>") and "<title>Shipping Slip - 16452.1</title>" in slip_doc)
# ── 3b. print copies (Dhruvil 2026-09-25): the page opened FOR PRINTING closes its tab after the print dialog;
#        the FILED pages keep plain auto-print (History > Open still shows them); 'Print both' is ONE page
ps_copy = A.build_slip_html(sample(), print_copy=True); pl_copy = A.build_labels_html(sample(), print_copy=True)
both_doc = A.build_slip_labels_html(sample())
check("filed pages: plain auto-print, never window.close (opened again from History they stay on screen)",
      all("setTimeout(function(){window.print();},600);" in x and "window.close" not in x and "_pccPrintThenClose" not in x for x in (slip_doc, labels_doc)))
check("print-then-close: window.print() first, the tab closes only when print() HELD the script (> 400 ms: Chrome's dialog was up), a second after it returns",
      A._PRINT_THEN_CLOSE_FN == ("function _pccPrintThenClose(){var t0=Date.now();window.print();"
                                 "if(Date.now()-t0>400){setTimeout(function(){window.close();},1000);}}")
      and A._PRINT_THEN_CLOSE_JS == A._PRINT_THEN_CLOSE_FN + "\nsetTimeout(_pccPrintThenClose,600);")
check("print copies (slip, labels, both) carry print-then-close and NOT the plain auto-print timer",
      all(A._PRINT_THEN_CLOSE_JS in x and "setTimeout(function(){window.print();},600);" not in x for x in (ps_copy, pl_copy, both_doc)))
check("a print copy IS the filed page apart from its print trigger (same markup, same CSS, same codes)",
      ps_copy.replace(A._slip_print_js(True), "") == slip_doc.replace(A._slip_print_js(False), "")
      and pl_copy.replace(A._slip_print_js(True), "") == labels_doc.replace(A._slip_print_js(False), ""))
_sheets = lambda doc: re.findall(r'<div class="sheet">.*?\n</div>(?=\n<div class="sheet">|\n<script>)', doc, re.S)
check("'Print both' = ONE page: the slip sheet first, then every box's label sheet, exactly as the two filed pages carry them",
      _sheets(both_doc) == _sheets(slip_doc) + _sheets(labels_doc) and len(_sheets(both_doc)) == 3
      and both_doc.count('id="lbl-ship"') == 1 and both_doc.count('id="lbl-shipl"') == 2
      and both_doc.index('id="lbl-ship"') < both_doc.index('id="lbl-shipl"'))
check("'Print both' page: one footer-QR script, the slip + labels title, the hint bar counts 1 + N pages, Apex CSS once",
      both_doc.count("getElementById('qr-ship')") == 1 and "<title>Shipping Slip + Labels - 16452.1</title>" in both_doc
      and "Shipping Slip + Labels &mdash; 3 pages" in both_doc and both_doc.count(".label{width:4in;height:6in;") == 1)
_ofl2 = A.SLIP_FONT_LINK
try:
    A.SLIP_FONT_LINK = True
    _onc = A.build_slip_labels_html(sample())
    check("with SLIP_FONT_LINK on, the print copy waits (bounded) for the web font and then prints-then-closes",
          "document.fonts.ready" in _onc and "setTimeout(_goPrint,2500)" in _onc and "_pccPrintThenClose();}" in _onc
          and "setTimeout(_pccPrintThenClose,600);" not in _onc)
finally:
    A.SLIP_FONT_LINK = _ofl2

# ══ 4. Values on the slip ════════════════════════════════════════════════════
check("slip header code shown", 'id="ship-code-display">4092-16452.1-DS-SLIP-20260924-1234<' in slip_doc)
check("slip fields as static text", 'id="ship-cust">377 Anduril<' in slip_doc and 'id="ship-wo">16452.1<' in slip_doc
      and 'id="ship-qty">15<' in slip_doc and 'id="ship-operator">Dhruvil Savaliya<' in slip_doc)
check("date shown as Chrome shows a date input (MM/DD/YYYY)", 'id="ship-date">09/24/2026<' in slip_doc)
check("WWYY date code on the slip (week 39 of 2026 -> 3926)", 'aria-label="Date Code in WWYY format">3926<' in slip_doc)
check("totals per calcBoxTotals: qty '15', weight '3.25 lbs'",
      'id="ship-total-qty" style="font-weight:900;font-size:10px;font-family:Arial,sans-serif;">15<' in slip_doc
      and 'id="ship-total-weight" style="font-weight:900;font-size:10px;font-family:Arial,sans-serif;">3.25 lbs<' in slip_doc)
check("customer parts radio Yes checked, No not", 'id="ship-cp-yes" value="yes" checked onclick="return false"' in slip_doc
      and 'id="ship-cp-no" value="no" onclick="return false"' in slip_doc)
check("FAIR + C of C checked, SI/PPAP not; never 'disabled'",
      'id="ship-fair" checked onclick' in slip_doc and 'id="ship-cofc" checked onclick' in slip_doc
      and 'id="ship-si" onclick' in slip_doc and 'id="ship-ppap" onclick' in slip_doc and "disabled" not in slip_doc)
check("ship via / notes", '>UPS GND</div>' in slip_doc and 'id="ship-notes">2 boxes<' in slip_doc)
e = sample(); e["boxes"] = []; e["customer"] = '<b>&"x'; e["customer_parts"] = "No"
ed = A.build_slip_html(e)
check("no boxes -> no rows, totals show em dash", '<tbody id="ship-box-tbody"></tbody>' in ed and ed.count(">\u2014<") == 2)
check("HTML-escaped values", "&lt;b&gt;&amp;&quot;x" in ed and "<b>&" not in ed)
check("customer parts No checked by default", 'id="ship-cp-no" value="no" checked' in ed)
check("datecode falls back to the slip date when blank", A.build_slip_html(dict(sample(), datecode_wwyy="")).count(">3926<") == 1)

# ══ 5. Labels: one page per box, per-box values, barcodes ════════════════════
check("labels: one #lbl-shipl per box (2)", labels_doc.count('id="lbl-shipl"') == 2 and labels_doc.count('<div class="sheet">') == 2)
check("labels: Box 1 of 2 / Box 2 of 2", 'id="shipl-box-current-print">1<' in labels_doc and 'id="shipl-box-current-print">2<' in labels_doc
      and labels_doc.count('id="shipl-box-total-print">2<') == 2)
check("labels: per-box qty + weight", 'id="shipl-qty" style="text-align:center;">10<' in labels_doc
      and 'id="shipl-qty" style="text-align:center;">5<' in labels_doc
      and 'id="shipl-weight" style="width:80px;border:none;border-bottom:1.5px solid #555;font-size:18px;font-weight:900;font-family:Arial,sans-serif;outline:none;text-align:center;">1.25<' in labels_doc)
check("labels: page 2 carries box 2's OWN qty (5) and weight (2), page 1 box 1's (10 / 1.25) — nothing copied from box 1; every page is 'label active'",
      [re.search(r'id="shipl-weight"[^>]*>([^<]*)<', pg).group(1) for pg in label_pages] == ["1.25", "2"]
      and [re.search(r'id="shipl-qty"[^>]*>([^<]*)<', pg).group(1) for pg in label_pages] == ["10", "5"]
      and [re.search(r'id="bc-qty"[^>]*>', pg) is not None and A.code128_svg(q, "bc-qty") in pg for pg, q in zip(label_pages, ("10", "5"))] == [True, True]
      and all(pg.count('<div class="label active" id="lbl-shipl">') == 1 for pg in label_pages) and len(label_pages) == 2)
check("labels: shared fields copied to every box (Apex shiplLoadBox)", labels_doc.count('id="shipl-cust" style="text-align:center;">377 Anduril<') == 2
      and labels_doc.count('id="shipl-note" style="text-align:center;">fragile<') == 2)
check("labels: header code shown (hidden in print by Apex's rule)", labels_doc.count('id="shipl-code-display">4092-16452.1-XX-LABEL-20260924-4321<') == 2)
check("labels: barcodes present for cpn/po/wo/eng/rev/date/qty",
      all(f'<svg id="{i}" width="' in labels_doc for i in ("bc-cpn", "bc-po", "bc-wo", "bc-eng", "bc-rev", "bc-date", "bc-qty")))
check("labels: date barcode encodes the ISO value (input.value), text shows MM/DD/YYYY",
      A.code128_svg("2026-09-24", "bc-date") in labels_doc and 'id="shipl-date" style="text-align:center;">09/24/2026<' in labels_doc)
check("labels: qty barcode per box", A.code128_svg("10", "bc-qty") in labels_doc and A.code128_svg("5", "bc-qty") in labels_doc)
l2 = A.build_labels_html(dict(sample(), po="", rev="C\u00e9", cpn=" 125 ", boxes=[{"qty": "", "weight": ""}]))
check("labels: empty value -> empty <svg> (no barcode), like Apex", '<svg id="bc-po"></svg>' in l2 and '<svg id="bc-qty"></svg>' in l2)
check("labels: non-printable chars stripped / value trimmed before encoding",
      A.code128_svg("C", "bc-rev") in l2 and A.code128_svg("125", "bc-cpn") in l2)
l3 = A.build_labels_html(dict(sample(), boxes=[]))
check("labels: no boxes recorded -> one label with the slip qty", l3.count('id="lbl-shipl"') == 1 and 'id="shipl-qty" style="text-align:center;">15<' in l3)
l5 = A.build_labels_html(dict(sample(), boxes=[{"qty": str(i), "weight": ""} for i in range(1, 6)]))
check("labels: N=5 boxes -> 5 pages", l5.count('id="lbl-shipl"') == 5 and l5.count('id="shipl-box-total-print">5<') == 5)

# ── 5b. barcodes under their values (Dhruvil 2026-09-25) ─────────────────────
_rows = re.findall(r'<span class="fi-lbl shipl-lbl">([^<]*)</span><span[^>]*id="shipl-(\w+)"[^>]*>[^<]*</span></div>\n'
                   r'      <div class="barcode-row bc-under">' + BC_WRAP.pattern + '</div>', label_pages[0], re.S)
check("every label barcode row is the aligned form (7 per label, none left in Apex's plain form)",
      labels_doc.count('<div class="barcode-row bc-under">') == 14 and '<div class="barcode-row">' not in labels_doc
      and len(BC_WRAP.findall(labels_doc)) == 14)
check("each barcode row's invisible title is the SAME text as its value row's title (so the same width), in Apex's row order",
      [(t, fid, sp) for t, fid, sp, w, svg in _rows] == [(t, fid, t) for t, fid in (("Customer P/N", "cpn"), ("Customer PO #", "po"),
       ("Work Order/Lot", "wo"), ("Engineering", "eng"), ("Revision", "rev"), ("Date", "date"), ("Qty", "qty"))])
check("each barcode's box keeps its full width + Apex's 8 px either side (the quiet zone): min-width = <svg> width + 16 (16 for an empty value), never wider than the row",
      all(int(w) == int(re.search(r'width="(\d+)px"', svg).group(1)) + 16 for *_x, w, svg in _rows)
      and '<span class="bc-box" style="min-width:min(16px,100%)"><svg id="bc-po"></svg></span>' in
          A.build_labels_html(dict(sample(), po="")))
check("the barcodes themselves are untouched (the exact JsBarcode <svg> inside the box) and aria-hidden hides the title copy",
      all(A.code128_svg(v, i) in labels_doc for v, i in (("125-1030-970/C", "bc-cpn"), ("PO-778", "bc-po"), ("16452.1", "bc-wo")))
      and labels_doc.count('<span class="bc-sp" aria-hidden="true">') == 14)

# ══ 6. Code128 = JsBarcode (auto code sets, DOM layout) ══════════════════════
from barcode.charsets import code128 as C128
CODES = list(C128.CODES); STOP = C128.STOP + "11"
def decode128(bits):
    assert len(bits) % 11 == 2 and bits.endswith(STOP), "bad length/stop"
    syms = [CODES.index(bits[i:i + 11]) for i in range(0, len(bits) - 13, 11)]
    start, data, chk = syms[0], syms[1:-1], syms[-1]
    assert (start + sum(i * v for i, v in enumerate(data, 1))) % 103 == chk, "checksum"
    cs = {103: "A", 104: "B", 105: "C"}[start]; out = ""; shift = None
    for v in data:
        s = shift or cs; shift = None
        if s == "C":
            if v < 100: out += f"{v:02d}"; continue
            if v == 100: cs = "B"; continue
            if v == 101: cs = "A"; continue
        else:
            if v < 96: out += chr(v + 32) if (s == "B" or v < 64) else chr(v - 64); continue
            if v == 98: shift = "B" if s == "A" else "A"; continue
            if v == 99: cs = "C"; continue
            if v == 100 and s == "A": cs = "B"; continue
            if v == 101 and s == "B": cs = "A"; continue
        raise AssertionError(f"unexpected symbol {v} in set {s}")
    return syms, out
ok = True
for v in ["16452.1", "125-1030-970/C", "2026-09-24", "10", "PO-778", "C", "a b", "ABC123DEF", "4092", "x", "12", "1"]:
    try:
        syms, txt = decode128(A.code128_modules(v))
        ok = ok and txt == v
    except Exception as ex:
        ok = False; print("      decode failed for", repr(v), ex)
check("Code128 modules decode back to the value (start/checksum/stop valid) for 12 samples", ok)
syms, _ = decode128(A.code128_modules("16452.1"))
check("JsBarcode auto set choice reproduced: '16452.1' = StartC 16 45 CODE-A '2' '.' '1'",
      syms[:-1] == [105, 16, 45, 101, 18, 14, 17])      # [:-1] drops the checksum symbol
import barcode as _pb
check("pure-digit / pure-alpha values match python-barcode's own Code128 bars",
      A.code128_modules("2026") == _pb.Code128("2026").build()[0] and A.code128_modules("ABC") == _pb.Code128("ABC").build()[0])
svg = A.code128_svg("10", "bc-qty")
bits = A.code128_modules("10")
runs = len(re.findall(r"1+", bits))
check("SVG has JsBarcode's attributes: width=ceil(n*1.3)px, height 26px, viewBox, transparent background",
      f'width="{-(-len(bits) * 13 // 10)}px" height="26px" x="0px" y="0px" viewBox="0 0 {-(-len(bits) * 13 // 10)} 26"' in svg
      and 'style="fill:transparent;"' in svg and '<g transform="translate(0, 0)" style="fill:#000000;">' in svg)
check("SVG: one <rect> per bar run at 1.3 px per module, 26 px high",
      svg.count('height="26"></rect>') == runs and '<rect x="0" y="0" width="2.6" height="26"></rect>' in svg)
check("SVG scales like JsBarcode's (viewBox present, no explicit distortion)", "preserveAspectRatio" not in svg)
# every bar's position and width, derived from the module string (NOT from code128_svg itself): rect k = the k-th
# run of '1's, x = first module x 1.3, width = run length x 1.3; the last rect is Code128's 2-module stop bar and
# ends on the last module (a stop bar one module early merges with the preceding space -> unreadable)
ok = True
for v in ("10", "16452.1", "2026-09-24", "125-1030-970/C", "PO-778"):
    bits = A.code128_modules(v)
    exp_runs = [(m.start(), len(m.group(0))) for m in re.finditer(r"1+", bits)]
    got = [(float(x), float(w)) for x, w in re.findall(r'<rect x="([\d.]+)" y="0" width="([\d.]+)" height="26"></rect>', A.code128_svg(v, "b"))]
    ok = ok and len(got) == len(exp_runs) > 10 and all(abs(x - s * 1.3) < 1e-6 and abs(w - n * 1.3) < 1e-6 for (x, w), (s, n) in zip(got, exp_runs))
    ok = ok and bits.endswith("11") and got and abs(got[-1][1] - 2.6) < 1e-6 and abs(got[-1][0] + got[-1][1] - len(bits) * 1.3) < 1e-6
check("SVG bar geometry = the module string bar for bar (x / width of EVERY rect, 5 values); the stop bar ends on the last module", ok)

# ══ 7. Codes, date code, totals helpers ══════════════════════════════════════
check("slip_form_code ENG-WO-INITIALS-SLIP-YYYYMMDD-RAND4",
      A.slip_form_code("4092", "16452.1", "Dhruvil Savaliya", "2026-09-24", 1234) == "4092-16452.1-DS-SLIP-20260924-1234")
check("initials: first letters of each name part, upper; blank -> XX",
      A.slip_form_code("1", "2", "anita patel", "2026-01-02", 1000).split("-")[2] == "AP"
      and A.slip_form_code("1", "2", "", "2026-01-02", 1000).split("-")[2] == "XX"
      and A.slip_form_code("1", "2", "Cher", "2026-01-02", 1000).split("-")[2] == "C")
check("blank ENG/WO -> 'ENG'/'WO'; whitespace -> '-'",
      A.slip_form_code("", "", "A B", "2026-01-02", 1000) == "ENG-WO-AB-SLIP-20260102-1000"
      and A.labels_form_code("40 92", "16452 .1", "2026-01-02", 1000).startswith("40-92-16452-.1-XX-LABEL-"))
check("labels code is always XX-LABEL", A.labels_form_code("4092", "16452.1", "2026-09-24", 4321) == "4092-16452.1-XX-LABEL-20260924-4321")
cre = re.compile(r"^4092-16452\.1-DS-SLIP-\d{8}-\d{4}$"); lre = re.compile(r"^4092-16452\.1-XX-LABEL-\d{8}-\d{4}$")
check("fresh codes match the regexes (date = today, random 4 digits)",
      cre.match(A.slip_form_code("4092", "16452.1", "Dhruvil Savaliya")) and lre.match(A.labels_form_code("4092", "16452.1"))
      and datetime.date.today().strftime("%Y%m%d") in A.slip_form_code("4092", "16452.1", "D S"))
check("random 4 digits in 1000..9999", all(1000 <= A.slip_code_rand() <= 9999 for _ in range(300)))
check("WWYY: 2026-09-24 -> 3926, 2026-01-01 -> 0126 (Apex calcDateCode)",
      A.slip_datecode_wwyy("2026-09-24") == "3926" and A.slip_datecode_wwyy("2026-01-01") == "0126")
check("WWYY: calendar-year YY exactly as Apex (2027-01-01 = ISO wk 53 -> 5327)", A.slip_datecode_wwyy("2027-01-01") == "5327")
check("WWYY accepts MM/DD/YYYY too; blank -> today", A.slip_datecode_wwyy("09/24/2026") == "3926" and len(A.slip_datecode_wwyy("")) == 4)
t = datetime.date.today()
if t.month not in (1, 12):
    dc = A.fmt_datecode()
    check("WWYY(today) is the app's YYWW fmt_datecode() swapped (mid-year)", A.slip_datecode_wwyy(t.isoformat()) == dc[2:] + dc[:2])
check("fmt_datecode() untouched (YYWW, ISO year)", re.fullmatch(r"\d{4}", A.fmt_datecode()) and A.fmt_datecode()[:2] == f"{t.isocalendar()[0] % 100:02d}")
check("calcBoxTotals: numeric prefixes like parseFloat, em dash when none",
      A.calc_box_totals([{"qty": "10", "weight": "1.25"}, {"qty": "5", "weight": "2"}, {"qty": "", "weight": ""}]) == ("15", "3.25 lbs")
      and A.calc_box_totals([{"qty": "1.5x", "weight": "abc"}]) == ("1.5", "\u2014")
      and A.calc_box_totals([]) == ("\u2014", "\u2014"))
check("toFixed(2) rounds half-up on the exact binary value (JS)", A._slip_js_to_fixed(0.125) == "0.13" and A._slip_js_to_fixed(1.005) == "1.00"
      and A._slip_js_to_fixed(0) == "0.00")
check("JS number strings: 15 -> '15', 3.5 -> '3.5'", A._slip_js_num(15.0) == "15" and A._slip_js_num(3.5) == "3.5")
_apex_list = lambda key: re.findall(r"'([^']*)'", re.search(key + r":\[([^\]]*)\]", apex_src).group(1)) if apex_src else []
check("box sizes start as Apex's DEFAULTS.boxsize verbatim; Ship Via = Apex's DEFAULTS.shipvia + 'Hand Carry' (ours, 2026-09-25) in alphabetical place",
      A.SLIP_BOXSIZE_OPTIONS == _apex_list("boxsize") and len(A.SLIP_BOXSIZE_OPTIONS) == 10
      and [x for x in A.SLIP_SHIPVIA_OPTIONS if x != "Hand Carry"] == _apex_list("shipvia")
      and A.SLIP_SHIPVIA_OPTIONS == ['Courier', 'Customer Pickup', 'Deliver', 'FedEx GND', 'FedEx Priority', 'Hand Carry', 'UPS 2nd Day', 'UPS GND'])
check("slip_data_template has every documented key",
      set(A.slip_data_template()) >= {"customer", "cpn", "po", "wo", "eng", "rev", "qty", "date", "datecode_wwyy", "prepared_by",
                                      "notes", "boxes", "customer_parts", "fair", "si", "ppap", "cofc", "ship_via", "label_notes",
                                      "slip_code", "labels_code", "printed", "outputs"})
check("SLIP_FOLDER constant", A.SLIP_FOLDER == "Shipping Slip")

# ══ 8. Files: placement, naming, never overwrite, leaf-only, offline fallback ═
import tempfile as _tf
root = _tf.mkdtemp(prefix="pcc_slip_")
faketmp = _tf.mkdtemp(prefix="pcc_slip_tmp_")
_orig_gettempdir = _tf.gettempdir
_tf.gettempdir = lambda: faketmp           # keep the fallback test inside our temp dir
try:
    wo = Path(root) / "16452.1"; wo.mkdir()
    d = sample(); d["slip_code"] = ""; d["labels_code"] = ""
    r = A.write_slip_files(str(wo), d, "both")
    sub = wo / "Shipping Slip"
    fd = A.fmt_file_date()
    check("files land in <WO>\\Shipping Slip\\ (leaf created)", sub.is_dir() and Path(r["slip_html"]).parent == sub and Path(r["labels_html"]).parent == sub)
    check("naming <eng>-<m-d-yy>-Shipping_Slip.html / -Shipping_Labels.html",
          Path(r["slip_html"]).name == f"4092-{fd}-Shipping_Slip.html" and Path(r["labels_html"]).name == f"4092-{fd}-Shipping_Labels.html")
    check("not a fallback; folder reported", r["fallback"] is False and r["folder"] == str(sub))
    check("fresh codes stamped into slip_data and used in the pages",
          cre.match(d["slip_code"]) and lre.match(d["labels_code"])
          and d["slip_code"] in open(r["slip_html"], encoding="utf-8").read()
          and d["labels_code"] in open(r["labels_html"], encoding="utf-8").read())
    check("outputs stamped into slip_data", d["outputs"] == {"slip_html": r["slip_html"], "labels_html": r["labels_html"], "pack_html": None})
    first_slip = open(r["slip_html"], encoding="utf-8").read()
    r2 = A.write_slip_files(str(wo), d, "both")
    check("reprint same day -> unique names, first files untouched",
          r2["slip_html"] != r["slip_html"] and Path(r2["slip_html"]).name.startswith(f"4092-{fd}-Shipping_Slip_")
          and open(r["slip_html"], encoding="utf-8").read() == first_slip and len(list(sub.glob("*.html"))) == 4)
    check("reprint draws a NEW code (once per print)", d["slip_code"] not in first_slip)
    r3 = A.write_slip_files(str(wo), d, "slip")
    check("which='slip' writes only the slip", r3["slip_html"] and r3["labels_html"] is None and len(list(sub.glob("*Labels*.html"))) == 2)
    r4 = A.write_slip_files(str(wo), d, "labels")
    check("which='labels' writes only the labels", r4["labels_html"] and r4["slip_html"] is None)
    try:
        A.write_slip_files(str(wo), d, "nope"); bad = False
    except ValueError:
        bad = True
    check("bad `which` raises", bad)
    # a file already there with the clean name is never overwritten
    wo2 = Path(root) / "16453.1"; (wo2 / "Shipping Slip").mkdir(parents=True)
    pre = wo2 / "Shipping Slip" / f"4092-{fd}-Shipping_Slip.html"; pre.write_text("KEEP", encoding="utf-8")
    r5 = A.write_slip_files(str(wo2), sample(), "slip")
    check("existing file with the clean name kept; new one gets a suffix", pre.read_text(encoding="utf-8") == "KEEP" and Path(r5["slip_html"]) != pre)
    # eng sanitised
    dd = sample(); dd["eng"] = 'AB/C:"x'
    r6 = A.write_slip_files(str(wo), dd, "slip")
    check("engineering # sanitised in the file name", Path(r6["slip_html"]).name.startswith("AB_C__x-"))
    # leaf-only: a missing WO folder is never created (no parents=True from a stored string)
    missing = Path(root) / "nope" / "16999.1"
    d7 = sample(); r7 = A.write_slip_files(str(missing), d7, "both")
    ftmp = Path(faketmp) / "PCC Shipping App - Print"
    check("missing WO folder -> nothing created there (no parents), fallback to the temp print folder",
          not (Path(root) / "nope").exists() and r7["fallback"] is True and Path(r7["slip_html"]).parent == ftmp
          and Path(r7["labels_html"]).parent == ftmp and r7["folder"] == str(ftmp))
    check("a temp-folder copy is NOT recorded as the page's output (outputs stay None; codes + page snapshots still stamped)",
          d7["outputs"] == {"slip_html": None, "labels_html": None, "pack_html": None} and cre.match(d7["slip_code"]) and d7["pages"]["slip"]["slip_code"] == d7["slip_code"]
          and A._is_temp_print_path(r7["slip_html"]) and not A._is_temp_print_path(r["slip_html"]))
    # the write itself refused AFTER the leaf resolved (share dropped between the check and the open, read-only
    # ACL, disk full): that page goes to the temp folder unrecorded, the page already on the share stays recorded
    wo9 = Path(root) / "16459.1"; wo9.mkdir()
    _orig_wnf = A._write_new_file
    def _deny_labels(path, text):
        if "Labels" in str(path) and Path(path).parent == wo9 / "Shipping Slip":
            raise PermissionError(13, "Access is denied", str(path))
        return _orig_wnf(path, text)
    A._write_new_file = _deny_labels
    try:
        d9 = sample(); r9b = A.write_slip_files(str(wo9), d9, "both"); raised = None
    except Exception as ex:
        r9b = {}; raised = ex
    finally:
        A._write_new_file = _orig_wnf
    check("write refused on the share for one page -> that page falls back to the temp folder (unrecorded), the other stays on the share (recorded), nothing raised",
          raised is None and r9b.get("fallback") is True and Path(r9b["slip_html"]).parent == wo9 / "Shipping Slip"
          and Path(r9b["labels_html"]).parent == ftmp and d9["outputs"] == {"slip_html": r9b["slip_html"], "labels_html": None, "pack_html": None}
          and r9b["folder"] == str(ftmp) and len(list((wo9 / "Shipping Slip").glob("*.html"))) == 1)
    if raised: print("      raised:", type(raised).__name__, raised)
    # … and when the LATER page can be written NOWHERE (share refused it AND the temp folder failed): no exception —
    # the slip already on the share keeps its record (the popup used to drop the whole dict and leave a page with an
    # unrecorded code in the WO folder), the failure is reported in 'errors', the labels code is not burnt
    wo13 = Path(root) / "16463.1"; wo13.mkdir()
    def _deny_labels_everywhere(path, text):
        if "Labels" in str(path):
            raise PermissionError(13, "Access is denied", str(path))
        return _orig_wnf(path, text)
    A._write_new_file = _deny_labels_everywhere
    try:
        d13 = sample(); d13["labels_code"] = "OLD-LABELS-CODE"; opened13 = []
        r13 = A.slip_print_pages(str(wo13), d13, "both", open_file=lambda p: opened13.append(p)); raised13 = None
    except Exception as ex:
        r13 = {}; raised13 = ex
    finally:
        A._write_new_file = _orig_wnf
    check("later page unwritable everywhere -> no exception: the slip is written on the share, recorded, opened and stamped; the labels page is reported in 'errors', not opened, not stamped, its code untouched",
          raised13 is None and Path(r13["slip_html"]).parent == wo13 / "Shipping Slip" and d13["outputs"] == {"slip_html": r13["slip_html"], "labels_html": None, "pack_html": None}
          and r13["labels_html"] is None and len(r13["errors"]) == 1 and r13["errors"][0][0] == "Shipping Labels" and "PermissionError" in r13["errors"][0][1]
          and opened13 == [r13["print_html"]] and Path(r13["print_html"]).name == f"4092-{fd}-Print_Slip.html" and Path(r13["print_html"]).parent == ftmp
          and d13["slip_code"] in open(r13["print_html"], encoding="utf-8").read() and 'id="lbl-shipl"' not in open(r13["print_html"], encoding="utf-8").read()
          and d13["printed"]["labels"] is None and bool(d13["printed"]["slip"]) and d13["labels_code"] == "OLD-LABELS-CODE"
          and d13["pages"]["labels"] is None and d13["pages"]["slip"]["slip_code"] == d13["slip_code"])
    if raised13: print("      raised:", type(raised13).__name__, raised13)
    def _deny_all(path, text): raise PermissionError(13, "Access is denied", str(path))
    A._write_new_file = _deny_all
    try:
        A.write_slip_files(str(wo13), sample(), "both"); raised_first = None
    except OSError as ex:
        raised_first = ex
    finally:
        A._write_new_file = _orig_wnf
    check("the FIRST page failing everywhere still raises (nothing of this print is on disk to lose); a normal write reports errors == []",
          isinstance(raised_first, PermissionError) and len(list((wo13 / "Shipping Slip").glob("*.html"))) == 1
          and A.write_slip_files(str(wo13), sample(), "slip")["errors"] == [])
    r8 = A.write_slip_files(None, sample(), "slip")
    check("no WO folder at all (offline) -> temp print folder, still writes", r8["fallback"] and Path(r8["slip_html"]).is_file())
    # a WO folder whose leaf can't be made (a FILE named 'Shipping Slip') -> fallback, nothing broken
    wo3 = Path(root) / "16454.1"; wo3.mkdir(); (wo3 / "Shipping Slip").write_text("x", encoding="utf-8")
    r9 = A.write_slip_files(str(wo3), sample(), "slip")
    check("leaf blocked by a file of the same name -> fallback to temp, file left alone",
          r9["fallback"] and (wo3 / "Shipping Slip").read_text(encoding="utf-8") == "x")

    # ══ 9. Re-emit from the History record (generator side) ═════════════════
    wo = Path(root) / "16455.1"; wo.mkdir()
    d = sample()
    r = A.write_slip_files(str(wo), d, "both")
    slip_txt = open(r["slip_html"], encoding="utf-8").read(); lab_txt = open(r["labels_html"], encoding="utf-8").read()
    rec = {"event_type": "close", "eng": "4092", "wo": "16455.1", "serials": [], "wo_folder": str(wo),
           "outputs": {}, "slip": json.loads(json.dumps(d))}   # as it would come back from History JSON
    os.remove(r["slip_html"]); os.remove(r["labels_html"])
    res = A.reemit_paperwork(rec)
    check("re-emit recreates BOTH missing pages under the ORIGINAL names",
          res["ok"] and sorted(Path(p).name for p in res["wrote"]) == sorted([Path(r["slip_html"]).name, Path(r["labels_html"]).name])
          and not res["errors"])
    check("re-emitted pages are byte-identical (recorded code, no new random digits)",
          open(r["slip_html"], encoding="utf-8").read() == slip_txt and open(r["labels_html"], encoding="utf-8").read() == lab_txt)
    check("re-emit needs no serials (slip pages carry no serials): the serial-less record was not refused",
          res["ok"] and not any("serials" in e for e in res["errors"]))
    res2 = A.reemit_paperwork(rec)
    check("re-emit again -> both kept, nothing written, nothing overwritten",
          not res2["wrote"] and len(res2["kept"]) == 2 and not res2["errors"] and open(r["slip_html"], encoding="utf-8").read() == slip_txt)
    os.remove(r["labels_html"])
    res3 = A.reemit_paperwork(rec)
    check("only the missing page is rebuilt", [Path(p).name for p in res3["wrote"]] == [Path(r["labels_html"]).name] and len(res3["kept"]) == 1)
    shutil.rmtree(wo / "Shipping Slip")
    res4 = A.reemit_paperwork(rec)
    check("subfolder gone but WO folder present -> leaf recreated, both rebuilt", res4["ok"] and len(res4["wrote"]) == 2 and (wo / "Shipping Slip").is_dir())
    # never printed = no recorded path AND no printed stamp: not fabricated even though the WO folder is right there
    rec_np = dict(rec, slip=dict(rec["slip"], outputs={"slip_html": None, "labels_html": None}, printed={"slip": None, "labels": None}))
    before = sorted(p.name for p in (wo / "Shipping Slip").glob("*"))
    res6 = A.reemit_paperwork(rec_np)
    check("pages never printed are not fabricated (WO folder present): nothing written, the 'no data to rebuild from' refusal, folder unchanged",
          res6["wrote"] == [] and res6["errors"] == ["This record has no COC/Packing List/Shipping Slip data to rebuild from."]
          and sorted(p.name for p in (wo / "Shipping Slip").glob("*")) == before)
    shutil.rmtree(wo)
    res5 = A.reemit_paperwork(rec)
    check("WO folder gone -> refuses (no phantom folders), errors reported",
          not res5["wrote"] and len(res5["errors"]) == 2 and not wo.exists() and not (Path(root) / "16455.1").exists())
    # printed while the share was down: temp copy, NO recorded path, printed stamps -> Re-emit is what files
    # the pages into the WO folder (never back into the temp folder), byte-identical to what printed
    wo8 = Path(root) / "16458.1"; wo8.mkdir()
    d8 = sample(); d8["date"] = "2026-09-10"; opened8 = []          # a slip date that is NOT today: the name must come from it
    r8b = A.slip_print_pages(str(Path(root) / "gone" / "16458.1"), d8, "both", open_file=lambda p: opened8.append(p))
    tmp_slip = open(r8b["slip_html"], encoding="utf-8").read(); n_tmp = len(list(ftmp.glob("*")))
    check("fallback print: ONE print copy (slip + labels) opened from the temp folder, both printed stamps, outputs None (no filed copy yet)",
          r8b["fallback"] and opened8 == [r8b["print_html"]] and Path(opened8[0]).parent == ftmp and "Print_Slip_and_Labels" in Path(opened8[0]).name
          and d8["printed"]["slip"] and d8["printed"]["labels"] and d8["outputs"] == {"slip_html": None, "labels_html": None, "pack_html": None})
    rec8 = {"event_type": "close", "eng": "4092", "wo": "16458.1", "serials": [], "wo_folder": str(wo8),
            "outputs": {"slip_html": None, "labels_html": None}, "slip": json.loads(json.dumps(d8))}
    res8 = A.reemit_paperwork(rec8)
    exp_slip = wo8 / "Shipping Slip" / "4092-9-10-26-Shipping_Slip.html"
    check("re-emit of a fallback-printed record files BOTH pages into <WO>\\Shipping Slip under SLIP-DATE names (9-10-26, not today's), never the temp folder, byte-identical to the temp print",
          res8["ok"] and len(res8["wrote"]) == 2 and all(Path(p).parent == wo8 / "Shipping Slip" for p in res8["wrote"])
          and exp_slip.is_file() and (wo8 / "Shipping Slip" / "4092-9-10-26-Shipping_Labels.html").is_file() and fd != "9-10-26"
          and exp_slip.read_text(encoding="utf-8") == tmp_slip and len(list(ftmp.glob("*"))) == n_tmp)
    res8c = A.reemit_paperwork(rec8)
    check("re-emit again: the filed copies are recognised by their form codes — both kept, nothing written twice",
          not res8c["wrote"] and len(res8c["kept"]) == 2 and "4092-9-10-26-Shipping_Slip.html" in res8c["kept"][0] and len(list((wo8 / "Shipping Slip").glob("*.html"))) == 2)
    rec8b = dict(rec8, slip=dict(rec8["slip"], outputs={"slip_html": r8b["slip_html"], "labels_html": None}, printed={"slip": "9/24/2026 10:00 AM", "labels": None}))
    shutil.rmtree(wo8 / "Shipping Slip", ignore_errors=True)
    res8b = A.reemit_paperwork(rec8b)
    check("a recorded path inside a temp print folder is ignored (never 'kept' there, never rebuilt there): the page is filed in the WO folder",
          res8b["ok"] and [Path(p).parent for p in res8b["wrote"]] == [wo8 / "Shipping Slip"] and not res8b["kept"] and Path(r8b["slip_html"]).is_file())
    # ONLINE print at 9:00, then the SAME page reprinted the SAME DAY while the share was down (weight corrected to
    # 2.0): the offline print has no recorded path and no name of its own, and the clean <eng>-<date> name already
    # belongs to the 9:00 file — the re-emit must still FILE the paper that shipped, not "keep" the 9:00 page
    wo12 = Path(root) / "16462.1"; wo12.mkdir()
    d12 = sample(); d12["date"] = datetime.date.today().strftime("%Y-%m-%d"); d12["boxes"] = [{"size": "10x10x5", "qty": "15", "weight": "1.5"}]
    r12 = A.slip_print_pages(str(wo12), d12, "slip", open_file=lambda p: None)
    f12 = Path(r12["slip_html"]); code_a = d12["slip_code"]; txt_a = f12.read_text(encoding="utf-8")
    d12["boxes"][0]["weight"] = "2.0"
    def _deny12(path, text):
        if Path(path).parent == wo12 / "Shipping Slip":
            raise PermissionError(13, "Access is denied", str(path))
        return _orig_wnf(path, text)
    A._write_new_file = _deny12
    try:
        r12b = A.slip_print_pages(str(wo12), d12, "slip", open_file=lambda p: None)
    finally:
        A._write_new_file = _orig_wnf
    code_b = d12["slip_code"]; txt_b = open(r12b["slip_html"], encoding="utf-8").read()
    check("(setup) 9:00 online print filed under the clean name; the 9:30 offline reprint went to the temp folder unrecorded, with a new code and weight 2.0",
          f12.name == f"4092-{fd}-Shipping_Slip.html" and r12b["fallback"] and d12["outputs"]["slip_html"] is None and code_a != code_b
          and ">2.0<" in txt_b and ">1.5<" in txt_a and ">2.0<" not in txt_a)
    rec12 = {"event_type": "close", "eng": "4092", "wo": "16462.1", "serials": [], "wo_folder": str(wo12), "outputs": {}, "slip": json.loads(json.dumps(d12))}
    res12 = A.reemit_paperwork(rec12)
    filed = [p for p in (wo12 / "Shipping Slip").glob("*.html") if p != f12]
    filed1 = filed[0] if len(filed) == 1 else None
    check("re-emit FILES the page that shipped (its code, weight 2.0) under the next free name (_HHMMSS, the same-day reprint rule) — never 'kept' because the 9:00 file wears the clean name; the 9:00 file untouched",
          res12["ok"] and len(res12["wrote"]) == 1 and Path(res12["wrote"][0]).name.startswith(f"4092-{fd}-Shipping_Slip_") and filed1 is not None
          and filed1.read_text(encoding="utf-8") == txt_b and code_b in txt_b and f12.read_text(encoding="utf-8") == txt_a and not res12["kept"] and not res12["errors"])
    if res12["kept"]: print("      kept:", res12["kept"])
    res12b = A.reemit_paperwork(rec12)
    check("re-emit again: the filed copy is recognised by its form code — kept, nothing written twice",
          not res12b["wrote"] and len(res12b["kept"]) == 1 and filed1 is not None and filed1.name in res12b["kept"][0]
          and len(list((wo12 / "Shipping Slip").glob("*.html"))) == 2)
    check("_reemit_filed_page: finds the file carrying the code, ignores other kinds, None for a blank code / no match / unreadable folder",
          filed1 is not None and A._reemit_filed_page(wo12 / "Shipping Slip", "Shipping_Slip", code_b) == filed1 and A._reemit_filed_page(wo12 / "Shipping Slip", "Shipping_Slip", code_a) == f12
          and A._reemit_filed_page(wo12 / "Shipping Slip", "Shipping_Labels", code_b) is None and A._reemit_filed_page(wo12 / "Shipping Slip", "Shipping_Slip", "") is None
          and A._reemit_filed_page(wo12 / "Shipping Slip", "Shipping_Slip", "NOPE-1234") is None and A._reemit_filed_page(wo12 / "nowhere", "Shipping_Slip", code_b) is None)
    # ONE record, TWO prints: the operator corrects a weight and reprints the slip only — the labels page on
    # paper still says the old weight, so that is what a re-emit of the labels page must reproduce
    wo10 = Path(root) / "16460.1"; wo10.mkdir()
    d10 = sample(); d10["boxes"] = [{"size": "10x10x5", "qty": "5", "weight": "1.5"}]
    r10 = A.write_slip_files(str(wo10), d10, "both")
    lab_orig = open(r10["labels_html"], encoding="utf-8").read()
    d10["boxes"][0]["weight"] = "2.0"
    r10b = A.write_slip_files(str(wo10), d10, "slip")
    check("reprinting one page leaves the other page's code and path untouched",
          d10["outputs"]["labels_html"] == r10["labels_html"] and d10["labels_code"] in lab_orig and r10b["labels_html"] is None)
    check("each page freezes what it was rendered from (pages snapshot: labels weight 1.5, slip weight 2.0, codes, no nested bookkeeping)",
          d10["pages"]["labels"]["boxes"][0]["weight"] == "1.5" and d10["pages"]["slip"]["boxes"][0]["weight"] == "2.0"
          and d10["pages"]["labels"]["labels_code"] == d10["labels_code"] and d10["pages"]["slip"]["slip_code"] == d10["slip_code"]
          and not ({"pages", "printed", "outputs"} & set(d10["pages"]["slip"])))
    rec10 = {"event_type": "close", "eng": "4092", "wo": "16460.1", "serials": [], "wo_folder": str(wo10), "outputs": {}, "slip": json.loads(json.dumps(d10))}
    os.remove(r10["labels_html"])
    res10 = A.reemit_paperwork(rec10)
    check("re-emit reproduces the labels page AS PRINTED (weight 1.5), not the form as edited afterwards (2.0)",
          res10["ok"] and [Path(p).name for p in res10["wrote"]] == [Path(r10["labels_html"]).name]
          and open(r10["labels_html"], encoding="utf-8").read() == lab_orig and ">1.5<" in lab_orig)
    wo11 = Path(root) / "16461.1"; wo11.mkdir()
    res11 = A.reemit_paperwork(dict(rec10, wo_folder=str(wo11), slip=dict(rec10["slip"], pages=None,
                                    outputs={"slip_html": None, "labels_html": None}, printed={"slip": "9/24/2026 10:00 AM", "labels": None})))
    check("a record without snapshots (older build) still rebuilds the printed page from its slip_data",
          res11["ok"] and len(res11["wrote"]) == 1 and Path(res11["wrote"][0]).parent == wo11 / "Shipping Slip"
          and ">2.0</div>" in open(res11["wrote"][0], encoding="utf-8").read())
    check("old-style record wording unchanged: serials but no COC / Packing List / slip -> the pre-slip 'nothing to rebuild' message",
          A.reemit_paperwork({"event_type": "close", "serials": ["S1"], "wo_folder": str(wo10)})["errors"]
          == ["This record has no COC/Packing List data to rebuild from."])
    wo6 = Path(root) / "16456.1"; wo6.mkdir()          # fresh, empty WO folder (no same-named files)
    rec_ns = dict(rec, coc={"po": "x", "ship_date": "2026-09-24"}, wo_folder=str(wo6))
    res7 = A.reemit_paperwork(rec_ns)
    check("COC without serials still errors (unchanged rule) while the slip pages rebuild",
          any("serials" in e for e in res7["errors"]) and len(res7["wrote"]) == 2 and res7["ok"] is False
          and all(Path(p).parent == wo6 / "Shipping Slip" for p in res7["wrote"]))
    check("non-close record refused", A.reemit_paperwork(dict(rec, event_type="draft"))["errors"])
    wo7 = Path(root) / "16457.1"; wo7.mkdir()
    check("outputs may also sit on the record root (implementer-2 compatible)",
          len(A.reemit_paperwork(dict(rec, wo_folder=str(wo7), outputs=rec["slip"]["outputs"],
                                      slip=dict(rec["slip"], outputs={})))["wrote"]) == 2)
finally:
    _tf.gettempdir = _orig_gettempdir
    shutil.rmtree(root, ignore_errors=True); shutil.rmtree(faketmp, ignore_errors=True)

# ══ 10. Draft round-trip / close record / History UI ═════════════════════════
# (implementer 2 — SlipLabelsPopup, self.slip_data in DRAFTS, close_record["slip"],
#  HistoryDetail buttons — appends its checks below this line.)
# Headless: the popup's Tk-free half (slip_prefill / slip_required_missing /
# slip_print_pages / SlipLabelsPopup._collect on a stand-in) is exercised for
# real; the Tk-bound wiring (main-window button, load/unload/close, History
# buttons) is pinned by source text.
import inspect, types
SRC_ALL = open(os.path.join(SCRIPTS, "app_logic.py"), encoding="utf-8").read()
TODAY = datetime.date.today().strftime("%Y-%m-%d")

# ── 10a. pre-fill: sources + precedence (folder > COC/Packing > Monday) ──────
pi = {"customer": "377 Anduril", "pn": "125-1030-970/C", "eng": "4092", "rev": "C", "po": "PO-778", "cpn": "125-1030-970/C"}
mh = {"po": "PO-MONDAY", "rev": "B", "cpn": "CPN-M", "customer": "Monday Cust", "wo": "16452.1", "eng": "4092"}
d = A.slip_prefill(pi, mh, None, None, "Dhruvil Savaliya", "16452.1", 15)
check("prefill: folder fields authoritative (customer/eng/rev), PO + customer P/N from the Monday-merged parsed_info",
      d["customer"] == "377 Anduril" and d["eng"] == "4092" and d["rev"] == "C" and d["po"] == "PO-778"
      and d["cpn"] == "125-1030-970/C" and d["wo"] == "16452.1")
check("prefill: qty = scanned serial count, date = today, WWYY from that date, prepared by = operator",
      d["qty"] == "15" and d["date"] == TODAY and d["datecode_wwyy"] == A.slip_datecode_wwyy(TODAY)
      and d["prepared_by"] == "Dhruvil Savaliya")
check("prefill: one box row carrying the qty; options at their Apex defaults; nothing printed",
      d["boxes"] == [{"size": "", "qty": "15", "weight": ""}] and d["customer_parts"] == "No"
      and not any(d[k] for k in ("fair", "si", "ppap", "cofc")) and d["ship_via"] == "" and d["label_notes"] == ""
      and d["slip_code"] == "" and d["labels_code"] == "" and d["printed"] == {"slip": None, "labels": None, "pack": None}
      and d["outputs"] == {"slip_html": None, "labels_html": None, "pack_html": None})
check("prefill has exactly the template's keys", set(d) == set(A.slip_data_template()))
d2 = A.slip_prefill({"customer": "N/A", "pn": "", "eng": "N/A", "rev": "N/A"}, {"po": "PO-M", "customer": "Monday Cust"},
                    {"po": "PO-COC", "pn": "PN-COC", "rev": "R2", "qty": "7", "job": "4099"},
                    {"customer": "Pack Cust", "wo": "WO-PK", "eng": "4100"}, "", "", 0)
check("prefill: 'N/A' folder values give way to COC/Packing entries (COC PO beats Monday PO)",
      d2["po"] == "PO-COC" and d2["cpn"] == "PN-COC" and d2["eng"] == "4099" and d2["rev"] == "R2"
      and d2["qty"] == "7" and d2["customer"] == "Pack Cust" and d2["wo"] == "WO-PK" and d2["prepared_by"] == "")
d3 = A.slip_prefill({}, {"customer": "Monday Cust", "cpn": "CPN-M", "po": "PO-M", "rev": "B", "eng": "E"}, None, None, "", "", 0)
check("prefill: Monday fills what nothing else has", d3["customer"] == "Monday Cust" and d3["cpn"] == "CPN-M"
      and d3["po"] == "PO-M" and d3["rev"] == "B" and d3["eng"] == "E" and d3["qty"] == "" and d3["boxes"][0]["qty"] == "")
check("prefill with nothing at all is still a complete, JSON-safe record",
      json.loads(json.dumps(A.slip_prefill())) == A.slip_prefill() and A.slip_prefill()["date"] == TODAY)
# a saved slip_data wins outright (re-open shows what was saved) and is normalised to the template shape
ex = sample(); ex["printed"]["slip"] = "09/24/2026 10:00 AM"; ex["outputs"]["slip_html"] = r"X:\s.html"
ex["fair"] = 1; ex["customer_parts"] = "yes"; del ex["label_notes"]; ex["boxes"][0]["qty"] = 10; ex["boxes"].append(None)
d4 = A.slip_prefill(pi, mh, {"po": "OTHER"}, None, "Somebody Else", "other-wo", 3, existing=ex)
check("existing slip_data wins over every other source (customer/qty/wo/prepared_by untouched)",
      d4["customer"] == "377 Anduril" and d4["qty"] == "15" and d4["wo"] == "16452.1" and d4["prepared_by"] == "Dhruvil Savaliya"
      and d4["po"] == "PO-778")
check("existing: normalised — bools are bools, Yes/No canonical, missing key from the template, box values strings (full first row kept), junk box dropped to blanks",
      d4["fair"] is True and d4["customer_parts"] == "Yes" and d4["label_notes"] == ""
      and d4["boxes"][0] == {"size": "10x10x5", "qty": "10", "weight": "1.25"} and d4["boxes"][1] == {"size": "", "qty": "5", "weight": "2"}
      and d4["boxes"][2] == {"size": "", "qty": "", "weight": ""} and len(d4["boxes"]) == 3)
check("existing: every box row (size, qty, weight) survives a re-open / draft restore unchanged (compared against the fixture, not against slip_prefill itself)",
      A.slip_prefill(existing=sample())["boxes"] == sample()["boxes"] and A.slip_prefill(existing=json.loads(json.dumps(sample())))["boxes"] == sample()["boxes"]
      and sample()["boxes"][0]["size"] == "10x10x5")
check("existing: codes / printed / outputs carried", d4["slip_code"] == ex["slip_code"] and d4["labels_code"] == ex["labels_code"]
      and d4["printed"] == {"slip": "09/24/2026 10:00 AM", "labels": None, "pack": None} and d4["outputs"] == {"slip_html": r"X:\s.html", "labels_html": None, "pack_html": None})
check("existing from a draft JSON round-trips to the same record",
      A.slip_prefill(existing=json.loads(json.dumps(sample()))) == A.slip_prefill(existing=sample()))
check("existing with no boxes keeps no boxes (popup then shows one blank row)", A.slip_prefill(existing=dict(sample(), boxes=[]))["boxes"] == [])

# ── 10b. Apex's required fields (getRequiredFields ship / shipl) ─────────────
base = sample()
check("complete record: nothing missing for slip / labels / both",
      A.slip_required_missing(base, "slip") == [] and A.slip_required_missing(base, "labels") == [] and A.slip_required_missing(base) == [])
check("slip requires Customer, WO, Qty, Prepared By — in form order, blanks and whitespace alike",
      A.slip_required_missing(dict(base, customer="", wo=" ", qty="", prepared_by="  "), "slip")
      == ["Customer", "Work Order / Lot", "Qty Shipped", "Prepared By"])
check("slip: Prepared By alone", A.slip_required_missing(dict(base, prepared_by=""), "slip") == ["Prepared By"])
check("labels require Customer, WO and a qty on EVERY box (Prepared By not required for labels)",
      A.slip_required_missing(dict(base, customer="", prepared_by="",
                                   boxes=[{"size": "", "qty": "", "weight": ""}, {"size": "", "qty": "5", "weight": ""},
                                          {"size": "", "qty": " ", "weight": "1"}]), "labels")
      == ["Customer", "Box 1 qty", "Box 3 qty"])
check("labels with no box rows fall back to the slip qty", A.slip_required_missing(dict(base, boxes=[], qty=""), "labels") == ["Qty Shipped"]
      and A.slip_required_missing(dict(base, boxes=[], qty="3"), "labels") == [])
check("labels: slip qty blank but every box has one -> ok (Apex checks the per-box field)",
      A.slip_required_missing(dict(base, qty=""), "labels") == [])
check("both = union without duplicates", A.slip_required_missing(dict(base, customer="", qty="", boxes=[]))
      == ["Customer", "Qty Shipped"] and A.slip_required_missing(dict(base, customer="", boxes=[{"qty": ""}]))
      == ["Customer", "Box 1 qty"])
check("None / empty data: everything missing, never raises", A.slip_required_missing(None, "slip") == ["Customer", "Work Order / Lot", "Qty Shipped", "Prepared By"]
      and A.slip_required_missing({}, "labels") == ["Customer", "Work Order / Lot", "Qty Shipped"])
check("required-field tables match Apex getRequiredFields (ship / shipl)",
      [k for k, _ in A.SLIP_REQUIRED_SLIP] == ["customer", "wo", "qty", "prepared_by"] and [k for k, _ in A.SLIP_REQUIRED_LABELS] == ["customer", "wo"])

# ── 10c. the popup's form read-back (_collect) on a Tk-free stand-in ─────────
class V:                                   # the only thing _collect needs from a Tk var
    def __init__(s, v): s.v = v
    def get(s): return s.v
st = types.SimpleNamespace()
st._base = dict(sample(), slip_code="CODE-S", labels_code="CODE-L",
                printed={"slip": "t1", "labels": None, "junk": 1}, outputs={"slip_html": "p1", "labels_html": None})
st._v = {k: V(" " + str(sample()[k]) + " ") for k in A.SlipLabelsPopup.FIELDS}
st._v["date"] = V("09/24/2026"); st._v["datecode_wwyy"] = V("")
st._v["customer_parts"] = V("yes")
st._v["fair"] = V(1); st._v["si"] = V(0); st._v["ppap"] = V(True); st._v["cofc"] = V(False)
st._box_rows = [(V("10x10x5"), V(" 10 "), V("1.25")), (V(""), V("5"), V("2"))]
got = A.SlipLabelsPopup._collect(st)
check("_collect: text fields stripped, date stored as YYYY-MM-DD, WWYY recomputed when blank",
      got["customer"] == "377 Anduril" and got["notes"] == "2 boxes" and got["ship_via"] == "UPS GND"
      and got["date"] == "2026-09-24" and got["datecode_wwyy"] == "3926")
check("_collect: Yes/No + flags canonical, boxes in row order as strings",
      got["customer_parts"] == "Yes" and got["fair"] is True and got["si"] is False and got["ppap"] is True and got["cofc"] is False
      and got["boxes"] == [{"size": "10x10x5", "qty": "10", "weight": "1.25"}, {"size": "", "qty": "5", "weight": "2"}])
check("_collect: codes / printed / outputs come from the base record only (junk keys dropped)",
      got["slip_code"] == "CODE-S" and got["labels_code"] == "CODE-L" and got["printed"] == {"slip": "t1", "labels": None, "pack": None}
      and got["outputs"] == {"slip_html": "p1", "labels_html": None, "pack_html": None})
check("_collect: exact template shape, JSON-safe, feeds the generators", set(got) == set(A.slip_data_template())
      and json.loads(json.dumps(got)) == got and "10x10x5" in A.build_slip_html(got) and A.build_labels_html(got).count('class="sheet"') == 2)
_pg = {"slip": {"customer": "as printed"}, "labels": None}
check("_collect / slip_prefill(existing) carry the per-page snapshots like codes / printed / outputs",
      A.SlipLabelsPopup._collect(types.SimpleNamespace(_base=dict(st._base, pages=_pg), _v=st._v, _box_rows=st._box_rows))["pages"] == dict(_pg, pack=None)
      and A.slip_prefill(existing=dict(sample(), pages=_pg))["pages"] == dict(_pg, pack=None)
      and A.slip_prefill(existing=dict(sample(), pages="junk"))["pages"] == {"slip": None, "labels": None, "pack": None})
# the box rows are rebuilt INCREMENTALLY: the spinbox's <FocusOut> / Return — with an unchanged OR a changed count —
# must not tear down the rows that stay (Tk sets the new focus BEFORE it delivers FocusOut: destroying the row that
# was just clicked stranded the keyboard on the popup, review rounds 1 + 2). Stand-in: fake row frames that record
# destroy(), fake i/N labels, an _add_box_row that appends fakes, a box frame that refuses to be walked.
class _Row:
    def __init__(s, n): s.name = f".box.row{n}"; s.dead = False
    def destroy(s): s.dead = True
    def __str__(s): return s.name
class _Lbl:
    def __init__(s): s.text = ""
    def config(s, **k): s.text = k.get("text", s.text)
class _NoTouch:
    def winfo_children(s): raise AssertionError("box frame was walked / torn down")
def rebuild_stand_in(rows):
    st = types.SimpleNamespace(_box_rows=[], _box_widgets=[], _box_frame=_NoTouch(), _nbox=V(str(len(rows) or 1)),
                               _tot=types.SimpleNamespace(config=lambda **k: None), _pl_btn=types.SimpleNamespace(config=lambda **k: None),
                               _focus=None, _idle=[])
    st._update_totals = lambda: A.SlipLabelsPopup._update_totals(st)
    st._labels_btn_text = lambda: A.SlipLabelsPopup._labels_btn_text(st)
    st.focus_get = lambda: st._focus
    st.after_idle = lambda f: st._idle.append(f)
    def add(size, qty, wt):
        n = len(st._box_rows) + 1
        st._box_rows.append((V(size), V(qty), V(wt)))
        st._box_widgets.append((_Row(n), _Lbl(), types.SimpleNamespace(focus_set=lambda n=n: st._idle.append(("focus", n)))))
    st._add_box_row = add
    for s_, q_, w_ in rows: add(s_, q_, w_)
    return st
st3 = rebuild_stand_in([("a", "1", "1"), ("b", "2", "2")])
before = list(st3._box_rows); wid_before = list(st3._box_widgets)
A.SlipLabelsPopup._rebuild_boxes(st3)
check("_rebuild_boxes with an unchanged count leaves every row widget alone (FocusOut-safe), renumbers i/N",
      st3._box_rows == before and st3._box_widgets == wid_before and not any(r.dead for r, _l, _s in st3._box_widgets)
      and [l.text for _r, l, _s in st3._box_widgets] == ["1/2", "2/2"] and st3._idle == [])
st3._nbox = V("3"); st3._focus = ".box.row2.!entry"        # count typed into the spinbox, then a click into Box 2's Qty
A.SlipLabelsPopup._rebuild_boxes(st3)
check("a LARGER count keeps the existing rows (same var objects, same widgets — the clicked entry survives), appends a blank row, renumbers i/N",
      st3._box_rows[:2] == before and st3._box_widgets[:2] == wid_before and len(st3._box_rows) == 3 and st3._box_rows[2][1].get() == ""
      and not any(r.dead for r, _l, _s in st3._box_widgets) and [l.text for _r, l, _s in st3._box_widgets] == ["1/3", "2/3", "3/3"] and st3._idle == [])
st3._nbox = V("1"); st3._focus = ".box.row3.!entry"        # focus sat in a row that goes away
gone = st3._box_widgets[1:]
A.SlipLabelsPopup._rebuild_boxes(st3)
check("a SMALLER count destroys only the rows past it (row 1 untouched, values kept) and, the focused row being gone, moves focus to the last row's Size on idle",
      st3._box_rows == before[:1] and st3._box_widgets == wid_before[:1] and all(r.dead for r, _l, _s in gone) and st3._box_rows[0][0].get() == "a"
      and [l.text for _r, l, _s in st3._box_widgets] == ["1/1"] and len(st3._idle) == 1 and callable(st3._idle[0]))
if st3._idle and callable(st3._idle[0]): st3._idle[0]()
check("… the idle callback focuses the surviving last row's Size combobox", bool(st3._idle) and st3._idle[-1] == ("focus", 1))
st3._nbox = V("2"); st3._focus = ".box.row1.!entry"
A.SlipLabelsPopup._rebuild_boxes(st3)
check("shrinking with the focus in a row that STAYS schedules no refocus", len(st3._box_rows) == 2 and len(st3._idle) == 2)
st3._nbox = V("150"); A.SlipLabelsPopup._rebuild_boxes(st3)
check("count clamped to 99 (spinbox var written back)", len(st3._box_rows) == 99 and st3._nbox.get() == "99" if hasattr(st3._nbox, "set") else len(st3._box_rows) == 99)
st4 = rebuild_stand_in([]); st4._nbox = V("2")
A.SlipLabelsPopup._rebuild_boxes(st4, initial=[{"size": "s", "qty": "1", "weight": "2"}, {"qty": "3"}, None])
check("initial (saved boxes) seeds the rows: values carried, missing keys blank, capped at the count, labelled i/N",
      [(s.get(), q.get(), w.get()) for s, q, w in st4._box_rows] == [("s", "1", "2"), ("", "3", "")] and [l.text for _r, l, _s in st4._box_widgets] == ["1/2", "2/2"])
st5 = rebuild_stand_in([]); st5._nbox = V("1")
A.SlipLabelsPopup._rebuild_boxes(st5, initial=[])
check("no saved boxes -> one blank row", [(s.get(), q.get(), w.get()) for s, q, w in st5._box_rows] == [("", "", "")])
st2 = types.SimpleNamespace(_base=None, _v=dict(st._v, date=V("not a date"), datecode_wwyy=V("0126")), _box_rows=[])
got2 = A.SlipLabelsPopup._collect(st2)
check("_collect: no base -> blank codes; unparseable date kept as typed; typed WWYY kept; no boxes -> []",
      got2["slip_code"] == "" and got2["date"] == "not a date" and got2["datecode_wwyy"] == "0126" and got2["boxes"] == []
      and got2["printed"] == {"slip": None, "labels": None, "pack": None})

# ── 10d. the Print action, headless (stub opener) ───────────────────────────
root = _tf.mkdtemp(prefix="pcc_slip_ui_")
faketmp2 = _tf.mkdtemp(prefix="pcc_slip_ui_tmp_")
_tf.gettempdir = lambda: faketmp2          # print copies go to OUR temp dir, never the real one
ftmp2 = Path(faketmp2) / "PCC Shipping App - Print"
try:
    wo = Path(root) / "16452.1"; wo.mkdir()
    ts_re = re.compile(r"^\d{1,2}/\d{1,2}/\d{4} \d{2}:\d{2} [AP]M$")      # fmt_date() + ' ' + fmt_time()
    dp = sample(); opened = []
    r = A.slip_print_pages(str(wo), dp, "both", open_file=lambda p: opened.append(p))
    check("print both: both pages FILED in <WO>\\Shipping Slip; ONE print copy (slip, then labels) opened from the temp print folder — never the filed pages",
          opened == [r["print_html"]] and Path(r["print_html"]).parent == ftmp2 and Path(r["print_html"]).name == f"4092-{A.fmt_file_date()}-Print_Slip_and_Labels.html"
          and Path(r["slip_html"]).parent == wo / "Shipping Slip" and Path(r["labels_html"]).parent == wo / "Shipping Slip"
          and r["opened"] == opened and r["open_errors"] == [] and r["fallback"] is False
          and open(r["print_html"], encoding="utf-8").read() == A.build_slip_labels_html(dp))
    check("print both: printed times stamped in the app's date+time style; outputs + fresh codes stamped",
          ts_re.match(dp["printed"]["slip"] or "") and ts_re.match(dp["printed"]["labels"] or "")
          and dp["outputs"] == {"slip_html": r["slip_html"], "labels_html": r["labels_html"], "pack_html": None}
          and cre.match(dp["slip_code"]) and lre.match(dp["labels_code"]))
    dq = sample(); dq["printed"] = {"slip": None, "labels": None}
    def boom(p): raise OSError("no browser")
    r2 = A.slip_print_pages(str(wo), dq, "slip", open_file=boom)
    check("opener failure: file written and reported, but NOT stamped as printed",
          r2["open_errors"] and r2["open_errors"][0][0] == "Shipping Slip" and "no browser" in r2["open_errors"][0][1]
          and dq["printed"]["slip"] is None and Path(r2["slip_html"]).is_file() and r2["opened"] == [] and dq["outputs"]["slip_html"] == r2["slip_html"])
    dl = sample(); dl["printed"] = {"slip": "KEEP", "labels": None}; dl["slip_code"] = "OLD-SLIP-CODE"
    r3 = A.slip_print_pages(str(wo), dl, "labels", open_file=lambda p: None)
    check("labels only: slip stamp and slip code untouched, labels stamped", dl["printed"]["slip"] == "KEEP" and ts_re.match(dl["printed"]["labels"] or "")
          and dl["slip_code"] == "OLD-SLIP-CODE" and r3["slip_html"] is None)
    dm = sample(); dm["printed"] = "garbage"
    A.slip_print_pages(str(wo), dm, "slip", open_file=lambda p: None)
    check("a corrupt 'printed' is repaired to the template shape", isinstance(dm["printed"], dict) and set(dm["printed"]) == {"slip", "labels", "pack"})
    check("printed summary line", A.slip_printed_summary(dl).startswith("Slip printed KEEP  ·  Labels printed ")
          and A.slip_printed_summary(None) == "Slip not printed  ·  Labels not printed  ·  Packing label not printed")
    check("print: same-day reprint never overwrote (5 distinct pages on disk)", len(list((wo / "Shipping Slip").glob("*.html"))) == 5)
    check("print copies never land in the WO folder (only filed pages there); one print copy per print that opened or tried to",
          not list((wo / "Shipping Slip").glob("*Print_*")) and len(list(ftmp2.glob("*-Print_*.html"))) == 4)
    # the temp folder refuses the print copy -> the filed page(s) are opened one by one, as before 2026-09-25
    _wnf0 = A._write_new_file
    A._write_new_file = lambda path, text: (_ for _ in ()).throw(PermissionError(13, "denied", str(path))) if "Print_" in str(path) else _wnf0(path, text)
    try:
        dpf = sample(); opf = []
        rpf = A.slip_print_pages(str(wo), dpf, "both", open_file=lambda p: opf.append(p))
    finally:
        A._write_new_file = _wnf0
    check("no print copy possible -> the filed slip and labels are opened instead (both stamped), print_html None",
          rpf["print_html"] is None and opf == [rpf["slip_html"], rpf["labels_html"]] and ts_re.match(dpf["printed"]["slip"] or "")
          and ts_re.match(dpf["printed"]["labels"] or ""))
    # the day-old print copies are tidied; fresh ones, filed-name pages and other files are left alone
    import time as _time
    old = ftmp2 / "4092-9-1-26-Print_Slip.html"; old.write_text("x", encoding="utf-8")
    old42 = ftmp2 / "labels4x2_16452.1_20260901_101010.html"; old42.write_text("x", encoding="utf-8")
    keep = ftmp2 / "4092-9-1-26-Shipping_Slip.html"; keep.write_text("x", encoding="utf-8")
    other = ftmp2 / "notes.txt"; other.write_text("x", encoding="utf-8")
    for f_ in (old, old42, keep, other): os.utime(f_, (_time.time() - 2 * 86400, _time.time() - 2 * 86400))
    A.slip_print_pages(str(wo), sample(), "slip", open_file=lambda p: None)
    check("print copies / 4x2 renders older than a day are tidied from the temp print folder; a filed-name page, other files and today's copies stay",
          not old.exists() and not old42.exists() and keep.exists() and other.exists() and len(list(ftmp2.glob("*-Print_*.html"))) == 5)

    # ── the popup's Print action ITSELF on a stand-in: a print sets self.result, so Cancel / X afterwards still
    #    returns the printed record (History must match the paper) — the one line that rule rests on
    class _MB:                                    # the popup's message boxes, recorded (never shown)
        def __init__(s): s.calls = []; s.answer = True
        def showwarning(s, t, m, **k): s.calls.append(("warn", t, m))
        def showerror(s, t, m, **k): s.calls.append(("error", t, m))
        def askyesno(s, t, m, **k): s.calls.append(("ask", t, m)); return s.answer
    def popup_stand_in(base, sc=0, wo_folder=None):
        st = types.SimpleNamespace(_base=base, _sc=sc, _wo_folder=wo_folder, result=None, rebuilt=0,
                                   _v={k: V(str(base.get(k) or "")) for k in A.SlipLabelsPopup.FIELDS},
                                   _box_rows=[(V(b["size"]), V(b["qty"]), V(b["weight"])) for b in base["boxes"]],
                                   _status=types.SimpleNamespace(config=lambda **k: None))
        st._v["customer_parts"] = V(base["customer_parts"])
        for k in ("fair", "si", "ppap", "cofc"): st._v[k] = V(base[k])
        def _rb(initial=None): st.rebuilt += 1
        st._rebuild_boxes = _rb; st._regrab = lambda: None
        st._collect = lambda: A.SlipLabelsPopup._collect(st)
        return st
    mb = _MB(); _mb0 = A.messagebox; A.messagebox = mb
    opened_p = []; _spp = A.slip_print_pages
    A.slip_print_pages = lambda wf, d, which: _spp(wf, d, which, open_file=lambda p: opened_p.append(p))   # never os.startfile in a test
    try:
        stp = popup_stand_in(sample(), sc=15, wo_folder=str(wo))
        A.SlipLabelsPopup._print(stp, "both")
        check("_print (both): typed count applied, validation passed, no dialog, pages written + ONE print copy opened, and self.result IS the printed record (self._base too) — Cancel afterwards keeps it",
              mb.calls == [] and len(opened_p) == 1 and stp.result is not None and stp.result is stp._base and stp.rebuilt == 1
              and cre.match(stp.result["slip_code"]) and ts_re.match(stp.result["printed"]["slip"] or "") and ts_re.match(stp.result["printed"]["labels"] or "")
              and Path(stp.result["outputs"]["slip_html"]).parent == wo / "Shipping Slip" and stp.result["boxes"] == sample()["boxes"])
        stq = popup_stand_in(dict(sample(), customer=""), sc=15, wo_folder=str(wo))
        A.SlipLabelsPopup._print(stq, "slip")
        check("_print with a missing required field: the warning names it, nothing written / opened, result stays None, base untouched",
              mb.calls[-1][0] == "warn" and "Customer" in mb.calls[-1][2] and stq.result is None and stq._base["customer"] == "" and len(opened_p) == 1)
        stz = popup_stand_in(sample(), sc=3, wo_folder=str(wo)); mb.answer = False
        A.SlipLabelsPopup._print(stz, "slip")
        check("_print: Qty Shipped vs scanned serials asks (COC-style); 'No' prints nothing", mb.calls[-1][0] == "ask" and "3 serial" in mb.calls[-1][2] and stz.result is None and len(opened_p) == 1)
        mb.answer = True
        A.SlipLabelsPopup._print(stz, "slip")
        check("… 'Yes' prints (the check warns, never blocks)", stz.result is stz._base and len(opened_p) == 2 and stz.result["printed"]["labels"] is None)
        many = popup_stand_in(dict(sample(), boxes=[{"size": "", "qty": "", "weight": ""}] * 20), sc=0, wo_folder=str(wo))
        A.SlipLabelsPopup._print(many, "labels")
        check("_print: the missing-fields message is capped (8 shown, '… and 12 more' — no 100-line message box for a stray 99)",
              mb.calls[-1][0] == "warn" and "… and 12 more" in mb.calls[-1][2] and mb.calls[-1][2].count("Box ") == 8 and many.result is None)
        A.slip_print_pages = lambda wf, d, which: (_ for _ in ()).throw(OSError("disk full"))
        stx = popup_stand_in(sample(), sc=15, wo_folder=str(wo))
        A.SlipLabelsPopup._print(stx, "slip")
        check("_print: a write that raises shows 'Could not create the page(s)' and keeps result None (nothing printed)",
              mb.calls[-1][0] == "error" and "Could not create the page(s)" in mb.calls[-1][2] and "disk full" in mb.calls[-1][2] and stx.result is None)
        A.slip_print_pages = lambda wf, d, which: {"slip_html": "x", "labels_html": None, "fallback": False, "folder": "f",
                                                   "errors": [("Shipping Labels", "PermissionError: denied")], "opened": ["x"], "open_errors": []}
        sty = popup_stand_in(sample(), sc=15, wo_folder=str(wo))
        A.SlipLabelsPopup._print(sty, "both")
        check("_print: a later page that could be written nowhere is reported ('Could not create: … Shipping Labels') while the written page's record is KEPT (result set)",
              mb.calls[-1][0] == "error" and "Could not create:" in mb.calls[-1][2] and "Shipping Labels: PermissionError" in mb.calls[-1][2] and sty.result is sty._base)
    finally:
        A.messagebox = _mb0; A.slip_print_pages = _spp

    # ── 10e. draft round-trip (slip_data travels like coc_data / packing_data) ─
    p = A.save_shipment_draft(str(wo), [], ["S1", "S2"], {"po": "x"}, None, operator="Op", draft_id="20260924_120000", slip_data=dp)
    dr = A.load_shipment_draft(str(wo), path=p)
    check("draft carries slip_data verbatim (JSON) beside coc_data / packing_data",
          dr["slip_data"] == json.loads(json.dumps(dp)) and dr["coc_data"] == {"po": "x"} and dr["packing_data"] is None and dr["serials"] == ["S1", "S2"])
    p0 = A.save_shipment_draft(str(wo), [], ["S1"], None, None, operator="Op")          # pre-2026-09-24 call shape
    check("old call signature still works and yields slip_data None (old drafts restore as 'never opened')",
          p0 and A.load_shipment_draft(str(wo))["slip_data"] is None)
    check("restore -> re-open: slip_prefill(existing=draft copy) equals the in-memory record",
          A.slip_prefill(pi, mh, None, None, "Z", "zz", 1, existing=dr["slip_data"]) == A.slip_prefill(existing=dp))
    check("empty slip_data is stored as None (nothing to restore)", json.load(open(A.save_shipment_draft(str(wo), [], ["S"], None, None, draft_id="x", slip_data={}), encoding="utf-8"))["slip_data"] is None)
    pu = A.save_shipment_draft(str(wo), [], ["S1"], None, None, draft_id="unprinted", slip_data=sample())
    check("an UNPRINTED slip form (popup Save, then Save Draft) round-trips through the draft too — printed {None, None}, outputs None, every field back",
          A.load_shipment_draft(str(wo), path=pu)["slip_data"] == json.loads(json.dumps(sample())) and sample()["printed"] == {"slip": None, "labels": None, "pack": None}
          and A.load_shipment_draft(str(wo), path=pu)["slip_data"]["boxes"] == sample()["boxes"])
    # the History record built from the popup's data still re-emits (end-to-end with 10c/10d output)
    rec = {"event_type": "close", "eng": "4092", "wo": "16452.1", "serials": [], "wo_folder": str(wo),
           "outputs": {"slip_html": dp["outputs"]["slip_html"], "labels_html": dp["outputs"]["labels_html"]}, "slip": dr["slip_data"]}
    os.remove(dp["outputs"]["labels_html"])
    rr = A.reemit_paperwork(rec)
    check("close record shape (slip + mirrored outputs) re-emits the missing page only", [Path(x).name for x in rr["wrote"]] == [Path(dp["outputs"]["labels_html"]).name]
          and len(rr["kept"]) == 1 and not rr["errors"])
    # the offline Outbox README: an offline print puts the pages in a FIFTH subfolder — the sync note must name it
    ob = Path(root) / "outbox"; ob.mkdir()
    for s in (A.PICTURES_FOLDER, A.SERIALS_FOLDER, A.COC_FOLDER, A.PACKING_LIST_FOLDER): (ob / s).mkdir()
    A.write_outbox_readme(ob, "Div", "WO1", "Op", [("COC xlsx", None)])
    t4 = (ob / "_OFFLINE_SHIPMENT_README.txt").read_text(encoding="utf-8")
    A.write_slip_files(str(ob), sample(), "slip")
    A.write_outbox_readme(ob, "Div", "WO1", "Op", [("COC xlsx", None), ("Shipping Slip html", "x/y.html")])
    t5 = (ob / "_OFFLINE_SHIPMENT_README.txt").read_text(encoding="utf-8")
    check("offline README: 'four subfolders' as before without a slip; 'five … Packing List, Shipping Slip' once the pages were printed there",
          "Copy the four subfolders" in t4 and "Shipping Slip" not in t4 and (ob / A.SLIP_FOLDER).is_dir()
          and "Copy the five subfolders" in t5 and "Packing List, Shipping Slip). They match" in t5 and "Shipping Slip html: y.html" in t5)
finally:
    _tf.gettempdir = _orig_gettempdir
    shutil.rmtree(root, ignore_errors=True); shutil.rmtree(faketmp2, ignore_errors=True)

# ── 10f. Tk-bound wiring, pinned by source text ──────────────────────────────
def src(obj):
    try: return inspect.getsource(obj)
    except Exception: return ""
app_init = src(A.ShippingApp.__init__); footer = src(A.ShippingApp._build_footer); chips = src(A.ShippingApp._update_status_chips)
load = src(A.ShippingApp._load_workorder); load_off = src(A.ShippingApp._load_workorder_offline); reset = src(A.ShippingApp._reset_form)
draftw = src(A.ShippingApp._save_draft_work); closew = src(A.ShippingApp._close_shipment_work); close = src(A.ShippingApp._close_shipment)
opener = src(A.ShippingApp._open_slip_popup); popup = src(A.SlipLabelsPopup); hist = src(A.HistoryDetail._build); hreemit = src(A.HistoryDetail._reemit)
check("main window: self.slip_data initialised beside coc_data", "self.slip_data=None" in app_init and "self.coc_data=None" in app_init)
check("main window (Dhruvil 2026-09-25): the printer-icon-only button sits ALONE in the footer's bottom-left corner (grid column 0, sticky w), its name on hover, wired to _open_slip_popup",
      'self.slip_btn=_btn(f,"🖨",self._open_slip_popup,' in footer and 'self.slip_btn.grid(row=0,column=0,sticky="w")' in footer
      and '_HoverTip(self.slip_btn,"Shipping Labels and Slip")' in footer and "Shipping Slip & Labels" not in footer and "🚚" not in footer)
check("main window: the Save buttons stay a centred group (column 1 between two equal 'uniform' side columns); Save and Close packed first so a narrow window never trims it",
      'f.grid_columnconfigure(0,weight=1,uniform="footer_side")' in footer and 'f.grid_columnconfigure(2,weight=1,uniform="footer_side")' in footer
      and "inner.grid(row=0,column=1)" in footer and footer.index('"✓   Save and Close Shipment"') < footer.index("draft_wrap=")
      and footer.count('pack(side="right",padx=6)') == 2 and "Save Draft — finish later" in footer)
check("main window: button enabled only while a WO is loaded — blue, green once a page was sent to print, no text (icon only); no getattr guard — the footer is built before the first chips refresh",
      'state="normal" if self.wo_folder else "disabled"' in chips and "b=self.slip_btn" in chips
      and 'b._base_bg=(C_GREEN if printed else C_BLUE) if self.wo_folder else "#BFC7C3"' in chips and "text=" not in chips.split("b=self.slip_btn")[1]
      and 'getattr(self,"slip_btn"' not in chips and app_init.index("self._build_ui()") < app_init.index("self._uc()"))
tip = src(A._HoverTip)
check("hover tip: shown after a short rest, hidden on leave / click, bound with add='+' so _btn's hover colouring survives",
      all(s in tip for s in ('widget.bind("<Enter>",self._schedule,add="+")', 'widget.bind("<Leave>",self._hide,add="+")',
                             'widget.bind("<ButtonPress>",self._hide,add="+")', "wm_overrideredirect(True)", "self.w.after(self.delay,self._show)")))
check("main window: slip_data cleared on every load / unload path",
      load.count("self.slip_data=None") >= 1 and "self.slip_data=None" in load_off and "self.slip_data=None" in reset)
check("main window: draft restore reads slip_data; switching WO guards unsaved slip data",
      'self.slip_data=draft.get("slip_data")' in load and "or self.slip_data" in load)
check("opener: passes the Monday hit, COC, Packing, operator, roster, the WO folder and the SHARED box sizes; keeps the result",
      all(s in opener for s in ("_monday_hit", "coc=self.coc_data", "packing=self.packing_data", "operators=self.operators",
                                "wo_folder=self.wo_folder", "box_sizes=load_box_sizes(online=None if self._online else False)",
                                "online=None if self._online else False)", "self.slip_data=p.result")))
check("popup title (Dhruvil 2026-09-25): 'Shipping Labels and Slip' in the title bar, the header and its error boxes",
      'self.title("Shipping Labels and Slip")' in popup and 'text="Shipping Labels and Slip"' in popup
      and popup.count('messagebox.showerror("Shipping Labels and Slip",') == 2 and '"Shipping Slip & Labels"' not in popup)
check("dropdowns pick-only like Apex's <select>s (a click anywhere opens the list): Prepared By / Ship Via readonly unless their list is empty, every box Size readonly; the wheel never re-picks a value",
      'state="readonly" if vals else "normal"' in popup and 'cv_=[_SLIP_VIA_NONE]+SLIP_SHIPVIA_OPTIONS' in popup
      and 'values=vals,font=("Segoe UI",10),width=20,height=min(25,len(vals)),state="readonly"' in popup and popup.count("self._no_wheel(") == 2
      and '"break")[1]' in src(A.SlipLabelsPopup._no_wheel) and "yview_scroll" in src(A.SlipLabelsPopup._no_wheel))
check("Size list: Apex's blank '--' first, the shared sizes, '+ Add new size…' last; picking it never keeps it as the value and opens the add dialog after the list closes",
      '<<ComboboxSelected>>' in popup and "self._size_picked(sv,sz)" in popup
      and "self.after_idle(lambda: self._add_size(sv,sz))" in src(A.SlipLabelsPopup._size_picked)
      and 'sv.set(getattr(sz,"_last","") or "")' in src(A.SlipLabelsPopup._size_picked)
      and "add_box_size(size,online=self._online_hint)" in src(A.SlipLabelsPopup._add_size) and "self._regrab()" in src(A.SlipLabelsPopup._add_size))
check("Size list is never cut short: 20 wide (the Segoe UI 14 list needs 143 px for '+ Add new size…'), as many rows as entries up to 25 — refreshed after an add; the header over it widened to match",
      "cb.config(values=vals,height=min(25,len(vals)))" in src(A.SlipLabelsPopup._add_size) and '("Size",22)' in popup)
_as = src(A.SlipLabelsPopup._add_size)
check("after the add-size dialog the keyboard focus goes back to that Size row in EVERY branch (simpledialog leaves it on the popup: Tab then landed on 'Print Slip')",
      "finally:" in _as and _as.index("finally:") < _as.index("sz.focus_set()") and _as.count("return") >= 2
      and _as.index("try:") < _as.index("simpledialog.askstring("))
check("no dead _slp alias (the footer button binds _open_slip_popup directly; _cp/_plp stay for their chip bindings)",
      not hasattr(A.ShippingApp, "_slp") and A.ShippingApp._cp is A.ShippingApp._open_coc_popup)
check("Save Draft: slip_data written to the draft and into the draft_save history event",
      "slip_data=self.slip_data" in draftw and '"slip": self.slip_data or None' in draftw)
check("Save and Close: safety draft + keep-draft re-save carry slip_data; close_record['slip'] + mirrored outputs; offline README rows",
      closew.count("slip_data=self.slip_data") == 2 and '"slip": self.slip_data or None' in closew
      and 'slip_out = ((self.slip_data or {}).get("outputs") or {})' in closew
      and '"slip_html": slip_out.get("slip_html") or None' in closew and '"labels_html": slip_out.get("labels_html") or None' in closew
      and "rec.update(close_record)" in closew and '("Shipping Slip html",slip_out.get("slip_html"))' in closew
      and "if self.slip_data:" in closew.split("readme_rows=")[1])
check("close-as-delete records had_slip + the printed pages' paths, and tells the operator they stay in the WO folder",
      '"had_slip": bool(self.slip_data)' in close and '"slip_outputs": slip_out or None' in close and '"slip_printed":' in close
      and close.count("stay in the WO folder") == 1 and close.count("{slip_note}") == 2)
check("popup: Print and Save first apply a box count TYPED into the spinbox (a tk.Button click fires no FocusOut)",
      src(A.SlipLabelsPopup._print).split("\n")[1].strip().startswith("self._rebuild_boxes()")
      and src(A.SlipLabelsPopup._save).split("\n")[1].strip().startswith("self._rebuild_boxes()")
      and "winfo_children" not in src(A.SlipLabelsPopup._rebuild_boxes)      # rows are adjusted, never torn down (FocusOut-safe for any count)
      and "self._add_box_row(" in src(A.SlipLabelsPopup._rebuild_boxes) and "fr_.destroy()" in src(A.SlipLabelsPopup._rebuild_boxes))
check("popup: the fallback warning names the page(s) and points at Re-emit; a later page written nowhere is reported without losing the written page",
      "Re-emit paperwork files the page into the WO folder" in popup and 'if res.get("errors"):' in src(A.SlipLabelsPopup._print)
      and 0 <= src(A.SlipLabelsPopup._print).find("self._base=d; self.result=d") < src(A.SlipLabelsPopup._print).find('if res.get("errors"):'))
check("box-count spinbox validator: blank or 1-2 digits only (a scanned barcode + Return can't build 99 rows)",
      all(A._slip_box_count_ok(s) for s in ("", "1", "9", "99", "07")) and not any(A._slip_box_count_ok(s) for s in ("100", "16452", "a", "1a", "-1", " 1", "1.5")))
check("popup: the spinbox validates every keystroke with it; initial focus is the Customer entry or Box 1's Size — never the spinbox",
      'validate="key",validatecommand=(self.register(_slip_box_count_ok),"%P")' in popup
      and "self._box_widgets[0][2]" in popup and "spn).focus_set()" not in popup and "spn.focus_set()" not in popup)
check("'Missing required fields' list capped at 8 + '… and N more' (a 99-box form gave a 100-line message box), used by _print",
      A._slip_missing_text(["Customer", "Box 1 qty"]) == "Customer\n  •  Box 1 qty" and A._slip_missing_text([]) == ""
      and A._slip_missing_text([f"Box {i} qty" for i in range(2, 100)]).endswith("Box 9 qty\n  •  … and 90 more")
      and A._slip_missing_text([f"Box {i} qty" for i in range(2, 100)]).count("\n") == 8 and A._slip_missing_text(list("abcdefgh")).count("\n") == 7
      and "_slip_missing_text(missing)" in src(A.SlipLabelsPopup._print))
_near = lambda a, b: a is not None and abs(a - b) < 1e-9
check("scroll-into-view maths: a row below the fold scrolls up just enough, one above scrolls down, visible -> None, nothing to scroll -> None, junk -> None, clamped to [0, 1]",
      _near(A._slip_scroll_fraction(731, 36, 0, 603, 817), (731 + 36 - 603) / 817)
      and _near(A._slip_scroll_fraction(10, 30, 200, 603, 817), 10 / 817)
      and A._slip_scroll_fraction(100, 30, 0, 603, 817) is None and A._slip_scroll_fraction(731, 36, 0, 900, 817) is None
      and A._slip_scroll_fraction("x", 1, 0, 1, 1) is None and A._slip_scroll_fraction(5000, 36, 0, 603, 817) == 1.0)
check("popup: sized to the whole one-box form (up to 980 px, screen-limited, never the old 780x760) and <FocusIn> anywhere in the body scrolls the focused row into view",
      "min(980,sh-90)" in popup and 'self.bind("<FocusIn>",lambda e: self._scroll_into_view(e.widget))' in popup
      and "cv.yview_moveto(f)" in src(A.SlipLabelsPopup._scroll_into_view) and "_slip_scroll_fraction(" in src(A.SlipLabelsPopup._scroll_into_view)
      and '"780x760"' not in popup)
check("HistoryDetail 'Output files': gated on real paths / printed stamps (a draft carrying an unprinted slip form gets no card of '—' rows); the four legacy rows only with outputs / file_status; the printed key carried explicitly",
      'if out or fs or any(sout.values()) or spr.get("slip") or spr.get("labels") or spr.get("pack"):' in hist and "if out or fs:" in hist
      and '("Shipping Slip","slip_html",None,"slip")' in hist and '("Shipping Labels","labels_html",None,"labels")' in hist
      and 'pkey.split("_")' not in hist and "if prkey and not p:" in hist and hist.count('sec("Output files")') == 1)
check("one escape helper (_slip_esc) for both pages' sheets (shared with the 'Print both' copy), the barcode rows' title copies and the page title; the summary docstring names no tooltip",
      A._slip_esc('<b>&"x') == "&lt;b&gt;&amp;&quot;x" and A._slip_esc(None) == "" and A._slip_esc(1.5) == "1.5"
      and "E = _slip_esc" in src(A._slip_sheet) and "E = _slip_esc" in src(A._labels_sheets)
      and "_html.escape" not in "".join(src(f) for f in (A._slip_sheet, A._labels_sheets, A.build_slip_html, A.build_labels_html,
                                                          A.build_slip_labels_html, A._slip_barcode_row))
      and "_slip_esc(title)" in src(A._slip_barcode_row) and "_slip_esc(title)" in src(A._slip_page)
      and "tooltip" not in (A.slip_printed_summary.__doc__ or ""))
check("HistoryDetail: the 'Open …' buttons sit on their own row (no longer clipped off the fixed-width action bar); window 820 wide",
      'ob.pack(side="bottom"' in hist and "_btn(ob,lbl" in hist and hist.index("ob=None") > hist.index("Re-emit paperwork")
      and '"820x680"' in src(A.HistoryDetail.__init__))
check("HistoryDetail: a page printed with the share down shows as 'not filed yet'; a delete shows the pages left behind",
      "not filed yet" in hist and '"Slip page left in WO folder"' in hist and '"Had Shipping Slip"' in hist)
check("HistoryDetail re-emit also offers a page that was printed offline (printed stamp, no recorded path)",
      'spr.get("slip")' in hreemit and 'spr.get("labels")' in hreemit)
check("popup: transient + grab, Escape + window-X cancel, re-grab after child dialogs, scanner-safe Return on buttons only",
      all(s in popup for s in ("self.transient(parent)", "self.grab_set()", '"<Escape>"', '"WM_DELETE_WINDOW"', "def _regrab", "self._regrab()",
                               'ps_b.bind("<Return>"', "self.wait_window()")))
check("popup: Apex lists (Size from the shared list, Apex's as the fallback) + spinbox 1..99 + live totals + the five buttons",
      all(s in popup for s in ("SLIP_BOXSIZE_OPTIONS", "SLIP_SHIPVIA_OPTIONS", "from_=1,to=99", "calc_box_totals",
                               '"🖨  Print Slip"', "Print Labels (", '"🖨  Print both"', '"✓  Save"', '"Cancel"')))
check("popup: Apex validation + print through the Tk-free helpers; share-fallback warning; nothing on a worker thread",
      all(s in popup for s in ("slip_required_missing(d,which)", "slip_print_pages(self._wo_folder,d,which)", "NOT saved to the share",
                               "slip_prefill(pi,monday,coc,packing,operator,wo,sc,existing=ed,division=division)")) and "threading" not in popup)
check("popup: date -> WWYY recomputed live; Prepared By offers the operator roster",
      "slip_datecode_wwyy(self._v[\"date\"].get())" in popup and "cv_=self._operators" in popup)
check("HistoryDetail: 'Open Shipping Slip' / 'Open Shipping Labels' next to 'Open COC', reading slip outputs or the mirrored ones",
      '"Open Shipping Slip","slip_html"' in hist and '"Open Shipping Labels","labels_html"' in hist and '"Open COC","coc_xlsx"' in hist
      and "p=out.get(key) or sout.get(key)" in hist)
check("HistoryDetail: slip block (boxes, totals, ship via, flags, codes, printed times) + re-emit gate covers slip",
      all(s in hist for s in ('sec("Shipping Slip & Labels")', "calc_box_totals(bx)", '"Ship Via"', '"Slip code"', '"Labels code"',
                              '"Slip printed"', '"Labels printed"', 'r.get("coc") or r.get("packing") or r.get("slip")')))
check("HistoryDetail re-emit lists the slip pages only when they were printed", '"Shipping Slip"' in hreemit and '"Shipping Labels"' in hreemit
      and 'sout.get("slip_html") or out.get("slip_html")' in hreemit)
check("no GUI side effects at import (no Tk root created)", not any(isinstance(v, A.tk.Tk) for v in vars(A).values()))

# ══ 11. Shared box sizes (Dhruvil 2026-09-25: "we keep getting new sizes") ══════
# Hermetic: BOX_SIZES_FILE and the offline cache point into a temp dir — nothing touches the share.
bx_root = Path(_tf.mkdtemp(prefix="pcc_boxsizes_"))
_bsf0, _cache0 = A.BOX_SIZES_FILE, A.CACHE_DIR
A.BOX_SIZES_FILE = str(bx_root / "share" / "box_sizes.json"); A.CACHE_DIR = bx_root / "cache"
(bx_root / "share").mkdir()
class VS(V):                                   # a Tk var stand-in that can also be set
    def set(s, v): s.v = v
def _wj(obj):
    with open(A.BOX_SIZES_FILE, "w", encoding="utf-8") as fh: json.dump(obj, fh)
try:
    APEX10 = ['5x5x5', '5x5x10', '10x10x5', '10x10x10', '12x10x8', '14x10x10', '16x12x10', '20x16x12', 'Pallet', 'Other']
    l0 = A.load_box_sizes(online=True)
    check("no shared file yet -> Apex's list (a copy: the constant can't be changed through it)", l0 == APEX10 and l0 is not A.SLIP_BOXSIZE_OPTIONS)
    sz1, ok1, er1 = A.add_box_size(" 18 X 12 x 6 ", online=True)
    on_disk = json.load(open(A.BOX_SIZES_FILE, encoding="utf-8"))
    check("add: typed '18 X 12 x 6' is filed as '18x12x6', ahead of Pallet / Other, written to the shared file and cached",
          ok1 and er1 is None and sz1 == APEX10[:8] + ["18x12x6", "Pallet", "Other"] and on_disk == sz1
          and json.load(open(A.CACHE_DIR / "box_sizes.json", encoding="utf-8")) == sz1 and A.load_box_sizes(online=True) == sz1)
    sz2, ok2, _ = A.add_box_size("18x12X6", online=True)
    check("the same size again (any case) is not added twice", ok2 and sz2 == sz1 and json.load(open(A.BOX_SIZES_FILE, encoding="utf-8")) == sz1)
    _wj(sz1[:-2] + ["20x20x20", "Pallet", "Other"])   # another station just added one
    sz3, ok3, _ = A.add_box_size("22x22x22", online=True)
    check("the list is re-read from the share before adding: a size another station added a moment ago is kept",
          ok3 and "20x20x20" in sz3 and "22x22x22" in sz3 and json.load(open(A.BOX_SIZES_FILE, encoding="utf-8")) == sz3 and sz3[-2:] == ["Pallet", "Other"])
    before_off = open(A.BOX_SIZES_FILE, encoding="utf-8").read()
    sz4, ok4, er4 = A.add_box_size("9x9x9", online=False)
    check("offline: not saved (the reason is returned), the shared file untouched, the form still gets the size for this shipment",
          ok4 is False and "not reachable" in er4 and "9x9x9" in sz4 and open(A.BOX_SIZES_FILE, encoding="utf-8").read() == before_off)
    check("offline load -> the last good copy cached on this PC", A.load_box_sizes(online=False) == sz3)
    _wj(["12x10x8", "", "  12X10x8 ", 7, "--", "+ Add new size…", "Pallet"])
    check("a hand-edited / junk file is cleaned: blanks, non-strings, reserved entries and duplicates dropped, order kept",
          A.load_box_sizes(online=True) == ["12x10x8", "Pallet"])
    with open(A.BOX_SIZES_FILE, "w", encoding="utf-8") as fh: fh.write("{not json")
    check("an unreadable shared file -> the cached copy (never an error, never empty)", A.load_box_sizes(online=True) == ["12x10x8", "Pallet"])
    shutil.rmtree(A.CACHE_DIR, ignore_errors=True)
    check("unreadable file and no cache -> Apex's list", A.load_box_sizes(online=True) == APEX10)
    try:
        A.save_box_sizes([""]); empty_raised = False
    except ValueError:
        empty_raised = True
    check("the list can never be saved empty", empty_raised)
    _deny_share = A._atomic_write_json
    A._atomic_write_json = lambda target, obj: (_ for _ in ()).throw(PermissionError(13, "Access is denied", str(target)))
    try:
        os.remove(A.BOX_SIZES_FILE)
        sz5, ok5, er5 = A.add_box_size("30x30x30", online=True)
    finally:
        A._atomic_write_json = _deny_share
    check("share reachable but the write refused -> not saved, the error text returned, the size still offered for this shipment",
          ok5 is False and "PermissionError" in er5 and "30x30x30" in sz5)
    check("normalize: spaces collapsed, X / × / * between digits -> x; other text kept",
          A.normalize_box_size(" 18 X 12 x 6 ") == "18x12x6" and A.normalize_box_size("18×12×6") == "18x12x6"
          and A.normalize_box_size("10*10*5") == "10x10x5" and A.normalize_box_size("Pallet   48x40") == "Pallet 48x40"
          and A.normalize_box_size(None) == "")
    check("box_size_problem: blank, over 24 characters and the reserved entries are refused; a normal size passes",
          A.box_size_problem("") and A.box_size_problem("x" * 25) and A.box_size_problem("--") and A.box_size_problem("-- Select --")
          and A.box_size_problem("+ Add new size…") and A.box_size_problem("+ add something") and A.box_size_problem("18x12x6") is None)
    try:
        A.add_box_size("   ", online=True); blank_raised = False
    except ValueError:
        blank_raised = True
    check("add_box_size refuses a blank size (the popup checks first)", blank_raised)
    check("stored values: '--' / '+ Add new size…' are never a size; Ship Via's '-- Select --' is blank",
          A._slip_size_value("--") == "" and A._slip_size_value("+ Add new size…") == "" and A._slip_size_value(" 18x12x6 ") == "18x12x6")
    stc = types.SimpleNamespace(_v={k: VS(str(sample().get(k) or "")) for k in A.SlipLabelsPopup.FIELDS},
                                _box_rows=[(V("--"), V("5"), V("1")), (V("18x12x6"), V("5"), V("2"))], _base=sample())
    stc._v["ship_via"] = V("-- Select --"); stc._v["customer_parts"] = V("No")
    for k in ("fair", "si", "ppap", "cofc"): stc._v[k] = V(False)
    dc_ = A.SlipLabelsPopup._collect(stc)
    check("_collect: Apex's blank options read back as blank (Size '--', Ship Via '-- Select --'); a picked size is kept",
          dc_["ship_via"] == "" and [b["size"] for b in dc_["boxes"]] == ["", "18x12x6"])
    sv_ = VS("12x10x8"); cb_ = types.SimpleNamespace(_last="12x10x8"); calls_ = []
    stp_ = types.SimpleNamespace(_sizes=["12x10x8"], after_idle=lambda fn: calls_.append(fn), _add_size=lambda sv, sz: None)
    check("Size values: Apex's '--' first, the shared list, '+ Add new size…' last",
          A.SlipLabelsPopup._size_values(stp_) == ["--", "12x10x8", "+ Add new size…"])
    sv_.set("+ Add new size…"); A.SlipLabelsPopup._size_picked(stp_, sv_, cb_)
    check("picking '+ Add new size…' puts the row back to its last size and schedules the add dialog (never keeps the entry)",
          sv_.get() == "12x10x8" and len(calls_) == 1)
    sv_.set("10x10x5"); A.SlipLabelsPopup._size_picked(stp_, sv_, cb_)
    check("picking a real size remembers it as the row's last size", cb_._last == "10x10x5" and len(calls_) == 1)
    admin = src(A.AdminPanel)
    check("Admin > Box sizes tab: add / remove go through the merging writers (fresh read, never the tab's stale copy), never the last one, a failed save reported; the tab never writes the list itself",
          '("bx","  Box sizes  ")' in admin and 'elif k=="bx": self._tbx()' in admin
          and "fresh=load_box_sizes()" in src(A.AdminPanel._abx) and "add_box_size(s)" in src(A.AdminPanel._abx)
          and "remove_box_size(s)" in src(A.AdminPanel._rbx) and "Keep at least one box size." in src(A.AdminPanel._rbx)
          and "NOT saved" in src(A.AdminPanel._bx_not_saved) and "save_box_sizes" not in admin and "_box_sizes_with(" not in admin)
    # the writers never start from a stale / cached copy
    _wj(["12x10x8", "18x12x6", "Pallet", "Other"])
    with open(A.BOX_SIZES_FILE, "w", encoding="utf-8") as fh: fh.write("{half-written")
    junk = open(A.BOX_SIZES_FILE, encoding="utf-8").read()
    sz6, ok6, er6 = A.add_box_size("40x40x40", online=True)
    check("add: a shared file that is there but unreadable is NEVER overwritten (the cached copy would erase other stations' sizes) — not saved, the reason returned, the size still offered",
          ok6 is False and "could not be read" in er6 and open(A.BOX_SIZES_FILE, encoding="utf-8").read() == junk and "40x40x40" in sz6)
    _wj(["12x10x8", "18x12x6", "22x22x22", "Pallet", "Other"])      # the share NOW (a station added 22x22x22 after the admin opened the tab)
    rz1, rk1, re1 = A.remove_box_size("18X12x6", online=True)
    check("remove: reads the share fresh (the size another station just added stays), matches any case, saves",
          rk1 and re1 is None and rz1 == ["12x10x8", "22x22x22", "Pallet", "Other"] and json.load(open(A.BOX_SIZES_FILE, encoding="utf-8")) == rz1)
    before_rm = open(A.BOX_SIZES_FILE, encoding="utf-8").read()
    rz2, rk2, _ = A.remove_box_size("99x99x99", online=True)
    check("remove: a size already gone -> nothing written, saved", rk2 and rz2 == rz1 and open(A.BOX_SIZES_FILE, encoding="utf-8").read() == before_rm)
    rz3, rk3, re3 = A.remove_box_size("12x10x8", online=False)
    check("remove offline: refused, the reason returned, the shared file untouched", rk3 is False and "not reachable" in re3
          and open(A.BOX_SIZES_FILE, encoding="utf-8").read() == before_rm)
    _wj(["Other"])
    rz4, rk4, re4 = A.remove_box_size("other", online=True)
    check("remove: never leaves the list empty", rk4 is False and re4 == "Keep at least one box size." and json.load(open(A.BOX_SIZES_FILE, encoding="utf-8")) == ["Other"])
    with open(A.BOX_SIZES_FILE, "w", encoding="utf-8") as fh: fh.write("{half-written")
    rz5, rk5, re5 = A.remove_box_size("Other", online=True)
    check("remove: an unreadable shared file is never overwritten", rk5 is False and "could not be read" in re5
          and open(A.BOX_SIZES_FILE, encoding="utf-8").read() == "{half-written")
    os.remove(A.BOX_SIZES_FILE)
    rz6, rk6, _ = A.remove_box_size("Other", online=True)
    check("remove from Apex's starting list (no shared file yet): the file is created with the other nine",
          rk6 and rz6 == APEX10[:-1] and json.load(open(A.BOX_SIZES_FILE, encoding="utf-8")) == APEX10[:-1])
finally:
    A.BOX_SIZES_FILE, A.CACHE_DIR = _bsf0, _cache0
    shutil.rmtree(bx_root, ignore_errors=True)
check("BOX_SIZES_FILE sits beside the operator roster on the share", A.BOX_SIZES_FILE == A.APP_DATA_DIR + "\\box_sizes.json")

# ══ 12. Second-chance Monday lookup (Dhruvil 2026-09-25: "Customer PO is not filling") ══
wf_ = Path(r"\\PCC2\X\377 Anduril\134-2201-970 [4065 Rev. A00 Rev. A00]\Production Data\16588.1")
check("monday_folder_wo: the WO number a folder name carries",
      A.monday_folder_wo("16588.1") == "16588.1" and A.monday_folder_wo("16588.1 REWORK") == "16588.1"
      and A.monday_folder_wo("16588.1_A") == "16588.1" and A.monday_folder_wo("") == "" and A.monday_folder_wo(None) == "")
h1 = {"wo": "16588.1", "eng": "4065", "po": "USPO61707"}
check("monday_hit_matches_folder: the same WO (or WO + '.', ' ', '_' suffix) or the job folder's Eng# matches; anything else doesn't",
      A.monday_hit_matches_folder(h1, wf_) and A.monday_hit_matches_folder({"wo": "16588", "eng": ""}, wf_)
      and A.monday_hit_matches_folder({"wo": "99999.1", "eng": "4065"}, wf_)
      and not A.monday_hit_matches_folder({"wo": "16589.1", "eng": "4066"}, wf_)
      and not A.monday_hit_matches_folder({"wo": "", "eng": "N/A"}, wf_) and not A.monday_hit_matches_folder(None, wf_))
check("pick_monday_hit_for_folder: exactly one item with the folder's WO -> it; none or two -> None (never a guess)",
      A.pick_monday_hit_for_folder([h1], wf_) is h1 and A.pick_monday_hit_for_folder([], wf_) is None
      and A.pick_monday_hit_for_folder([h1, dict(h1, po="OTHER")], wf_) is None
      and A.pick_monday_hit_for_folder([{"wo": "16589.1"}], wf_) is None and A.pick_monday_hit_for_folder([{"wo": ""}], wf_) is None
      and A.pick_monday_hit_for_folder([{"wo": "16589.1"}, h1], wf_) is h1 and A.pick_monday_hit_for_folder(None, wf_) is None)
check("_load_workorder: the first hit is kept only if it is this folder's job; with none — and only when Monday ANSWERED the first time — the answer in hand is tried first, then Monday once more with the folder's own WO (not when that is what was typed); taken only if unique, before the confirm (its PO is shown there)",
      "if self._monday_hit and not monday_hit_matches_folder(self._monday_hit, wf):" in load
      and 'monday_hits=hits; monday_ok=getattr(self,"_monday_last_ok",False)' in load
      and "if self._monday_hit is None and MONDAY_LOOKUP_ENABLED and MONDAY_API_KEY and monday_ok:" in load
      and "self._monday_hit=pick_monday_hit_for_folder(monday_hits, wf)" in load
      and 'if self._monday_hit is None and fwo and fwo.lower()!=(wo or "").strip().lower():' in load
      and "self._monday_hit=pick_monday_hit_for_folder(self._monday_lookup_ui(fwo), wf)" in load
      and load.index("pick_monday_hit_for_folder(monday_hits, wf)") < load.index("pick_monday_hit_for_folder(self._monday_lookup_ui(fwo), wf)")
      < load.index("dl=ConfirmDialog(self,wo,wf,div,monday=self._monday_hit)")
      and load.index("monday_hit_matches_folder(") < load.index("fwo=monday_folder_wo(wf.name)"))
mui = src(A.ShippingApp._monday_lookup_ui)
check("_monday_lookup_ui records whether Monday ANSWERED (reset first — the busy early-return counts as a failure — then set from the worker's status)",
      mui.index("self._monday_last_ok=False") < mui.index('if getattr(self,"_monday_busy",False):')
      and "monday_lookup(term,status=st)" in mui and 'self._monday_last_ok=bool(box["done"] and st.get("ok"))' in mui)
_mq0 = A._monday_query
try:
    stt = {}
    A._monday_query = lambda term, timeout: (_ for _ in ()).throw(TimeoutError("timed out"))
    r_to = A.monday_lookup("16588", status=stt)
    check("monday_lookup status: a timeout -> [] and ok False (no second lookup after it)", r_to == [] and stt == {"ok": False})
    A._monday_query = lambda term, timeout: {"data": {"by_wo": {"items": []}, "by_eng": {"items": []}}}
    stt = {}; r_nm = A.monday_lookup("16588", status=stt)
    check("monday_lookup status: Monday answered 'no match' -> [] and ok True (the second lookup may run)", r_nm == [] and stt == {"ok": True})
    A._monday_query = lambda term, timeout: {"errors": [{"message": "rate limited"}]}
    stt = {}; A.monday_lookup("16588", status=stt)
    check("monday_lookup status: a GraphQL error / rate limit -> ok False", stt == {"ok": False})
    stt = {}; A.monday_lookup("   ", status=stt)
    check("monday_lookup status: nothing to look up -> ok False; the old call shape (no status) still works",
          stt == {"ok": False} and A.monday_lookup("16588") == [])
finally:
    A._monday_query = _mq0

# ══ 13. The paperwork carries the BOM rev (Dhruvil 2026-10-05: "It should always carry BOM rev. If monday has
#        BOM rev column then it should always prioritize that to folder name") ══════════════════════════════
check("folder names: the BOM rev is the FIRST rev ('Customer PN [PCC Eng - BOM Rev - Assy Rev]'); dotted / hyphenated revs read whole; '-' placeholders skipped",
      [A.bom_rev_from_inside(s) for s in ("3486LF Rev. A Rev. 1", "4079 Rev. A Rev. A4", "3414 Rev. 2.02 Rev. ", "4033 Rev. 1.1.11 Rev. 1.1.10",
                                          "3344LF Rev. 1.1 Rev. - ", "2429-0 Rev. 5-4-23", "3156LF Rev. - A Rev. - A", "3694 Rev. B Assy Rev",
                                          "3694 Rev C", "4081 Rev. B00", "3706LF Rev. 5r2 Rev. -", "3267 Rev.  Rev. ", "3156LF Rev. - Rev. -", "", None)]
      == ["A", "A", "2.02", "1.1.11", "1.1", "5-4-23", "A", "B", "C", "B00", "5r2", "", "", "", ""])
check("rev_from_inside (the LAST rev) reads dotted revs whole too: 'Rev. 1.1.11 Rev. 1.1.10' -> '1.1.10' (was '1'), 'Rev. 2.1 Rev. 2.1' -> '2.1' (was '2')",
      A.rev_from_inside("4033 Rev. 1.1.11 Rev. 1.1.10") == "1.1.10" and A.rev_from_inside("3342LF Rev. 2.1 Rev. 2.1") == "2.1"
      and A.rev_from_inside("3486LF Rev. A Rev. 1") == "1" and A.rev_from_inside("4081 Rev. B00") == "B00" and A.rev_from_inside("3267 Rev.  Rev. ") == "")
check("Monday's rev (its BOM Rev column): placeholders are none, a typed 'Rev' word is dropped, two DIFFERENT revs are none (never a guess)",
      [A.monday_bom_rev(v) for v in ("C", " A00 ", "-", "--", "N/A", "", None, "Rev 03", "REV. B", "Revision 2", "Rev", "A, A", "A, Rev A", "A, B", "8-24-21")]
      == ["C", "A00", "", "", "", "", "", "03", "B", "2", "", "A", "A", "", "8-24-21"])
bx13 = Path(_tf.mkdtemp(prefix="pcc_bomrev_"))
try:
    def wo13(job, wo="16000.1"):
        p = bx13 / "198 Prism" / job / "Production Data" / wo; p.mkdir(parents=True); return p
    check("parse_wo_path takes the folder name's BOM rev (first rev, read whole)",
          A.parse_wo_path(wo13("1427765_3486LF Rev. A Rev. 1"))["rev"] == "A"
          and A.parse_wo_path(wo13("301 [3414 Rev. 2.02 Rev. ]"))["rev"] == "2.02"
          and A.parse_wo_path(wo13("GSR10 [4033 Rev. 1.1.11 Rev. 1.1.10]"))["rev"] == "1.1.11"
          and A.parse_wo_path(wo13("PC1060C_3694 Rev C"))["rev"] == "C"
          and A.parse_wo_path(wo13("15102203 Water Draw Box [3156LF Rev. - Rev. -]"))["rev"] == "")
finally:
    shutil.rmtree(bx13, ignore_errors=True)
def merged(folder, hit):
    st = types.SimpleNamespace(_monday_hit=hit)
    info = dict(folder); A.ShippingApp._merge_monday(st, info); return info
f13 = {"customer": "Rad Source", "pn": "1427765", "eng": "3486LF", "rev": "A"}
check("Monday's BOM rev WINS over the folder name's; the folder's is the fallback when Monday has none / a placeholder / two different revs / no hit",
      merged(dict(f13, rev="1"), {"rev": "A"})["rev"] == "A" and merged(f13, {"rev": "B"})["rev"] == "B"
      and merged(f13, {"rev": "-"})["rev"] == "A" and merged(f13, {"rev": ""})["rev"] == "A" and merged(f13, {"rev": "A, B"})["rev"] == "A"
      and merged(f13, {"rev": "Rev 03"})["rev"] == "03" and merged(f13, None)["rev"] == "A" and merged(dict(f13, rev=""), {"rev": "C"})["rev"] == "C")
check("…while customer / eng / P/N stay folder-first (Monday fills only their blanks) and PO / CPN come from Monday as before",
      merged(dict(f13, customer="N/A", eng=""), {"customer": "198 Prism Systems", "eng": "3694", "po": "26-040", "cpn": "PC1060"})
      == {"customer": "Prism Systems", "pn": "1427765", "eng": "3694", "rev": "A", "po": "26-040", "cpn": "PC1060"}
      and merged(f13, {"customer": "Other", "eng": "9999", "rev": "A"})["customer"] == "Rad Source"
      and merged(f13, {"customer": "Other", "eng": "9999", "rev": "A"})["eng"] == "3486LF")
check("slip pre-fill: the job's rev first (= Monday's BOM rev, else the folder's), the last-resort Monday rev goes through the same rule",
      A.slip_prefill({}, {"rev": "-"}, None, None, "", "", 0)["rev"] == "" and A.slip_prefill({}, {"rev": "Rev 03"}, None, None, "", "", 0)["rev"] == "03"
      and A.slip_prefill({"rev": "A"}, {"rev": "B"}, None, None, "", "", 0)["rev"] == "A")
check("wiring: parse_wo_path uses bom_rev_from_inside; the load status line names Monday's BOM rev only when it is the one used; the COC still defaults to the job's rev",
      'r["rev"]=bom_rev_from_inside(inside)' in src(A.parse_wo_path)
      and "if monday_bom_rev(h.get(\"rev\")): pre.append(f\"BOM Rev {monday_bom_rev(h.get('rev'))}\")" in load
      and 'self._rv=tk.StringVar(value=pi.get("rev",""))' in src(A.COCPopup.__init__))

# ══ 14. PACKING LABEL — Apex #lbl-pack, 3 x 2 in (Dhruvil 2026-10-19: "I need to replicate whatever is in apex") ══
def pk_sample(mode="itar"):
    d = sample(); d.update(pack_qty="10", pack_cage="1WXP2", pack_mode=mode, pack_date="10/19/2026", pack_datecode="2643",
                           pack_code="4092-16452.1-PACK-20261019-4321")
    return d
pk_doc = A.build_pack_html(pk_sample()); pk_ind = A.build_pack_html(pk_sample("industrial"))
if APEX:
    _drop_pk = lambda t, a: apex_drop(t, a) or "pack-mode" in a.get("class", "").split() or "pack-fields" in a.get("class", "").split()
    _vi0 = VALUE_IDS
    VALUE_IDS = _vi0 + ("pack-",)
    try:
        def _pk_proj(p, apex):
            out = []
            for t in project(p, apex):
                if apex and t[1] == "lbl-pack": t = (t[0], t[1], t[2] + " active") + t[3:]
                out.append(t)
            return out
        pexp = _pk_proj(seq_of(apex_src, "lbl-pack", drop=_drop_pk), True)
        pgot = _pk_proj(seq_of(pk_doc, "lbl-pack", leaf=app_leaf), False)
    finally:
        VALUE_IDS = _vi0
    check("packing label: Apex #lbl-pack element sequence (tag/id/class/style/depth/attrs/static text) reproduced — the ITAR/Industrial buttons + WO/Eng pickers (screen-only in Apex) left out, inputs as spans",
          pexp == pgot and len(pexp) > 30)
    if pexp != pgot: print("      first difference:", first_diff(pexp, pgot))
    check("packing label: static texts compared (P/N, QTY, CAGE, PO, D/C, DATE, WO#/Lot, Printed Circuits Corp.)",
          {t for *_, t in pexp} >= {"P/N", "QTY", "CAGE", "PO", "D/C", "DATE", "WO#/Lot", "Printed Circuits Corp."})
    R2, _acss = apex_rules()
    pk_styles = "\n".join(re.findall(r"<style[^>]*>(.*?)</style>", pk_doc, re.S))
    pk_sel = [".label.label-kit,.label.label-pack,.label.label-spack", ".pack-hdr", ".pack-co", ".pack-body", ".pack-fld", ".pack-r", ".pack-k",
              ".pack-v", ".pack-vbig", ".pack-vdisp", ".pack-fld-bc", ".pack-fld-bc svg", '#lbl-pack[data-mode="itar"] .pack-ind-only',
              '#lbl-pack[data-mode="industrial"] .pack-itar-only', ".kit-selects,.pack-fields,.pack-mode", "#pack-code-display,#shipl-code-display"]
    _miss = [r for s_ in pk_sel for r in R2.get(s_, ["<" + s_ + " not in Apex>"]) if r not in pk_styles]
    check(f"packing label: every Apex rule it uses present VERBATIM ({len(pk_sel)} selectors incl. its print rules and the hidden form code)", not _miss)
    for x in _miss[:5]: print("      missing:", x[:90])
    _sz = [r for sel, rs in R2.items() if "#lbl-pack.active" in sel and "pack-3x" not in sel for r in rs]
    check("packing label: Apex's 3 x 2 print-size rules for #lbl-pack.active present verbatim", len(_sz) >= 2 and all(r in pk_styles for r in _sz))
check("packing label page: @page 3in x 2in margin 0, the wrapper sheet 3 x 2 in print, the 4 x 6 page rules untouched for the slip pages",
      "@media print{@page{size:3in 2in;margin:0}}" in pk_doc and "@media print{.sheet{width:3in;height:2in;}}" in pk_doc
      and "@page{size:4in 6in;margin:0}" not in pk_doc and "@page{size:4in 6in;margin:0}" in slip_doc and "size:3in 2in" not in slip_doc)
check("packing label: hint names the 3 x 2 Shipping printer; one page; title 'Packing Label - <WO>'; filed page auto-prints, print copy prints-then-closes",
      "Printer: 3 &times; 2 Shipping &middot; 3 &times; 2 in" in pk_doc and "Packing Label &mdash; 1 page" in pk_doc
      and "<title>Packing Label - 16452.1</title>" in pk_doc and "setTimeout(function(){window.print();},600);" in pk_doc
      and A._PRINT_THEN_CLOSE_JS in A.build_pack_html(pk_sample(), print_copy=True) and pk_doc.count('id="lbl-pack"') == 1)
check("packing label values as static text: P/N, QTY, CAGE, PO, D/C, today's DATE, WO, the form code (hidden in print by Apex's rule)",
      all(x in pk_doc for x in ('id="pack-cpn">125-1030-970/C<', 'id="pack-qty">10<', 'id="pack-cage">1WXP2<', 'id="pack-po">PO-778<',
                                'id="pack-dc">2643<', 'id="pack-date-disp">10/19/2026<', 'id="pack-wo-disp">16452.1<',
                                'id="pack-code-display">4092-16452.1-PACK-20261019-4321<')))
check("packing label barcodes: JsBarcode width 1 / height 12 for every value, the same ids as Apex; an empty value -> no barcode",
      all(A.code128_svg(v, i, 1, 12) in pk_doc for v, i in (("125-1030-970/C", "bc-pack-cpn"), ("10", "bc-pack-qty"), ("1WXP2", "bc-pack-cage"),
          ("PO-778", "bc-pack-po"), ("2643", "bc-pack-dc"), ("10/19/2026", "bc-pack-date"), ("16452.1", "bc-pack-wo")))
      and 'height="12px"' in A.code128_svg("10", "x", 1, 12) and '<svg id="bc-pack-po"></svg>' in A.build_pack_html(dict(pk_sample(), po=""))
      and A.code128_svg("16452.1", "bc-wo") == A.code128_svg("16452.1", "bc-wo", 1.3, 26))
check("packing label mode = Apex's data-mode: ITAR/MAG shows CAGE + D/C, Industrial shows DATE (Apex's own hide rules do it)",
      'id="lbl-pack" data-mode="itar"' in pk_doc and 'id="lbl-pack" data-mode="industrial"' in pk_ind
      and A.pack_mode_for("MAG - ITAR") == "itar" and A.pack_mode_for("Industrial") == "industrial" and A.pack_mode_for(None) == "itar")
check("packing label form code = Apex generateCode('pack'): ENG-WO-PACK-YYYYMMDD-RAND4; D/C = Apex calcDateCode (YY + ISO week)",
      A.pack_form_code("4092", "16452.1", "2026-10-19", 4321) == "4092-16452.1-PACK-20261019-4321"
      and A.pack_form_code("", "", "2026-10-19", 1000) == "ENG-WO-PACK-20261019-1000"
      and A.pack_datecode_yyww("2026-10-19") == "2643" and A.pack_datecode_yyww("2027-01-01") == "2753")
check("packing label data: QTY starts at 1 and CAGE at 1WXP2 (Apex's values); the pre-fill takes the mode from the division; Apex's required fields WO + P/N",
      A.slip_data_template()["pack_qty"] == "1" and A.slip_data_template()["pack_cage"] == "1WXP2"
      and A.slip_prefill({}, {}, None, None, "", "16452.1", 3, division="Industrial")["pack_mode"] == "industrial"
      and A.slip_prefill(existing=sample(), division="MAG - ITAR")["pack_mode"] == "itar"
      and A.slip_required_missing(dict(sample(), cpn="", wo=""), "pack") == ["Work Order / Lot", "Customer P/N"]
      and A.slip_required_missing(dict(sample(), customer=""), "pack") == [])
pk_root = Path(_tf.mkdtemp(prefix="pcc_pack_")); pk_tmp = _tf.mkdtemp(prefix="pcc_pack_tmp_")
_tf.gettempdir = lambda: pk_tmp
try:
    wo_pk = pk_root / "16452.1"; wo_pk.mkdir()
    dpk = sample(); opened_pk = []
    rpk = A.slip_print_pages(str(wo_pk), dpk, "pack", open_file=lambda p: opened_pk.append(p))
    fdp = A.fmt_file_date()
    check("print 'pack': the label is FILED in <WO>\\Shipping Slip as <eng>-<date>-Packing_Label.html, ONE print copy opened from the temp folder, only the packing label stamped",
          Path(rpk["pack_html"]).parent == wo_pk / "Shipping Slip" and Path(rpk["pack_html"]).name == f"4092-{fdp}-Packing_Label.html"
          and opened_pk == [rpk["print_html"]] and "Print_Packing_Label" in Path(rpk["print_html"]).name
          and rpk["slip_html"] is None and rpk["labels_html"] is None and dpk["printed"]["pack"] and not dpk["printed"]["slip"]
          and dpk["outputs"]["pack_html"] == rpk["pack_html"] and re.match(r"^4092-16452\.1-PACK-\d{8}-\d{4}$", dpk["pack_code"]))
    check("print 'pack': today's date + date code frozen into the record and its page snapshot (History / Re-emit reproduce the paper)",
          dpk["pack_date"] == datetime.date.today().strftime("%m/%d/%Y") and dpk["pack_datecode"] == A.pack_datecode_yyww()
          and dpk["pages"]["pack"]["pack_code"] == dpk["pack_code"] and dpk["pages"]["pack"]["pack_date"] == dpk["pack_date"]
          and A.slip_printed_summary(dpk).endswith("Packing label printed " + dpk["printed"]["pack"]))
    pk_txt = open(rpk["pack_html"], encoding="utf-8").read()
    rec_pk = {"event_type": "close", "eng": "4092", "wo": "16452.1", "serials": [], "wo_folder": str(wo_pk), "outputs": {},
              "slip": json.loads(json.dumps(dpk))}
    os.remove(rpk["pack_html"])
    rr_pk = A.reemit_paperwork(rec_pk)
    check("re-emit rebuilds a missing packing label under its ORIGINAL name, byte-identical (recorded code + frozen date); only that page",
          rr_pk["ok"] and [Path(x).name for x in rr_pk["wrote"]] == [f"4092-{fdp}-Packing_Label.html"]
          and open(rpk["pack_html"], encoding="utf-8").read() == pk_txt)
    check("the QTY-vs-serials prompt is for the slip only (the packing label's QTY is its own) and the popup has its Print Packing Label button",
          'if which!="pack" and self._sc>0' in src(A.SlipLabelsPopup._print) and 'self._print("pack")' in popup
          and '"pack_qty","pack_cage"' in popup and "division=self.division_var.get()" in opener)
    hist2 = src(A.HistoryDetail._build)
    check("History: 'Open Packing Label' button, packing label rows (QTY / CAGE or DATE / code / printed), Output files row; re-emit lists it",
          '("Open Packing Label","pack_html")' in hist2 and '"Packing label printed"' in hist2 and '("Packing Label","pack_html",None,"pack")' in hist2
          and 'docs.append("Packing Label")' in hreemit)
finally:
    _tf.gettempdir = _orig_gettempdir
    shutil.rmtree(pk_root, ignore_errors=True); shutil.rmtree(pk_tmp, ignore_errors=True)

print("ALL PASSED" if not fail else f"{fail} FAILED")
sys.exit(1 if fail else 0)
