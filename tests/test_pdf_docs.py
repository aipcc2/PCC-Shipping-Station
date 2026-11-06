"""
Tests for pdf_docs.py (the PDFs the app draws itself, no Excel).
Run from the repo folder:   py -3.12 -m unittest discover -s tests -v
Needs reportlab (installed by build_exe.bat). pypdf is optional: with it, the text on each page is checked too.
"""
import re
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

try:
    import pdf_docs
except ImportError:  # reportlab not installed
    pdf_docs = None
try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

LOGO = Path(__file__).resolve().parent.parent / "logo.png"
SERIALS = [f"PCC2611-{i:04d}" for i in range(1, 31)]
COC = {"ship_date": "11/5/2026", "po": "2859984-1", "pn": "15102203 Water Draw Box", "rev": "B", "qty": 30,
       "datecode": "2645", "job": "4099LF", "type": "CCA", "assembly_process": "Lead Free PCB Assembly"}
PL = {"customer": "Anduril Industries", "packing_no": "PL-2207", "po": "2859984-1",
      "pn": "15102203 Water Draw Box Assembly With A Very Long Name", "wo": "16691.1", "eng": "4099LF",
      "date": "11/5/2026", "datecode": "2645"}


@unittest.skipIf(pdf_docs is None, "reportlab is not installed")
class PdfDocsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def check(self, path, expected_pages, must_contain):
        path = Path(path)
        self.assertTrue(path.is_file())
        data = path.read_bytes()
        self.assertTrue(data.startswith(b"%PDF-"))
        if PdfReader is None:
            self.assertEqual(len(re.findall(rb"/Type\s*/Page[^s]", data)), expected_pages)
            return
        reader = PdfReader(str(path))
        self.assertEqual(len(reader.pages), expected_pages)
        text = "\n".join(page.extract_text() for page in reader.pages)
        for s in must_contain:
            self.assertIn(s, text)

    def test_pdf_sits_next_to_the_xlsx_with_the_same_name(self):
        self.assertEqual(pdf_docs.pdf_path_for(r"C:\WO\16691.1\Certificate of Compliance\4099-11-5-26-COC_142233.xlsx"),
                         Path(r"C:\WO\16691.1\Certificate of Compliance\4099-11-5-26-COC_142233.pdf"))

    def test_coc(self):
        out = pdf_docs.write_coc_pdf(self.dir / "4099-11-5-26-COC.xlsx", COC, SERIALS)
        self.assertTrue(out.endswith("4099-11-5-26-COC.pdf"))
        self.check(out, 3, ["CERTIFICATE OF COMPLIANCE", "2859984-1", "15102203 Water Draw Box Rev.- B",
                            "30ea.", "2645 (YYWW)", "4099LF", "PCC2611-0001", "PCC2611-0030", "Hung Ly"])

    def test_packing_list(self):
        out = pdf_docs.write_packing_list_pdf(self.dir / "4099-11-5-26-Packing_List.xlsx", PL, SERIALS, logo_path=LOGO)
        self.check(out, 3, ["PACKING LIST", "Anduril Industries", "PL-2207", "16691.1", "TOTAL UNITS", "PCC2611-0030"])

    def test_serial_list(self):
        out = pdf_docs.write_serials_pdf(self.dir / "4099-11-5-26 serials.xlsx", "4099LF", "16691.1", SERIALS,
                                         "11/5/2026", "02:15 PM", "Dhruvil")
        self.check(out, 2, ["Serial Numbers - WO 16691.1 (Eng #4099LF)", "Operator: Dhruvil", "Total serials: 30", "PCC2611-0030"])

    def test_no_serials_and_odd_values_still_make_a_pdf(self):
        coc = dict(COC, rev="", qty="", po="PO#/ü 55")
        self.check(pdf_docs.write_coc_pdf(self.dir / "a.xlsx", coc, []), 1, ["Rev.-"])
        pl = {k: "" for k in PL}
        self.check(pdf_docs.write_packing_list_pdf(self.dir / "b.xlsx", pl, [], logo_path=None), 1, ["TOTAL UNITS"])
        self.check(pdf_docs.write_serials_pdf(self.dir / "c.xlsx", "", "", []), 1, ["Total serials: 0"])


if __name__ == "__main__":
    unittest.main()
