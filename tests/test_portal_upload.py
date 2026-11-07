"""
Tests for portal_upload.py. Run from the repo folder:   py -m unittest discover -s tests -v
No network: uploads are replaced with a fake.
"""
import json
import os
import sys
import tempfile
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import portal_upload as pu  # noqa: E402

KEY = "k" * 40
CONFIG = {"enabled": True, "key": KEY, "url": "https://example.test/fn"}
COC = {"po": "2859984-1", "ship_date": "11/5/2026", "job": "4099"}
PL = {"po": "2859984-1", "packing_no": "PL-2207", "date": "11/5/2026"}


class Recorder:
    """Stands in for the network: remembers every call and answers like the portal."""
    def __init__(self, status=200, body=None):
        self.calls = []
        self.status = status
        self.body = body if body is not None else {"ok": True, "document_id": "doc-1"}

    def __call__(self, url, headers, body):
        self.calls.append((url, headers, body))
        return self.status, self.body


class Env(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        base = Path(self.tmp.name)
        self.ind = base / "Production" / "Industrial"
        self.itar = base / "Production" / "MAG - ITAR"
        self.q = base / "queue"
        self.roots = dict(industrial_root=str(self.ind), itar_root=str(self.itar))

    def tearDown(self):
        self.tmp.cleanup()

    def wo(self, root, customer="89 Anduril", job="4099 Board [4099 Rev. B]", wo="16691.1"):
        wf = root / customer / job / "Production Data" / wo
        (wf / "Certificate of Compliance").mkdir(parents=True)
        (wf / "Packing List").mkdir(parents=True)
        coc = wf / "Certificate of Compliance" / "4099-11-5-26-COC.pdf"
        pl = wf / "Packing List" / "4099-11-5-26-Packing_List.pdf"
        coc.write_bytes(b"%PDF-1.7 test")
        pl.write_bytes(b"%PDF-1.7 test")
        return wf, coc, pl

    def enqueue(self, division, wf, coc, pl, **kw):
        return pu.enqueue_shipment(division, wf, wf.name, "4099", coc_pdf=coc, coc_data=COC, packing_pdf=pl,
                                   packing_data=PL, serials=["SN1", "SN2"], config=kw.pop("config", CONFIG),
                                   qdir=self.q, **self.roots, **kw)


class GuardTests(Env):
    """Industrial AND MAG - ITAR jobs go to the portal (Sahil, 2026-11-07). Still blocked: unknown divisions,
    files outside the division's own share folder, and anything carrying an export-control marking."""

    def test_industrial_job_is_allowed(self):
        wf, coc, pl = self.wo(self.ind)
        self.assertIsNone(pu.block_reason("Industrial", [wf, coc, pl], ["2859984-1"], **self.roots))

    def test_mag_itar_job_is_allowed_and_queued(self):
        wf, coc, pl = self.wo(self.itar)
        status, msg = self.enqueue("MAG - ITAR", wf, coc, pl)
        self.assertEqual(status, "queued")
        jobs = pu.load_jobs(self.q)
        self.assertEqual(len(jobs), 2)
        self.assertEqual({j["division"] for j in jobs}, {"MAG - ITAR"})

    def test_mag_itar_job_uploads_with_its_division(self):
        wf, coc, pl = self.wo(self.itar)
        self.enqueue("MAG - ITAR", wf, coc, None)
        send = Recorder()
        job = pu.process_job(pu.load_jobs(self.q)[0], CONFIG, send=send, qdir=self.q, **self.roots)
        self.assertEqual(job["status"], "done")
        self.assertIn(b'name="division"\r\n\r\nMAG - ITAR', send.calls[0][2])

    def test_folder_must_match_the_division(self):
        iwf, icoc, ipl = self.wo(self.itar)
        self.assertEqual(self.enqueue("Industrial", iwf, icoc, ipl)[0], "blocked")
        wf, coc, pl = self.wo(self.ind, customer="90 Other")
        self.assertEqual(self.enqueue("MAG - ITAR", wf, coc, pl)[0], "blocked")
        self.assertEqual(pu.load_jobs(self.q), [])

    def test_real_network_paths(self):
        industrial = r"\\PCC2\PCC-I-Shares\Production\Industrial\89 Anduril\4099 [4099 Rev. B]\Production Data\16691.1"
        itar = r"\\PCC2\PCC-I-Shares\Production\MAG - ITAR\12 Customer\5001 [5001 Rev. A]\Production Data\17001.1"
        self.assertIsNone(pu.block_reason("Industrial", [industrial]))
        self.assertIsNone(pu.block_reason("MAG - ITAR", [itar]))
        self.assertIsNone(pu.block_reason("MAG - ITAR", ["//pcc2/pcc-i-shares/production/mag - itar/x/y"]))
        for division, path in [("Industrial", itar),
                               ("MAG - ITAR", industrial),
                               ("Industrial", r"C:\Users\op\AppData\Local\PCC Shipping App\OfflineOutbox\x"),
                               ("Industrial", r"\\PCC2\PCC-I-Shares\Production\Industrial-Old\x"),
                               ("Industrial", industrial + r"\Export Controlled copy"),
                               ("MAG - ITAR", itar.replace("12 Customer", "12 Customer (CUI)"))]:
            with self.subTest(division=division, path=path):
                self.assertIsNotNone(pu.block_reason(division, [path]))

    def test_only_the_two_divisions(self):
        for division in ["industrial", "Industrial ", "MAG-ITAR", "", None]:
            with self.subTest(division=division):
                self.assertIsNotNone(pu.block_reason(division, []))

    def test_export_control_markings_are_blocked(self):
        wf, coc, pl = self.wo(self.itar)
        for text in ["EXPORT CONTROLLED", "USML Cat XI", "Contains technical data", "CUI//SP-EXPT",
                     "22 CFR 120.10", "International Traffic in Arms Regulations"]:
            with self.subTest(text=text):
                self.assertIsNotNone(pu.block_reason("MAG - ITAR", [wf], [text], **self.roots))
        for ok in ["ITAR", "MAG - ITAR", "Guitar pedal board", "Cuisinart", "PL-2207"]:
            with self.subTest(ok=ok):
                self.assertIsNone(pu.block_reason("MAG - ITAR", [wf], [ok], **self.roots))

    def test_edited_queue_file_cannot_upload_from_the_wrong_place(self):
        """A queued job edited to point outside its division's folder: the worker refuses and never calls the network."""
        wf, coc, pl = self.wo(self.ind)
        self.enqueue("Industrial", wf, coc, pl)
        iwf, icoc, ipl = self.wo(self.itar)
        send = Recorder()
        for job in pu.load_jobs(self.q):
            job["pdf_path"] = str(icoc)
            job = pu.process_job(job, CONFIG, send=send, qdir=self.q, **self.roots)
            self.assertEqual(job["status"], "blocked")
        self.assertEqual(send.calls, [])

    def test_edited_marking_cannot_upload(self):
        wf, coc, pl = self.wo(self.ind)
        self.enqueue("Industrial", wf, coc, pl)
        send = Recorder()
        for job in pu.load_jobs(self.q):
            job["meta"]["po_number"] = "EXPORT CONTROLLED 55"
            self.assertEqual(pu.process_job(job, CONFIG, send=send, qdir=self.q, **self.roots)["status"], "blocked")
        self.assertEqual(send.calls, [])

    def test_upload_job_itself_refuses(self):
        wf, coc, pl = self.wo(self.ind)
        job = {"division": "Industrial", "wo_folder": str(wf), "pdf_path": str(coc),
               "doc_type": "CoC", "meta": {"po_number": "USML 4"}}
        send = Recorder()
        with self.assertRaises(pu.UploadBlocked):
            pu.upload_job(job, CONFIG, send=send, **self.roots)
        self.assertEqual(send.calls, [])


class QueueTests(Env):
    def test_queues_coc_packing_list_and_serial_list(self):
        wf, coc, pl = self.wo(self.ind)
        sl = wf / "4099-11-5-26 serials.pdf"
        sl.write_bytes(b"%PDF-1.7 test")
        status, msg = self.enqueue("Industrial", wf, coc, pl, serials_pdf=sl)
        self.assertEqual(status, "queued")
        self.assertEqual(msg, "Sending to the portal: CoC, Packing List, Serial list.")
        jobs = sorted(pu.load_jobs(self.q), key=lambda j: j["doc_type"])
        self.assertEqual([j["doc_type"] for j in jobs], ["CoC", "Packing slip", "Serial list"])
        self.assertEqual(jobs[0]["meta"], {"work_order": "16691.1", "engineering_number": "4099", "po_number": "2859984-1",
                                           "packing_list_no": "PL-2207", "ship_date": "11/5/2026", "serials": "SN1, SN2"})
        self.assertTrue(jobs[0]["pdf_path"].endswith("4099-11-5-26-COC.pdf"))

    def test_off_until_set_up(self):
        wf, coc, pl = self.wo(self.ind)
        for config in [{**CONFIG, "enabled": False}, {**CONFIG, "key": "short"}, {**CONFIG, "url": "http://x"}]:
            with self.subTest(config=config):
                self.assertEqual(self.enqueue("Industrial", wf, coc, pl, config=config)[0], "off")
        self.assertEqual(pu.load_jobs(self.q), [])

    def test_never_raises(self):
        status, msg = pu.enqueue_shipment("Industrial", None, None, None, coc_pdf=object(), config=CONFIG, qdir=self.q, **self.roots)
        self.assertIn(status, ("blocked", "error"))


class UploadTests(Env):
    def run_one(self, send, now=None):
        wf, coc, pl = self.wo(self.ind)
        self.enqueue("Industrial", wf, coc, None)
        job = pu.load_jobs(self.q)[0]
        return pu.process_job(job, CONFIG, send=send, now=now, qdir=self.q, **self.roots)

    def test_success_sends_pdf_with_upload_key_and_fields(self):
        send = Recorder()
        job = self.run_one(send)
        self.assertEqual(job["status"], "done")
        self.assertEqual(job["document_id"], "doc-1")
        url, headers, body = send.calls[0]
        self.assertEqual(url, CONFIG["url"])
        self.assertEqual(headers["x-shipping-key"], KEY)
        self.assertNotIn("x-station-id", headers)
        self.assertNotIn(b'name="station"', body)
        for part in [b'name="division"\r\n\r\nIndustrial', b'name="doc_type"\r\n\r\nCoC', b'name="work_order"\r\n\r\n16691.1',
                     b'name="po_number"\r\n\r\n2859984-1', b'filename="4099-11-5-26-COC.pdf"', b"%PDF-1.7 test"]:
            self.assertIn(part, body)
        saved = json.loads((self.q / f"{job['id']}.json").read_text())
        self.assertEqual(saved["status"], "done")

    def test_portal_down_retries_later(self):
        now = time.time()
        job = self.run_one(Recorder(status=0, body={"error": "URLError"}), now=now)
        self.assertEqual(job["status"], "pending")
        self.assertEqual(job["next_try_at"], now + 60)

    def test_work_order_not_in_portal_is_quiet_then_skipped(self):
        """Not every work order is on the portal: retry quietly for a day, then skip — never a failure."""
        answer = Recorder(status=404, body={"ok": False, "code": "work_order_not_in_portal", "retry": True})
        job = self.run_one(answer)
        self.assertEqual(job["status"], "pending")
        self.assertEqual(pu.summary(pu.load_jobs(self.q)), {"waiting": 0, "failed": 0, "sent_today": 0})
        job["created_at"] -= 25 * 3600
        job = pu.process_job(job, CONFIG, send=answer, qdir=self.q, **self.roots)
        self.assertEqual(job["status"], "skipped")
        self.assertIn("Not in the Customer Portal", job["last_error"])
        self.assertEqual(pu.summary(pu.load_jobs(self.q)), {"waiting": 0, "failed": 0, "sent_today": 0})

    def test_other_waiting_jobs_still_count(self):
        self.run_one(Recorder(status=503, body={}))
        self.assertEqual(pu.summary(pu.load_jobs(self.q))["waiting"], 1)

    def test_refusal_is_marked_failed(self):
        job = self.run_one(Recorder(status=409, body={"ok": False, "code": "po_mismatch", "error": "PO doesn't match", "retry": False}))
        self.assertEqual(job["status"], "failed")
        self.assertEqual(job["last_error"], "PO doesn't match")

    def test_gives_up_after_seven_days(self):
        wf, coc, pl = self.wo(self.ind)
        self.enqueue("Industrial", wf, coc, None)
        job = pu.load_jobs(self.q)[0]
        job["created_at"] -= 8 * 86400
        job = pu.process_job(job, CONFIG, send=Recorder(status=503, body={}), qdir=self.q, **self.roots)
        self.assertEqual(job["status"], "failed")
        self.assertIn("7 days", job["last_error"])

    def test_missing_pdf_fails_without_sending(self):
        wf, coc, pl = self.wo(self.ind)
        self.enqueue("Industrial", wf, coc, None)
        coc.unlink()
        send = Recorder()
        job = pu.process_job(pu.load_jobs(self.q)[0], CONFIG, send=send, qdir=self.q, **self.roots)
        self.assertEqual(job["status"], "failed")
        self.assertEqual(send.calls, [])

    def test_upload_error_retries(self):
        def boom(url, headers, body):
            raise OSError("network down")
        job = self.run_one(boom)
        self.assertEqual(job["status"], "pending")

    def test_classify(self):
        self.assertEqual(pu.classify(200, {"ok": True}), "done")
        self.assertEqual(pu.classify(200, {"ok": False}), "failed")
        self.assertEqual(pu.classify(503, {}), "retry")
        self.assertEqual(pu.classify(429, {}), "retry")
        self.assertEqual(pu.classify(401, {"ok": False}), "failed")
        self.assertEqual(pu.classify(403, {"ok": False, "code": "itar_division"}), "failed")


if __name__ == "__main__":
    unittest.main()
