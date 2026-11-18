import json, sys, tempfile, unittest
from datetime import date
from pathlib import Path


from claude_token_report import cli as ct
from claude_token_report import pricing


def rec(mid, model="claude-sonnet-4-5", out=100, ts="2026-11-16T02:00:00Z", **kw):
    u = {"input_tokens": 10, "output_tokens": out, "cache_read_input_tokens": 1000,
         "cache_creation_input_tokens": 500,
         "cache_creation": {"ephemeral_5m_input_tokens": 500, "ephemeral_1h_input_tokens": 0}}
    u.update(kw)
    return json.dumps({"type": "assistant", "timestamp": ts, "requestId": "r" + mid,
                       "message": {"id": mid, "model": model, "usage": u}})


class T(unittest.TestCase):
    def test_fmt_man(self):
        self.assertEqual(ct.fmt_man(120000), "12万")
        self.assertEqual(ct.fmt_man(18000), "1.8万")
        self.assertEqual(ct.fmt_man(1028000), "102.8万")
        self.assertEqual(ct.fmt_man(9999), "9,999")
        self.assertEqual(ct.fmt_sen(18000), "18千")
        self.assertEqual(ct.fmt_sen(2_688_000), "2,688千")
        self.assertEqual(ct.fmt_sen(999), "999")
        self.assertEqual(ct.UNITS["raw"](1234567), "1,234,567")

    def test_dedupe_and_skip(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "a.jsonl").write_text("\n".join([
                rec("m1", out=5), rec("m1", out=100),  # 重複: 最後を採用
                "broken{", rec("m2", model="<synthetic>"),
                json.dumps({"type": "user"}),
            ]))
            recs = list(ct.iter_records(d))
            self.assertEqual(len(recs), 1)
            self.assertEqual(recs[0][3], 100)

    def test_cost(self):
        # sonnet-4-5: in3 out15 w5m3.75 read0.3
        c = pricing.cost_usd("claude-sonnet-4-5-20250929", 1_000_000, 1_000_000, 1_000_000, 1_000_000, 0)
        self.assertAlmostEqual(c, 3 + 15 + 0.3 + 3.75)
        self.assertIsNone(pricing.cost_usd("unknown-model", 1, 1, 1, 1, 1))

    def test_prefix_specificity(self):
        self.assertEqual(pricing.lookup("claude-opus-5-5")[0], 4)
        self.assertEqual(pricing.lookup("claude-opus-5")[0], 5)
        self.assertEqual(pricing.lookup("claude-opus-4-1-20250805")[0], 15)

    def test_haiku_threshold(self):
        self.assertEqual(pricing.lookup("claude-haiku-5-5", 100_000)[0], 0.10)
        self.assertEqual(pricing.lookup("claude-haiku-5-5", 100_001)[0], 0.50)

    def test_aggregate_range(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "a.jsonl").write_text("\n".join([
                rec("m1"), rec("m2", ts="2026-11-17T02:00:00Z")]))
            rows, unk = ct.aggregate(ct.iter_records(d), since=date(2026, 11, 17))
            self.assertEqual(len(rows), 1)
            self.assertFalse(unk)

    def test_excel(self):
        import zipfile, re
        import xml.dom.minidom as m
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "a.jsonl").write_text("\n".join([
                rec("m1"), rec("m2", model="claude-haiku-5-5", ts="2026-11-17T02:00:00Z"),
                rec("m3", model="unknown-x")]))
            detail, unk = ct.aggregate_detail(ct.iter_records(d))
            self.assertEqual(unk, {"unknown-x"})
            out = Path(d) / "o.xlsx"
            ct.write_excel(out, detail, 150.0)
            z = zipfile.ZipFile(out)
            for n in z.namelist():
                m.parseString(z.read(n))
            meta = z.read("xl/worksheets/sheet2.xml").decode()
            self.assertIn("VLOOKUP($B2,単価表!", meta)
            self.assertIn("<f>(C2*H2+F2*I2+G2*J2+E2*K2+D2*L2)/1000000</f>", meta)
            # sonnet-4-5: 10*3 + 500*3.75 + 1000*0.3 + 100*15 = 3705 / 1e6
            self.assertIn("<v>0.003705</v>", meta)


if __name__ == "__main__":
    unittest.main()
