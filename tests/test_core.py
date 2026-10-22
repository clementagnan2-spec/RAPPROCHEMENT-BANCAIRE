import os, sys, unittest
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import core


class TestCore(unittest.TestCase):
    def test_formules_equilibre(self):
        s = core.new_state()
        s["report"]["compta"] = {"montant": 1000, "sens": "Crédit"}
        s["report"]["banque"] = {"montant": 1000, "sens": "Crédit"}
        r = core.calc(core.left_rows(s), core.right_rows(s))
        self.assertEqual((r["D51"], r["E51"], r["E52"], r["D52"]), (0, 1000, 0, 1000))
        self.assertEqual((r["D53"], r["E53"]), (1000, 1000))
        self.assertEqual(r["E56"], "OK")
        self.assertEqual(r["E57"], "OK")

    def test_pointage(self):
        s = core.new_state()
        s["gl"] = [{"date": "01/01/2025", "lib": "A", "debit": 100, "credit": 0, "pointe": True},
                   {"date": "02/01/2025", "lib": "B", "debit": 0, "credit": 40, "pointe": False}]
        self.assertEqual(len(core.left_rows(s)), 2)  # report + 1 non pointée
        r = core.calc(core.left_rows(s), core.right_rows(s))
        self.assertEqual(r["E51"], 40)

    def test_nombres_dates(self):
        self.assertEqual(core.parse_num("1 234 567,50"), 1234567.5)
        self.assertEqual(core.parse_num("1,234,567"), 1234567)
        self.assertEqual(core.parse_num("(500)"), -500)
        self.assertEqual(core.parse_date("2025-12-31"), "31/12/2025")


if __name__ == "__main__":
    unittest.main()
