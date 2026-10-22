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

    def _balanced(self):
        s = core.new_state()
        s["meta"]["au"] = "31/12/2025"
        s["report"]["compta"] = {"montant": 1000, "sens": "Débit"}
        s["report"]["banque"] = {"montant": 1000, "sens": "Crédit"}
        return s

    def test_validation_et_mot_de_passe(self):
        s = self._balanced()
        self.assertTrue(core.validate_check(s)[0])
        s["bank"].append({"date": "01/12/2025", "lib": "frais", "debit": 0, "credit": 50})
        self.assertFalse(core.validate_check(s)[0])          # déséquilibré -> refus
        s["bank"] = []
        core.set_password(s, "secret"); core.lock_state(s)
        self.assertTrue(s["lock"]); self.assertTrue(core.check_password(s, "secret")); self.assertFalse(core.check_password(s, "x"))
        core.unlock_state(s)
        self.assertIsNone(s["lock"]); self.assertEqual(len(s["journal"]), 2)

    def test_sauvegarde_verrou_et_nouvelle_periode(self):
        import tempfile
        s = self._balanced()
        s["gl"] = [{"date": "01/12/2025", "lib": "A", "debit": 5, "credit": 0, "pointe": True},
                   {"date": "02/12/2025", "lib": "B", "debit": 0, "credit": 7, "pointe": False}]
        core.set_password(s, "pw"); core.lock_state(s)
        p = os.path.join(tempfile.mkdtemp(), "x.rapp"); core.save_project(s, p)
        t = core.load_project(p)
        self.assertTrue(t["lock"]); self.assertTrue(core.check_password(t, "pw"))
        n = core.next_period(t)
        self.assertEqual([g["lib"] for g in n["gl"]], ["B"]); self.assertIsNone(n["lock"]); self.assertTrue(core.check_password(n, "pw"))

    def test_modele_et_impression(self):
        import tempfile
        p = os.path.join(tempfile.mkdtemp(), "m.xlsx"); core.export_template(p)
        self.assertEqual(core.import_grand_livre(p), [])      # modèle vide : rien à importer
        from openpyxl import load_workbook
        wb = load_workbook(p); ws = wb["Grand livre"]
        ws.append(["03/12/2025", "CHQ1", "REGLEMENT <FOURNISSEUR>", None, 850000]); wb.save(p)
        r = core.import_grand_livre(p)
        self.assertEqual((r[0]["date"], r[0]["credit"]), ("03/12/2025", 850000))
        s = core.new_state(); s["gl"] = r
        h = core.build_print_html(s)
        self.assertIn("850", h); self.assertIn("&lt;FOURNISSEUR&gt;", h); self.assertIn("non validé", h)

    def test_import_cumulatif_et_modification(self):
        s = core.new_state()
        jan = [{"date": "05/01/2026", "lib": "CHQ A", "debit": 0, "credit": 100, "pointe": False},
               {"date": "06/01/2026", "lib": "CHQ B", "debit": 0, "credit": 50, "pointe": False},
               {"date": "06/01/2026", "lib": "CHQ B", "debit": 0, "credit": 50, "pointe": False}]   # 2 lignes identiques légitimes
        add, sk = core.merge_import(s["gl"], jan); s["gl"] += add
        self.assertEqual((len(add), sk), (3, 0))
        s["gl"][0]["pointe"] = True
        fev = [{"date": "03/02/2026", "lib": "REMISE C", "debit": 70, "credit": 0, "pointe": False}]
        add, sk = core.merge_import(s["gl"], fev); s["gl"] += add
        self.assertEqual(len(s["gl"]), 4); self.assertTrue(s["gl"][0]["pointe"])          # rien écrasé, pointage conservé
        add, sk = core.merge_import(s["gl"], jan + fev)                                   # ré-import : tout est déjà là
        self.assertEqual((len(add), sk), (0, 4))
        self.assertIsNone(core.edit_gl_line(s, 1, "07/01/2026", "CHQ B corrigé", 0, 55))
        self.assertEqual((s["gl"][1]["credit"], s["gl"][1]["date"]), (55, "07/01/2026")); self.assertIn("Écriture modifiée", s["journal"][-1])
        self.assertIn("invalide", core.edit_gl_line(s, 1, "xx", "a", 0, 5))
        self.assertIn("débit OU crédit", core.edit_gl_line(s, 1, "07/01/2026", "a", 5, 5))


if __name__ == "__main__":
    unittest.main()
