"""The demo-kit reset program stays within the 5-DML limit and clears Decided By / Decided At."""
import csv, subprocess, sys, tempfile, unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[3] / "docs" / "demo-kit" / "reset-demo.py"


class ResetDemoTests(unittest.TestCase):
    def emit(self):
        with tempfile.TemporaryDirectory() as d:
            seed = Path(d) / "rate_seed.csv"
            with open(seed, "w", newline="") as f:
                w = csv.writer(f)
                w.writerow(["LINE_ITEM", "RATE"])
                w.writerow(["North America Revenue", "0.01"])
            r = subprocess.run([sys.executable, str(SCRIPT), "--workbook", "wb", "--rate-seed", str(seed)],
                               capture_output=True, text=True)
        self.assertEqual(r.returncode, 0, r.stderr)
        return r.stdout

    def test_program_compiles_and_has_at_most_five_dml(self):
        prog = self.emit()
        compile(prog, "reset-program", "exec")
        self.assertLessEqual(prog.count("sigma.data.sql.execute("), 5)

    def test_clears_decided_columns_on_migrated_rows(self):
        prog = self.emit()
        self.assertIn("get_column('ap_dby')", prog)
        self.assertIn("get_column('ap_dat')", prog)
        self.assertIn('SET "{a_dby}" = NULL, "{a_dat}" = NULL WHERE "{a_req}" IN (', prog)

    def test_still_deletes_non_migrated_approvals(self):
        self.assertIn('DELETE FROM {a_src.ref()} WHERE "{a_req}" IS NULL OR "{a_req}" NOT IN (', self.emit())


if __name__ == "__main__":
    unittest.main()
