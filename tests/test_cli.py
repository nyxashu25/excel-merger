import contextlib
import io
import unittest

from excel_merger.cli import main
from helpers import TempDirCase, make_csv, make_xlsx, read_xlsx


def run(argv):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        try:
            code = main(argv)
        except SystemExit as exc:  # argparse exits on usage errors
            code = exc.code
    return code, out.getvalue(), err.getvalue()


class CliTests(TempDirCase):
    def test_success_prints_a_summary(self):
        a = make_xlsx(self.path("a.xlsx"), {"S": [["v"], [1], [2]]})
        b = make_csv(self.path("b.csv"), [["v"], ["3"]])
        out = self.path("out.xlsx")
        code, stdout, stderr = run([str(a), str(b), "-o", str(out)])
        self.assertEqual(code, 0, stderr)
        self.assertIn("a.xlsx [S]  2 rows", stdout)
        self.assertIn("Merged 2 file(s), 2 sheet(s): 3 rows, 3 columns", stdout)
        self.assertEqual(read_xlsx(out)["Merged"][0], ("v", "Source file", "Source sheet"))

    def test_singular_nouns(self):
        a = make_xlsx(self.path("a.xlsx"), {"S": [["v"], [1]]})
        _, stdout, _ = run([str(a), "-o", str(self.path("o.xlsx")), "--no-source"])
        self.assertIn("a.xlsx [S]  1 row\n", stdout)
        self.assertIn(": 1 row, 1 column", stdout)

    def test_quiet_suppresses_output(self):
        a = make_xlsx(self.path("a.xlsx"), {"S": [["v"], [1]]})
        code, stdout, _ = run([str(a), "-o", str(self.path("o.xlsx")), "-q"])
        self.assertEqual((code, stdout), (0, ""))

    def test_errors_go_to_stderr_with_exit_code_1(self):
        code, stdout, stderr = run([str(self.path("nope.xlsx")), "-o", str(self.path("o.xlsx"))])
        self.assertEqual(code, 1)
        self.assertEqual(stdout, "")
        self.assertIn("error: Not found", stderr)

    def test_flags_are_passed_through(self):
        a = make_xlsx(self.path("a.xlsx"), {"S": [["Email", "n"], ["x", 1], ["x", 1]], "Other": [["Email", "n"], ["y", 2]]})
        out = self.path("o.xlsx")
        code, stdout, _ = run([str(a), "-o", str(out), "--dedupe", "--no-source", "--sheet", "S"])
        self.assertEqual(code, 0)
        self.assertIn("1 duplicates removed", stdout)
        self.assertEqual(read_xlsx(out)["Merged"], [("Email", "n"), ("x", 1)])

    def test_usage_errors_use_exit_code_2(self):
        code, _, stderr = run(["-o", "x.xlsx"])  # no inputs
        self.assertEqual(code, 2)
        a = make_xlsx(self.path("a.xlsx"), {"S": [["v"], [1]]})
        code, _, stderr = run([str(a), "-o", str(self.path("o.csv")), "--delimiter", ";;"])
        self.assertEqual(code, 2)
        self.assertIn("single character", stderr)

    def test_warnings_are_reported_but_do_not_fail(self):
        a = make_xlsx(self.path("a.xlsx"), {"Data": [["v"], [1]]})
        b = make_xlsx(self.path("b.xlsx"), {"Other": [["v"], [2]]})
        code, _, stderr = run([str(a), str(b), "-o", str(self.path("o.xlsx")), "--sheet", "Data"])
        self.assertEqual(code, 0)
        self.assertIn("warning:", stderr)
        self.assertIn("b.xlsx", stderr)


if __name__ == "__main__":
    unittest.main()
