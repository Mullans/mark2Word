"""CLI output path resolution tests."""

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from mark2word.cli import main, resolve_output_path
from mark2word.errors import Mark2WordError


class OutputPathTests(unittest.TestCase):
    def test_single_input_default_output(self):
        md = Path("report.md")
        self.assertEqual(
            resolve_output_path(md, None, multiple_inputs=False),
            Path("report.docx"),
        )

    def test_single_input_explicit_file(self):
        md = Path("report.md")
        out = Path("build/final.docx")
        self.assertEqual(
            resolve_output_path(md, out, multiple_inputs=False),
            out,
        )

    def test_single_input_output_directory(self):
        with TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "out"
            result = resolve_output_path(
                Path("report.md"), out_dir, multiple_inputs=False
            )
            self.assertEqual(result, out_dir / "report.docx")

    def test_multiple_inputs_require_directory(self):
        with self.assertRaises(Mark2WordError):
            resolve_output_path(
                Path("a.md"), Path("only.docx"), multiple_inputs=True
            )

    def test_multiple_inputs_to_directory(self):
        with TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "out"
            result = resolve_output_path(Path("a.md"), out_dir, multiple_inputs=True)
            self.assertEqual(result, out_dir / "a.docx")


class CliIntegrationTests(unittest.TestCase):
    def test_i_and_o_flags(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            md = root / "sample.md"
            out = root / "custom.docx"
            md.write_text("# Hello\n", encoding="utf-8")
            main(["-i", str(md), "-o", str(out)])
            self.assertTrue(out.exists())

    def test_multiple_inputs_output_directory(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            one = root / "one.md"
            two = root / "two.md"
            out_dir = root / "docx"
            one.write_text("# One\n", encoding="utf-8")
            two.write_text("# Two\n", encoding="utf-8")
            main(["-i", str(one), "-i", str(two), "-o", str(out_dir)])
            self.assertTrue((out_dir / "one.docx").exists())
            self.assertTrue((out_dir / "two.docx").exists())


if __name__ == "__main__":
    unittest.main()
