import unittest
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

from mark2word import load_theme


class ThemeLoadingTests(unittest.TestCase):
    def test_load_theme_searches_explicit_theme_directory(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            doc_dir = root / "documents"
            theme_dir = root / "themes"
            doc_dir.mkdir()
            theme_dir.mkdir()

            md_path = doc_dir / "resume.md"
            md_path.write_text("---\nextends: base-theme.yaml\n---\n# Resume\n", encoding="utf-8")
            (theme_dir / "base-theme.yaml").write_text(
                'font: Aptos\nbody: { space_after: "3pt" }\n',
                encoding="utf-8",
            )

            external, glob = load_theme(
                {"extends": "base-theme.yaml", "size": 12},
                md_path,
                theme_dirs=[theme_dir],
            )

            self.assertEqual(external["font"], "Aptos")
            self.assertEqual(external["body"]["space_after"], "3pt")
            self.assertEqual(glob, {"size": 12})


class CliTests(unittest.TestCase):
    def test_mark2word_help(self):
        result = subprocess.run(
            [
                sys.executable,
                "-c",
                "import sys; sys.argv=['mark2word', '--help']; from mark2word.cli import main; main()",
            ],
            capture_output=True,
            check=False,
            text=True,
        )

        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("Convert a styled-Markdown dialect", result.stdout)


if __name__ == "__main__":
    unittest.main()
