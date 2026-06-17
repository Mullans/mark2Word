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
    def test_module_file_runs_as_direct_script(self):
        script = Path(__file__).parents[1] / "src" / "mark2word" / "__init__.py"

        result = subprocess.run(
            [sys.executable, str(script), "--help"],
            capture_output=True,
            check=False,
            text=True,
        )

        self.assertEqual(result.returncode, 0)
        self.assertIn("Convert a styled-Markdown dialect", result.stdout)


if __name__ == "__main__":
    unittest.main()
