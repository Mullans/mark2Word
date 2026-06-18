"""Parser validation, theme errors, and CLI error handling."""

import io
import unittest
from contextlib import redirect_stderr
from pathlib import Path
from tempfile import TemporaryDirectory

from mark2word import (
    FrontmatterError,
    ParseError,
    ThemeError,
    build,
    load_theme,
    parse_blocks,
    register_block_parser,
    reset_plugins,
    split_frontmatter,
)
from mark2word.cli import EXIT_FAILURE, main


class TableValidationTests(unittest.TestCase):
    def test_ragged_table_row_raises_parse_error(self):
        with self.assertRaises(ParseError) as ctx:
            parse_blocks("| a | b |\n| - | - |\n| only |\n")
        self.assertIn("expected 2", str(ctx.exception))
        self.assertEqual(ctx.exception.line_no, 3)

    def test_extra_columns_raises_parse_error(self):
        with self.assertRaises(ParseError) as ctx:
            parse_blocks("| a | b |\n| - | - |\n| x | y | z |\n")
        self.assertEqual(ctx.exception.line_no, 3)


class ThemeErrorTests(unittest.TestCase):
    def test_missing_extends_raises_theme_error(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            md = root / "doc.md"
            md.write_text("---\nextends: missing.yaml\n---\n", encoding="utf-8")
            with self.assertRaises(ThemeError) as ctx:
                load_theme({"extends": "missing.yaml"}, md)
            self.assertIn("missing theme", str(ctx.exception))

    def test_unknown_page_size_raises_theme_error(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "bad-page.docx"
            with self.assertRaises(ThemeError) as ctx:
                build(
                    glob={"page": {"size": "legal"}},
                    external={},
                    md_body="Body\n",
                    out_path=out,
                )
            self.assertIn("unknown page size", str(ctx.exception))


class FrontmatterTests(unittest.TestCase):
    def test_invalid_yaml_raises_frontmatter_error(self):
        with self.assertRaises(FrontmatterError) as ctx:
            split_frontmatter("---\nfont: [\n---\nBody\n")
        self.assertIn("invalid YAML", str(ctx.exception))


class PluginRegistryTests(unittest.TestCase):
    def test_reset_plugins_clears_registrations(self):
        def parse_banner(stripped, line_no, ctx):
            if stripped.startswith("!!!"):
                return {"type": "body", "text": stripped[3:].strip(), "region": list(ctx["region_stack"])}
            return None

        register_block_parser(parse_banner, prepend=True)
        self.addCleanup(reset_plugins)
        self.assertEqual(parse_blocks("!!! hi\n")[0]["text"], "hi")
        reset_plugins()
        self.assertEqual(parse_blocks("!!! hi\n")[0]["text"], "!!! hi")


class CliErrorTests(unittest.TestCase):
    def _run_cli(self, args: list[str]) -> tuple[int, str]:
        stderr = io.StringIO()
        with redirect_stderr(stderr):
            code = main(args)
        return code, stderr.getvalue()

    def test_unclosed_region_exits_failure(self):
        with TemporaryDirectory() as tmp:
            md = Path(tmp) / "bad.md"
            md.write_text("<!-- region: x -->\n# Hi\n", encoding="utf-8")
            code, err = self._run_cli(["-i", str(md)])
            self.assertEqual(code, EXIT_FAILURE)
            self.assertIn("unclosed region", err)

    def test_unclosed_code_fence_exits_failure(self):
        with TemporaryDirectory() as tmp:
            md = Path(tmp) / "bad.md"
            md.write_text("```\nno close\n", encoding="utf-8")
            code, err = self._run_cli(["-i", str(md)])
            self.assertEqual(code, EXIT_FAILURE)
            self.assertIn("unclosed fenced code block", err)

    def test_ragged_table_exits_failure(self):
        with TemporaryDirectory() as tmp:
            md = Path(tmp) / "bad.md"
            md.write_text("| a | b |\n| - | - |\n| x |\n", encoding="utf-8")
            code, err = self._run_cli(["-i", str(md)])
            self.assertEqual(code, EXIT_FAILURE)
            self.assertIn("columns", err)

    def test_missing_image_exits_failure(self):
        with TemporaryDirectory() as tmp:
            md = Path(tmp) / "bad.md"
            md.write_text("![x](missing.png)\n", encoding="utf-8")
            code, err = self._run_cli(["-i", str(md)])
            self.assertEqual(code, EXIT_FAILURE)
            self.assertIn("image not found", err)

    def test_invalid_frontmatter_exits_failure(self):
        with TemporaryDirectory() as tmp:
            md = Path(tmp) / "bad.md"
            md.write_text("---\nfont: [\n---\n# Hi\n", encoding="utf-8")
            code, err = self._run_cli(["-i", str(md)])
            self.assertEqual(code, EXIT_FAILURE)
            self.assertIn("invalid YAML", err)


if __name__ == "__main__":
    unittest.main()
