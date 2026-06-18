"""Tests for blockquote, hr, pagebreak, internal links, theme fidelity, CLI."""

import base64
import io
import unittest
import zipfile
from contextlib import redirect_stderr
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document

from mark2word import build, parse_blocks
from mark2word.slug import slugify
from mark2word.cli import EXIT_FAILURE, EXIT_SUCCESS, main
from mark2word.theme import split_frontmatter

MINIMAL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


class MarkdownFeatureTests(unittest.TestCase):
    def test_blockquote_parses(self):
        blocks = parse_blocks("> quoted text\n> second line\n")
        self.assertEqual(blocks[0]["type"], "blockquote")
        self.assertIn("quoted text", blocks[0]["text"])

    def test_horizontal_rule_in_body(self):
        front, body = split_frontmatter("---\ntitle: Doc\n---\n---\n")
        self.assertEqual(front["title"], "Doc")
        blocks = parse_blocks(body)
        self.assertEqual(blocks[0]["type"], "hr")

    def test_pagebreak_comment(self):
        blocks = parse_blocks("Before\n<!-- pagebreak -->\nAfter\n")
        self.assertEqual(blocks[1]["type"], "pagebreak")

    def test_blockquote_renders_in_docx(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "bq.docx"
            build(
                glob={"blockquote": {"indent_left": "18pt", "italic": True}},
                external={},
                md_body="> A quoted line\n",
                out_path=out,
            )
            doc = Document(out)
            self.assertIn("quoted", doc.paragraphs[0].text.lower())

    def test_hr_renders_border(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "hr.docx"
            build(glob={}, external={}, md_body="---\n", out_path=out)
            with zipfile.ZipFile(out) as z:
                xml = z.read("word/document.xml")
            self.assertIn(b"w:pBdr", xml)

    def test_internal_link_to_heading(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "links.docx"
            build(
                glob={},
                external={},
                md_body="# My Section\n\nSee [link](#my-section).\n",
                out_path=out,
            )
            with zipfile.ZipFile(out) as z:
                doc_xml = z.read("word/document.xml").decode("utf-8")
            self.assertIn('w:anchor="my-section"', doc_xml)

    def test_unknown_internal_link_raises(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "bad.docx"
            with self.assertRaises(Exception):
                build(
                    glob={},
                    external={},
                    md_body="[nope](#missing)\n",
                    out_path=out,
                )

    def test_bold_hyperlink(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "bold-link.docx"
            build(
                glob={},
                external={},
                md_body="[**bold link**](https://example.com)\n",
                out_path=out,
            )
            with zipfile.ZipFile(out) as z:
                xml = z.read("word/document.xml")
            self.assertIn(b"w:hyperlink", xml)
            self.assertIn(b"w:b", xml)

    def test_indent_right(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "ind.docx"
            build(
                glob={"body": {"indent_right": "24pt"}},
                external={},
                md_body="Indented\n",
                out_path=out,
            )
            doc = Document(out)
            self.assertEqual(doc.paragraphs[0].paragraph_format.right_indent.pt, 24)

    def test_table_spacing(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "tspace.docx"
            build(
                glob={"table": {"space_before": 8, "space_after": 6}},
                external={},
                md_body="Above\n| A | B |\n| - | - |\n| 1 | 2 |\nBelow\n",
                out_path=out,
            )
            doc = Document(out)
            table = doc.tables[0]
            self.assertEqual(
                table.rows[0].cells[0].paragraphs[0].paragraph_format.space_before.pt, 8
            )

    def test_image_width_and_caption(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            img = root / "pic.png"
            img.write_bytes(MINIMAL_PNG)
            out = root / "img.docx"
            build(
                glob={"image": {"width": "50pt", "alt_mode": "caption", "align": "center"}},
                external={},
                md_body=f"![Swatch label]({img.name})\n",
                out_path=out,
                md_path=root / "doc.md",
            )
            doc = Document(out)
            self.assertEqual(len(doc.inline_shapes), 1)
            self.assertTrue(any("Swatch" in p.text for p in doc.paragraphs))

    def test_code_lang_theme_override(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "code.docx"
            build(
                glob={"code": {"font": "Consolas", "langs": {"python": {"color": "0000FF"}}}},
                external={},
                md_body="```python\nx = 1\n```\n",
                out_path=out,
            )
            doc = Document(out)
            run = doc.paragraphs[0].runs[0]
            self.assertEqual(str(run.font.color.rgb), "0000FF")

    def test_page_header_footer_from_theme(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "chrome.docx"
            build(
                glob={
                    "title": "My Doc",
                    "page": {"footer": "My Doc || Page {page}"},
                },
                external={},
                md_body="# Title\n",
                out_path=out,
            )
            with zipfile.ZipFile(out) as z:
                footer = z.read("word/footer1.xml").decode("utf-8")
            self.assertIn("My Doc", footer)
            self.assertIn("fldChar", footer)

    def test_slugify(self):
        self.assertEqual(slugify("Hello World!"), "hello-world")


class CliFeatureTests(unittest.TestCase):
    def test_version_flag(self):
        with self.assertRaises(SystemExit) as ctx:
            main(["--version"])
        self.assertEqual(ctx.exception.code, 0)

    def test_check_theme_ok(self):
        theme = Path("skills/mark2word/assets/base-theme.yaml")
        if not theme.exists():
            self.skipTest("skill assets not present")
        code = main(["--check-theme", str(theme), "--theme-dir", "skills/mark2word/assets"])
        self.assertEqual(code, EXIT_SUCCESS)

    def test_check_theme_bad_page_size(self):
        with TemporaryDirectory() as tmp:
            bad = Path(tmp) / "bad.yaml"
            bad.write_text("page:\n  size: legal\n", encoding="utf-8")
            stderr = io.StringIO()
            with redirect_stderr(stderr):
                code = main(["--check-theme", str(bad)])
            self.assertEqual(code, EXIT_FAILURE)
            self.assertIn("unknown page size", stderr.getvalue())


if __name__ == "__main__":
    unittest.main()
