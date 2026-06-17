"""Tests for list format parsing."""

import unittest

from mark2word.errors import ThemeError
from mark2word.list_format import parse_format_string, resolve_level_numbering


class ListFormatTests(unittest.TestCase):
    def test_bare_keyword_uses_trailing_period(self):
        self.assertEqual(
            parse_format_string("1", default_num_fmt="decimal", default_lvl_text="%1."),
            ("decimal", "%1."),
        )

    def test_template_with_dot(self):
        self.assertEqual(
            parse_format_string("1.", default_num_fmt="decimal", default_lvl_text="%1"),
            ("decimal", "%1."),
        )

    def test_lower_letter_template(self):
        self.assertEqual(
            parse_format_string("a.", default_num_fmt="decimal", default_lvl_text="%1"),
            ("lowerLetter", "%1."),
        )

    def test_roman_template_with_suffix(self):
        self.assertEqual(
            parse_format_string("roman )", default_num_fmt="decimal", default_lvl_text="%1"),
            ("lowerRoman", "%1 )"),
        )

    def test_section_prefix(self):
        self.assertEqual(
            parse_format_string("Section 1:", default_num_fmt="decimal", default_lvl_text="%1"),
            ("decimal", "Section %1:"),
        )

    def test_explicit_num_fmt_and_template(self):
        self.assertEqual(
            resolve_level_numbering({"num_fmt": "upperRoman", "template": "(%1)"}, ordered=True, ilvl=0),
            ("upperRoman", "(%1)"),
        )

    def test_partial_explicit_raises(self):
        with self.assertRaises(ThemeError):
            resolve_level_numbering({"num_fmt": "decimal"}, ordered=True, ilvl=0)

    def test_word_lvl_text_maps_nested_placeholder(self):
        from mark2word.list_format import word_lvl_text

        self.assertEqual(word_lvl_text("%1.", 0), "%1.")
        self.assertEqual(word_lvl_text("%1.", 1), "%2.")
        self.assertEqual(word_lvl_text("Section %1:", 2), "Section %3:")
        self.assertEqual(
            parse_format_string("◦", default_num_fmt="bullet", default_lvl_text="•"),
            ("bullet", "◦"),
        )
        with self.assertRaises(ThemeError):
            parse_format_string("bogus", default_num_fmt="decimal", default_lvl_text="%1")


if __name__ == "__main__":
    unittest.main()
