"""CLI output path resolution tests."""

import os
import stat
import unittest
from contextlib import contextmanager
from pathlib import Path
from tempfile import TemporaryDirectory

from mark2word.cli import EXIT_FAILURE, UsageError, main, resolve_output_path
from mark2word.errors import Mark2WordError
from mark2word.paths import ensure_output_writable


@contextmanager
def _exclusive_file_lock(path: Path):
    """Hold an exclusive lock similar to an open Word document."""
    if os.name == "nt":
        import ctypes

        GENERIC_READ = 0x80000000
        GENERIC_WRITE = 0x40000000
        OPEN_EXISTING = 3
        FILE_ATTRIBUTE_NORMAL = 0x80
        INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateFileW(
            str(path),
            GENERIC_READ | GENERIC_WRITE,
            0,
            None,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            None,
        )
        if handle in (INVALID_HANDLE_VALUE, -1):
            raise OSError(kernel32.GetLastError(), "CreateFileW", str(path))
        try:
            yield
        finally:
            kernel32.CloseHandle(handle)
    else:
        import fcntl

        with open(path, "r+b") as fh:
            fcntl.flock(fh.fileno(), fcntl.LOCK_EX)
            yield


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
        with self.assertRaises(UsageError):
            resolve_output_path(
                Path("a.md"), Path("only.docx"), multiple_inputs=True
            )

    def test_multiple_inputs_to_directory(self):
        with TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "out"
            result = resolve_output_path(Path("a.md"), out_dir, multiple_inputs=True)
            self.assertEqual(result, out_dir / "a.docx")


class InputPathTests(unittest.TestCase):
    def test_missing_markdown_fails_before_convert(self):
        with TemporaryDirectory() as tmp:
            missing = Path(tmp) / "missing.md"
            self.assertEqual(main(["-i", str(missing)]), EXIT_FAILURE)

    def test_missing_theme_fails_before_convert(self):
        with TemporaryDirectory() as tmp:
            md = Path(tmp) / "sample.md"
            md.write_text("---\nextends: missing-theme.yaml\n---\n# Body\n", encoding="utf-8")
            self.assertEqual(main(["-i", str(md)]), EXIT_FAILURE)


class OutputWritableTests(unittest.TestCase):
    def test_ensure_output_writable_new_file(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "new.docx"
            ensure_output_writable(out)
            self.assertFalse(out.exists())

    def test_locked_output_raises(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "locked.docx"
            out.write_bytes(b"PK")
            with _exclusive_file_lock(out):
                with self.assertRaises(Mark2WordError) as ctx:
                    ensure_output_writable(out)
                self.assertIn("another program", str(ctx.exception))

    def test_read_only_output_raises(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "readonly.docx"
            out.write_bytes(b"PK")
            out.chmod(stat.S_IREAD)
            try:
                with self.assertRaises(Mark2WordError) as ctx:
                    ensure_output_writable(out)
                self.assertIn("read-only", str(ctx.exception).lower())
            finally:
                out.chmod(stat.S_IWRITE | stat.S_IREAD)

    def test_locked_output_fails_before_write(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            md = root / "sample.md"
            out = root / "locked.docx"
            md.write_text("# Hello\n", encoding="utf-8")
            out.write_bytes(b"PK")
            with _exclusive_file_lock(out):
                self.assertEqual(main(["-i", str(md), "-o", str(out)]), EXIT_FAILURE)

    def test_batch_aborts_when_any_output_locked(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            one = root / "one.md"
            two = root / "two.md"
            out_dir = root / "out"
            out_dir.mkdir()
            one_out = out_dir / "one.docx"
            one.write_text("# One\n", encoding="utf-8")
            two.write_text("# Two\n", encoding="utf-8")
            one_out.write_bytes(b"unchanged")
            with _exclusive_file_lock(one_out):
                self.assertEqual(
                    main(["-i", str(one), "-i", str(two), "-o", str(out_dir)]),
                    EXIT_FAILURE,
                )
            self.assertFalse((out_dir / "two.docx").exists())
            self.assertEqual(one_out.read_bytes(), b"unchanged")


class CliIntegrationTests(unittest.TestCase):
    def test_i_and_o_flags(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            md = root / "sample.md"
            out = root / "custom.docx"
            md.write_text("# Hello\n", encoding="utf-8")
            self.assertEqual(main(["-i", str(md), "-o", str(out)]), 0)
            self.assertTrue(out.exists())

    def test_multiple_inputs_output_directory(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            one = root / "one.md"
            two = root / "two.md"
            out_dir = root / "docx"
            one.write_text("# One\n", encoding="utf-8")
            two.write_text("# Two\n", encoding="utf-8")
            self.assertEqual(main(["-i", str(one), "-i", str(two), "-o", str(out_dir)]), 0)
            self.assertTrue((out_dir / "one.docx").exists())
            self.assertTrue((out_dir / "two.docx").exists())


if __name__ == "__main__":
    unittest.main()
