"""Command-line interface for mark2word."""

from __future__ import annotations

import argparse
import warnings
from pathlib import Path

from mark2word.emit import build
from mark2word.errors import Mark2WordError
from mark2word.parser import parse_to_ast
from mark2word.theme import discover_theme_dirs, load_theme, split_frontmatter


def validate_document(md_path: Path, theme_dirs: list[Path]) -> None:
    text = md_path.read_text(encoding="utf-8")
    front, body = split_frontmatter(text)
    load_theme(front, md_path, theme_dirs=theme_dirs)
    parse_to_ast(body)


def _is_explicit_output_file(path: Path) -> bool:
    if path.exists():
        return path.is_file()
    return path.suffix.lower() == ".docx"


def resolve_output_path(
    md_path: Path,
    output: Path | None,
    *,
    multiple_inputs: bool,
) -> Path:
    """Resolve the .docx path for a single markdown input."""
    if multiple_inputs:
        if output is None:
            return md_path.with_suffix(".docx")
        if _is_explicit_output_file(output):
            raise Mark2WordError(
                f"--output must be a directory when converting multiple inputs: {output}"
            )
        output.mkdir(parents=True, exist_ok=True)
        return output / f"{md_path.stem}.docx"

    if output is None:
        return md_path.with_suffix(".docx")
    if _is_explicit_output_file(output):
        output.parent.mkdir(parents=True, exist_ok=True)
        return output
    output.mkdir(parents=True, exist_ok=True)
    return output / f"{md_path.stem}.docx"


def convert_one(
    md_path: Path,
    out_path: Path,
    theme_dirs: list[Path],
    *,
    check: bool = False,
    verbose: bool = False,
) -> None:
    if check:
        validate_document(md_path, theme_dirs)
        print(f"OK {md_path}")
        return
    text = md_path.read_text(encoding="utf-8")
    front, body = split_frontmatter(text)
    external, glob = load_theme(front, md_path, theme_dirs=theme_dirs)
    build(glob, external, body, out_path, verbose=verbose, md_path=md_path)


def _collect_inputs(args: argparse.Namespace) -> list[Path]:
    inputs = list(args.inputs or [])
    inputs.extend(args.positional_inputs or [])
    if not inputs:
        raise Mark2WordError("at least one input file is required (use -i/--input)")
    return inputs


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description="Convert a styled-Markdown dialect into a Word document."
    )
    parser.add_argument(
        "-i",
        "--input",
        dest="inputs",
        action="append",
        type=Path,
        help="Markdown input file. Repeat for multiple files.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help=(
            "Output .docx file or directory. With one input, may be an explicit "
            ".docx path or a directory (uses the markdown basename). With multiple "
            "inputs, must be a directory."
        ),
    )
    parser.add_argument(
        "--theme-dir",
        action="append",
        default=[],
        type=Path,
        help="Additional folder to search for relative frontmatter extends paths.",
    )
    parser.add_argument(
        "--no-auto-theme-dir",
        action="store_true",
        help="Do not search .mark2word/themes near the input file.",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Validate frontmatter, theme, and markdown without writing a .docx.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Emit warnings such as ignored ordered-list numbers.",
    )
    parser.add_argument(
        "positional_inputs",
        nargs="*",
        type=Path,
        help=argparse.SUPPRESS,
    )
    args = parser.parse_args(argv)

    if args.verbose:
        warnings.simplefilter("always")

    try:
        input_files = _collect_inputs(args)
    except Mark2WordError as exc:
        parser.error(str(exc))

    multiple = len(input_files) > 1

    for md_path in input_files:
        theme_dirs = list(args.theme_dir)
        if not args.no_auto_theme_dir:
            for discovered in discover_theme_dirs(md_path):
                if discovered not in theme_dirs:
                    theme_dirs.append(discovered)
        try:
            out_path = resolve_output_path(
                md_path, args.output, multiple_inputs=multiple
            )
        except Mark2WordError as exc:
            parser.error(str(exc))
        convert_one(
            md_path,
            out_path,
            theme_dirs,
            check=args.check,
            verbose=args.verbose,
        )
