"""Markdown to Word conversion with YAML theming."""

from mark2word.ast import Block, InlineRun
from mark2word.cli import main
from mark2word.emit import build
from mark2word.errors import (
    FrontmatterError,
    Mark2WordError,
    ParseError,
    RegionError,
    ThemeError,
)
from mark2word.parser import parse_blocks, parse_inline, parse_to_ast
from mark2word.plugins import register_block_emitter, register_block_parser
from mark2word.theme import Resolver, deep_merge, discover_theme_dirs, load_theme, split_frontmatter

__all__ = [
    "Block",
    "FrontmatterError",
    "InlineRun",
    "Mark2WordError",
    "ParseError",
    "RegionError",
    "Resolver",
    "ThemeError",
    "build",
    "deep_merge",
    "discover_theme_dirs",
    "load_theme",
    "main",
    "parse_blocks",
    "parse_inline",
    "parse_to_ast",
    "register_block_emitter",
    "register_block_parser",
    "split_frontmatter",
]

if __name__ == "__main__":
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    main()
