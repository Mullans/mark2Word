"""Theme loading, cascade resolution, and unit helpers."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import yaml
from docx.shared import Inches, Pt, RGBColor

from mark2word.errors import FrontmatterError, ThemeError

DEFAULTS: dict[str, Any] = {
    "font": "Calibri",
    "size": 11,
    "color": "000000",
    "body": {},
    "list": {"indent_left": "18pt", "indent_hanging": "18pt"},
    "heading": {"bold": True},
    "code": {"font": "Consolas", "size": 9},
}

PROP_KEYS = {
    "font", "size", "color", "bold", "italic", "align", "line",
    "space_before", "space_between", "space_after",
    "indent_left", "indent_hanging", "indent_first_line", "border_bottom",
    "fill",
}

HEADINGS = {f"h{i}" for i in range(1, 7)}
TEXT_ELEMENTS = {"body", "list", "code"}
TABLE_ELEMENTS = {"table", "th", "td"}

PAGE_SIZES = {
    "letter": (Inches(8.5), Inches(11)),
    "a4": (Inches(8.27), Inches(11.69)),
}

FRONTMATTER_RE = re.compile(
    r"^\ufeff?---[ \t]*\r?\n(.*?)^---[ \t]*\r?\n?(.*)$",
    re.DOTALL | re.MULTILINE,
)


def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = dict(base)
    for key, value in override.items():
        if (
            key in result
            and isinstance(result[key], dict)
            and isinstance(value, dict)
        ):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def to_length(v):
    if isinstance(v, (int, float)):
        return Pt(v)
    s = str(v).strip().lower()
    if s.endswith("pt"):
        return Pt(float(s[:-2]))
    if s.endswith("in"):
        return Inches(float(s[:-2]))
    return Pt(float(s))


def hex_color(v):
    return RGBColor.from_string(str(v).lstrip("#"))


def pt_value(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().lower()
    if s.endswith("pt"):
        return float(s[:-2])
    if s.endswith("in"):
        return float(s[:-2]) * 72.0
    return float(s)


def split_frontmatter(text: str) -> tuple[dict[str, Any], str]:
    m = FRONTMATTER_RE.match(text)
    if not m:
        return {}, text
    parsed = yaml.safe_load(m.group(1))
    body = m.group(2)
    if parsed is None:
        return {}, body
    if not isinstance(parsed, dict):
        raise FrontmatterError(
            f"frontmatter must be a YAML mapping, got {type(parsed).__name__}"
        )
    return parsed, body


def discover_theme_dirs(md_path: Path) -> list[Path]:
    dirs: list[Path] = []
    seen: set[Path] = set()
    for parent in [md_path.parent, *md_path.parent.parents]:
        candidate = (parent / ".mark2word" / "themes").resolve()
        if candidate.is_dir() and candidate not in seen:
            seen.add(candidate)
            dirs.append(candidate)
    return dirs


def _theme_candidates(ext: Path, md_path: Path, theme_dirs=None):
    if ext.is_absolute():
        yield ext
        return
    yield md_path.parent / ext
    for theme_dir in theme_dirs or []:
        yield Path(theme_dir).expanduser() / ext


def _load_theme_file(
    ext: str | Path,
    md_path: Path,
    theme_dirs=None,
    seen: set[Path] | None = None,
) -> dict[str, Any]:
    seen = seen or set()
    checked = [path.resolve() for path in _theme_candidates(Path(ext), md_path, theme_dirs)]
    ext_path = next((path for path in checked if path.exists()), None)
    if ext_path is None:
        searched = ", ".join(str(path) for path in checked)
        raise FileNotFoundError(f"extends references missing theme: {ext}; searched: {searched}")
    if ext_path in seen:
        raise ThemeError(f"theme inheritance cycle detected at: {ext_path}")
    seen.add(ext_path)
    data = yaml.safe_load(ext_path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise ThemeError(f"theme file must be a YAML mapping: {ext_path}")
    parent_ext = data.pop("extends", None)
    if parent_ext is not None:
        parent = _load_theme_file(parent_ext, md_path, theme_dirs, seen)
        data = deep_merge(parent, data)
    return data


def load_theme(
    front: dict[str, Any],
    md_path: Path,
    theme_dirs=None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    external: dict[str, Any] = {}
    front = dict(front)
    ext = front.pop("extends", None)
    if ext is not None:
        external = _load_theme_file(ext, md_path, theme_dirs)
    return external, front


def _bare(layer: dict[str, Any]) -> dict[str, Any]:
    return {k: v for k, v in layer.items() if k in PROP_KEYS}


def _key_order(element: str) -> list[str]:
    if element in HEADINGS:
        return ["heading", element]
    if element in TEXT_ELEMENTS:
        return ["text", element]
    if element in TABLE_ELEMENTS:
        if element == "table":
            return ["table"]
        return ["table", element]
    return [element]


def _layer_contribution(layer: dict[str, Any], element: str) -> dict[str, Any]:
    out = dict(_bare(layer))
    for key in _key_order(element):
        sub = layer.get(key)
        if isinstance(sub, dict):
            out.update({k: v for k, v in sub.items() if k in PROP_KEYS})
    return out


class Resolver:
    def __init__(self, external: dict[str, Any], glob: dict[str, Any]):
        self.global_layers = [DEFAULTS, external, glob]
        self.region_sources = [external, glob]
        self._cache: dict[tuple[str, tuple[str, ...]], dict[str, Any]] = {}

    @staticmethod
    def _descend(layer: dict[str, Any], path: list[str]) -> dict[str, Any] | None:
        node: Any = layer
        for name in path:
            if not isinstance(node, dict):
                return None
            node = node.get("$" + name)
        return node if isinstance(node, dict) else None

    def resolve(self, element: str, region_path: tuple[str, ...] | None = None) -> dict[str, Any]:
        key = (element, region_path or ())
        if key not in self._cache:
            self._cache[key] = self._resolve_uncached(element, region_path)
        return dict(self._cache[key])

    def _resolve_uncached(
        self, element: str, region_path: tuple[str, ...] | None = None
    ) -> dict[str, Any]:
        style: dict[str, Any] = {}
        for layer in self.global_layers:
            style.update(_layer_contribution(layer, element))
        for depth in range(1, len(region_path or ()) + 1):
            prefix = region_path[:depth]
            for src in self.region_sources:
                block = self._descend(src, list(prefix))
                if block is not None:
                    style.update(_layer_contribution(block, element))
        return style

    def _list_block_from_layers(self, region_path: tuple[str, ...] | None) -> dict[str, Any]:
        """Merge raw `list` theme blocks (including nesting meta keys)."""
        block: dict[str, Any] = {}
        for layer in self.global_layers:
            sub = layer.get("list")
            if isinstance(sub, dict):
                block = deep_merge(block, sub)
        for depth in range(1, len(region_path or ()) + 1):
            prefix = region_path[:depth]
            for src in self.region_sources:
                node = self._descend(src, list(prefix))
                if isinstance(node, dict):
                    sub = node.get("list")
                    if isinstance(sub, dict):
                        block = deep_merge(block, sub)
        return block

    def resolve_list_style(
        self,
        list_level: int,
        region_path: tuple[str, ...] | None = None,
    ) -> dict[str, Any]:
        """Resolve list styling with per-level indents and optional level overrides."""
        style = dict(self.resolve("list", region_path))
        list_block = self._list_block_from_layers(region_path)
        levels = list_block.get("levels") or {}
        level_ov = levels.get(list_level) or levels.get(str(list_level))
        if isinstance(level_ov, dict):
            style.update({k: v for k, v in level_ov.items() if k in PROP_KEYS})

        if not isinstance(level_ov, dict) or "indent_left" not in level_ov:
            base = pt_value(style.get("indent_left", 0))
            step_raw = list_block.get("indent_step")
            if step_raw is None:
                step_raw = style.get("indent_hanging") or style.get("indent_left") or 18
            step = pt_value(step_raw)
            if style.get("indent_left") is not None or list_block.get("indent_step") is not None or list_level > 0:
                style["indent_left"] = base + list_level * step
        return style
