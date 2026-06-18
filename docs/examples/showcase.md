---
extends: showcase-theme.yaml
title: mark2word Style Gallery
$pullquote:
  color: "943634"
---

<!-- region: header-region -->
# mark2word Style Gallery
Themed Markdown → polished Word · docs/examples/showcase.md
<!-- /region -->

## At a Glance

This document is a living catalog of what mark2word can render. Jump to any section:

- [Inline markup](#inline-markup) — emphasis, links, dual-align
- [Lists](#lists) — nested bullets and numbers
- [Images & tables](#images-and-tables) — layout and cell styling
- [Code](#code-blocks) — fenced blocks with language themes
- [Theme regions](#theme-regions) — local overrides without touching body text

The **page header and footer** on every sheet come from the theme (`page.header` / `page.footer` with `{title}`, `{page}`, and `{pages}`). The centered line at the very end of this file is a **body footer region** — content you control in Markdown, not Word page chrome.

## Document Purpose

mark2word separates *what you write* from *how it looks*. YAML frontmatter, external themes, and invisible HTML comments carry styling; the Markdown body stays readable in GitHub, Obsidian, or any previewer.

> **Why it matters.** Authors draft in plain Markdown. Designers tune one theme file. Agents can regenerate content without rebuilding Word XML by hand — and the same skeleton can become a memo, a proposal, or a report by swapping YAML.

For __human workflows__, that means polish without fighting the Word ribbon. For __agent workflows__, it means structured output that compiles to a client-ready `.docx`.

---

### Heading Level Three

#### Heading Level Four

##### Heading Level Five

###### Heading Level Six

## Inline Markup

A body paragraph with **bold**, *italic*, ***bold italic***, _underscore italic_, __underscore bold__, and `inline code`. External links work: [**mark2word on GitHub**](https://github.com/Mullans/mark2Word). Internal links jump to bookmarks on headings — try [the lists section](#lists) or [code blocks](#code-blocks).

Dual-aligned lines split on `||` (ignored inside backticks):

Left-aligned label || Right-aligned value

<!-- region: pullquote -->
The same Markdown skeleton can look like a memo, a proposal, or a résumé — swap the theme, keep the words.
<!-- /region -->

## Lists

**Unordered lists** use theme `ul` levels (`*`, `o`, `-` at depth):

- Top-level bullet
- second item
  - nested level one
    - nested level two
- back to top level

**Ordered lists** use `ol` levels (`1.`, `a.`, `i.`):

1. First top-level item
2. Second item
   1. Nested ordered (letter format at depth 1 in theme)
   2. Another nested item
3. Third top-level item

A body paragraph between lists **restarts numbering** — useful when prose interrupts a sequence.

9. This run starts at nine (markdown sets the seed value)
10. and continues from there

## Images and Tables

Images resolve relative to this file. Alt text can feed Word accessibility and/or a caption (`image.alt_mode` in theme):

![Accent swatch used in headings and table headers](sample.png)

<!-- region: feature-grid -->
| Feature | Markdown | Theme keys |
| - | - | - |
| Page chrome | *(none — theme only)* | `page.header`, `page.footer`, `{title}` |
| Headings | `#` … `######` | `heading`, `h1`–`h6`, bookmarks for `#links` |
| Blockquote | `> quote` | `blockquote`, `indent_left` / `indent_right` |
| Horizontal rule | `---` on its own line | `hr`, `border_bottom` |
| Page break | `<!-- pagebreak -->` | *(invisible in MD preview)* |
| Lists | `-` / `1.` | `list`, `ol`, `ul`, `levels` |
| Regions | `<!-- region: name -->` | `$name` nested blocks |
| Tables | pipe syntax | `table`, `th`, `td`, `border`, `padding` |
| Code | fenced blocks | `code`, `code.langs.{lang}` |
| Images | `![alt](path)` | `image`, `width`, `max_width`, `alt_mode` |
<!-- /region -->

## Code Blocks

Language tags select optional overrides under `code.langs` in the theme. Pipes inside fences are **not** dual-align markers:

```python
# showcase-theme.yaml → code.langs.python
def convert(md_path: Path) -> Path:
    return md_path.with_suffix(".docx")
```

```yaml
# Chained extends in showcase.md → showcase-theme.yaml → showcase-base.yaml
page:
  footer: "Style Gallery || Page {page} of {pages}"
code:
  langs:
    yaml: { color: "006666" }
```

And there can `be inline code sections` too, and page-breaks like this:

<!-- pagebreak -->

## Theme Regions

Regions wrap stretches of Markdown and apply `$region-name` style blocks from the theme or frontmatter. They nest and override globals without inline formatting.

<!-- region: theme-sample -->
**Theme tips**

- Quote hex colors in YAML — `color: "5B5B5B"`
- Use `text` for shared body/list defaults
- Chain `extends` across theme files (see `showcase-base.yaml`)
<!-- /region -->

<!-- region: theme-sample2 -->
- Layer `$regions` for local personality
- Override one key without losing the rest of the cascade
- Validate with `uv run mark2word --check-theme showcase-theme.yaml`
<!-- /region -->

<!-- region: theme-list-sample -->
Per-level list colors in a region:

- level zero
  - level one
    - level two
- back to zero
<!-- /region -->

## Appendix: Zero-Padded Numbers

<!-- region: appendix -->
The `$appendix` region sets `ol.levels.0.format: "01."` for formal enumerations:

1. First appendix item
2. Second appendix item
3. Third appendix item
<!-- /region -->

---

*End of gallery body. The line below is a styled region, not the Word page footer (see theme `page.footer`).*

<!-- region: footer-region -->
mark2word showcase || generated from docs/examples/showcase.md
<!-- /region -->
