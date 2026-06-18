# mark2Word

Convert styled Markdown to Word (`.docx`). YAML frontmatter and external themes control typography and layout; the Markdown body stays readable in any viewer.

## Installing

Clone the repo and run `uv sync`. CLI entry points: `mark2word`, `sync-skill`.

## Running

```bash
uv run mark2word YOUR_FILE.md
uv run mark2word --theme-dir docs/examples YOUR_FILE.md
uv run mark2word -i FILE1.md -i FILE2.md -o OUT_FOLDER
uv run mark2word -i FILE.md -o OUTPUT.docx
uv run mark2word --check YOUR_FILE.md
uv run mark2word --check-theme my-theme.yaml --theme-dir .mark2word/themes
uv run mark2word --version
uv run sync-skill
```

| Flag | Purpose |
|------|---------|
| `-i` / `--input` | Markdown input (repeat for batch) |
| `-o` / `--output` | Output `.docx` or directory (directory required for batch) |
| `--theme-dir` | Folder for relative `extends` paths (repeatable) |
| `--no-auto-theme-dir` | Skip auto-discovery of `.mark2word/themes` |
| `--check` | Validate frontmatter, theme, and markdown; no `.docx` written |
| `--check-theme` | Validate a theme YAML file (including `extends` chain) |
| `--verbose` | Warn on ordered-list number mismatches |
| `-V` / `--version` | Print version |

Exit codes: `0` success, `1` error, `2` usage error.

The converter validates inputs and outputs before building (missing files, theme errors, read-only or locked outputs).

## Markdown Elements

| Syntax | Output |
|--------|--------|
| `#` … `######` | Headings (Word bookmarks for internal links) |
| Plain line | Body paragraph (one line = one paragraph) |
| `- item` / `1. item` | Bulleted / numbered lists (native Word multilevel; 2 spaces per nest level) |
| `> quote` | Blockquote (shaded cell with optional side borders) |
| `**b**` `*i*` `` `code` `` | Inline emphasis and code |
| `[text](url)` | Hyperlink (bold/italic inside link text supported) |
| `[text](#slug)` | Internal link to a heading bookmark |
| `left \|\| right` | Dual-aligned line (ignored inside backticks) |
| `![alt](path.png)` | Embedded image (path relative to the markdown file) |
| Pipe table + `\| - \|` separator | Table |
| ` ```lang ` … ` ``` ` | Fenced code block |
| `---` / `***` / `___` alone on a line | Horizontal rule (body only — frontmatter stripped first) |
| `<!-- region: name -->` … `<!-- /region -->` | Named style region |
| `<!-- pagebreak -->` | Page break |

**Invisible in normal Markdown preview:** YAML frontmatter, HTML comments (`region`, `pagebreak`), and all Word-only styling.

**Internal link slugs** are derived from heading text (e.g. `## My Section` → `#my-section`). Duplicate headings get `-2`, `-3`, etc.

## Frontmatter and Themes

Frontmatter sits between `---` fences at the top of the file. Use it for inline styling, region overrides, or to reference an external theme (or both).

```yaml
---
extends: base-theme.yaml
title: Quarterly Report
font: Calibri
body: { space_after: 2pt }
$callout: { color: "943634" }
---
```

### Theme inheritance (`extends`)

Relative paths resolve from the markdown file's directory, then each `--theme-dir`. Theme YAML files may also declare `extends` (chained inheritance). Child keys deep-merge over parents. Cycles are rejected.

```text
document.md  →  extends: report-theme.yaml
report-theme.yaml  →  extends: base-theme.yaml
```

Auto-discovery: unless `--no-auto-theme-dir` is set, `.mark2word/themes` is searched near the input file and in parent directories.

### Target keys

| Key | Applies to |
|-----|------------|
| `body` | Body paragraphs |
| `blockquote` | `>` blockquote lines (shaded cell with optional side borders) |
| `code` | Shared code typography (font, size, color) |
| `code_block` | Fenced code blocks (` ```lang ` … ` ``` `) |
| `code_inline` | Inline `` `code` `` backticks |
| `image` | `![alt](path)` images |
| `hr` | Horizontal rules |
| `table`, `th`, `td` | Pipe tables |
| `list`, `ol`, `ul` | List settings (`ol`/`ul` override shared `list`) |
| `text` | Shared defaults for body and list |
| `heading`, `h1`–`h6` | Headings |
| `$name` | Region matched by `<!-- region: name -->` |

### Style keys

- `font`, `size`, `color`, `bold`, `italic`
- `align` — `left`, `center`, `right`, `justify`
- `line` — multiple (e.g. `1.1`) or exact points (`13pt`)
- `space_before`, `space_after`, `space_between`
- `indent_left`, `indent_right`, `indent_first_line`, `indent_hanging`
- `border_bottom`, `border_left`, `border_right` — `{ size: 0.5pt, color: "2B579A" }` (use `none` to disable)
- `fill` — background color; use `none` to disable. Applies to table cells (`th`/`td`), fenced code blocks (paragraph shading), inline code (run highlight), and blockquotes (cell background)
- `padding` — blockquote cell or table cell insets: `{ top, bottom, left, right }` in points or inches

**Blockquote** (under `blockquote`):

- Uses a single-cell table aligned to the text column; `indent_left` / `indent_right` offset the bar from the margin
- `space_before` is placed above the quote (outside the shaded cell), not inside it

**Image** (under `image`):

- `width`, `max_width` — size in points or inches; `max_width` scales proportionally
- `align` — paragraph alignment for the image (and caption when shown)
- `alt_mode` — `doc` (Word accessibility metadata, default), `caption` (visible caption paragraph), `both`, `none`

**Table** (under `table`, `th`, or `td`):

- `border` — on `table`: `{ size: 0.5pt, color: "CCCCCC" }`
- `padding` — on `th`/`td`: `{ top: 4pt, bottom: 4pt, left: 6pt, right: 6pt }`
- `space_before`, `space_after` — on `table`

**Code** — shared typography under `code`; block-specific keys under `code_block` and `code_inline`. Per-language overrides match the fence language tag:

```yaml
code:
  font: Consolas
  size: 9
code_block:
  fill: "F5F5F5"
  langs:
    python: { color: "000080" }
    yaml: { color: "008080" }
code_inline:
  fill: "F5F5F5"
```

Built-in defaults (`defaults.yaml`) supply baseline code fills, blockquote styling, and list indents; override in themes or frontmatter as needed.

### List numbering

Settings under `list` apply to both ordered and unordered lists. Use `ol` or `ul` for kind-specific overrides. Indents and formats compile into native Word multilevel definitions.

```yaml
list:
  indent_left: 14pt
  indent_hanging: 9pt
  indent_step: 14pt
  levels:
    2: { color: "943634" }
ol:
  levels:
    0: { format: "1." }
    1: { format: "a." }
ul:
  levels:
    0: { format: "•" }
    1: { format: "◦" }
```

Per-level keys: `format`, or `num_fmt` + `template` (both required if either is set), plus any style key for list item text.

| `format` | Examples |
|----------|----------|
| `1` / `1.` | 1, 2, 3 / 1., 2., 3. |
| `01` / `01.` | 01, 02, 03 |
| `a`, `a.`, `A`, `(A)` | a, b, c / A, B, C / (A), (B) |
| `i`, `roman`, `I`, `Roman` | i, ii, iii / I, II, III |
| `Section 1:` | Section 1:, Section 2:, … |
| `•`, `-`, `*`, `bullet` | Bullets |

### Page settings

```yaml
title: My Document          # for {title} in header/footer; falls back to first h1
page:
  size: letter                # or a4
  margin: { top: 0.5in, bottom: 0.5in, left: 0.7in, right: 0.7in }
  header: "{title}"
  footer: "Confidential || Page {page} of {pages}"
```

Header/footer placeholders: `{page}`, `{pages}`, `{title}`. Use `left || right` for dual-aligned lines.

Theme page chrome is document-wide. Content footers in the markdown body (e.g. a `$footer-region` region) are separate from `page.footer`.

### Regions

Define `$region-name` in frontmatter or theme, then wrap markdown. Top-level region keys (`font`, `color`, …) apply to body text inside the region; nested keys like `body` or `h2` target specific elements.

```markdown
<!-- region: callout -->
Important note with **emphasis**.
<!-- /region -->
```

Nested regions supported; inner overrides beat outer.

### Style priority

1. Built-in defaults (`src/mark2word/defaults.yaml`) — always merged first  
2. External theme chain (`extends`)  
3. Frontmatter globals  
4. Active region path (outer → inner; regions use theme + frontmatter only, not built-in defaults)  

Within a layer, later keys override earlier: top-level props → category → subcategory (e.g. `code` → `code_block`, `heading` → `h2`, `text` → `body`). Region top-level keys (`font`, `size`, `color`, …) apply to all elements in that region, including body paragraphs.

## Example

See `docs/examples/showcase.md` with `docs/examples/showcase-theme.yaml` for a full feature demo.

```markdown
---
extends: base-theme.yaml
title: Sample Memo
page:
  footer: "Draft || Page {page}"
heading: { bold: true, color: "2B579A" }
$header: { align: center }
---

<!-- region: header -->
# Sample Memo
Author || 2026-06-17
<!-- /region -->

## Overview

Body with **bold**, [a link](https://example.com), and [a section jump](#overview).

> A blockquote for emphasis.

---

| Item | Status |
| - | - |
| Lists | Done |

<!-- pagebreak -->

## Appendix

1. First
   1. Nested
2. Second
```
