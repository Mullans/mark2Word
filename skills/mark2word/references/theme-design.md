# Theme Design

Read this when creating, editing, reviewing, or substantially customizing a Mark2Word theme.

## Location

Create project themes in a local folder such as:

```text
.mark2word/themes/
```

Pass that folder during conversion:

```bash
python /path/to/mark2word/scripts/mark2word.py \
  --theme-dir /path/to/mark2word/assets \
  --theme-dir /path/to/project/.mark2word/themes \
  /path/to/project/document.md
```

The converter also auto-discovers `.mark2word/themes` near the input file unless `--no-auto-theme-dir` is set.

## Starting Point

Use `assets/base-theme.yaml` as the baseline. For a custom theme, copy it into the project theme folder, rename it, and edit:

```text
.mark2word/themes/executive-resume.yaml
```

Point the Markdown document at it:

```yaml
---
extends: executive-resume.yaml
---
```

## Theme Inheritance (`extends`)

Themes compose through `extends`. It works in **both** places:

1. **Markdown frontmatter** — primary entry point for a document.
2. **Theme YAML files** — chain base → specialized layers.

Each child theme is deep-merged over its parent. Later layers win on conflicting keys.

**Example chain** (from `docs/examples/` in the mark2Word repo):

```yaml
# showcase-base.yaml — shared defaults
font: Calibri
size: 10
heading: { bold: true, color: "2B579A" }
```

```yaml
# showcase-theme.yaml — document-specific layer
extends: showcase-base.yaml
size: 11
h1: { size: 22 }
table: { size: 10 }
th: { bold: true, fill: "2B579A", color: "FFFFFF" }
```

```yaml
# showcase.md frontmatter — per-document tweaks
---
extends: showcase-theme.yaml
$pullquote: { color: "943634" }
---
```

Relative `extends` paths resolve from the Markdown file's directory, then each `--theme-dir` in order. Cycles are detected and rejected with a clear error.

You do **not** need to flatten inheritance into a single file — layered themes are the intended pattern.

## Style Priority

Within the merged theme, styles resolve in this order (later wins):

1. Built-in defaults
2. External theme chain (base → … → leaf theme file)
3. Frontmatter globals (keys not starting with `$`)
4. Active region path — outer regions first, inner regions last

Within a layer, specificity applies: `h2` beats `heading`, `body` beats `text`, `ol`/`ul` beat shared `list`.

## Theme Shape

A theme is YAML. Top-level properties apply globally; element blocks override them for specific content.

```yaml
page:
  size: letter
  margin: { top: 0.5in, bottom: 0.5in, left: 0.7in, right: 0.7in }

font: Calibri
size: 10
color: "000000"

text: { line: 1.1 }
body: { space_after: 2pt }
list:
  space_before: 2pt
  space_between: 2pt
  space_after: 6pt
  indent_left: 9pt
  indent_hanging: 9pt
  indent_step: 9pt

heading: { bold: true, color: "2B579A" }
h1: { size: 14, color: "000000" }
h2: { size: 12, space_before: 12pt, space_after: 4pt, border_bottom: { size: 0.5pt, color: "2B579A" } }

code: { font: Consolas, size: 9, color: "1F497D" }

table: { font: Calibri, size: 10 }
th: { bold: true, color: "FFFFFF", fill: "2B579A", align: center }
td: { size: 10 }
```

## Target Keys

| Key | Applies to |
|-----|------------|
| `body` | Ordinary paragraphs |
| `text` | Shared defaults for `body` and list items |
| `list` | Both ordered and unordered lists |
| `ol` | Ordered lists only (overrides `list`) |
| `ul` | Unordered lists only (overrides `list`) |
| `heading` | Shared defaults for `h1`–`h6` |
| `h1` … `h6` | Individual heading levels |
| `code` | Fenced and inline code |
| `table` | Whole table defaults |
| `th` / `td` | Header / data cells |
| `$name` | Named region matched by `<!-- region: name -->` |

Region blocks accept the same nested keys (`body`, `h2`, `list`, etc.) scoped to that region.

## Style Keys

Supported style properties:

- `font` — font family (`Calibri`, `Aptos`, `Consolas`, …)
- `size` — point size (`10` or `10pt`)
- `color` — hex color (`"2B579A"` or `"#2B579A"`)
- `bold`, `italic` — `true` / `false`
- `align` — `left`, `center`, `right`, `justify`
- `line` — line-spacing multiple (`1.1`) or exact points (`13pt`)
- `space_before`, `space_after` — paragraph spacing
- `space_between` — spacing between adjacent list items
- `indent_left`, `indent_first_line`, `indent_hanging` — paragraph indents
- `border_bottom` — bottom border, e.g. `{ size: 0.5pt, color: "2B579A" }`
- `fill` — table cell background hex (for `th` / `td`)

Lengths: bare points (`10`), explicit points (`10pt`), or inches (`0.5in`).

## Page Settings

Supported page sizes: `letter`, `a4`.

```yaml
page:
  size: letter
  margin:
    top: 0.5in
    bottom: 0.5in
    left: 0.7in
    right: 0.7in
```

## Lists and Numbering

List appearance splits into two mechanisms:

1. **Paragraph/run styling** — colors, fonts, spacing from theme keys.
2. **Word multilevel numbering** — indents and bullet/number formats compiled into native Word list definitions so editing in Word behaves correctly.

Shared list keys under `list` apply to both kinds. Use `ol` or `ul` for kind-specific overrides (same relationship as `heading` → `h1`–`h6`).

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

### List meta keys (under `list`, `ol`, or `ul`)

- `indent_left`, `indent_hanging`, `indent_step` — nesting indents (compiled into Word numbering)
- `levels` — per-depth overrides keyed by `0`, `1`, `2`, …

### Per-level keys inside `levels`

- `format` — shorthand or template (see table below)
- `num_fmt` + `template` — explicit Word numbering (both required if either is set)
- Any [style key](#style-keys) such as `color` or `font` (applied to list item text)

Default ordered format is `1.` (displays `1.`, `2.`, `3.`). Use template `1` without a trailing period when you want `1`, `2`, `3`.

| `format` value | Numbering | Examples |
|----------------|-----------|----------|
| `1` | decimal | 1, 2, 3 |
| `1.` | decimal | 1., 2., 3. |
| `a` / `alph` | lower letter | a, b, c |
| `a.` | lower letter | a., b., c. |
| `A` / `Alph` | upper letter | A, B, C |
| `(A)` | upper letter | (A), (B), (C) |
| `i` / `roman` | lower Roman | i, ii, iii |
| `roman )` | lower Roman | i ), ii ), iii ) |
| `I` / `Roman` | upper Roman | I, II, III |
| `Section 1:` | decimal | Section 1:, Section 2:, … |
| `•`, `-`, `*` | bullet | literal bullet character |
| `bullet` | bullet | Word bullet format |

Explicit control:

```yaml
ol:
  levels:
    2:
      num_fmt: upperRoman
      template: "(%1)"
```

**Markdown nesting:** indent list items with 2 spaces per level. A body paragraph between items starts a new list run (ordered numbering restarts).

## Regions

Regions provide local style overrides. Define the region in the theme or frontmatter:

```yaml
$header:
  align: center
  h1: { space_after: 1.5pt }
  text: { size: 9, space_after: 0.75pt }

$pullquote:
  body: { align: center, italic: true, space_before: 8pt }
```

Use in Markdown:

```markdown
<!-- region: header -->
# Greg Davies
Chicago, IL || greg@example.com
<!-- /region -->
```

Nested regions are supported. Inner regions override outer regions on conflicting keys.

## Tables and Code

Pipe tables pick up `table`, `th`, and `td` styles. Header row is the first row; the separator row (`| - | - |`) is required.

Fenced code blocks and inline `` `code` `` use the `code` target. Language tags on fences are parsed but not yet used for syntax highlighting.

## Design Workflow

1. Identify the document type: resume, report, proposal, memo, etc.
2. Copy `assets/base-theme.yaml` into `.mark2word/themes/<name>.yaml` or chain from it with `extends`.
3. Adjust global typography: `font`, `size`, `color`, `text.line`, page margins.
4. Tune headings, lists, tables, and code styling.
5. Add `$region` blocks only for repeated local treatments (headers, pull quotes, footers).
6. Convert a representative document; use `--check` first if iterating on YAML syntax.
7. Inspect the `.docx` in Word — especially nested lists, tables, and region boundaries.

## Common Mistakes

- **Missing `--theme-dir`** — if `extends: base-theme.yaml` cannot be found, include the skill `assets/` folder and any project theme folder.
- **Overusing regions** — prefer element styles until a local block is genuinely needed.
- **Unquoted hex colors** — quote values like `"2B579A"` so YAML does not treat them as numbers.
- **Flattening theme chains** — nested `extends` in theme files is supported; use layers instead of one giant file.
- **Expecting `indent_right`** — not currently supported; use dual-align lines or margins instead.
- **Ordered list gaps** — markdown `1.` then `3.` is accepted but `--verbose` warns; Word continues the sequence.
