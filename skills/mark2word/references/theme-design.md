# Theme Design

Read this when creating, editing, reviewing, or substantially customizing a Mark2Word theme.

Validate a theme before converting:

```bash
uv run mark2word --check-theme .mark2word/themes/my-theme.yaml --theme-dir .mark2word/themes
```

## Location

Create project themes in:

```text
.mark2word/themes/
```

Pass during conversion:

```bash
python /path/to/mark2word/scripts/mark2word.py \
  --theme-dir /path/to/mark2word/assets \
  --theme-dir /path/to/project/.mark2word/themes \
  /path/to/project/document.md
```

Auto-discovery: `.mark2word/themes` near the input file (unless `--no-auto-theme-dir`).

## Starting Point

Copy `assets/base-theme.yaml` into the project theme folder:

```text
.mark2word/themes/executive-resume.yaml
```

Point the document at it:

```yaml
---
extends: executive-resume.yaml
---
```

## Theme Inheritance (`extends`)

Themes compose through `extends` in **both** markdown frontmatter and theme YAML files. Child keys deep-merge over parents; later layers win. Cycles are rejected.

**Example chain** (`docs/examples/` in the mark2Word repo):

```yaml
# showcase-base.yaml
font: Calibri
heading: { bold: true, color: "2B579A" }
```

```yaml
# showcase-theme.yaml
extends: showcase-base.yaml
size: 11
h1: { size: 22 }
th: { bold: true, fill: "2B579A", color: "FFFFFF" }
```

```yaml
# showcase.md frontmatter
---
extends: showcase-theme.yaml
$pullquote: { color: "943634" }
---
```

Layered themes are the intended pattern — do not flatten into one file.

## Style Priority

1. Built-in defaults  
2. External theme chain  
3. Frontmatter globals  
4. Active region path (outer → inner)  

Within a layer: `h2` > `heading`, `body` > `text`, `ol`/`ul` > `list`.

## Theme Shape

```yaml
title: My Document

page:
  size: letter
  margin: { top: 0.5in, bottom: 0.5in, left: 0.7in, right: 0.7in }
  header: "{title}"
  footer: "Confidential || Page {page} of {pages}"

font: Calibri
size: 10
color: "000000"

text: { line: 1.1 }
body: { space_after: 2pt }

blockquote:
  indent_left: 18pt
  italic: true

list:
  space_between: 2pt
  indent_left: 9pt
  indent_hanging: 9pt
  indent_step: 9pt

heading: { bold: true, color: "2B579A" }
h2: { border_bottom: { size: 0.5pt, color: "2B579A" } }

code:
  font: Consolas
  size: 9
  langs:
    python: { color: "000080" }

image:
  max_width: 5in
  alt_mode: caption
  align: center

hr:
  border_bottom: { size: 0.5pt, color: "999999" }

table:
  space_before: 6pt
  space_after: 6pt
  border: { size: 0.5pt, color: "CCCCCC" }
th:
  bold: true
  fill: "2B579A"
  color: "FFFFFF"
  padding: { top: 4pt, bottom: 4pt, left: 6pt, right: 6pt }
td: { size: 10 }
```

## Target Keys

| Key | Applies to |
|-----|------------|
| `body` | Ordinary paragraphs |
| `text` | Shared defaults for `body` and list items |
| `blockquote` | `>` blockquote lines |
| `code` | Fenced and inline code |
| `image` | `![alt](path)` images |
| `hr` | Horizontal rules (`---` in body) |
| `list` | Both ordered and unordered lists |
| `ol` / `ul` | Kind-specific list overrides |
| `heading` | Shared heading defaults |
| `h1` … `h6` | Individual heading levels |
| `table` | Whole table |
| `th` / `td` | Header / data cells |
| `$name` | Region matched by `<!-- region: name -->` |

Regions accept nested keys (`body`, `h2`, `list`, etc.) scoped to that region.

## Style Keys

- `font`, `size`, `color`, `bold`, `italic`
- `align` — `left`, `center`, `right`, `justify`
- `line` — multiple (`1.1`) or exact points (`13pt`)
- `space_before`, `space_after`, `space_between`
- `indent_left`, `indent_right`, `indent_first_line`, `indent_hanging`
- `border_bottom` — `{ size: 0.5pt, color: "2B579A" }`
- `fill` — table cell background

**Image** (under `image`):

| Key | Purpose |
|-----|---------|
| `width` | Fixed picture width |
| `max_width` | Scale down proportionally if wider |
| `align` | Caption alignment when alt is shown |
| `alt_mode` | `doc` (accessibility, default), `caption`, `both`, `none` |

**Table** (under `table`, `th`, `td`):

| Key | Purpose |
|-----|---------|
| `border` | On `table`: `{ size, color }` |
| `padding` | On `th`/`td`: `{ top, bottom, left, right }` in points |
| `space_before`, `space_after` | On `table` |

**Code** per-language (fence tag matches key under `code.langs`):

```yaml
code:
  font: Consolas
  langs:
    python: { color: "000080" }
    yaml: { color: "008080" }
```

Lengths: bare points (`10`), `10pt`, or `0.5in`.

## Page Settings

Sizes: `letter`, `a4`.

```yaml
title: My Document    # {title} in header/footer; else first h1

page:
  size: letter
  margin:
    top: 0.5in
    bottom: 0.5in
    left: 0.7in
    right: 0.7in
  header: "{title}"
  footer: "Draft || Page {page} of {pages}"
```

Placeholders: `{page}`, `{pages}`, `{title}`. Dual-align with `left || right`.

Theme page chrome is document-wide. Body `$footer` regions in markdown are content footers, not Word page footers.

## Invisible Markdown (Word-only)

Not shown in normal Markdown preview:

- YAML frontmatter
- `<!-- region: … -->`, `<!-- /region -->`, `<!-- pagebreak -->`

Horizontal rules use body `---` after frontmatter is stripped — not confused with frontmatter fences.

Do not use `\newpage` (visible in MD renderers). Use `<!-- pagebreak -->`.

## Internal Links

Headings get Word bookmarks (slug from text). Link with `[text](#heading-slug)`. Unknown anchors fail at conversion.

## Blockquotes and Horizontal Rules

- **Blockquote:** lines starting with `>` (blank line ends the quote unless next line continues with `>`)
- **HR:** `---`, `***`, or `___` alone on a line in the document body

## Lists and Numbering

1. **Paragraph/run styling** — colors, fonts, spacing from theme keys.  
2. **Word multilevel numbering** — indents and formats compiled into native list definitions.

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

### List meta keys

- `indent_left`, `indent_hanging`, `indent_step`
- `levels` — keyed by `0`, `1`, `2`, …

### Per-level keys

- `format` — shorthand (see table)
- `num_fmt` + `template` — explicit Word control (both required if either set)
- Any style key for list item text

Default ordered format: `1.`

| `format` | Examples |
|----------|----------|
| `1` / `1.` | 1, 2, 3 / 1., 2., 3. |
| `01` / `01.` | 01, 02, 03 |
| `a`, `a.`, `A`, `(A)` | Letters |
| `i`, `roman`, `I`, `Roman` | Roman numerals |
| `Section 1:` | Section 1:, Section 2:, … |
| `•`, `-`, `*`, `bullet` | Bullets |

Explicit control:

```yaml
ol:
  levels:
    2:
      num_fmt: upperRoman
      template: "(%1)"
```

**Nesting:** 2 spaces per level. Body paragraph between items starts a new list run.

## Regions

```yaml
$header:
  align: center
  h1: { space_after: 1.5pt }

$pullquote:
  body:
    align: center
    indent_left: 24pt
    indent_right: 24pt
    italic: true
```

```markdown
<!-- region: header -->
# Greg Davies
Chicago, IL || greg@example.com
<!-- /region -->
```

Nested regions supported.

## Tables and Code

Pipe tables: first row = header, separator row required (`| - | - |`).

Fenced blocks use `code` styling; language tag selects `code.langs.{lang}` override when present.

## Design Workflow

1. Identify document type (resume, report, memo, …).  
2. Copy or chain from `assets/base-theme.yaml`.  
3. Set global typography and page settings.  
4. Tune headings, lists, tables, code, images.  
5. Add `$region` blocks only for repeated local treatments.  
6. Run `--check-theme` on the theme, `--check` on a sample document.  
7. Convert and inspect in Word.

## Common Mistakes

- **Missing `--theme-dir`** — include skill `assets/` and project theme folder when using `extends: base-theme.yaml`.  
- **Unquoted hex colors** — use `"2B579A"` so YAML does not parse as a number.  
- **Flattening theme chains** — use layered `extends` instead of one giant file.  
- **Ordered list gaps** — `1.` then `3.` is accepted; `--verbose` warns.  
- **Confusing page chrome with body footers** — theme `page.footer` vs markdown `$footer` region.  
- **Internal link slug mismatch** — slug is lowercase, hyphenated heading text.
