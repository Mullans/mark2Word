# mark2Word

A tool to convert Markdown files to Word documents, but with all of the styling thrown in too.

## Installing

Until I set up packaging, just clone the repo, then run `uv sync` from inside it to install the dependencies.

## Running

```bash
uv run mark2word YOUR_FILE.md
# Keep your themes in a different directory
uv run mark2word --theme-dir docs/examples YOUR_FILE.md
# Specify the output file name (default just swaps extension)
uv run mark2word --input YOUR_FILE.md --output YOUR_DOC.docx
# Convert multiple files at once
uv run mark2word --input FILE1.md --input FILE2.md --output OUT_FOLDER
# Copy the script into the skill
uv run sync-skill
```

## Markdown Elements

- `#` .. `######` - headings (h1..h6)
- `- item`   /   `1. item` - bulleted / numbered list items (real Word lists)
- `text` - body paragraph (one line == one paragraph)
- `**b**` `*i*` `***bi***` - inline emphasis (**b** *i* ***bi***)
- `\\*`  - * (literal asterisk)
- `left || right` - dual-aligned line (left/right side text is left/right aligned)
- `<!-- region: name -->` - open a region (resolves to the $name style block)
- `<!-- /region -->` - close the most recently opened region

## Frontmatter (extra styling)

The frontmatter is at the top of the markdown document between `---` fences at the top of the file. Any styling you want to use goes here, or you can reference an external style YAML document (or both).

```yaml
extends: base-theme.yaml
font: Calibri
body: { space_after: 2pt }
$my-notes: { size: 14 }
```

### Elements

#### Target Keys

These keys describe the target to apply the style to.

- `body` - general body/paragraph text
- `ol` - ordered-list-only settings
- `ul` - unordered-list-only settings
- `list` - shared list settings for both ordered and unordered lists
- `text` - either `body` or `list`
- `h1`-`h6` - `#`-`######`
- `heading` - any headings items (`h1`-`h6`)
- `$summary` - the region named `summary` (see [Regions](#regions))

#### Style Keys

These keys are applied to their current scope. If they aren't within a target, they apply to the entire document.

- `font` - font family (Calibri)
- `size` - point size (10 or 10pt)
- `color` - color hex ("2B579A" or "000000")
- `bold`, `italic` - true / false
- `align` - left | center | right | justify
- `line` - Either the line-spacing multiple (1.1) or the exact spacing in points (13pt)
- `space_before`, `space_after` - gap above/below this element
- `space_between` - gap between items of the same type in pts (mainly lists)
- `indent_left`, `indent_first_line`, `indent_hanging` - left/first-line/hanging indent
- `fill` - table cell background color hex (for `th` / `td`)

List keys (under `list`, `ol`, or `ul`):

- `indent_left`, `indent_hanging`, `indent_step` - list nesting indents (compiled into Word multilevel numbering)
- `levels` - per-level overrides keyed by depth (`0`, `1`, `2`, …)

Per-level keys inside `levels`:

- `format` - shorthand or template (see [List numbering formats](#list-numbering-formats))
- `num_fmt` + `template` - explicit Word numbering control (both required if either is set)
- any [style key](#style-keys) such as `color` or `font` (applied to list item text)

Settings under `list` apply to both ordered and unordered lists (like `heading` applies to all heading levels). Use `ol` or `ul` for kind-specific overrides.

List indents and numbering formats are compiled into Word multilevel list definitions so numbering and nesting behave natively when you edit the document in Word later. Theme keys still control appearance; complexity stays on our side.

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

### List numbering formats

Use `format` for shorthand or templates. The default ordered format is `1.` (`1.`, `2.`, `3.`). Use `1` in a template string when you want numbers without a trailing period.

| `format` value | Numbering | Display examples |
|----------------|-----------|------------------|
| `1` | decimal | 1, 2, 3 (via template `1` → `%1`) |
| `1.` | decimal | 1., 2., 3. |
| `a` / `alph` | lower letter | a, b, c |
| `a.` | lower letter | a., b., c. |
| `A` / `Alph` | upper letter | A, B, C |
| `(A)` | upper letter | (A), (B), (C) |
| `i` / `roman` | lower Roman | i, ii, iii |
| `roman )` | lower Roman | i ), ii ), iii ) |
| `I` / `Roman` | upper Roman | I, II, III |
| `Section 1:` | decimal | Section 1:, Section 2:, … |

For full control, set explicit Word properties:

```yaml
ol:
  levels:
    2:
      num_fmt: upperRoman
      template: "(%1)"
```

Bullet lists use the same `format` key (`•`, `-`, `*`, or `bullet`).

Spacing/indent sizes can be given in points or inches (`10` or `10pt` = 10 points, `10in` = 10 inches).

#### Special Keys

- `extends` - Path to a yaml style file (relative to the markdown file's location). Not required, but must be a valid file if present.
- `border_bottom` - If present, add a bottom border line (ie. `{ size: 0.5pt, color: "2B579A" }` to add a blue underline that's 0.5pt thick).
 - `page` - If present, can define page size and margin.

```YAML
page:
  size: letter
  margin: { top: 0.5in, bottom: 0.5in, left: 0.7in, right: 0.7in }
```

### Regions

Regions let you change style in local areas, so they take all the same keys as the global scope. You can define regions in the frontmatter by adding a key starting with `$`, and then by putting `<!-- region: REGION_NAME -->` at the start and `<!-- /region -- >` at the end of any part of the markdown you want to group together.

```markdown
---
$main: { color: "2B579A"}
---
<!-- region: main -->
# Header Text

lorum ipsum etc...
<!-- /region -->
```

### Style Priority

Styles resolve in order of scope: defaults, external (via `extends`), frontmatter (global), frontmatter (region).
Within a scope, more specific rules also get priority over less specific ones. For example, style for `h1` would take priority over rules for `heading`.

## Example

```markdown
---
extends: base-theme.yaml
page:
  size: letter
font: Calibri
size: 12
heading: { bold: true, color: "2B579A" }
h2: { size: 14, space_after: 1.5pt }
$header:
  align: center
  h2: { space_after: 0pt }
---
## Some Title
<!-- region: header -->
## Centered Title
Lorem ipsum dolor sit something something...
<!-- /region -->
```
