# Theme Design

Read this only when creating, editing, reviewing, or substantially customizing a Mark2Word theme.

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

## Starting Point

Use `assets/base-theme.yaml` as the baseline. For a custom theme, copy the base theme into the project theme folder, rename it, and edit the copy:

```text
.mark2word/themes/executive-resume.yaml
```

Then point the Markdown document at it:

```yaml
---
extends: executive-resume.yaml
---
```

The converter resolves `extends` from the Markdown frontmatter. Do not rely on nested `extends` inside a theme file.

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
list: { space_before: 2pt, space_between: 2pt, space_after: 6pt, indent_left: 9pt, indent_hanging: 9pt }

heading: { bold: true, color: "2B579A" }
h1: { size: 14, color: "000000" }
h2: { size: 12, space_before: 12pt, space_after: 4pt, border_bottom: { size: 0.5pt, color: "2B579A" } }
```

## Target Keys

Use these target blocks:

- `body`: ordinary paragraphs.
- `list`: bulleted and numbered list items.
- `text`: shared defaults for `body` and `list`.
- `h1` through `h6`: heading levels.
- `heading`: shared defaults for all headings.
- `$name`: a named region matched by `<!-- region: name -->`.

More specific targets override broader targets. For example, `h2` overrides `heading`, and `body` overrides `text`.

## Style Keys

Supported style properties:

- `font`: font family, such as `Calibri` or `Aptos`.
- `size`: point size, such as `10` or `10pt`.
- `color`: hex color, such as `"2B579A"` or `"#2B579A"`.
- `bold`, `italic`: `true` or `false`.
- `align`: `left`, `center`, `right`, or `justify`.
- `line`: line-spacing multiple such as `1.1`, or exact point spacing such as `13pt`.
- `space_before`, `space_after`: paragraph spacing.
- `space_between`: spacing between adjacent list items.
- `indent_left`, `indent_first_line`, `indent_hanging`: paragraph indents.
- `border_bottom`: bottom paragraph border, such as `{ size: 0.5pt, color: "2B579A" }`.

Lengths can be bare points (`10`), explicit points (`10pt`), or inches (`0.5in`).

## Page Settings

Only `letter` page size is currently supported:

```yaml
page:
  size: letter
  margin:
    top: 0.5in
    bottom: 0.5in
    left: 0.7in
    right: 0.7in
```

## Regions

Regions provide local style overrides. Define the region in the theme:

```yaml
$header:
  align: center
  h1: { space_after: 1.5pt }
  text: { size: 9, space_after: 0.75pt }
```

Use it in Markdown:

```markdown
<!-- region: header -->
# Greg Davies
Chicago, IL || greg@example.com
<!-- /region -->
```

Nested regions are supported. Inner regions override outer regions.

## Design Workflow

1. Identify the document type: resume, report, proposal, notes, or another format.
2. Copy `assets/base-theme.yaml` into `.mark2word/themes/<theme-name>.yaml`.
3. Adjust global typography first: `font`, `size`, `color`, `text.line`, and page margins.
4. Tune headings and list spacing next.
5. Add `$region` blocks only for repeated local treatments, such as resume headers or callout sections.
6. Convert a representative document and inspect the output `.docx`.

## Common Mistakes

- Missing `--theme-dir`: if `extends` cannot be found, include both the skill `assets` folder and the project theme folder.
- Overusing regions: prefer element styles until a local block is genuinely needed.
- Forgetting quotes around colors: quote hex values to keep YAML from treating them as numbers.
- Expecting nested theme inheritance: put `extends` in the Markdown frontmatter, not inside another theme.
