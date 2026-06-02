# mark2Word

This is my quick tool to convert markdown files to Word documents.

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
- `list` - bulleted/numbered lists
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
