---
name: mark2word
description: Use when converting styled Markdown, YAML-frontmatter documents, or .md files into Microsoft Word .docx with mark2word, including selecting or creating project-local themes.
---

# Mark2Word

## Overview

Mark2Word converts styled Markdown into a Word `.docx`. Styling lives in YAML frontmatter, external theme files, and named regions — the Markdown body stays readable in any viewer.

### Supported content

- Headings (`#` … `######`) with Word bookmarks for internal links
- Body paragraphs (one line = one paragraph)
- Bulleted (`-` / `*`) and ordered (`1.`) lists with **nested levels** (2 spaces per indent)
- Blockquotes (`> quoted text`, multi-line)
- Inline **bold**, *italic*, `code`, and [hyperlinks](url) (emphasis inside links works)
- Internal links: `[text](#heading-slug)`
- Dual-aligned lines: `Left || Right` (ignored inside backticks)
- Fenced code blocks (`` ```lang `` … `` ``` ``)
- Pipe tables with separator row
- Images: `![alt](path.png)` (paths relative to the markdown file)
- Horizontal rules: `---`, `***`, or `___` on their own line (body only; frontmatter stripped first)
- Named regions: `<!-- region: name -->` … `<!-- /region -->` (nested supported)
- Page breaks: `<!-- pagebreak -->` (HTML comment, invisible in MD viewers)

### Supported styling

- Global typography, spacing, alignment, borders, `indent_left` / `indent_right`
- Element targets: `body`, `text`, `blockquote`, `code`, `image`, `hr`, `heading`, `h1`–`h6`, `list`, `ol`, `ul`, `table`, `th`, `td`
- Per-level list numbering and colors via `list` / `ol` / `ul` → `levels` (including zero-padded `01.`)
- Code per-language overrides: `code.langs.{lang}` matched to fence language tags
- Image sizing: `width`, `max_width`, `align`, `alt_mode` (`doc` | `caption` | `both` | `none`)
- Table `border`, cell `padding`, `space_before` / `space_after`
- Region blocks (`$name`) with nested element overrides
- Page size (`letter`, `a4`), margins, **header/footer** with `{page}`, `{pages}`, `{title}` and `||` dual-align
- **Chained themes** via `extends` in frontmatter *and* in theme YAML files

Word-only features (frontmatter, HTML comments) do not appear in normal Markdown preview.

## Resources

- `scripts/mark2word.py` — bundled converter (`python-docx`, `pyyaml`)
- `assets/base-theme.yaml` — default base theme; pass as `--theme-dir` when documents use `extends: base-theme.yaml`
- `references/theme-design.md` — full theme authoring guide (read when creating or editing themes)
- `docs/examples/showcase.md` — feature demo (in the mark2Word repo)

## Convert A Document

Prefer the user's project environment when one exists.

**From the mark2Word repo:**

```bash
uv run mark2word --theme-dir docs/examples path/to/document.md
```

**Bundled skill script** (absolute paths):

```bash
python /path/to/mark2word/scripts/mark2word.py \
  --theme-dir /path/to/mark2word/assets \
  /path/to/project/document.md \
  /path/to/project/document.docx
```

Omit output path to write `<basename>.docx` beside the input.

**Batch:** repeat `-i`; `-o` must be a directory:

```bash
uv run mark2word -i a.md -i b.md -o ./out/
```

**Validate document** (frontmatter, theme chain, markdown parse):

```bash
uv run mark2word --check document.md
```

**Validate theme YAML** (extends chain, page size, list formats, keys):

```bash
uv run mark2word --check-theme .mark2word/themes/my-theme.yaml --theme-dir .mark2word/themes
```

## CLI Options

| Flag | Purpose |
|------|---------|
| `-i` / `--input` | Markdown input (repeat for batch) |
| `-o` / `--output` | Output `.docx` file or directory |
| `--theme-dir` | Folder for relative `extends` paths (repeatable; first match wins) |
| `--no-auto-theme-dir` | Skip auto-discovery of `.mark2word/themes` near the input |
| `--check` | Validate document without writing `.docx` |
| `--check-theme` | Validate theme YAML (including `extends` chain) |
| `--verbose` | Warn on ordered-list number mismatches |
| `-V` / `--version` | Print version |

Exit codes: `0` success, `1` error, `2` usage error.

Inputs, theme chains, and output paths are checked **before** conversion.

## Theme Resolution

```yaml
---
extends: base-theme.yaml
title: My Report
font: Calibri
$header: { align: center }
---
```

**Search order** for relative `extends`:

1. Markdown file's directory  
2. Each `--theme-dir`, in order  

**Auto-discovery:** `.mark2word/themes` near the input and in parent dirs (unless `--no-auto-theme-dir`).

**Chained inheritance** in theme files:

```text
showcase.md  →  extends: showcase-theme.yaml
showcase-theme.yaml  →  extends: showcase-base.yaml
```

**Style priority:** defaults → theme chain → frontmatter globals → region path (outer → inner).

Read `references/theme-design.md` before designing themes.

## Project Local Themes

```text
/path/to/project/.mark2word/themes/my-theme.yaml
```

```yaml
---
extends: my-theme.yaml
---
```

Copy or chain from `assets/base-theme.yaml`. Include `--theme-dir` for skill `assets/` when the chain references `base-theme.yaml`.

## Quick Markdown Reference

```markdown
---
extends: my-theme.yaml
title: My Doc
page:
  footer: "Confidential || Page {page} of {pages}"
---

<!-- region: header -->
# Title
Subtitle || contact@example.com
<!-- /region -->

## Section

See [Overview](#section) for details.

Body with **bold**, *italic*, `code`, and [link](https://example.com).

> Blockquote line one
> line two

Left side || Right side

- Bullet
  - Nested
1. Ordered
   1. Nested
2. Continues

| Col A | Col B |
| - | - |
| one | two |

![Diagram](images/diagram.png)

---

```python
print("hello")
```

<!-- pagebreak -->

## Next Section
```

**Lists:** a body paragraph between items starts a new list run (ordered numbering restarts). `--verbose` warns when markdown numbers skip (e.g. `1.` then `3.`).

**Internal links:** slug from heading text (`## My Section` → `#my-section`).

## Verification

1. Confirm `.docx` exists and is non-empty.  
2. Spot-check in Word: headings, nested lists, tables, regions, header/footer.  
3. Re-run with `--check` or `--check-theme` if validation is needed.

**Common fixes**

- Missing theme — add `--theme-dir` for skill `assets/` and project theme folder.  
- Output write errors — close the file in Word or remove read-only flag.  
- Missing images — paths relative to the markdown file, or absolute.  
- Unknown `#anchor` — slug must match an existing heading.
