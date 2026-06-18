---
name: mark2word
description: Use when converting styled Markdown, YAML-frontmatter documents, or .md files into Microsoft Word .docx with mark2word, including selecting or creating project-local themes.
---

# Mark2Word

## Overview

Mark2Word converts a styled Markdown file into a Word `.docx`. Styling lives in YAML frontmatter, external theme files, and named regions — the Markdown body stays readable in any viewer.

**Supported content**

- Headings (`#` … `######`)
- Body paragraphs (one line = one paragraph)
- Bulleted lists (`-` / `*`) and ordered lists (`1.`) with **nested levels** (2 spaces per indent)
- Inline **bold**, *italic*, `code`, and [hyperlinks](url)
- Dual-aligned lines: `Left text || Right text` (ignored inside backticks)
- Fenced code blocks (` ``` `)
- Pipe tables
- Embedded images (`![alt](path.png)`) — paths relative to the Markdown file
- Named regions: `<!-- region: name -->` … `<!-- /region -->` (nested regions supported)

**Supported styling**

- Global typography, spacing, alignment, borders
- Per-element targets: `body`, `text`, `heading`, `h1`–`h6`, `list`, `ol`, `ul`, `code`, `table`, `th`, `td`
- Per-level list numbering and colors via `list` / `ol` / `ul` → `levels`
- Region blocks (`$name`) with nested element overrides
- Page size (`letter`, `a4`) and margins
- **Chained themes** via `extends` in frontmatter *and* in theme YAML files

## Resources

- `scripts/mark2word.py` — bundled converter (needs `python-docx` and `pyyaml`)
- `assets/base-theme.yaml` — default base theme; pass its folder as `--theme-dir` when documents reference `extends: base-theme.yaml`
- `references/theme-design.md` — full theme authoring guide. Read when creating, editing, or reviewing themes.
- `docs/examples/showcase.md` — living sample of features (in the mark2Word repo)

## Convert A Document

Prefer the user's project environment when one exists.

**From the mark2Word repo** (recommended when available):

```bash
uv run mark2word --theme-dir docs/examples path/to/document.md
```

**Using the bundled skill script** (absolute paths from any working directory):

```bash
python /path/to/mark2word/scripts/mark2word.py \
  --theme-dir /path/to/mark2word/assets \
  /path/to/project/document.md \
  /path/to/project/document.docx
```

When `output` is omitted, the converter writes `<basename>.docx` beside the input.

**Multiple files** — repeat `--input` / `-i`; `--output` must be a directory:

```bash
uv run mark2word -i a.md -i b.md -o ./out/
```

**Validate without writing** (frontmatter, theme chain, Markdown parse):

```bash
uv run mark2word --check document.md
```

## CLI Options

| Flag | Purpose |
|------|---------|
| `-i` / `--input` | Markdown input (repeat for batch) |
| `-o` / `--output` | Output `.docx` file or directory |
| `--theme-dir` | Folder searched for relative `extends` paths (repeatable; first match wins) |
| `--no-auto-theme-dir` | Skip auto-discovery of `.mark2word/themes` near the input |
| `--check` | Parse and validate only; no `.docx` written |
| `--verbose` | Warn on ordered-list number mismatches |

Exit codes: `0` success, `1` conversion/theme/parse error, `2` usage error.

The converter checks inputs and outputs **before** building: missing files, unreadable themes, read-only or locked output paths produce clear errors early.

## Theme Resolution

Documents reference a theme in frontmatter:

```yaml
---
extends: base-theme.yaml
font: Calibri
$header: { align: center }
---
```

**Search order** for relative `extends` paths:

1. The Markdown file's directory
2. Each `--theme-dir` folder, in the order provided

**Auto-discovery:** unless `--no-auto-theme-dir` is set, the converter also searches `.mark2word/themes` in the input file's directory and each parent directory.

**Chained inheritance:** theme YAML files may also declare `extends`. Parent themes load first; child keys override via deep merge. Example chain (see repo `docs/examples/`):

```text
showcase.md  →  extends: showcase-theme.yaml
showcase-theme.yaml  →  extends: showcase-base.yaml
```

Put shared defaults in the base file; layer document-specific overrides in intermediate or frontmatter keys. Cycles are detected and rejected.

**Style priority** (lowest → highest): built-in defaults → external theme chain → frontmatter globals → active region path (outer → inner).

Read `references/theme-design.md` before designing or editing themes.

## Project Local Themes

When a user wants a custom look, create a project folder such as:

```text
/path/to/project/.mark2word/themes/my-theme.yaml
```

Point the document at it:

```yaml
---
extends: my-theme.yaml
---
```

Copy `assets/base-theme.yaml` as a starting point, or chain from it:

```yaml
extends: base-theme.yaml
h2: { size: 14, color: "2B579A" }
```

Always include `--theme-dir` pointing at `assets/` when the chain references `base-theme.yaml` from the skill.

## Quick Markdown Reference

```markdown
---
extends: my-theme.yaml
---

<!-- region: header -->
# Title
Subtitle || contact@example.com
<!-- /region -->

## Section

Body with **bold**, *italic*, `code`, and [a link](https://example.com).

Left side || Right side

- Bullet one
  - Nested bullet
1. Ordered one
   1. Nested ordered (restarts per level)
2. Ordered two

| Header | Value |
| - | - |
| Cell | Data |

![Diagram](images/diagram.png)

```python
# Fenced code block
print("hello")
```
```

**List notes:** a body paragraph between list items starts a new list run (ordered numbering restarts). Use `--verbose` to see warnings when markdown numbers skip (e.g. `1.` then `3.`).

## Verification

After conversion:

1. Confirm the `.docx` exists and is non-empty.
2. Open in Word and spot-check headings, nested lists, tables, and any themed regions.
3. If theme or parse issues are suspected, re-run with `--check`.

**Common fixes**

- `extends references missing theme` — add `--theme-dir` for both skill `assets/` and the project theme folder.
- Output write errors — close the target file in Word or remove read-only flag; the converter validates writability up front.
- Missing images — ensure paths are relative to the Markdown file or use absolute paths.
