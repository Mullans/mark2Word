---
name: mark2word
description: Use when converting styled Markdown, YAML-frontmatter documents, or .md files into Microsoft Word .docx with mark2word, including selecting or creating project-local themes.
---

# Mark2Word

## Overview

Use the bundled converter to turn a styled Markdown file into a Word `.docx`. The converter supports YAML frontmatter, external YAML themes, headings, lists, inline emphasis, dual-aligned `left || right` lines, and named regions.

## Resources

- `scripts/mark2word.py`: the converter. It can run directly with Python.
- `assets/base-theme.yaml`: the default base theme. Always include this folder as a theme directory when documents use `extends: base-theme.yaml`.
- `references/theme-design.md`: detailed theme authoring guidance. Read it only when the user wants to create, edit, review, or substantially customize a theme.

## Convert A Document

Use a Python environment with `python-docx` and `pyyaml` installed. Prefer the user's project environment if one exists.

Run from any working directory with absolute paths:

```bash
python /path/to/mark2word/scripts/mark2word.py \
  --theme-dir /path/to/mark2word/assets \
  /path/to/project/document.md \
  /path/to/project/document.docx
```

If the project has local themes, include that folder too:

```bash
python /path/to/mark2word/scripts/mark2word.py \
  --theme-dir /path/to/mark2word/assets \
  --theme-dir /path/to/project/.mark2word/themes \
  /path/to/project/document.md
```

When `output_file` is omitted, the converter writes a `.docx` beside the input Markdown.

## Theme Resolution

Frontmatter can reference a theme:

```yaml
---
extends: base-theme.yaml
---
```

Relative `extends` paths are searched in this order:

1. The Markdown file's directory.
2. Each `--theme-dir` folder, in the order provided.

Use this to keep `base-theme.yaml` inside the skill while letting project documents still say `extends: base-theme.yaml`.

## Project Local Themes

When a user asks for a custom theme, create or update a project-local folder such as:

```text
/path/to/project/.mark2word/themes/
```

Read `references/theme-design.md` before designing or editing that theme. Keep normal conversion tasks on this page; do not load the theme reference unless theme work is requested.

## Verification

After conversion, confirm the `.docx` exists and is non-empty.