---
extends: showcase-theme.yaml
$pullquote:
  color: "943634"
---

<!-- region: header -->
# mark2word Style Gallery
A sample of themed Markdown → Word conversion
<!-- /region -->

## Document Purpose

This document demonstrates how mark2word can take a Markdown document with a little extra invisible style and make it a polished Word file. For __human work__, it can be an easy way to separate drafting your documents and polishing the style, letting you focus on each separately before you're ready to compile them together into a Word document. For __working with agents__, Markdown provides a simple yet expressive output format, and keeping the style and text separate lets the agent focus on the content of your document rather than writing another Python script or messing with XML.
Plus, since the styling only uses the front-matter, YAML, or comments, the Markdown files still render in any viewer.

### Heading Level Three

#### Heading Level Four

##### Heading Level Five

###### Heading Level Six

## Inline Markup

A plain body paragraph with **bold**, *italic*, ***bold italic***, _underscore italic_, __underscore bold__, and `inline code`. Adding links also works: [mark2word on GitHub](https://github.com/Mullans/mark2Word). You can even use `||` in the middle of a line to split it into left- and right-aligned sides.

Left-aligned Text || Right-aligned Text

<!-- region: pullquote -->
You can even add in custom styling: the same Markdown skeleton can look like a memo, a proposal, or a résumé depending on the YAML you attach.
<!-- /region -->

## Lists

**Unordered Lists**

- Markdown unordered lists
- work just like you would expect,
  - and nested bullets inherit list indentation.
  - Their styling comes from the `list` theme block.
- Making it easy to use!

**Ordered Lists**

1. Markdown ordered lists
2. also work as expected.
   1. Nested lists too!
3. But look what happens if you split the list.

A body paragraph between lists restarts numbering.

9. The first number of the list sets the start value for the ordered run,
10. and the numbering continues within the new run.

## Images

Markdown embedded images work like this:
![Color swatch](sample.png)

## Tables

Styled elements reference table:

| Element | What you write | What the theme controls |
| - | - | - |
| Headings | `#` through `######` | `heading`, `h1`–`h6`, fonts, color, borders |
| Body | plain lines | `body`, `text`, spacing, alignment |
| Lists | `- item` / `1. item` | `list`, indents, `space_between` |
| Regions | `<!-- region: name -->` | `$name` blocks with nested targets |
| Tables | pipe tables | `table`, `th`, `td`, `fill` |
| Code | fenced blocks | `code` font, size, color |

## Code Block

The delimiter below is preserved inside fenced code:

```python
# Pipes in code are not dual-alignment markers
label = "left || right"
print(label)
```

## Region/Theme Tips

<!-- region: theme-sample -->
**Theme tips**

- Quote hex colors in the YAML theme - `color: "5B5B5B"`
- Use `text` for shared body/list defaults
<!-- /region -->
<!-- region: theme-sample2 -->
- Layer `$regions` for local personality
- Override one margin without losing the rest via frontmatter
<!-- /region -->

<!-- region: theme-list-sample -->
- You can even
  - have different
    - styles for
  - each level
- of list nesting!
<!-- /region -->
<!-- region: footer -->
mark2word showcase || generated from docs/examples/showcase.md
<!-- /region -->
