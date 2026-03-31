---
name: docx-builder
description: Use when creating branded Word documents from markdown, restyling existing documents, or generating .docx files with custom fonts, colours, and layouts
---

# docx-builder

Generate professionally branded Word documents from markdown using the `docx_builder` Python library.

## Overview

The library is at `{skill_dir}/lib/`. Run scripts via `{skill_dir}/lib/.venv/bin/python`.

Read the library source to understand the API — it is self-documenting. The key classes are `Brand`, `DocBuilder`, `FontDiscovery`, and the `render_markdown` function.

## Before Generating: Writing Quality

**REQUIRED:** Before rendering any prose to .docx, invoke these skills on the content:

1. **stop-slop** — eliminate AI writing patterns
2. **elements-of-style:writing-clearly-and-concisely** — apply Strunk's rules

The library handles formatting. You handle writing quality. Do not skip this step.

## Brand Extraction

When the user provides a website URL for branding, follow [references/brand-extraction.md](references/brand-extraction.md).

Key points the checklist covers:
- Extract colours from CSS (primary, accent, body text, backgrounds, borders)
- Identify fonts from Google Fonts links, @font-face, CSS font-family rules
- **Check font availability** with `FontDiscovery.get_family(name)` before using any font
- If the website font is not installed, show the user what IS installed and pick together
- Download the logo (check for light/dark variants — use the one suited for white backgrounds)
- Save the Brand config as `brand.json` in the project for reuse

## Font Weights

Never use `run.bold = True`. The library uses actual font weight variants.

When creating a Brand, pick `emphasis_weight` and `heading_weight` based on what the font family offers. Use `FontDiscovery.get_family(name)` to see available weights:

```python
from docx_builder.fonts import FontDiscovery
family = FontDiscovery.get_family("Avenir Next")
print(family)  # shows all available weights
```

Common choices:
- `emphasis_weight=500` (Medium) for inline emphasis like status markers
- `heading_weight=700` (Bold) for headings

## Document Layout

See [references/layouts.md](references/layouts.md) for common professional patterns.

Quick defaults for a whitepaper:
- A4, margins: left/right 3.0cm, top 3.5cm, bottom 2.54cm
- Header: logo left, title right (table layout)
- Footer: page number centered, "Confidential" right (tab stops)
- Cover: logo, title, accent line, subtitle, confidential notice
- Justified body text, British English spell-check

## Common Mistakes

- Skipping stop-slop on prose content
- Using the website's font without checking if it's installed locally
- Using `run.bold = True` instead of font weight variants
- Forgetting to pick the correct logo variant (dark-on-white for white page backgrounds)
- Not setting custom margins (defaults are too narrow for professional whitepapers)
