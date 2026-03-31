# Common Document Layouts

## Page Sizes and Margins

**Whitepaper / Report:**
- A4, left/right 3.0cm, top 3.5cm (room for header), bottom 2.54cm
- Justified body text, 1.3 line spacing

**Letter / Proposal:**
- A4 or Letter, left/right 2.54cm, top/bottom 2.54cm

## Headers

**Logo + title (most common for branded documents):**
Uses a borderless 2-column table. Tab stops cannot vertically align inline images with text.
```python
builder.setup_header(logo=True, title="Document Title")
```

**Title only:**
Uses tab stops via `run.add_tab()`. No table needed.
```python
builder.setup_header(logo=False, title="Document Title")
```

**No header:**
```python
# Simply don't call setup_header()
```

## Footers

Footers use tab stops with `run.add_tab()` for text-only layouts.

**Page number centered, label right (whitepaper):**
```python
builder.setup_footer(center="page_number", right="Confidential")
```

**Title left, page number right (report):**
```python
builder.setup_footer(left="Document Title", right="page_number")
```

**Page X of Y (formal):**
```python
builder.setup_footer(right="page_x_of_y")
```

## Cover Pages

```python
builder.add_cover(
    title="Security Whitepaper",
    subtitle="Version 1.0 — March 2026",
    confidential=True,  # adds "Confidential" near bottom
)
```

The cover page uses the brand logo (if set), title in primary colour, accent line, and subtitle in body colour.

## Table of Contents

```python
builder.add_toc(title="Table of Contents", levels=2)
```

Inserts a TOC field. Users update it in Word with Ctrl+A then F9.

## Tab Stops vs Tables

- **Tab stops** (`run.add_tab()`): for text-only left/center/right alignment
- **Borderless tables**: when any zone contains an image (tabs can't vertically align images with text)
- Never use `add_run("\t")` — use `run.add_tab()` which generates proper `<w:tab/>` XML
