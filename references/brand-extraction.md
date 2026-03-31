# Brand Extraction Checklist

Follow these steps when extracting brand from a website.

## 1. Fetch the Website

Use WebFetch on the homepage. Also fetch any linked CSS files.

## 2. Extract Colours

Look for, in order of reliability:
- CSS custom properties (`--primary-color`, `--brand-color`)
- Inline styles on key elements (header, hero, CTAs)
- Linked stylesheet colour values

Identify these roles:
- **Primary** — most prominent dark/brand colour (headings, covers, CTA backgrounds)
- **Accent** — highlight colour (underlines, hover states, accent elements)
- **Body** — text colour for paragraphs (usually a dark grey, not black)
- **Light background** — alternating table rows, section backgrounds (very light grey/blue)
- **Border** — dividers, card borders (light grey)

## 3. Extract Fonts

Check for:
- Google Fonts `<link>` tags (extract family names and weights loaded)
- `@font-face` declarations in CSS
- `font-family` rules on `body`, headings, and buttons

Note which weights are loaded (300, 400, 500, 600, 700, etc.).

## 4. Check Font Availability

```python
from docx_builder.fonts import FontDiscovery

family = FontDiscovery.get_family("Font Name From Website")
if family:
    print(f"Installed: {family}")
else:
    print("Not installed. Available alternatives:")
    for f in FontDiscovery.list_families():
        print(f"  {f}")
```

If the website font is not installed, present alternatives to the user. Good system font substitutes for common web fonts:
- Nunito Sans / Inter → Avenir Next, Helvetica Neue
- Open Sans / Lato → Helvetica Neue, Gill Sans
- Roboto / Source Sans → Avenir Next, Futura
- Playfair Display / Merriweather → Georgia, Charter

## 5. Find the Logo

Look in this order:
- `<header>` / `<nav>` — `<img>` or `<svg>` elements
- `<link rel="icon">` or `<link rel="apple-touch-icon">`
- `<meta property="og:image">`

If the site has light and dark logo variants (common), download the **dark-on-white** variant for documents with white page backgrounds.

Prefer SVG over PNG. The library converts SVG to PNG automatically.

## 6. Create and Save Brand Config

```python
from pathlib import Path
from docx_builder import Brand

brand = Brand(
    primary="#1C2541",
    accent="#6FFFE9",
    body="#3E465E",
    light_bg="#F2F5F8",
    border="#E2E8F0",
    font_family="Avenir Next",
    emphasis_weight=500,
    heading_weight=700,
    logo_path=Path("path/to/logo.svg"),
)
brand.save(Path("brand.json"))
```

Save `brand.json` in the project directory so it can be reused in future sessions.
