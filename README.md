# docx-skill

A Claude Code plugin for generating professional Word documents.

Ask Claude to write a report, convert a markdown file, or draft a proposal. This skill handles the formatting: proper typography, cover pages, headers, footers, styled tables, and table of contents. Provide a website URL and it extracts your colours, fonts, and logo too.

## Writing quality

LLMs produce recognisable writing. Hollow antithesis ("It's not X. It's Y."), breathless transitions, significance inflation, staccato fragments for manufactured emphasis. Readers spot it immediately.

This skill requires Claude to run [stop-slop](https://github.com/hardikpandya/stop-slop) on all prose before rendering it to .docx. stop-slop catches specific AI writing patterns: throat-clearing openers, binary contrasts, dramatic fragmentation, false agency, vague declaratives. It scores text on five dimensions (directness, rhythm, trust, authenticity, density) and revises until the score passes.

The result reads like a person wrote it. The document looks like a designer formatted it.

## What it does

- Generates .docx files from any source: markdown files, scratch, conversation
- Cover pages, headers with logo, footers with page numbers, table of contents
- Discovers installed fonts and their weight variants (Thin through Black)
- Uses actual font weights for emphasis (e.g. "Avenir Next Medium") instead of faking bold
- Optionally extracts colours, fonts, and logo from a website
- Cleans AI writing patterns from prose via stop-slop before rendering

## Install

```bash
# Add as a plugin marketplace
/plugin marketplace add your-org/docx-skill
```

Or clone directly:

```bash
git clone https://github.com/your-org/docx-skill.git ~/.claude/skills/docx-skill
```

No further setup needed. The skill creates its Python environment automatically the first time Claude uses it.

### Prerequisites

- Python 3.12+
- Cairo (for SVG logo conversion): `brew install cairo`

## Usage

Ask Claude to create a Word document. The skill triggers when you mention creating, generating, or restyling `.docx` files.

```
Write a two-page executive summary of our product and generate it as a Word document
```

```
Convert docs/whitepaper.md to a formatted .docx with a cover page
```

```
Create a proposal document using our style from https://www.example.com
```

## How the library works

The Python library at `lib/` has four modules:

**`fonts.py`** scans macOS font directories and parses OpenType metadata. You ask it "what weights does Avenir Next have?" and it tells you Regular (400), Medium (500), Demi Bold (600), Bold (700).

**`brand.py`** stores document style config: colours (primary, accent, body, light background, border), a font family with chosen emphasis and heading weights, and an optional logo path. Saves to and loads from JSON.

**`builder.py`** wraps python-docx with document building methods. Sets up page size, styles, cover pages, headers (borderless table when a logo is present, tab stops when text-only), footers, and table of contents.

**`markdown_parser.py`** converts markdown to docx content. Handles headings, paragraphs, bullet lists, tables, code blocks (skipped), and horizontal rules (page breaks). Emphasis markers render using the configured emphasis font weight, not a bold flag.

## Running tests

```bash
cd lib
.venv/bin/pytest tests/ -v
```

## Project structure

```
.claude-plugin/
  plugin.json             # Plugin manifest
skills/
  docx-skill/
    SKILL.md              # Skill prompt (loaded by Claude Code)
    references/
      brand-extraction.md # Brand extraction checklist
      layouts.md          # Header/footer/cover patterns
      writing-quality.md  # stop-slop integration
lib/
  pyproject.toml          # Python dependencies
  docx_builder/           # Python library
    fonts.py              # Font discovery
    brand.py              # Style config
    builder.py            # Document builder
    markdown_parser.py    # Markdown to docx
  tests/                  # 39 tests
```

## License

MIT
