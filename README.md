# docx-skill

A Claude Code plugin that generates professional Word documents.

Tell Claude to write a report, convert a markdown file, or build a proposal as a .docx. Claude adds cover pages, headers, footers, styled tables, and a table of contents. Point Claude at a website and it pulls colours, fonts, and a logo for branding.

## Writing quality

LLM-generated prose follows recognisable patterns: hollow antithesis ("It's not X. It's Y."), significance inflation, staccato fragments. A well-formatted document with AI-sounding prose still looks AI-generated.

This plugin integrates [stop-slop](https://github.com/hardikpandya/stop-slop) to clean prose before generating the .docx. stop-slop bans specific phrases and structures, then scores text on five dimensions (directness, rhythm, trust, authenticity, density). Claude rewrites until the score passes 35/50.

## Features

- Generates .docx from markdown files, conversation, or from scratch
- Cover pages, headers with logo, footers with page numbers, table of contents
- Scans installed fonts and their weight variants (Thin through Black)
- Uses font weight variants for emphasis (e.g. "Avenir Next Medium") instead of the bold flag
- Pulls colours, fonts, and logo from a website for branding (optional)

## Install

```bash
/plugin marketplace add your-org/docx-skill
```

Or clone:

```bash
git clone https://github.com/your-org/docx-skill.git ~/.claude/skills/docx-skill
```

Claude creates the Python environment on first use. No manual setup.

### Prerequisites

- Python 3.12+
- Cairo (for SVG logo conversion): `brew install cairo`

## Usage

The skill activates when you ask Claude to create, generate, or restyle a `.docx` file.

```
Write a two-page executive summary of our product and generate it as a Word document
```

```
Convert docs/whitepaper.md to a formatted .docx with a cover page
```

```
Create a proposal using our style from https://www.example.com
```

## Library

Four Python modules at `lib/`:

**`fonts.py`** lists weight variants for any installed font family. Avenir Next, for example, has Regular (400), Medium (500), Demi Bold (600), Bold (700). Claude picks appropriate weights for headings and emphasis based on what the font offers.

**`brand.py`** stores document style: five colours (primary, accent, body, light background, border), a font family with emphasis and heading weights, and an optional logo path. Persists to JSON for reuse across sessions.

**`builder.py`** wraps python-docx. Configures page size, styles, cover pages, headers (borderless table when a logo is present, tab stops for text-only), footers, and table of contents.

**`markdown_parser.py`** converts markdown to docx content: headings, paragraphs, bullet lists, tables, code blocks (skipped), horizontal rules (page breaks). Emphasis markers use the configured font weight, not a bold flag.

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
