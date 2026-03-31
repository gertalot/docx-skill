# Writing Quality Integration

The docx-builder library handles formatting. You handle writing quality. These are separate concerns.

## Before Rendering to .docx

**REQUIRED** — invoke both skills on all prose content:

1. **stop-slop** — cuts AI writing patterns (throat-clearing, binary contrasts, dramatic fragmentation, false agency, vague declaratives)
2. **elements-of-style:writing-clearly-and-concisely** — applies Strunk's rules (active voice, positive form, concrete language, omit needless words)

## When to Apply

- When the user provides markdown content: run both skills on it before rendering
- When you write new prose: apply both skills during writing, then render
- When restyling an existing document: ask the user if they want a writing quality pass too

## What to Check

From stop-slop's quick checks:
- Any "not X, it's Y" contrasts? State Y directly.
- Paragraph ends with punchy one-liner? Vary it.
- Vague declarative ("The implications are significant")? Name the specific implication.
- Any adverbs? Kill them.

Score on 5 dimensions (directness, rhythm, trust, authenticity, density). Revise below 35/50.
