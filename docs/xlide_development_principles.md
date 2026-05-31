# XLIDE Development Principles

## Deterministic Logic Only

All code and analyzer logic must be deterministic. The same inputs must produce
the same outputs for tokens, parse trees, symbols, completions, diagnostics,
workbook operations, and generated metadata.

Do not implement heuristic, fuzzy, probability-based, or "looks like" behavior
unless the operator explicitly and unwaveringly agrees before implementation.

When behavior is uncertain:

- Prefer no diagnostic over a guessed diagnostic.
- Mark unresolved grammar or API behavior as pending instead of inventing it.
- Use explicit parsers, curated metadata, or documented rules.
- Keep provenance clear for native/runtime/host metadata.
- Treat ambiguous resolution as unknown, not as a best guess.

## Shoot For The Moon

For complicated work, create or update a roadmap Markdown file first. The
roadmap should name the intended complete system, phases, risks, and definition
of done.

Implementation should then proceed iteratively through production-quality
vertical slices. Each slice should fit the larger roadmap and avoid throwaway
MVP shortcuts or half measures.
