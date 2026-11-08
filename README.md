# ooxml-writer

**Edit OOXML instead of only rendering it.**

Renderers exist now: pptx-renderer, ooxml-pages, OpenOOXML. None of them closes the loop. You can draw a slide, you
cannot change one and save a PPTX that PowerPoint opens unchanged.

The gap is the **write path**, and it is not a rendering problem. An OOXML file is a zip of parts joined by
relationships, and the usual approach (parse into a model, rebuild from the model) loses everything the model did not
represent: theme variants, extension lists, vendor-specific parts, the exact ordering a diff tool notices.

## The approach here

The package is the source of truth.

- Parts are byte arrays until something asks to parse one.
- Only edited parts are re-serialised. Everything else is copied verbatim.
- Part order is preserved, because readers do not care but humans diffing two archives do.
- The text surgery on runs is string-surgical on purpose: a full XML round-trip reorders namespace declarations and
  rewrites self-closing tags, which is precisely the fidelity loss this project exists to avoid.

The property that makes it honest is testable in one line: **open a file, save it without edits, and every part must be
byte-identical.** That test is in this repository.

## Status

**Stage 0: scaffolding.** `OoxmlPackage` opens a package, exposes parts, replaces one, and saves with untouched parts
byte-identical. `replace_text_run` edits the nth `<a:t>` run of a slide and leaves the rest of the XML exactly as it was.
Seven tests: round-trip identity, order preservation, targeted run replacement, missing-part refusal, malformed package
refusal, and `diff` naming changed parts.

No schema validation, no shape or layout editing, no charts or notes yet.

```bash
pip install -e ".[dev]" && ruff check . && python -m pytest -q
```

## Roadmap

1. **Byte-faithful package** (done).
2. **Run-level editing**: replace, split, merge runs; keep `<a:rPr>` formatting attached to the right text.
3. **Shape and text-box editing** without touching unrelated XML.
4. **Round-trip corpus**: real decks from real tools, saved and re-opened, with a per-part diff report.
5. **Fidelity scoreboard**: elements preserved vs lost, named individually, published.

## Non-goals

- Not a renderer. Use the existing ones for that.
- No rebuilding an archive from a document model.
- No silent repairs: an unparseable part is reported, not rewritten into something plausible.

## Licence

MIT.
