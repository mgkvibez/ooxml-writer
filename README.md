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

## What works today

- Package: open a zip into named parts, read and replace one part, save, and diff two packages by part name. Every part
  you did not touch is copied byte-for-byte, and a test proves it: a binary PNG and `[Content_Types].xml` survive an edit
  unchanged.
- Paragraphs: read a slide's text the way a person sees it, not the way the file stores it. `paragraphs(xml)` flattens
  every `<a:r>` run into the paragraph's readable text while keeping each run's span; `find_paragraph` finds a phrase
  split across runs; `replace_paragraph_text` and `set_first_matching_paragraph` change the text, keep the first run's
  formatting, and empty the runs after it.

### The trap this project exists for

PowerPoint does not store `Hello world` as one string. It stores `Hello ` and `wor` and `ld` in three runs, because each
run carries its own formatting. So `xml.replace("Hello world", ...)` finds nothing on a real slide, and a naive reader
concludes the text is not there. Both operations must work on the paragraph's concatenated text while the write happens
at run granularity. The tests pin this literally: `assert "Hello world" not in SLIDE` passes, and `find_paragraph` still
finds it.

Text is escaped on the way in (`<`, `>`, `&`), so an edit cannot corrupt the part, and padded text gets
`xml:space="preserve"`. That detail shipped a real bug here: the writer added the attribute and the reader's regex no
longer matched its own output, so the text read back empty. The padded-text test caught it; the reader now accepts
attributes on `<a:t>`.

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
