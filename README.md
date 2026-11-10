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
- Phrase replacement: `replace_phrase` changes exactly the characters of a phrase wherever it sits, even when PowerPoint
  split it across three runs, and leaves the runs either side of it untouched. `replace_paragraph_text` is the blunt
  version (one paragraph becomes one run); this is the one a find-and-replace in a deck actually needs.
- Round-trip report: `python -m ooxml_writer.roundtrip deck.pptx` saves the deck, opens the saved file again, and prints
  every part with what happened to it: identical with its size, changed with both sizes, added, or dropped. It exits
  non-zero if anything differs, so it works as a gate rather than as a reassurance.

  Archive order is reported separately, and that is not pedantry. A package whose parts are all byte-identical but written
  in a different order passes every content check while producing two archives that a person comparing them sees as
  entirely different files. The report names that case instead of hiding it behind 41 "identical" lines.

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

## Tables

`table_cells` reads every table as rows of flattened cell text, and `set_table_cell_text` writes one cell while keeping
its run properties: the new text lands in the first run and the rest are emptied rather than deleted, because deleting a
run drops the formatting the cell depends on. Everything outside the cell stays byte-identical, which the tests assert by
comparing the whole part against the original with only that one string changed.

An earlier version of this module said Word and PowerPoint store tables the same way. **They do not.** WordprocessingML
uses `<w:tbl>`/`<w:tr>`/`<w:tc>` with `<w:r><w:t>` runs; DrawingML, which every PPTX slide uses, calls the same shapes
`<a:tbl>`/`<a:tr>`/`<a:tc>` with `<a:r><a:t>`. The patterns now capture the namespace prefix and every fragment written
back reuses it, so a cell edit stays DrawingML. The old code wrote `<w:r>` into a PowerPoint table, which is exactly the
kind of file that makes PowerPoint offer to repair your presentation.

## Shapes, addressed by the name on screen

`shapes(xml)` lists what a slide contains as a person sees it: `Title 1`, `Content Placeholder 2`, `TextBox 3`, each
with its id, its paragraph count and its text one line per paragraph. `set_shape_text(package, part, name, text)` writes
one of them.

Three decisions worth stating:

1. **By name, not by index.** `replace_paragraph_text` already works by index, which is what a program has and what a
   person does not. The name in the selection pane is the thing a caller can actually state.
2. **A body of several paragraphs becomes one paragraph.** Replacing three bullets with a sentence and leaving two empty
   `<a:p>` elements behind looks broken, and an empty paragraph still draws its bullet glyph. The extra elements are
   removed rather than emptied, which costs only the run properties of paragraphs that are not being kept.
3. **A shape with no text body is an error, not a no-op.** Asking to set a picture's text is asking the wrong question,
   and `ShapeError` says so instead of silently changing nothing.

Ambiguity is refused too: two shapes that both answer to a name (or to it ignoring case) raise rather than guess, because
picking one silently is how a title gets overwritten by its subtitle.

## The fidelity scoreboard, and what it refuses to claim

`ooxml_writer.fidelity` counts every element in every part before and after an edit, by the prefix the part itself
declares, and names what went:

```bash
python -m ooxml_writer.fidelity deck.pptx
# 41 of 41 parts kept their structure, 2643 elements counted in the original.
```

After collapsing a three-bullet body into one sentence it prints `ppt/slides/slide1.xml  structure changed -2 a:p -2 a:r
-2 a:t`, which is exactly what that edit removes. Three limits are built into the report rather than hidden behind it:

1. It measures structure, not appearance. A shape moved, a colour changed, a sentence rewritten in place: identical
   census.
2. That blind spot is pinned by tests, in both directions. Rewriting a title's text changes the bytes and changes no
   element, and the scoreboard says so; the round-trip report is the one that notices.
3. Only two readings are compared, so an element that was replaced one for one reads as preserved. Counting cannot see
   identity, only name and number.

## A real deck, not a fixture written by the tests

`tests/fixtures/deck.pptx` is a two-slide deck built with python-pptx, committed as bytes so the suite needs no
PowerPoint toolchain (`tools/make-fixture.py` is the provenance record; the suite itself has no PowerPoint dependency).
It contains the three things a naive editor destroys: a phrase split across three runs, a DrawingML table, and a binary
PNG part.

The tests open it, cut its table cells, replace a phrase, edit a paragraph, and save, asserting after each edit that every
other part in the zip is byte-identical and that the edited part still parses. Independently, the result was opened with
**python-pptx**, a codebase this project shares nothing with: it reads `Hello there`, `Revenue is up 12%.`, the updated
table, and a clean zip whose only changed part is `ppt/slides/slide1.xml`.

## Status

**Stage 1: you can edit a real deck and PowerPoint's own tooling still reads it.** Package round-trip, paragraph reading,
run-level replacement, phrase-level replacement, tables in both namespaces, and shape text by name, all on a real
python-pptx deck with a per-part byte diff. The diff is now a published report: all 41 parts of that deck come back
identical, in their original archive order.

**Tests:** 72

No schema validation, no layout or geometry editing (a shape's position and size are untouched), no charts, notes or media
replacement yet, and no corpus of decks
written by PowerPoint itself (the fixture is python-pptx, which is honest about its scope: a deck saved by PowerPoint
carries extension lists and vendor parts this fixture does not).

```bash
pip install -e ".[dev]" && ruff check . && python -m pytest -q
```

## Roadmap

1. **Byte-faithful package** (done).
2. **Run-level editing** (done): replace a whole paragraph, or replace a phrase and keep the runs either side.
3. **Shape and text-box editing** (done): list a slide's shapes, edit one by name, refuse the ambiguous and the ones
   with no text body.
4. **Round-trip corpus**: real decks from real tools, saved and re-opened, with a per-part diff report.
5. **Fidelity scoreboard** (started): every part's element census before and after an edit, differences named. Not yet
   published as a table against a corpus, because the corpus (item 4) is still only the one python-pptx deck.

## Non-goals

- Not a renderer. Use the existing ones for that.
- No rebuilding an archive from a document model.
- No silent repairs: an unparseable part is reported, not rewritten into something plausible.

## Licence

MIT.
