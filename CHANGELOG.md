# Changelog

All notable changes to this project are documented here. The format is based
on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] - 2026-11-10

### Added

- Placeholder-aware API: `Slide#title` (the `:title` / `:ctrTitle`
  placeholder), `Slide#placeholders` returning `Rbpptx::Placeholders`
  (`placeholders[:body]` by type, `placeholders[1]` by `<p:ph idx>`),
  `Shape#placeholder?` and `Shape#placeholder_idx`.
- Speaker notes: `Slide#notes` / `Slide#notes=`, `Slide#notes_slide`
  returning `Rbpptx::NotesSlide`. Setting notes on a slide without a notes
  page creates one from the notes master; raises `Rbpptx::Error` when the
  presentation has no notes master.
- Group shapes and tables: `Slide#groups` (`Rbpptx::GroupShape`, nestable),
  `Slide#all_shapes` (every `<p:sp>`, descending into groups),
  `Slide#tables` / `Slide#all_tables` (`Rbpptx::Table`, `TableRow`,
  `TableCell`). Table cells support the same `text` / `text=` /
  `paragraphs` / `runs` API as shapes, plus `merge_origin?` / `spanned?`.
- Slide structure edits: `Presentation#move_slide`,
  `Presentation#delete_slide` (removes the notes page, section and
  custom-show entries, relationships, and content-type overrides; refuses
  when another part still references the slide unless `unlink: true` is
  given, which strips "jump to slide" hyperlinks first), and
  `Presentation#slide(index)` raising `Rbpptx::SlideNotFoundError`.

### Changed

- Allow `rubyzip` 3.x (`>= 2.3, < 4`). Development dependencies now allow
  `rbs` 4.x and `rdoc` 8.x.
- Re-serialized parts are written without pretty-printing, so edits no
  longer re-indent single-line XML.
- `text=` (shapes and table cells) keeps template formatting on empty
  fields: when the body has no runs, the first `<a:endParaRPr>` is used as
  the run template, and the first paragraph's `<a:pPr>` (alignment,
  indent) and `<a:endParaRPr>` are cloned onto every new paragraph.
  Previously, filling an empty form field fell back to PowerPoint's default
  18pt text and lost the field's alignment.

### Fixed

- Saving a package whose ZIP entries use data descriptors (general-purpose
  bit 3, as written by Google Slides exports and other streaming zippers)
  produced a file PowerPoint refused to open, even for a no-op save. The
  flag is now cleared on raw-copied entries, whose local headers carry the
  real CRC and sizes.
- Saving twice from the same `Presentation` lost the edits made before the
  first save; saving over the source file broke later lazy reads of
  not-yet-parsed slides.

## [0.1.0] - 2026-04-26

- Initial public release.

### Added

- `Rbpptx.open(path)` opens a `.pptx` package and returns a
  `Rbpptx::Presentation`. Block form auto-closes the underlying ZIP on
  return or exception, matching the `File.open` / `Zip::File.open` idiom.
- `Rbpptx::Presentation#slides` enumerates slides in presentation order
  (i.e. the order of `<p:sldId>` entries in `ppt/presentation.xml`,
  resolved via the presentation's own rels), parsing slide XML lazily.
- `Rbpptx::Slide#shapes` returns each `<p:sp>` child of the slide's
  `<p:spTree>` wrapped as `Rbpptx::Shape`, with its own dirty tracking so
  the slide knows whether its XML needs to be re-serialized on save.
- `Rbpptx::Shape#text` returns the shape's concatenated text body —
  paragraphs joined with `"\n"`, soft line breaks (`<a:br>`) also rendered
  as `"\n"`.
- `Rbpptx::Shape#text=` replaces the shape's text body. Lines (separated by
  `"\n"`) become individual `<a:p>` paragraphs. Run properties from the
  first existing `<a:rPr>` (font, size, color, language) are cloned onto
  each new run so template formatting carries through. Shapes without a
  `<p:txBody>` raise `Rbpptx::Error` rather than silently failing.
- `Rbpptx::Shape#placeholder_type` returns the placeholder role as a
  Symbol (`:ctrTitle`, `:subTitle`, `:title`, `:body`, `:obj`, `:ftr`,
  `:dt`, `:sldNum`, `:pic`, ...) by reading
  `<p:nvSpPr>/<p:nvPr>/<p:ph type="...">`. Returns `:body` when the
  placeholder element is present without a `type` attribute (matches the
  OOXML default), and `nil` for non-placeholder shapes.
- `Rbpptx::Shape#paragraphs` and `Rbpptx::Shape#runs` expose the OOXML
  paragraph and run granularity below the shape: `paragraphs` returns the
  shape's `<a:p>` children; `runs` flattens runs across paragraphs in
  document order.
- `Rbpptx::Paragraph` and `Rbpptx::Run` wrappers; `Run#text=` is a
  surgical text setter that rewrites only the run's `<a:t>` content.
  `<a:rPr>`, `<a:br>` soft breaks, sibling runs, paragraph properties,
  `<a:endParaRPr>`, and unknown OOXML extensions on the shape are all
  preserved. Use this instead of `Shape#text=` when the template's
  per-run formatting matters (e.g. replacing a `{{slot}}` marker inside a
  multi-run paragraph).
- `Rbpptx::Presentation#save(path)` writes the presentation to disk via a
  sibling temp file and atomic rename, so `pres.save(original_path)` is
  safe. Only slides whose shapes have been mutated are re-serialized;
  every other ZIP entry (themes, masters, layouts, media, fonts, rels,
  PowerPoint-extension XML) streams through with its already-deflated
  bytes via `Zip::OutputStream#copy_raw_entry` — no inflate, no
  re-deflate. The bytes of unmodified parts are bit-identical to the
  source ZIP.
- `Rbpptx::UnsupportedFormatError` raised by `Rbpptx.open` when the file
  is empty, not a ZIP container, or a legacy `.ppt` (BIFF/CFB) file.
  Legacy inputs are detected by the OLE compound-file magic and reported
  with a `libreoffice --headless --convert-to pptx` hint rather than
  surfacing an opaque `Zip::Error` from rubyzip.
- `Rbpptx::PresentationFormatError`, `Rbpptx::SlideFormatError`,
  `Rbpptx::SlideNotFoundError`, `Rbpptx::ClosedPresentationError` for
  malformed-input and use-after-close cases. All error messages carry the
  presentation path and (where relevant) the offending part name.
- RBS signatures for the public surface (`sig/rbpptx.rbs`).
- `benchmark/compare.rb` drives a three-way comparison against
  python-pptx (via `uv run --with python-pptx`) and docxtemplater +
  PizZip (via `node`). Three scenarios — read (open + walk all slides +
  concatenate text), fill_one (replace one shape + save), fill_many
  (replace 10 shapes + save) — configurable via `RBPPTX_BENCH_SAMPLE`,
  `RBPPTX_BENCH_WARMUP`, `RBPPTX_BENCH_ITERATIONS`. On the 18 MB sample
  deck rbpptx is 4.7x / 16.8x / 13.0x faster than python-pptx and
  2.0x / 4.5x / 3.4x faster than docxtemplater across the three
  scenarios.
- `benchmark/fixture_gen.py` synthesizes a sized deck with `{{slot_*}}`
  markers via python-pptx, so save-path benchmarks can isolate the cost
  of the XML rewrite from the cost of re-deflating embedded media.
