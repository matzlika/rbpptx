# Changelog

All notable changes to this project are documented here. The format is based
on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
