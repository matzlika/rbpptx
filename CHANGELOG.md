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
- `Rbpptx::Presentation#save(path)` writes the presentation to disk via a
  sibling temp file and atomic rename, so `pres.save(original_path)` is
  safe. Only slides whose shapes have been mutated are re-serialized;
  every other ZIP entry (themes, masters, layouts, media, fonts, rels,
  PowerPoint-extension XML) is streamed through byte-for-byte.
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
