# rbpptx

[![Gem Version](https://badge.fury.io/rb/rbpptx.svg?icon=si%3Arubygems)](https://badge.fury.io/rb/rbpptx)

Read-and-edit Ruby gem for `.pptx` (PowerPoint, OOXML) files, modeled after
[python-pptx](https://github.com/scanny/python-pptx).

`rbpptx` is built around a single design promise: **what we don't understand,
we don't touch**. The presentation is opened as a ZIP package, only the
parts you mutate are re-serialized, and every other entry — slide layouts,
masters, themes, embedded fonts, images, animations, hyperlink colors,
PowerPoint-extension XML — round-trips byte-for-byte. Inside a slide that
you do edit, only the `<a:p>` children of the targeted `<p:txBody>` are
rewritten; surrounding `<p:nvSpPr>`, `<p:spPr>`, `<a:bodyPr>`, `<a:lstStyle>`,
and any unknown OOXML extensions remain in place. The first existing
`<a:rPr>` (font, size, color, language) is cloned onto each new run, so
template formatting carries through to the replacement text.

Supported:

- open an existing `.pptx`, enumerate slides in presentation order
- read every shape's name, concatenated text, and placeholder type
  (`:ctrTitle`, `:subTitle`, `:body`, ...)
- walk into paragraphs and runs (`shape.paragraphs`, `shape.runs`,
  `paragraph.runs`)
- **bulldozer setter**: `shape.text = "..."` replaces the entire text body,
  cloning the first existing `<a:rPr>` onto each new run so template
  formatting carries through
- **surgical setter**: `run.text = "..."` rewrites only the run's `<a:t>`
  content, preserving `<a:rPr>`, soft line breaks, sibling runs, paragraph
  properties, and unknown OOXML extensions
- save back to a new path, or overwrite the original via temp-file + atomic
  rename — unmodified entries stream through untouched

Out of scope:

- creating a presentation from scratch — `Rbpptx.new` is reserved for a
  future release and currently raises `NotImplementedError`
- inserting / deleting / reordering / duplicating slides
- editing media (images, embedded fonts) or adding new shapes
- placeholder-aware API (`slide.title`, `slide.placeholders[:body]`)
- notes pages, comments, charts, SmartArt
- legacy `.ppt` (BIFF/CFB) input — `rbpptx` reads OOXML `.pptx` only.
  Convert first, e.g. `libreoffice --headless --convert-to pptx file.ppt`.
  `Rbpptx.open` detects the OLE compound-file magic and raises
  `Rbpptx::UnsupportedFormatError` with the conversion hint rather than
  surfacing an opaque ZIP parse error from rubyzip.

## Install

```bash
gem install rbpptx
```

Or in a `Gemfile`:

```ruby
gem "rbpptx"
```

Requires Ruby 3.2+, `nokogiri`, and `rubyzip`.

## Usage

### Reading

```ruby
require "rbpptx"

Rbpptx.open("deck.pptx") do |pres|
  pres.slides.each do |slide|
    puts slide.name                        # => "slide1", "slide2", ...
    slide.shapes.each do |shape|
      puts "  #{shape.name}: #{shape.text.inspect}"
    end
  end
end
```

`Rbpptx.open` yields a `Rbpptx::Presentation` and auto-closes the underlying
ZIP when the block returns (or raises). Without a block it returns the
presentation; call `pres.close` when done.

Slides are enumerated in the order they appear inside the presentation —
i.e. the order of `<p:sldId>` entries in `ppt/presentation.xml`, which is
not necessarily the slide-part filename order.

### Editing text and saving

The bulldozer setter replaces a shape's entire text body. Lines (separated
by `"\n"`) become individual `<a:p>` paragraphs:

```ruby
require "rbpptx"

Rbpptx.open("template.pptx") do |pres|
  pres.slides.first.shapes.first.text = "差し替えタイトル"
  pres.slides[1].shapes.find { |s| s.text.include?("会社名") }.text = "Acme Inc."
  pres.save("filled.pptx")
end
```

The surgical setter rewrites only the targeted run's `<a:t>`, preserving
all other formatting and structural XML. Use it for `{{slot}}`-style
template fills where a marker run sits inside a multi-run paragraph:

```ruby
Rbpptx.open("template.pptx") do |pres|
  pres.slides.each do |slide|
    slide.shapes.flat_map(&:runs).each do |run|
      run.text = run.text.gsub("{{name}}", "Acme")
    end
  end
  pres.save("filled.pptx")
end
```

Or target by placeholder role:

```ruby
Rbpptx.open("template.pptx") do |pres|
  pres.slides.first.shapes.find { |s| s.placeholder_type == :ctrTitle }.text = "Q2 Review"
  pres.slides.first.shapes.find { |s| s.placeholder_type == :subTitle }.text = "Acme Inc."
  pres.save("filled.pptx")
end
```

### Same-path overwrite

`Presentation#save` writes to a temp file in the target directory and
atomically renames it into place, so `pres.save(original_path)` is safe:

```ruby
Rbpptx.open("deck.pptx") do |pres|
  pres.slides.first.shapes.first.text = "Updated"
  pres.save("deck.pptx")    # overwrite in place
end
```

If the rename fails, the temp file is removed before the error propagates.

### Inspecting without parsing every slide

Slide XML is parsed lazily: enumerating slides, reading their `name`, or
checking `dirty?` does not pay the per-slide parse cost. The first call to
`slide.shapes` (or `slide.to_xml`) is what triggers the parse, and the
result is cached.

```ruby
Rbpptx.open("deck.pptx") do |pres|
  pres.slides.length                       # cheap
  pres.slides.map(&:name)                  # cheap
  pres.slides.first.shapes.first.text      # parses only slide 1
end
```

### Error handling

All `rbpptx` exceptions inherit from `Rbpptx::Error` and the messages carry
the presentation path and (where relevant) the offending part name:

| Error                              | Raised when                                                       |
| ---                                | ---                                                               |
| `Rbpptx::UnsupportedFormatError`   | file is empty, not a ZIP, or a legacy `.ppt` (BIFF/CFB) container |
| `Rbpptx::PresentationFormatError`  | `_rels/.rels`, `ppt/presentation.xml`, or its rels are malformed  |
| `Rbpptx::SlideFormatError`         | a slide part is missing or its XML cannot be parsed               |
| `Rbpptx::SlideNotFoundError`       | reserved for future index-based slide accessors                   |
| `Rbpptx::ClosedPresentationError`  | `slides` / `save` called after `close`                            |

```ruby
begin
  Rbpptx.open("deck.pptx") { |pres| pres.slides.each { |s| process(s) } }
rescue Rbpptx::UnsupportedFormatError => e
  warn "not a pptx: #{e.message}"
rescue Rbpptx::PresentationFormatError, Rbpptx::SlideFormatError => e
  warn "malformed: #{e.message}"
end
```

## Design Notes

- **Fidelity over feature coverage.** `rbpptx` is a small, surgical editor,
  not a PowerPoint object model. The bytes it doesn't understand are the
  bytes it doesn't touch — including PowerPoint-version-specific extensions
  (`p14:`, `p15:`, `ahyp:`, `mc:AlternateContent`) that a full object model
  would have to know about to round-trip safely.
- **Lazy parse, eager write.** Slides parse on first access; saves
  re-serialize only the slides whose shapes have been mutated, and stream
  every other ZIP entry through with its already-deflated bytes
  (`Zip::OutputStream#copy_raw_entry`) — no decompress, no re-deflate.
  Bit-identical preservation, not just content-identical.
- **rPr cloning, not rebuilding.** When a shape's text is replaced, the
  first existing `<a:rPr>` is deep-cloned onto each new run. A template's
  font, size, color, and `lang` attribute carry through without `rbpptx`
  needing to model those properties.
- **Atomic same-path save.** `save` writes to a sibling temp file and
  renames it into place, so a crash mid-write cannot leave a half-written
  `.pptx` at the target path.

## Benchmarks

The selling-point story for `rbpptx` is on real-world decks where the
package carries embedded media — themes, slide masters, slide layouts,
images, fonts. `python-pptx` parses these eagerly at `Presentation(path)`
time; `rbpptx` defers everything past `presentation.xml` and its rels, so
the open-and-read path stays cheap regardless of how much binary mass the
deck carries.

Real-world deck (18 MB, 21 slides, embedded fonts and images), 5 iterations:

| benchmark             | rbpptx (s) | python-pptx (s) | speedup |
| ---                   | ---:       | ---:            | ---:    |
| read (walk all text)  | **0.019**  | 0.089           | **4.6x faster** |
| fill_one (1 + save)   | **0.032**  | 0.553           | **17x faster**  |
| fill_many (10 + save) | **0.042**  | 0.556           | **13x faster**  |

Memory (RSS delta) on the read scenario:

| library      | rss Δ (kB) |
| ---          | ---:       |
| rbpptx       | **3,232**  |
| python-pptx  | 28,416     |

The fill scenarios used to tie when both libraries re-deflated every
ZIP entry on save. `rbpptx` now streams unmodified entries through with
`Zip::OutputStream#copy_raw_entry`, so the deflate cost is paid only on
the slide XML that actually changed — which is what unlocks the 13–17x
gap on read-modify-write. The advantage is real and reproducible:

```bash
bundle exec ruby benchmark/compare.rb
```

`uv` is required for the python-pptx subprocess; the comparison script
invokes `uv run --with python-pptx python3 benchmark/python_pptx_compare.py`
so no system-level Python install is mutated.

To benchmark against a synthetic deck with no embedded media (isolates the
XML-rewrite cost from binary deflate), generate a fixture first:

```bash
uv run --with python-pptx python3 benchmark/fixture_gen.py 50 samples/bench_50.pptx
RBPPTX_BENCH_SAMPLE=samples/bench_50.pptx bundle exec ruby benchmark/compare.rb
```

## Development

Development in this repository assumes Ruby 3.4.8 (`.ruby-version`).

```bash
bundle install

# Run tests
bundle exec ruby -Ilib -Itest test/rbpptx_test.rb

# Generate API docs
bundle exec rake rdoc
```

The test suite uses a real `.pptx` placed at
`samples/プロトタイピング プレゼンテーション.pptx` (gitignored). Sample-dependent
tests skip with a clear message when the file is absent, so the suite still
runs cleanly in a fresh checkout.
