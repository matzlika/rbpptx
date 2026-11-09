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
- placeholder-aware access: `slide.title`, `slide.placeholders[:body]`,
  `slide.placeholders[1]` (by `idx`)
- speaker notes: read and write `slide.notes`; a notes page is created from
  the notes master when the slide has none
- text inside group shapes (nested to any depth) and table cells:
  `slide.groups`, `slide.all_shapes`, `slide.tables`, `table.cell(r, c)`
- reorder and delete slides: `pres.move_slide`, `pres.delete_slide`
- save back to a new path, or overwrite the original via temp-file + atomic
  rename — unmodified entries stream through untouched

Out of scope:

- creating a presentation from scratch — `Rbpptx.new` is reserved for a
  future release and currently raises `NotImplementedError`
- inserting / duplicating slides
- editing media (images, embedded fonts) or adding new shapes
- adding / removing table rows and columns, table styling
- comments, charts, SmartArt
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

Requires Ruby 3.2+, `nokogiri`, and `rubyzip` (2.3+ or 3.x).

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

### Placeholders

Target shapes by placeholder role instead of by position:

```ruby
Rbpptx.open("template.pptx") do |pres|
  slide = pres.slides.first
  slide.title.text = "Q2 Review"                    # :title or :ctrTitle
  slide.placeholders[:subTitle].text = "Acme Inc."  # Symbol: by type
  slide.placeholders[1].text = "Acme Inc."          # Integer: by <p:ph idx>
  slide.placeholders.map(&:placeholder_type)        # => [:ctrTitle, :subTitle]
  pres.save("filled.pptx")
end
```

Integer keys follow python-pptx and match the placeholder's `idx`
attribute, not its position in the list.

### Speaker notes

```ruby
Rbpptx.open("deck.pptx") do |pres|
  pres.slides.each { |s| puts s.notes }   # "" when the slide has no notes page
  pres.slides[0].notes = "Talk about Q2.\nThen Q3."
  pres.save("deck.pptx")
end
```

Setting notes on a slide without a notes page creates one from the
presentation's notes master (new part, relationships, and content-type
override). Decks that have never had notes may lack a notes master; in that
case `notes=` raises `Rbpptx::Error` — add a note in PowerPoint once to
create it.

### Groups and tables

`slide.shapes` returns only the top-level `<p:sp>` shapes. Shapes inside
group shapes and the text of table cells are reachable separately:

```ruby
Rbpptx.open("deck.pptx") do |pres|
  slide = pres.slides[0]

  slide.groups.first.shapes        # direct children of a group
  slide.all_shapes                 # every <p:sp>, descending into nested groups
  slide.all_shapes.flat_map(&:runs).each { |r| r.text = r.text.gsub("{{name}}", "Acme") }

  table = slide.tables.first       # slide.all_tables also searches groups
  table.row_count                  # => 3
  table.cell(0, 1).text            # => "Q2"
  table.cell(1, 1).text = "42"     # same setters as Shape (text=, runs, paragraphs)
  table.rows.map { |row| row.cells.map(&:text) }
  table.cell(1, 0).spanned?        # covered by a merged cell?

  pres.save("deck.pptx")
end
```

### Reordering and deleting slides

```ruby
Rbpptx.open("deck.pptx") do |pres|
  pres.move_slide(0, -1)             # first slide to the end
  pres.move_slide(pres.slides[3], 0)
  pres.delete_slide(2)               # also removes its notes page
  pres.slide(0).name                 # raises SlideNotFoundError when out of range
  pres.save("deck.pptx")
end
```

`delete_slide` removes the slide's `<p:sldId>`, its relationship, its
custom-show and section entries, its content-type override, and the slide
and notes parts. Shared parts (layouts, images) are kept.

If another slide still points at the one being deleted — typically a
"jump to slide" link on an agenda slide — deletion is refused with
`Rbpptx::Error`, because PowerPoint treats a dangling reference as a
corrupt file. Pass `unlink: true` to remove those hyperlinks first (the
link text stays):

```ruby
pres.delete_slide(5, unlink: true)
```

References that are not hyperlinks are always refused.

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
| `Rbpptx::SlideNotFoundError`       | `slide(i)` / `move_slide` / `delete_slide` index is out of range  |
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
  The compressed payload is copied verbatim. Entries written by streaming
  zippers (e.g. Google Slides exports) get their "data descriptor" flag
  cleared, since the real CRC and sizes are written into the local header
  instead.
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

| benchmark             | rbpptx (s) | python-pptx (s) | docxtemplater (s) | rbpptx vs python-pptx | rbpptx vs docxtemplater |
| ---                   | ---:       | ---:            | ---:              | ---:                  | ---:                    |
| read (walk all text)  | **0.019**  | 0.089           | 0.037             | **4.7x faster**       | **2.0x faster**         |
| fill_one (1 + save)   | **0.033**  | 0.555           | 0.149             | **16.8x faster**      | **4.5x faster**         |
| fill_many (10 + save) | **0.042**  | 0.545           | 0.144             | **13.0x faster**      | **3.4x faster**         |

Memory (RSS delta) on the read scenario:

| library        | rss Δ (kB) |
| ---            | ---:       |
| rbpptx         | **2,956**  |
| python-pptx    | 28,416     |
| docxtemplater  | 34,816     |

The fill scenarios used to tie when each library re-deflated every ZIP
entry on save. `rbpptx` now streams unmodified entries through with
`Zip::OutputStream#copy_raw_entry`, so the deflate cost is paid only on
the slide XML that actually changed — which is what unlocks the 13–17x
gap against python-pptx and the 3–4.5x gap against docxtemplater on
read-modify-write. The advantage is real and reproducible:

```bash
cd benchmark && npm install && cd ..
RBPPTX_SAMPLE=/path/to/your/deck.pptx bundle exec ruby benchmark/compare.rb
```

`uv` is required for the python-pptx subprocess; the comparison script
invokes `uv run --with python-pptx python3 benchmark/python_pptx_compare.py`
so no system-level Python install is mutated. The Node.js side runs from
`benchmark/node_modules` (gitignored); rerun `npm install` after pulling
benchmark changes. Use any `.pptx` you have on hand — the table above
was measured against an 18 MB deck with 21 slides and embedded fonts.

To benchmark against a synthetic deck with no embedded media (isolates the
XML-rewrite cost from binary deflate), generate a fixture first:

```bash
uv run --with python-pptx python3 benchmark/fixture_gen.py 50 /tmp/bench_50.pptx
RBPPTX_SAMPLE=/tmp/bench_50.pptx bundle exec ruby benchmark/compare.rb
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

The bulk of the test suite runs against a small in-tree fixture
(~3 KB) built from scratch by `test/fixture_builder.rb` using only
`rubyzip` and hand-coded OOXML. No external sample is required for a
fresh checkout to get meaningful coverage of read, edit, save, and
fidelity behavior. A handful of dog-fooding tests (`test_sample_*`) opt
into any real-world `.pptx` you point at via the `RBPPTX_SAMPLE`
environment variable, and skip cleanly when it is unset:

```bash
RBPPTX_SAMPLE=/path/to/your/deck.pptx bundle exec ruby -Ilib -Itest test/rbpptx_test.rb
```
