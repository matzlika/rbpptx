require "minitest/autorun"
require "fileutils"
require "tmpdir"
require_relative "../lib/rbpptx"
require_relative "fixture_builder"

class RbpptxTest < Minitest::Test
  SAMPLE_PATH = ENV["RBPPTX_SAMPLE"]

  FIXTURE_DIR = File.expand_path("fixtures", __dir__)
  FIXTURE_PATH = File.join(FIXTURE_DIR, "built.pptx")
  FileUtils.mkdir_p(FIXTURE_DIR)
  Rbpptx::Test::FixtureBuilder.build(FIXTURE_PATH)

  # ---------- meta ----------

  def test_version_is_defined
    assert_match(/\A\d+\.\d+\.\d+\z/, Rbpptx::VERSION)
  end

  def test_new_is_not_implemented_yet
    assert_raises(NotImplementedError) { Rbpptx.new }
  end

  # ---------- input validation ----------

  def test_open_rejects_non_zip_with_helpful_message
    Dir.mktmpdir do |dir|
      path = File.join(dir, "garbage.pptx")
      File.binwrite(path, "not a pptx at all")

      err = assert_raises(Rbpptx::UnsupportedFormatError) { Rbpptx.open(path) }
      assert_includes err.message, path
      assert_includes err.message, "ZIP signature"
    end
  end

  def test_open_rejects_legacy_ppt_with_helpful_message
    Dir.mktmpdir do |dir|
      path = File.join(dir, "legacy.ppt")
      File.binwrite(path, "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1".b + ("\x00".b * 512))

      err = assert_raises(Rbpptx::UnsupportedFormatError) { Rbpptx.open(path) }
      assert_includes err.message, path
      assert_includes err.message, "legacy .ppt"
      assert_includes err.message, "convert-to pptx"
    end
  end

  # ---------- lifecycle (built fixture) ----------

  def test_close_is_idempotent
    pres = Rbpptx.open(FIXTURE_PATH)
    assert_equal true, pres.close
    assert_equal false, pres.close
    assert pres.closed?
  end

  def test_slides_after_close_raises
    pres = Rbpptx.open(FIXTURE_PATH)
    pres.close
    assert_raises(Rbpptx::ClosedPresentationError) { pres.slides }
  end

  def test_open_with_block_yields_presentation_and_auto_closes
    captured = nil
    Rbpptx.open(FIXTURE_PATH) do |pres|
      captured = pres
      refute pres.closed?
    end
    assert captured.closed?
  end

  def test_open_with_block_closes_on_exception
    captured = nil
    assert_raises(RuntimeError) do
      Rbpptx.open(FIXTURE_PATH) do |pres|
        captured = pres
        raise "boom"
      end
    end
    assert captured.closed?
  end

  def test_save_after_close_raises
    pres = Rbpptx.open(FIXTURE_PATH)
    pres.close
    Dir.mktmpdir do |dir|
      assert_raises(Rbpptx::ClosedPresentationError) { pres.save(File.join(dir, "x.pptx")) }
    end
  end

  # ---------- structural reads (built fixture) ----------

  def test_fixture_has_three_slides_in_presentation_order
    Rbpptx.open(FIXTURE_PATH) do |pres|
      assert_equal %w[slide1 slide2 slide3], pres.slides.map(&:name)
    end
  end

  def test_shape_name_is_extracted
    Rbpptx.open(FIXTURE_PATH) do |pres|
      assert_equal ["Title 1", "Subtitle 2"], pres.slides[0].shapes.map(&:name)
    end
  end

  def test_paragraphs_split_at_a_p_boundaries
    Rbpptx.open(FIXTURE_PATH) do |pres|
      body = pres.slides[1].shapes.first
      assert_equal 2, body.paragraphs.length
      assert_equal body.text, body.paragraphs.map(&:text).join("\n")
    end
  end

  def test_runs_flattens_across_paragraphs
    Rbpptx.open(FIXTURE_PATH) do |pres|
      body = pres.slides[1].shapes.first
      assert_equal 2, body.paragraphs.length
      assert_equal ["line one", "line two", "second paragraph"], body.runs.map(&:text)
    end
  end

  def test_paragraph_text_renders_br_as_newline_within_paragraph
    Rbpptx.open(FIXTURE_PATH) do |pres|
      first_para = pres.slides[1].shapes.first.paragraphs.first
      assert_equal "line one\nline two", first_para.text
    end
  end

  def test_placeholder_type_for_known_placeholders
    Rbpptx.open(FIXTURE_PATH) do |pres|
      assert_equal :ctrTitle, pres.slides[0].shapes[0].placeholder_type
      assert_equal :subTitle, pres.slides[0].shapes[1].placeholder_type
      assert_equal :body,     pres.slides[1].shapes[0].placeholder_type
    end
  end

  def test_placeholder_type_nil_for_non_placeholders
    Rbpptx.open(FIXTURE_PATH) do |pres|
      pres.slides[2].shapes.each do |shape|
        assert_nil shape.placeholder_type, "slide 3 shapes are intentionally non-placeholder"
      end
    end
  end

  def test_runs_skip_a_br_and_a_fld_elements
    Rbpptx.open(FIXTURE_PATH) do |pres|
      first_para = pres.slides[1].shapes.first.paragraphs.first
      run_count = first_para.runs.length
      child_count = first_para.instance_variable_get(:@node).element_children.length
      assert_operator child_count, :>, run_count,
                      "paragraph should have non-run children (e.g. <a:br>) excluded from runs"
    end
  end

  def test_paragraphs_returns_empty_for_shape_without_txbody
    Rbpptx.open(FIXTURE_PATH) do |pres|
      txless = pres.slides[2].shapes.first
      refute has_txbody?(txless), "slide 3 shape 0 is intentionally without <p:txBody>"

      assert_equal [], txless.paragraphs
      assert_equal [], txless.runs
      assert_equal "", txless.text
    end
  end

  # ---------- save / edit (built fixture) ----------

  def test_save_without_changes_preserves_all_entries_byte_for_byte
    Dir.mktmpdir do |dir|
      out = File.join(dir, "copy.pptx")
      Rbpptx.open(FIXTURE_PATH) { |pres| pres.save(out) }

      original_entries = zip_entries(FIXTURE_PATH)
      saved_entries = zip_entries(out)
      assert_equal original_entries.keys.sort, saved_entries.keys.sort
      original_entries.each do |name, bytes|
        assert_equal bytes, saved_entries.fetch(name), "entry #{name} differs after no-op save"
      end
    end
  end

  def test_setting_text_marks_slide_dirty_and_round_trips
    Dir.mktmpdir do |dir|
      out = File.join(dir, "edited.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        slide = pres.slides[0]
        refute slide.dirty?
        slide.shapes.first.text = "差し替えタイトル"
        assert slide.dirty?
        refute pres.slides[1].dirty?
        pres.save(out)
        refute slide.dirty?, "dirty flag should clear after save"
      end

      Rbpptx.open(out) do |reloaded|
        assert_equal "差し替えタイトル", reloaded.slides[0].shapes[0].text
        assert_equal "fixture subtitle",   reloaded.slides[0].shapes[1].text
      end
    end
  end

  def test_multi_line_text_becomes_multiple_paragraphs
    Dir.mktmpdir do |dir|
      out = File.join(dir, "multi.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.slides[0].shapes[0].text = "一行目\n二行目\n三行目"
        pres.save(out)
      end

      Rbpptx.open(out) do |reloaded|
        assert_equal "一行目\n二行目\n三行目", reloaded.slides[0].shapes[0].text
        assert_equal 3, reloaded.slides[0].shapes[0].paragraphs.length
      end
    end
  end

  def test_unmodified_slides_preserved_byte_for_byte_after_edit
    Dir.mktmpdir do |dir|
      out = File.join(dir, "partial.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.slides[0].shapes[0].text = "新タイトル"
        pres.save(out)
      end

      original = zip_entries(FIXTURE_PATH)
      saved = zip_entries(out)
      original.each do |name, bytes|
        next if name == "ppt/slides/slide1.xml"

        assert_equal bytes, saved.fetch(name), "entry #{name} should be unchanged"
      end
      refute_equal original["ppt/slides/slide1.xml"], saved["ppt/slides/slide1.xml"]
    end
  end

  def test_text_setter_preserves_first_run_rpr_for_template_formatting
    Dir.mktmpdir do |dir|
      out = File.join(dir, "fmt.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.slides[1].shapes.first.text = "新しい見出し"
        pres.save(out)
      end

      slide_xml = read_zip_entry(out, "ppt/slides/slide2.xml")
      assert_match(/Verdana/, slide_xml,    "Verdana font from rPr should survive")
      assert_match(/sz="2400"/, slide_xml,  "size 2400 from rPr should survive")
      assert_match(%r{<a:srgbClr val="2E74B5"}, slide_xml, "color from rPr should survive")
    end
  end

  def test_save_to_existing_path_overwrites_atomically
    Dir.mktmpdir do |dir|
      target = File.join(dir, "deck.pptx")
      FileUtils.cp(FIXTURE_PATH, target)

      Rbpptx.open(target) do |pres|
        pres.slides[0].shapes[0].text = "上書き"
        pres.save(target)
      end

      Rbpptx.open(target) do |reloaded|
        assert_equal "上書き", reloaded.slides[0].shapes[0].text
      end

      refute Dir.children(dir).any? { |f| f.include?("rbpptx-tmp") }, "no temp files should remain"
    end
  end

  # Regression: copy_raw_entry kept general-purpose bit 3 ("sizes follow in
  # a data descriptor") without writing a descriptor, so a no-op save of a
  # Google Slides export could not be opened by PowerPoint.
  def test_save_of_streaming_zip_writes_self_describing_local_headers
    Dir.mktmpdir do |dir|
      src = Rbpptx::Test::FixtureBuilder.build_streaming_copy(FIXTURE_PATH, File.join(dir, "streamed.pptx"))
      assert local_headers(src).all? { |h| h[:flags] & 8 != 0 }, "precondition: input uses data descriptors"

      out = File.join(dir, "saved.pptx")
      Rbpptx.open(src) do |pres|
        pres.slides[0].shapes[0].text = "edited"
        pres.save(out)
      end

      local_headers(out).each do |h|
        assert_equal 0, h[:flags] & 8, "#{h[:name]}: bit 3 set but no data descriptor is written"
        # Re-serialized parts are written by rubyzip with Zip64 placeholders;
        # only raw-copied entries are compared field by field.
        next if h[:name] == "ppt/slides/slide1.xml"

        assert_equal h[:central_crc], h[:crc], "#{h[:name]}: local header CRC"
        assert_equal h[:central_size], h[:compressed_size], "#{h[:name]}: local header size"
      end
      original = zip_entries(FIXTURE_PATH)
      saved = zip_entries(out)
      assert_equal original.keys.sort, saved.keys.sort
      original.each do |name, bytes|
        next if name == "ppt/slides/slide1.xml"

        assert_equal bytes, saved.fetch(name), "entry #{name} content should be unchanged"
      end
    end
  end

  def test_text_setter_on_shape_without_txbody_raises
    Rbpptx.open(FIXTURE_PATH) do |pres|
      txless = pres.slides[2].shapes.first
      refute has_txbody?(txless)

      err = assert_raises(Rbpptx::Error) { txless.text = "x" }
      assert_includes err.message, "txBody"
    end
  end

  # ---------- run-level surgical edits (built fixture) ----------

  def test_run_text_setter_preserves_other_runs_in_same_shape
    Dir.mktmpdir do |dir|
      out = File.join(dir, "run.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        body = pres.slides[1].shapes.first
        body.runs[0].text = "FIRST"
        assert pres.slides[1].dirty?
        pres.save(out)
      end

      Rbpptx.open(out) do |reloaded|
        runs = reloaded.slides[1].shapes.first.runs
        assert_equal ["FIRST", "line two", "second paragraph"], runs.map(&:text)
      end
    end
  end

  def test_run_text_setter_preserves_rpr_attributes_byte_level
    Dir.mktmpdir do |dir|
      out = File.join(dir, "rprkeep.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        body = pres.slides[1].shapes.first
        body.runs[0].text = "REPLACED"
        pres.save(out)
      end

      slide_xml = read_zip_entry(out, "ppt/slides/slide2.xml")
      assert_includes slide_xml, "REPLACED"
      refute_includes slide_xml, "line one"
      # Every rPr attribute on the original first run must still be present
      # somewhere in the slide XML after the surgical edit.
      assert_match(/sz="2400"/, slide_xml)
      assert_match(/Verdana/, slide_xml)
      assert_match(%r{<a:srgbClr val="2E74B5"}, slide_xml)
      assert_match(/lang="en-US"/, slide_xml)
      assert_match(/b="1"/, slide_xml, "bold attribute should survive")
    end
  end

  def test_run_text_setter_preserves_a_br_soft_breaks
    Dir.mktmpdir do |dir|
      out = File.join(dir, "br.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.slides[1].shapes.first.runs[0].text = "X"
        pres.save(out)
      end
      slide_xml = read_zip_entry(out, "ppt/slides/slide2.xml")
      assert_includes slide_xml, "<a:br>", "soft line break should survive run.text= edit"
    end
  end

  def test_run_text_setter_marks_only_owning_slide_dirty
    Rbpptx.open(FIXTURE_PATH) do |pres|
      pres.slides[1].shapes.first.runs[0].text = "edited"
      assert pres.slides[1].dirty?
      refute pres.slides[0].dirty?
      refute pres.slides[2].dirty?
    end
  end

  def test_run_level_edit_preserves_br_that_shape_edit_destroys
    Dir.mktmpdir do |dir|
      run_out = File.join(dir, "run.pptx")
      shape_out = File.join(dir, "shape.pptx")

      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.slides[1].shapes.first.runs[0].text = "Z"
        pres.save(run_out)
      end
      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.slides[1].shapes.first.text = "Z\nline two\nsecond paragraph"
        pres.save(shape_out)
      end

      run_xml = read_zip_entry(run_out, "ppt/slides/slide2.xml")
      shape_xml = read_zip_entry(shape_out, "ppt/slides/slide2.xml")

      assert_includes run_xml, "<a:br>",
                      "run-level edit must preserve <a:br> soft break"
      refute_includes shape_xml, "<a:br>",
                      "shape-level edit drops the <a:br>"
    end
  end

  # ---------- placeholder API (built fixture) ----------

  def test_slide_title_returns_title_or_ctr_title_placeholder
    Rbpptx.open(FIXTURE_PATH) do |pres|
      assert_equal "Hello rbpptx", pres.slides[0].title.text
      assert_nil pres.slides[2].title
    end
  end

  def test_placeholders_lookup_by_type_and_idx
    Rbpptx.open(FIXTURE_PATH) do |pres|
      phs = pres.slides[0].placeholders
      assert_equal 2, phs.size
      assert_equal "fixture subtitle", phs[:subTitle].text
      assert_equal "fixture subtitle", phs[1].text, "Integer key matches <p:ph idx>"
      assert_equal "Hello rbpptx", phs[0].text, "omitted idx defaults to 0"
      assert_nil phs[:body]
      assert_raises(TypeError) { phs["title"] }
      assert pres.slides[2].placeholders.empty?
    end
  end

  def test_placeholder_idx_and_predicate
    Rbpptx.open(FIXTURE_PATH) do |pres|
      title, sub = pres.slides[0].shapes
      assert title.placeholder?
      assert_equal 0, title.placeholder_idx
      assert_equal 1, sub.placeholder_idx
      refute pres.slides[2].shapes.last.placeholder?
      assert_nil pres.slides[2].shapes.last.placeholder_idx
    end
  end

  def test_placeholder_text_edit_round_trips
    Dir.mktmpdir do |dir|
      out = File.join(dir, "ph.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.slides[1].placeholders[:body].text = "本文"
        pres.save(out)
      end
      Rbpptx.open(out) { |pres| assert_equal "本文", pres.slides[1].placeholders[:body].text }
    end
  end

  # ---------- notes (built fixture) ----------

  def test_notes_reads_existing_notes_slide
    Rbpptx.open(FIXTURE_PATH) do |pres|
      assert_equal "speaker notes for slide 1", pres.slides[0].notes
      assert_kind_of Rbpptx::NotesSlide, pres.slides[0].notes_slide
      assert_equal :body, pres.slides[0].notes_slide.notes_placeholder.placeholder_type
    end
  end

  def test_notes_empty_when_slide_has_no_notes_slide
    Rbpptx.open(FIXTURE_PATH) do |pres|
      assert_nil pres.slides[1].notes_slide
      assert_equal "", pres.slides[1].notes
    end
  end

  def test_editing_existing_notes_round_trips_and_leaves_slide_untouched
    Dir.mktmpdir do |dir|
      out = File.join(dir, "notes.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.slides[0].notes = "新しいノート\n二行目"
        assert pres.slides[0].dirty?
        pres.save(out)
        refute pres.slides[0].dirty?
      end

      Rbpptx.open(out) { |pres| assert_equal "新しいノート\n二行目", pres.slides[0].notes }
      original = zip_entries(FIXTURE_PATH)
      saved = zip_entries(out)
      assert_equal original.keys.sort, saved.keys.sort
      original.each do |name, bytes|
        next if name == "ppt/notesSlides/notesSlide1.xml"

        assert_equal bytes, saved.fetch(name), "entry #{name} should be unchanged"
      end
    end
  end

  def test_setting_notes_creates_notes_slide_when_missing
    Dir.mktmpdir do |dir|
      out = File.join(dir, "newnotes.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.slides[1].notes = "created <&> notes"
        pres.save(out)
      end

      Rbpptx.open(out) do |pres|
        assert_equal "created <&> notes", pres.slides[1].notes
        assert_equal "speaker notes for slide 1", pres.slides[0].notes
        assert_equal "ppt/notesSlides/notesSlide2.xml", pres.slides[1].notes_slide.part_name
      end

      saved = zip_entries(out)
      assert_includes saved["[Content_Types].xml"], %(PartName="/ppt/notesSlides/notesSlide2.xml")
      notes_rels = saved.fetch("ppt/notesSlides/_rels/notesSlide2.xml.rels")
      assert_includes notes_rels, %(Target="../notesMasters/notesMaster1.xml")
      assert_includes notes_rels, %(Target="../slides/slide2.xml")
      assert_includes saved.fetch("ppt/slides/_rels/slide2.xml.rels"), %(Target="../notesSlides/notesSlide2.xml")
    end
  end

  def test_setting_notes_without_notes_master_raises
    Dir.mktmpdir do |dir|
      path = File.join(dir, "nomaster.pptx")
      rewrite_fixture(path) do |name, body|
        next nil if name.start_with?("ppt/notesMasters/")

        name == "ppt/_rels/presentation.xml.rels" ? body.sub(/<Relationship Id="rId10"[^>]*>/, "") : body
      end

      Rbpptx.open(path) do |pres|
        err = assert_raises(Rbpptx::Error) { pres.slides[1].notes = "x" }
        assert_includes err.message, "notes master"
      end
    end
  end

  # ---------- groups & tables (built fixture) ----------

  def test_shapes_excludes_grouped_shapes_but_all_shapes_descends
    Rbpptx.open(FIXTURE_PATH) do |pres|
      slide = pres.slides[2]
      assert_equal ["NoTextShape", "Plain Text"], slide.shapes.map(&:name)
      assert_equal ["NoTextShape", "Plain Text", "Grouped Text", "Nested Text"], slide.all_shapes.map(&:name)
    end
  end

  def test_groups_nest
    Rbpptx.open(FIXTURE_PATH) do |pres|
      outer = pres.slides[2].groups.first
      assert_equal "Outer Group", outer.name
      assert_equal ["Grouped Text"], outer.shapes.map(&:name)
      assert_equal ["Inner Group"], outer.groups.map(&:name)
      assert_equal "in group\nin nested group", outer.text
    end
  end

  def test_grouped_shape_text_edit_round_trips
    Dir.mktmpdir do |dir|
      out = File.join(dir, "grp.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        nested = pres.slides[2].groups.first.groups.first.shapes.first
        nested.runs.first.text = "edited nested"
        assert pres.slides[2].dirty?
        pres.save(out)
      end
      Rbpptx.open(out) { |pres| assert_equal "edited nested", pres.slides[2].all_shapes.last.text }
    end
  end

  def test_table_cells_read
    Rbpptx.open(FIXTURE_PATH) do |pres|
      table = pres.slides[2].tables.first
      assert_equal "Table 1", table.name
      assert_equal 2, table.row_count
      assert_equal 2, table.column_count
      assert_equal [["Name", "Value"], ["merged", ""]], table.rows.map { |r| r.cells.map(&:text) }
      assert table.cell(1, 0).merge_origin?
      assert table.cell(1, 1).spanned?
      refute table.cell(0, 0).spanned?
      assert_nil table.cell(5, 0)
      assert_equal [table.name], pres.slides[2].all_tables.map(&:name)
    end
  end

  def test_table_cell_text_edit_keeps_rpr_and_round_trips
    Dir.mktmpdir do |dir|
      out = File.join(dir, "tbl.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        table = pres.slides[2].tables.first
        table.cell(0, 1).text = "値"
        table.cell(1, 0).runs.first.text = "結合"
        pres.save(out)
      end

      Rbpptx.open(out) do |pres|
        table = pres.slides[2].tables.first
        assert_equal [["Name", "値"], ["結合", ""]], table.rows.map { |r| r.cells.map(&:text) }
      end
      xml = read_zip_entry(out, "ppt/slides/slide3.xml")
      assert_match(%r{<a:rPr lang="en-US" b="1"/><a:t>値</a:t>}, xml, "bold rPr should be cloned onto new run")
    end
  end

  def test_text_setter_on_empty_field_uses_end_para_rpr_and_ppr
    Dir.mktmpdir do |dir|
      out = File.join(dir, "field.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.slides[2].tables.first.cell(1, 1).text = "記入1\n記入2"
        pres.save(out)
      end

      xml = read_zip_entry(out, "ppt/slides/slide3.xml")
      cell = xml[%r{<a:tc hMerge="1">.*?</a:tc>}m]
      assert_equal 2, cell.scan(%r{<a:r><a:rPr lang="ja-JP" sz="1200"><a:latin typeface="Meiryo"/></a:rPr><a:t>記入\d</a:t></a:r>}).size,
                   "endParaRPr should become the run template when there are no runs"
      assert_equal 2, cell.scan(%r{<a:p><a:pPr algn="r"/>}).size, "first pPr should be kept on every paragraph"
      assert_equal 2, cell.scan("<a:endParaRPr").size
    end
  end

  # ---------- slide reorder / delete (built fixture) ----------

  def test_slide_accessor_and_out_of_range
    Rbpptx.open(FIXTURE_PATH) do |pres|
      assert_equal "slide3", pres.slide(-1).name
      assert_raises(Rbpptx::SlideNotFoundError) { pres.slide(3) }
    end
  end

  def test_move_slide_reorders_and_round_trips
    Dir.mktmpdir do |dir|
      out = File.join(dir, "moved.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        assert_equal %w[slide3 slide1 slide2], pres.move_slide(2, 0).map(&:name)
        assert_equal %w[slide1 slide2 slide3], pres.move_slide(pres.slides[0], -1).map(&:name)
        pres.move_slide(0, 1)
        assert_equal %w[slide2 slide1 slide3], pres.slides.map(&:name)
        assert_raises(Rbpptx::SlideNotFoundError) { pres.move_slide(0, 3) }
        pres.save(out)
      end

      Rbpptx.open(out) { |pres| assert_equal %w[slide2 slide1 slide3], pres.slides.map(&:name) }
      original = zip_entries(FIXTURE_PATH)
      saved = zip_entries(out)
      original.each do |name, bytes|
        next if name == "ppt/presentation.xml"

        assert_equal bytes, saved.fetch(name), "entry #{name} should be unchanged by a reorder"
      end
    end
  end

  def test_delete_slide_removes_slide_notes_and_bookkeeping
    Dir.mktmpdir do |dir|
      out = File.join(dir, "deleted.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        removed = pres.delete_slide(0)
        assert_equal "slide1", removed.name
        assert_equal %w[slide2 slide3], pres.slides.map(&:name)
        pres.save(out)
      end

      Rbpptx.open(out) { |pres| assert_equal %w[slide2 slide3], pres.slides.map(&:name) }
      saved = zip_entries(out)
      %w[ppt/slides/slide1.xml ppt/slides/_rels/slide1.xml.rels
         ppt/notesSlides/notesSlide1.xml ppt/notesSlides/_rels/notesSlide1.xml.rels].each do |gone|
        refute saved.key?(gone), "#{gone} should be removed"
      end
      refute_includes saved["[Content_Types].xml"], "/ppt/slides/slide1.xml"
      refute_includes saved["[Content_Types].xml"], "/ppt/notesSlides/notesSlide1.xml"
      refute_includes saved["ppt/_rels/presentation.xml.rels"], "slides/slide1.xml"
      pres_xml = saved["ppt/presentation.xml"]
      refute_match(/<p:sldId id="256"/, pres_xml)
      refute_match(/<p14:sldId id="256"/, pres_xml, "section list entry should be dropped")
      assert_match(/<p14:sldId id="257"/, pres_xml)
      assert saved.key?("ppt/notesMasters/notesMaster1.xml"), "shared parts stay"
    end
  end

  def test_delete_slide_refuses_when_hyperlinked_without_unlink
    Rbpptx.open(FIXTURE_PATH) do |pres|
      err = assert_raises(Rbpptx::Error) { pres.delete_slide(1) }
      assert_includes err.message, "ppt/slides/slide3.xml"
      assert_includes err.message, "unlink: true"
      assert_equal 3, pres.slides.size, "refusal must not mutate state"
      refute pres.slides[2].dirty?
    end
  end

  def test_delete_slide_with_unlink_strips_hyperlink_but_keeps_text
    Dir.mktmpdir do |dir|
      out = File.join(dir, "unlinked.pptx")
      Rbpptx.open(FIXTURE_PATH) do |pres|
        pres.delete_slide(pres.slides[1], unlink: true)
        assert pres.slides.last.dirty?
        pres.save(out)
      end

      Rbpptx.open(out) do |pres|
        assert_equal %w[slide1 slide3], pres.slides.map(&:name)
        assert_equal "plain shape, no placeholder", pres.slides[1].shapes.last.text
      end
      saved = zip_entries(out)
      refute_includes saved["ppt/slides/slide3.xml"], "hlinkClick"
      refute_includes saved["ppt/slides/_rels/slide3.xml.rels"], "slide2.xml"
      refute saved.key?("ppt/slides/slide2.xml")
    end
  end

  def test_delete_slide_refuses_non_hyperlink_reference_even_with_unlink
    Dir.mktmpdir do |dir|
      path = File.join(dir, "oddref.pptx")
      rewrite_fixture(path) do |name, body|
        name == "ppt/slides/slide3.xml" ? body.sub(%(<a:hlinkClick r:id="rId1"), %(<a:hlinkClick r:id="rId9")).sub("<p:spTree>", %(<p:spTree><p:custData r:id="rId1"/>)) : body
      end

      Rbpptx.open(path) do |pres|
        err = assert_raises(Rbpptx::Error) { pres.delete_slide(1, unlink: true) }
        assert_includes err.message, "custData"
        assert_equal 3, pres.slides.size
      end
    end
  end

  def test_save_twice_keeps_edits_and_save_over_source_keeps_reading
    Dir.mktmpdir do |dir|
      target = File.join(dir, "deck.pptx")
      second = File.join(dir, "second.pptx")
      FileUtils.cp(FIXTURE_PATH, target)

      Rbpptx.open(target) do |pres|
        pres.slides[0].shapes[0].text = "一回目"
        pres.save(target)
        # slide 2 is parsed for the first time *after* overwriting the source
        assert_equal "line one\nline two\nsecond paragraph", pres.slides[1].shapes.first.text
        pres.slides[1].notes = "後から追加"
        pres.save(second)
      end

      Rbpptx.open(second) do |pres|
        assert_equal "一回目", pres.slides[0].shapes[0].text, "earlier edit must survive a second save"
        assert_equal "後から追加", pres.slides[1].notes
      end
    end
  end

  # ---------- sample-specific (skip unless RBPPTX_SAMPLE is set) ----------
  #
  # These run only when the developer points RBPPTX_SAMPLE at a real-world
  # .pptx. Assertions are content-agnostic on purpose so any reasonable
  # deck can drive the dog-fooding pass.

  def test_sample_opens_with_at_least_one_slide
    skip_unless_sample
    Rbpptx.open(SAMPLE_PATH) do |pres|
      assert_operator pres.slides.length, :>=, 1
      pres.slides.each { |s| assert_match(/\Aslide\d+\z/, s.name) }
    end
  end

  def test_sample_walks_every_slide_and_shape_without_raising
    skip_unless_sample
    Rbpptx.open(SAMPLE_PATH) do |pres|
      pres.slides.each do |slide|
        slide.shapes.each do |shape|
          assert_kind_of String, shape.text
          assert(shape.placeholder_type.nil? || shape.placeholder_type.is_a?(Symbol))
        end
      end
    end
  end

  def test_sample_round_trips_a_no_op_save_byte_for_byte
    skip_unless_sample
    Dir.mktmpdir do |dir|
      out = File.join(dir, "roundtrip.pptx")
      Rbpptx.open(SAMPLE_PATH) { |pres| pres.save(out) }

      original = zip_entries(SAMPLE_PATH)
      saved = zip_entries(out)
      assert_equal original.keys.sort, saved.keys.sort
      original.each do |name, bytes|
        assert_equal bytes, saved.fetch(name), "entry #{name} differs after no-op save"
      end
    end
  end

  def test_sample_reads_notes_groups_and_placeholders
    skip_unless_sample
    Rbpptx.open(SAMPLE_PATH) do |pres|
      pres.slides.each do |slide|
        assert_kind_of String, slide.notes
        assert_operator slide.all_shapes.size, :>=, slide.shapes.size
        slide.all_tables.each { |t| t.rows.each { |r| r.cells.each { |c| assert_kind_of String, c.text } } }
        slide.placeholders.each { |ph| assert ph.placeholder? }
      end
    end
  end

  def test_sample_reorder_delete_and_notes_round_trip
    skip_unless_sample
    Dir.mktmpdir do |dir|
      out = File.join(dir, "edited.pptx")
      names = nil
      Rbpptx.open(SAMPLE_PATH) do |pres|
        pres.move_slide(0, -1)
        pres.delete_slide(pres.slides.size / 2, unlink: true)
        pres.slides[0].notes = "rbpptx notes"
        names = pres.slides.map(&:name)
        pres.save(out)
      end

      Rbpptx.open(out) do |pres|
        assert_equal names, pres.slides.map(&:name)
        assert_equal "rbpptx notes", pres.slides[0].notes
      end
      assert_package_consistent(out)
    end
  end

  private

  # Every rels target exists, and every part (except rels and the content
  # types part itself) has a content type.
  def assert_package_consistent(path)
    entries = zip_entries(path)
    ct = entries.fetch("[Content_Types].xml")
    defaults = ct.scan(/<Default\b[^>]*>/).map { |el| el[/Extension="([^"]+)"/, 1].to_s.downcase }
    overrides = ct.scan(/<Override\b[^>]*>/).map { |el| el[/PartName="([^"]+)"/, 1].to_s.delete_prefix("/") }

    entries.each_key do |name|
      next if name == "[Content_Types].xml"

      assert(overrides.include?(name) || defaults.include?(name[/\.([^.\/]+)\z/, 1].to_s.downcase),
             "#{name} has no content type")
    end
    overrides.each { |n| assert entries.key?(n), "content type override for missing part #{n}" }

    entries.each do |name, body|
      next unless name.end_with?(".rels")

      source_dir = File.dirname(File.dirname(name))
      source_dir = "" if source_dir == "."
      body.scan(/<Relationship\b[^>]*>/).each do |rel|
        next if rel.include?('TargetMode="External"')

        target = rel[/Target="([^"]+)"/, 1]
        resolved = target.start_with?("/") ? target.delete_prefix("/") : File.expand_path(target, "/#{source_dir}").delete_prefix("/")
        assert entries.key?(resolved), "#{name}: dangling target #{target}"
      end
    end
  end

  # Reads each entry's local file header next to its central directory record.
  def local_headers(path)
    raw = File.binread(path)
    Zip::File.open(path) do |zf|
      zf.entries.map do |e|
        _sig, _ver, flags, _meth, _t, _d, crc, csize = raw.byteslice(e.local_header_offset, 30).unpack("VvvvvvVV")
        { name: e.name, flags: flags, crc: crc, compressed_size: csize,
          central_crc: e.crc, central_size: e.compressed_size }
      end
    end
  end

  # Writes a variant of the built fixture; the block maps (name, body) to a
  # new body, or nil to drop the entry.
  def rewrite_fixture(path)
    entries = zip_entries(FIXTURE_PATH)
    Zip::OutputStream.open(path) do |out|
      entries.each do |name, body|
        new_body = yield(name, body.force_encoding("UTF-8"))
        next if new_body.nil?

        out.put_next_entry(name)
        out.write(new_body)
      end
    end
  end

  def zip_entries(path)
    require "zip"
    Zip::File.open(path) do |zf|
      zf.each_with_object({}) do |entry, h|
        next if entry.directory?

        h[entry.name] = entry.get_input_stream.read
      end
    end
  end

  def read_zip_entry(path, name)
    require "zip"
    Zip::File.open(path) { |zf| zf.find_entry(name).get_input_stream.read.force_encoding("UTF-8") }
  end

  def has_txbody?(shape)
    node = shape.instance_variable_get(:@node)
    !node.at_xpath("./p:txBody", "p" => Rbpptx::Shape::P_NS).nil?
  end

  def skip_unless_sample
    return if SAMPLE_PATH && File.exist?(SAMPLE_PATH)

    skip "set RBPPTX_SAMPLE to a .pptx path to enable dog-fooding tests"
  end
end
