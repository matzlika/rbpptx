require "minitest/autorun"
require "fileutils"
require "tmpdir"
require_relative "../lib/rbpptx"
require_relative "fixture_builder"

class RbpptxTest < Minitest::Test
  SAMPLE_PATH = File.expand_path(
    "../samples/prototyping_presentation.pptx", __dir__
  )

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

  # ---------- sample-specific (skip when missing) ----------
  #
  # These keep the dog-fooding signal on a real-world deck: 21 slides, embedded
  # fonts, multi-language text. They skip cleanly in CI / fresh checkouts where
  # the gitignored sample is absent.

  def test_sample_has_21_slides_in_order
    skip_unless_sample
    Rbpptx.open(SAMPLE_PATH) do |pres|
      assert_equal 21, pres.slides.length
      assert_equal (1..21).map { |i| "slide#{i}" }, pres.slides.map(&:name)
    end
  end

  def test_sample_first_slide_contains_title_text
    skip_unless_sample
    Rbpptx.open(SAMPLE_PATH) do |pres|
      texts = pres.slides.first.shapes.map(&:text)
      assert_includes texts.join, "プロトタイピング"
      assert_includes texts.join, "プレゼンテーション"
    end
  end

  def test_sample_second_slide_renders_paragraphs_with_newlines
    skip_unless_sample
    Rbpptx.open(SAMPLE_PATH) do |pres|
      toc_shape = pres.slides[1].shapes.find { |s| s.text.include?("製品紹介") }
      refute_nil toc_shape
      assert_includes toc_shape.text, "\n"
      assert_includes toc_shape.text, "プロジェクトのタイムライン"
    end
  end

  private

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
    skip "sample not present at #{SAMPLE_PATH}" unless File.exist?(SAMPLE_PATH)
  end
end
