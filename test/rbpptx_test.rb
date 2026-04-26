require "minitest/autorun"
require "fileutils"
require "tmpdir"
require_relative "../lib/rbpptx"

class RbpptxTest < Minitest::Test
  SAMPLE_PATH = File.expand_path(
    "../samples/プロトタイピング プレゼンテーション.pptx", __dir__
  )

  def test_version_is_defined
    assert_match(/\A\d+\.\d+\.\d+\z/, Rbpptx::VERSION)
  end

  def test_new_is_not_implemented_yet
    assert_raises(NotImplementedError) { Rbpptx.new }
  end

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

  def test_close_is_idempotent
    skip_unless_sample
    pres = Rbpptx.open(SAMPLE_PATH)
    assert_equal true, pres.close
    assert_equal false, pres.close
    assert pres.closed?
  end

  def test_slides_after_close_raises
    skip_unless_sample
    pres = Rbpptx.open(SAMPLE_PATH)
    pres.close
    assert_raises(Rbpptx::ClosedPresentationError) { pres.slides }
  end

  def test_open_with_block_yields_presentation_and_auto_closes
    skip_unless_sample
    captured = nil
    Rbpptx.open(SAMPLE_PATH) do |pres|
      captured = pres
      refute pres.closed?
    end
    assert captured.closed?
  end

  def test_open_with_block_closes_on_exception
    skip_unless_sample
    captured = nil
    assert_raises(RuntimeError) do
      Rbpptx.open(SAMPLE_PATH) do |pres|
        captured = pres
        raise "boom"
      end
    end
    assert captured.closed?
  end

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

  def test_shape_name_is_extracted
    skip_unless_sample
    Rbpptx.open(SAMPLE_PATH) do |pres|
      names = pres.slides.first.shapes.map(&:name)
      assert names.all? { |n| n.is_a?(String) && !n.empty? }
    end
  end

  def test_save_without_changes_preserves_all_entries_byte_for_byte
    skip_unless_sample
    Dir.mktmpdir do |dir|
      out = File.join(dir, "copy.pptx")
      Rbpptx.open(SAMPLE_PATH) { |pres| pres.save(out) }

      original_entries = zip_entries(SAMPLE_PATH)
      saved_entries = zip_entries(out)
      assert_equal original_entries.keys.sort, saved_entries.keys.sort
      original_entries.each do |name, bytes|
        assert_equal bytes, saved_entries.fetch(name), "entry #{name} differs after no-op save"
      end
    end
  end

  def test_setting_text_marks_slide_dirty_and_round_trips
    skip_unless_sample
    Dir.mktmpdir do |dir|
      out = File.join(dir, "edited.pptx")
      Rbpptx.open(SAMPLE_PATH) do |pres|
        slide = pres.slides.first
        refute slide.dirty?
        slide.shapes.first.text = "差し替えタイトル"
        assert slide.dirty?
        refute pres.slides[1].dirty?
        pres.save(out)
        refute slide.dirty?, "dirty flag should clear after save"
      end

      Rbpptx.open(out) do |reloaded|
        assert_equal "差し替えタイトル", reloaded.slides.first.shapes.first.text
        assert_includes reloaded.slides[1].shapes.map(&:text).join, "目次"
      end
    end
  end

  def test_multi_line_text_becomes_multiple_paragraphs
    skip_unless_sample
    Dir.mktmpdir do |dir|
      out = File.join(dir, "multi.pptx")
      Rbpptx.open(SAMPLE_PATH) do |pres|
        pres.slides.first.shapes.first.text = "一行目\n二行目\n三行目"
        pres.save(out)
      end

      Rbpptx.open(out) do |reloaded|
        assert_equal "一行目\n二行目\n三行目", reloaded.slides.first.shapes.first.text
      end
    end
  end

  def test_unmodified_slides_preserved_byte_for_byte_after_edit
    skip_unless_sample
    Dir.mktmpdir do |dir|
      out = File.join(dir, "partial.pptx")
      Rbpptx.open(SAMPLE_PATH) do |pres|
        pres.slides.first.shapes.first.text = "新タイトル"
        pres.save(out)
      end

      original = zip_entries(SAMPLE_PATH)
      saved = zip_entries(out)
      original.each do |name, bytes|
        next if name == "ppt/slides/slide1.xml"

        assert_equal bytes, saved.fetch(name), "entry #{name} should be unchanged"
      end
      refute_equal original["ppt/slides/slide1.xml"], saved["ppt/slides/slide1.xml"]
    end
  end

  def test_text_setter_preserves_first_run_rpr_for_template_formatting
    skip_unless_sample
    Dir.mktmpdir do |dir|
      out = File.join(dir, "fmt.pptx")
      Rbpptx.open(SAMPLE_PATH) do |pres|
        pres.slides[1].shapes.first.text = "新しい見出し"
        pres.save(out)
      end

      slide_xml = read_zip_entry(out, "ppt/slides/slide2.xml")
      assert_match(/Montserrat/, slide_xml, "Montserrat font from rPr should survive")
      assert_match(%r{<a:srgbClr val="FFFFFF"}, slide_xml, "white color from rPr should survive")
    end
  end

  def test_save_to_existing_path_overwrites_atomically
    skip_unless_sample
    Dir.mktmpdir do |dir|
      target = File.join(dir, "deck.pptx")
      FileUtils.cp(SAMPLE_PATH, target)

      Rbpptx.open(target) do |pres|
        pres.slides.first.shapes.first.text = "上書き"
        pres.save(target)
      end

      Rbpptx.open(target) do |reloaded|
        assert_equal "上書き", reloaded.slides.first.shapes.first.text
      end

      refute Dir.children(dir).any? { |f| f.include?("rbpptx-tmp") }, "no temp files should remain"
    end
  end

  def test_text_setter_on_shape_without_txbody_raises
    skip_unless_sample
    Rbpptx.open(SAMPLE_PATH) do |pres|
      txless = pres.slides.flat_map(&:shapes).find { |s| !has_txbody?(s) }
      skip "no shape without <p:txBody> found in sample" unless txless

      err = assert_raises(Rbpptx::Error) { txless.text = "x" }
      assert_includes err.message, "txBody"
    end
  end

  def test_save_after_close_raises
    skip_unless_sample
    pres = Rbpptx.open(SAMPLE_PATH)
    pres.close
    Dir.mktmpdir do |dir|
      assert_raises(Rbpptx::ClosedPresentationError) { pres.save(File.join(dir, "x.pptx")) }
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
    Zip::File.open(path) { |zf| zf.find_entry(name).get_input_stream.read }
  end

  def has_txbody?(shape)
    node = shape.instance_variable_get(:@node)
    !node.at_xpath("./p:txBody", "p" => Rbpptx::Shape::P_NS).nil?
  end


  def skip_unless_sample
    skip "sample not present at #{SAMPLE_PATH}" unless File.exist?(SAMPLE_PATH)
  end
end
