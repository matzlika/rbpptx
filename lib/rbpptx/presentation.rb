module Rbpptx
  # A PowerPoint presentation loaded from a +.pptx+ file.
  #
  # Reading is the only mode wired up at present: the package is opened, the
  # presentation part is located via +_rels/.rels+, the slide id list is read
  # in document order, and slides are resolved through the presentation's own
  # rels. Each {Slide} retains a handle on the underlying ZIP and parses its
  # XML lazily.
  class Presentation
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main".freeze
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships".freeze
    PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships".freeze
    OFFICE_DOC_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".freeze
    SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide".freeze

    LEGACY_PPT_SIGNATURE = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1".b.freeze
    ZIP_SIGNATURE = "PK\x03\x04".b.freeze

    # @return [String] absolute path the presentation was loaded from
    attr_reader :path

    # Opens an existing presentation. See {Rbpptx.open} for the public entry
    # point that supports the block-with-auto-close idiom.
    def self.open(path)
      new(path)
    end

    def initialize(path)
      @path = path.to_s
      @closed = false
      validate_signature!
      begin
        @zip = Zip::File.open(@path)
      rescue Zip::Error => e
        raise UnsupportedFormatError, "#{@path}: not a valid .pptx (#{e.message})"
      end
    end

    # @return [Array<Slide>] slides in presentation order (matches the order of
    #   +<p:sldId>+ entries inside +ppt/presentation.xml+)
    def slides
      ensure_open!
      @slides ||= build_slides
    end

    # Writes the presentation out to +path+, preserving every part that has
    # not been mutated byte-for-byte. Slides whose shapes have been edited are
    # re-serialized from their in-memory Nokogiri document; all other entries
    # (themes, masters, layouts, media, fonts, rels) are streamed straight
    # from the source ZIP without re-parsing, so unknown PowerPoint features
    # round-trip unchanged.
    #
    # +path+ may equal the original load path; the new file is written to a
    # temp file in the same directory and atomically renamed into place. On
    # success, dirty flags on each slide are cleared.
    #
    # @param path [String, #to_path]
    # @return [String] the path that was written
    def save(path)
      ensure_open!
      out_path = path.to_s
      overrides = (@slides || []).each_with_object({}) do |slide, h|
        h[slide.part_name] = slide.to_xml if slide.dirty?
      end

      tmp_path = "#{out_path}.rbpptx-tmp.#{Process.pid}.#{rand(1 << 32).to_s(16)}"
      begin
        Zip::OutputStream.open(tmp_path) do |out|
          @zip.each do |entry|
            next if entry.directory?

            out.put_next_entry(entry.name)
            if (override_xml = overrides[entry.name])
              out.write(override_xml)
            else
              entry.get_input_stream { |stream| IO.copy_stream(stream, out) }
            end
          end
        end
        File.rename(tmp_path, out_path)
      rescue StandardError
        File.unlink(tmp_path) if File.exist?(tmp_path)
        raise
      end

      (@slides || []).each(&:clear_dirty!)
      out_path
    end

    # Releases the underlying ZIP file. Idempotent.
    #
    # @return [Boolean] +true+ on the first call, +false+ on subsequent calls
    def close
      return false if @closed

      @zip&.close
      @zip = nil
      @closed = true
      true
    end

    # @return [Boolean]
    def closed?
      @closed
    end

    private

    def ensure_open!
      raise ClosedPresentationError, "#{@path}: presentation is closed" if @closed
    end

    def validate_signature!
      head = File.binread(@path, 8) rescue nil
      raise UnsupportedFormatError, "#{@path}: file is empty or unreadable" if head.nil? || head.empty?

      if head.start_with?(LEGACY_PPT_SIGNATURE)
        raise UnsupportedFormatError,
              "#{@path}: looks like a legacy .ppt (BIFF/CFB) file. " \
              "Convert to .pptx first, e.g. `libreoffice --headless --convert-to pptx`."
      end

      unless head.start_with?(ZIP_SIGNATURE)
        raise UnsupportedFormatError, "#{@path}: missing ZIP signature; not a .pptx package"
      end
    end

    def build_slides
      pres_part = locate_presentation_part
      pres_rels = parse_rels(rels_path_for(pres_part))
      pres_dir = File.dirname(pres_part)

      doc = parse_xml(pres_part)
      sld_ids = doc.xpath("/p:presentation/p:sldIdLst/p:sldId", "p" => P_NS)

      sld_ids.map do |sld_id|
        rid = sld_id.attribute_with_ns("id", R_NS)&.value
        target = pres_rels[rid] or raise PresentationFormatError, "#{@path}: rId #{rid.inspect} not found in presentation rels"
        Slide.new(zip: @zip, part_name: resolve_relative(pres_dir, target))
      end
    end

    def locate_presentation_part
      doc = parse_xml("_rels/.rels")
      rel = doc.at_xpath(
        "/pkg:Relationships/pkg:Relationship[@Type=$type]",
        { "pkg" => PKG_REL_NS },
        { "type" => OFFICE_DOC_REL_TYPE }
      )
      raise PresentationFormatError, "#{@path}: officeDocument relationship missing from _rels/.rels" unless rel

      target = rel["Target"] or raise PresentationFormatError, "#{@path}: officeDocument relationship has no Target"
      target.sub(%r{\A/}, "")
    end

    def parse_rels(rels_part)
      entry = @zip.find_entry(rels_part)
      return {} unless entry

      doc = Nokogiri::XML(entry.get_input_stream.read)
      doc.xpath("/pkg:Relationships/pkg:Relationship", "pkg" => PKG_REL_NS).each_with_object({}) do |r, h|
        h[r["Id"]] = r["Target"]
      end
    end

    def parse_xml(part_name)
      entry = @zip.find_entry(part_name) or raise PresentationFormatError, "#{@path}: missing part #{part_name}"
      doc = Nokogiri::XML(entry.get_input_stream.read)
      raise PresentationFormatError, "#{@path}: #{part_name}: #{doc.errors.first}" unless doc.errors.empty?

      doc
    end

    def rels_path_for(part_name)
      dir = File.dirname(part_name)
      base = File.basename(part_name)
      dir == "." ? "_rels/#{base}.rels" : "#{dir}/_rels/#{base}.rels"
    end

    def resolve_relative(base_dir, target)
      return target.sub(%r{\A/}, "") if target.start_with?("/")

      File.expand_path(target, "/#{base_dir}").sub(%r{\A/}, "")
    end
  end
end
