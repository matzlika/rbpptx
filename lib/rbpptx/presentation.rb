module Rbpptx
  # A PowerPoint presentation loaded from a +.pptx+ file.
  #
  # The package is opened, the presentation part is located via
  # +_rels/.rels+, the slide id list is read in document order, and slides
  # are resolved through the presentation's own rels. Each {Slide} parses
  # its XML lazily.
  #
  # Slide-level structure edits ({#move_slide}, {#delete_slide}) rewrite
  # only the bookkeeping parts that must change — +ppt/presentation.xml+,
  # its rels, and +[Content_Types].xml+ — and leave every other part as-is.
  class Presentation
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main".freeze
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships".freeze
    P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main".freeze
    PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships".freeze
    OFFICE_DOC_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".freeze
    SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide".freeze
    NOTES_SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide".freeze
    NOTES_MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster".freeze
    NOTES_SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml".freeze

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
        @package = Package.new(@path)
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

    # @param index [Integer] zero-based; negative counts from the end
    # @return [Slide]
    # @raise [SlideNotFoundError] if +index+ is out of range
    def slide(index)
      slides.fetch(index)
    rescue IndexError
      raise SlideNotFoundError, "#{@path}: no slide at index #{index} (#{slides.size} slides)"
    end

    # Moves a slide to a new position in the presentation order. Only
    # +<p:sldIdLst>+ in +ppt/presentation.xml+ is rewritten.
    #
    # @param slide_or_index [Slide, Integer] the slide to move
    # @param new_index [Integer] zero-based destination; negative counts from
    #   the end (+-1+ moves to the last position)
    # @return [Array<Slide>] slides in the new order
    # @raise [SlideNotFoundError] if either argument is out of range
    def move_slide(slide_or_index, new_index)
      target = resolve_slide(slide_or_index)
      size = slides.size
      to = new_index.negative? ? new_index + size : new_index
      raise SlideNotFoundError, "#{@path}: destination index #{new_index} out of range (#{size} slides)" unless (0...size).cover?(to)
      return slides if slides.index(target) == to

      others = slides - [target]
      node = @sld_id_nodes.fetch(target)
      node.unlink
      if to.zero?
        @sld_id_nodes.fetch(others.first).add_previous_sibling(node)
      else
        @sld_id_nodes.fetch(others[to - 1]).add_next_sibling(node)
      end

      @slides = others.insert(to, target)
      @package.touch(@presentation_part)
      @slides
    end

    # Removes a slide from the presentation, together with its notes page.
    #
    # The slide's +<p:sldId>+, its relationship from the presentation, its
    # entries in custom shows and sections, its content-type override, and
    # the slide/notes parts themselves are dropped. Shared parts (layouts,
    # media) stay in the package even if no other slide uses them.
    #
    # Other parts may still point at the slide — typically "jump to slide"
    # hyperlinks on an agenda slide. By default deletion is then refused,
    # because removing the slide would leave a dangling reference that
    # PowerPoint reports as a corrupt file. Pass +unlink: true+ to remove
    # those hyperlinks (+<a:hlinkClick>+ / +<a:hlinkHover>+; the link text
    # stays) and their relationships first. References that are not
    # hyperlinks always cause a refusal.
    #
    # @param slide_or_index [Slide, Integer]
    # @param unlink [Boolean] remove hyperlinks on other parts that target
    #   this slide
    # @return [Slide] the removed slide (detached; do not use it further)
    # @raise [SlideNotFoundError] if the index is out of range
    # @raise [Rbpptx::Error] if another part references the slide and the
    #   reference cannot be removed
    def delete_slide(slide_or_index, unlink: false)
      target = resolve_slide(slide_or_index)
      part = target.part_name
      notes_part = @package.related(part, NOTES_SLIDE_REL_TYPE).first

      referrers = @package.referrers_of(part) - [@presentation_part, part, notes_part]
      unless referrers.empty?
        unless unlink
          raise Error, "#{@path}: cannot delete #{part}; it is still referenced by #{referrers.join(', ')} " \
                       "(pass unlink: true to remove hyperlinks pointing at it)"
        end
        unlink_references!(referrers, part)
      end

      node = @sld_id_nodes.delete(target)
      sld_id = node["id"]
      rid = node.attribute_with_ns("id", R_NS)&.value
      node.remove

      pres_doc = @package.xml(@presentation_part)
      pres_doc.xpath("//p14:sldId[@id=$id]", { "p14" => P14_NS }, { "id" => sld_id }).each(&:remove)
      pres_doc.xpath("//p:custShow/p:sldLst/p:sld[@r:id=$rid]", { "p" => P_NS, "r" => R_NS }, { "rid" => rid }).each(&:remove)
      @package.touch(@presentation_part)

      @package.remove_relationship(@presentation_part, rid)
      @package.delete(part)
      @package.delete(notes_part) if notes_part

      @slides.delete(target)
      target
    end

    # Writes the presentation out to +path+, preserving every part that has
    # not been mutated byte-for-byte. Edited parts are re-serialized from
    # their in-memory Nokogiri documents; all other entries (themes, masters,
    # layouts, media, fonts) are streamed straight from the source ZIP
    # without re-parsing, so unknown PowerPoint features round-trip
    # unchanged.
    #
    # +path+ may equal the original load path; the new file is written to a
    # temp file in the same directory and atomically renamed into place. On
    # success, dirty flags on each slide are cleared.
    #
    # @param path [String, #to_path]
    # @return [String] the path that was written
    def save(path)
      ensure_open!
      out_path = @package.save(path.to_s)
      (@slides || []).each(&:clear_dirty!)
      out_path
    end

    # Releases the underlying ZIP file. Idempotent.
    #
    # @return [Boolean] +true+ on the first call, +false+ on subsequent calls
    def close
      return false if @closed

      @package&.close
      @package = nil
      @closed = true
      true
    end

    # @return [Boolean]
    def closed?
      @closed
    end

    # Creates an empty notes page for +slide+ from the notes master.
    #
    # @api private
    # @return [NotesSlide]
    # @raise [Rbpptx::Error] if the presentation has no notes master
    def create_notes_slide_for(slide)
      ensure_open!
      slides # ensures @presentation_part is resolved
      master = @package.related(@presentation_part, NOTES_MASTER_REL_TYPE).first
      unless master
        raise Error, "#{@path}: presentation has no notes master; cannot create a notes slide " \
                     "(add speaker notes to any slide in PowerPoint once to create one)"
      end

      n = 1
      n += 1 while @package.part?("ppt/notesSlides/notesSlide#{n}.xml")
      part = "ppt/notesSlides/notesSlide#{n}.xml"

      @package.add(part, Nokogiri::XML(new_notes_slide_xml))
      @package.add_relationship(part, NOTES_MASTER_REL_TYPE, master)
      @package.add_relationship(part, SLIDE_REL_TYPE, slide.part_name)
      @package.add_relationship(slide.part_name, NOTES_SLIDE_REL_TYPE, part)
      @package.add_content_type_override(part, NOTES_SLIDE_CONTENT_TYPE)

      NotesSlide.new(package: @package, part_name: part)
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

    def resolve_slide(slide_or_index)
      case slide_or_index
      when Integer then slide(slide_or_index)
      when Slide
        slides.find { |s| s.equal?(slide_or_index) } or
          raise SlideNotFoundError, "#{@path}: slide #{slide_or_index.part_name} is not part of this presentation"
      else
        raise TypeError, "expected Rbpptx::Slide or Integer, got #{slide_or_index.class}"
      end
    end

    HYPERLINK_ELEMENTS = %w[hlinkClick hlinkHover].freeze
    private_constant :HYPERLINK_ELEMENTS

    # Removes hyperlinks in +referrers+ that target +part+. Everything is
    # validated before anything is mutated, so a refusal leaves the package
    # untouched.
    def unlink_references!(referrers, part)
      plan = referrers.map do |source|
        doc = @package.xml(source)
        rels = @package.relationships(source).select { |r| r.part == part }
        uses = rels.to_h do |rel|
          attrs = doc.xpath("//@*[namespace-uri()=$ns and .=$rid]", nil, { "ns" => R_NS, "rid" => rel.id })
          foreign = attrs.reject { |a| a.parent.namespace&.href == A_NS && HYPERLINK_ELEMENTS.include?(a.parent.name) }
          unless foreign.empty?
            raise Error, "#{@path}: cannot delete #{part}; #{source} references it from " \
                         "<#{foreign.first.parent.name}>, which is not a hyperlink"
          end
          [rel.id, attrs.map(&:parent)]
        end
        [source, uses]
      end

      plan.each do |source, uses|
        uses.each do |rid, elements|
          elements.each(&:remove)
          @package.remove_relationship(source, rid)
        end
        owner = slides.find { |s| s.part_name == source }
        owner ? owner.mark_dirty! : @package.touch(source)
      end
    end

    def build_slides
      @presentation_part = locate_presentation_part
      pres_rels = @package.relationships(@presentation_part).to_h { |r| [r.id, r] }

      doc = @package.xml(@presentation_part)
      raise PresentationFormatError, "#{@path}: #{@presentation_part} has no root element" unless doc.root

      @sld_id_nodes = {}.compare_by_identity
      doc.xpath("/p:presentation/p:sldIdLst/p:sldId", "p" => P_NS).map do |sld_id|
        rid = sld_id.attribute_with_ns("id", R_NS)&.value
        rel = pres_rels[rid]
        raise PresentationFormatError, "#{@path}: rId #{rid.inspect} not found in presentation rels" unless rel&.part

        slide = Slide.new(package: @package, part_name: rel.part, presentation: self)
        @sld_id_nodes[slide] = sld_id
        slide
      end
    end

    def locate_presentation_part
      doc = @package.xml("_rels/.rels")
      rel = doc.at_xpath(
        "/pkg:Relationships/pkg:Relationship[@Type=$type]",
        { "pkg" => PKG_REL_NS },
        { "type" => OFFICE_DOC_REL_TYPE }
      )
      raise PresentationFormatError, "#{@path}: officeDocument relationship missing from _rels/.rels" unless rel

      target = rel["Target"] or raise PresentationFormatError, "#{@path}: officeDocument relationship has no Target"
      target.sub(%r{\A/}, "")
    end

    def new_notes_slide_xml
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:notes xmlns:a="#{A_NS}" xmlns:r="#{R_NS}" xmlns:p="#{P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>
      XML
    end
  end
end
