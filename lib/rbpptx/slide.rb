module Rbpptx
  # Common behavior for an XML part that holds a shape tree (+<p:cSld>/
  # <p:spTree>+): slides and notes slides.
  #
  # Parsing is deferred until the shape tree is first accessed and then
  # cached, so simply enumerating slides does not pay the per-slide parse
  # cost.
  class SlidePart
    include ShapeContainer

    # @return [String] the part name inside the package, e.g.
    #   +"ppt/slides/slide1.xml"+
    attr_reader :part_name

    # @param package [Package] the open package
    # @param part_name [String] absolute part name (no leading slash)
    # @api private
    def initialize(package:, part_name:)
      @package = package
      @part_name = part_name
      @dirty = false
    end

    # @return [Placeholders] placeholder shapes on this part, addressable as
    #   +placeholders[:body]+ (by type) or +placeholders[1]+ (by +idx+)
    def placeholders
      Placeholders.new(shapes)
    end

    # @return [Boolean] whether any shape on this part has been mutated since
    #   load or the last save
    def dirty?
      @dirty
    end

    # Marks the part as dirty. Called by {Shape} and {Run} setters; not part
    # of the public API.
    #
    # @api private
    def mark_dirty!
      @package.touch(@part_name)
      @dirty = true
    end

    # @api private
    def clear_dirty!
      @dirty = false
    end

    # @return [String] the part's XML, reflecting any in-memory mutations.
    #   The XML declaration is preserved.
    def to_xml
      doc.to_xml(save_with: Package::SAVE_OPTIONS)
    end

    private

    def doc
      @doc ||= @package.xml(@part_name, error_class: SlideFormatError)
    end

    def shape_tree_node
      doc.at_xpath("/*/p:cSld/p:spTree", "p" => P_NS)
    end

    def owner_part
      self
    end
  end

  # A single slide.
  class Slide < SlidePart
    NOTES_SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide".freeze

    # @param package [Package] the open package
    # @param part_name [String] absolute part name (no leading slash)
    # @param presentation [Presentation, nil] owner, needed to create a notes
    #   slide on demand
    # @api private
    def initialize(package:, part_name:, presentation: nil)
      super(package: package, part_name: part_name)
      @presentation = presentation
    end

    # @return [String] the slide's basename without extension, e.g. +"slide1"+
    def name
      File.basename(@part_name, ".xml")
    end

    # @return [Array<Shape>] +<p:sp>+ children of the slide's shape tree, in
    #   document order. Shapes inside groups are not included; see
    #   {#all_shapes} and {#groups}.
    def shapes
      @shapes ||= super
    end

    # @return [Shape, nil] the title placeholder (+:title+ or +:ctrTitle+),
    #   or +nil+ when the slide has none
    def title
      shapes.find { |s| %i[title ctrTitle].include?(s.placeholder_type) }
    end

    # @return [NotesSlide, nil] the slide's notes page, or +nil+ when the
    #   slide has none. Does not create one; see {#notes=}.
    def notes_slide
      return @notes_slide if defined?(@notes_slide)

      part = @package.related(@part_name, NOTES_SLIDE_REL_TYPE).first
      @notes_slide = part && NotesSlide.new(package: @package, part_name: part)
    end

    # @return [String] speaker notes text, or +""+ when the slide has no
    #   notes page
    def notes
      notes_slide&.text || ""
    end

    # Replaces the speaker notes. If the slide has no notes page yet, one is
    # created from the presentation's notes master.
    #
    # @param new_text [String]
    # @return [String]
    # @raise [Rbpptx::Error] if a notes page must be created but the
    #   presentation has no notes master, or the notes page has no body
    #   placeholder
    def notes=(new_text)
      ns = notes_slide
      unless ns
        raise Error, "#{@part_name}: cannot create a notes slide without a presentation" unless @presentation

        ns = @presentation.create_notes_slide_for(self)
        @notes_slide = ns
      end
      ns.text = new_text
    end

    # @return [Boolean] whether this slide or its notes page has been mutated
    #   since load or the last save
    def dirty?
      super || (defined?(@notes_slide) && @notes_slide&.dirty?) || false
    end

    # @api private
    def clear_dirty!
      @notes_slide&.clear_dirty! if defined?(@notes_slide)
      super
    end
  end

  # The notes page (+<p:notes>+) attached to a {Slide}.
  class NotesSlide < SlidePart
    # @return [Shape, nil] the body placeholder that holds the speaker notes
    def notes_placeholder
      shapes.find { |s| s.placeholder_type == :body }
    end

    # @return [String] the speaker notes text (+""+ if there is no body
    #   placeholder)
    def text
      notes_placeholder&.text || ""
    end

    # @param new_text [String]
    # @return [String]
    # @raise [Rbpptx::Error] if the notes page has no body placeholder
    def text=(new_text)
      ph = notes_placeholder or raise Error, "#{@part_name}: notes slide has no body placeholder"
      ph.text = new_text
    end
  end
end
