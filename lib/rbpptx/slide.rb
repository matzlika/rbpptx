module Rbpptx
  # A single slide loaded from a presentation's ZIP container.
  #
  # Parsing of the slide's XML payload is deferred until {#shapes} is first
  # called and then cached, so simply enumerating slides via
  # {Presentation#slides} does not pay the per-slide parse cost.
  class Slide
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main".freeze

    # @return [String] the part name of the slide inside the package, e.g.
    #   +"ppt/slides/slide1.xml"+
    attr_reader :part_name

    # @param zip [Zip::File] the open package
    # @param part_name [String] absolute part name (no leading slash)
    def initialize(zip:, part_name:)
      @zip = zip
      @part_name = part_name
      @dirty = false
    end

    # @return [String] the slide's basename without extension, e.g. +"slide1"+
    def name
      File.basename(@part_name, ".xml")
    end

    # @return [Array<Shape>] +<p:sp>+ children of the slide's shape tree, in
    #   document order
    def shapes
      @shapes ||= build_shapes
    end

    # @return [Boolean] whether any shape on this slide has been mutated since
    #   load
    def dirty?
      @dirty
    end

    # Marks the slide as dirty. Called by {Shape} setters; not part of the
    # public API.
    #
    # @api private
    def mark_dirty!
      @dirty = true
    end

    # @return [String] the slide's XML, reflecting any in-memory mutations.
    #   The XML declaration is preserved.
    def to_xml
      doc.to_xml
    end

    # @api private
    def clear_dirty!
      @dirty = false
    end

    private

    def doc
      @doc ||= parse_doc
    end

    def parse_doc
      entry = @zip.find_entry(@part_name)
      raise SlideFormatError, "missing slide part #{@part_name}" unless entry

      parsed = Nokogiri::XML(entry.get_input_stream.read)
      raise SlideFormatError, "#{@part_name}: #{parsed.errors.first}" unless parsed.errors.empty?

      parsed
    end

    def build_shapes
      doc.xpath("/p:sld/p:cSld/p:spTree/p:sp", "p" => P_NS).map { |sp| Shape.new(node: sp, slide: self) }
    end
  end
end
