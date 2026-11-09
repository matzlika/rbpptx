module Rbpptx
  # A single shape (+<p:sp>+) on a slide or notes slide.
  #
  # Two layers of text access are exposed (see {TextBody}):
  #
  # * {#text} / {#text=} — bulldozer-level: read the entire concatenated
  #   text body, or replace it wholesale. The setter cleans the text body's
  #   +<a:p>+ children and rebuilds them, cloning the first existing
  #   +<a:rPr>+ onto each new run so template formatting carries through.
  # * {#paragraphs} / {#runs} — surgical: walk into individual {Paragraph}
  #   and {Run} wrappers and mutate run text via {Run#text=}, which only
  #   rewrites +<a:t>+ content. Run properties, hyperlinks, soft line
  #   breaks, paragraph properties, and any unknown OOXML extensions on the
  #   shape are preserved.
  #
  # In addition, {#placeholder_type} and {#placeholder_idx} read the shape's
  # placeholder role from +<p:nvSpPr>/<p:nvPr>/<p:ph>+ so callers can target
  # shapes by semantic role (e.g. +:ctrTitle+, +:subTitle+, +:body+) instead
  # of by index. See also {Slide#placeholders} and {Slide#title}.
  class Shape
    include TextBody

    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main".freeze
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze

    # @param node [Nokogiri::XML::Element] the +<p:sp>+ element
    # @param slide [Slide, NotesSlide, nil] back-reference used to mark the
    #   owning part dirty when this shape is mutated
    def initialize(node:, slide: nil)
      @node = node
      @slide = slide
    end

    # @return [String, nil] value of +<p:cNvPr name="...">+, the human-readable
    #   shape name set by the authoring tool
    def name
      @node.at_xpath("./p:nvSpPr/p:cNvPr/@name", "p" => P_NS)&.value
    end

    # @return [Boolean] whether the shape is a placeholder
    def placeholder?
      !placeholder_node.nil?
    end

    # @return [Symbol, nil] the placeholder type as a Symbol (e.g.
    #   +:ctrTitle+, +:subTitle+, +:title+, +:body+, +:obj+, +:ftr+,
    #   +:dt+, +:sldNum+, +:pic+), or +nil+ if the shape is not a
    #   placeholder. Returns +:body+ when the shape has a +<p:ph>+ element
    #   with no +type+ attribute, matching the OOXML default.
    def placeholder_type
      ph = placeholder_node
      return nil unless ph

      type = ph["type"]
      type ? type.to_sym : :body
    end

    # @return [Integer, nil] the placeholder index (+<p:ph idx="...">+) that
    #   links the shape to its layout placeholder, or +nil+ if the shape is
    #   not a placeholder. Returns +0+ when +idx+ is omitted, matching the
    #   OOXML default.
    def placeholder_idx
      ph = placeholder_node
      return nil unless ph

      ph["idx"] ? Integer(ph["idx"], 10) : 0
    end

    private

    def placeholder_node
      @node.at_xpath("./p:nvSpPr/p:nvPr/p:ph", "p" => P_NS)
    end

    def text_body_node
      @node.at_xpath("./p:txBody", "p" => P_NS)
    end

    def text_body_label
      "shape has no <p:txBody>"
    end
  end
end
