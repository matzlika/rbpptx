require "cgi"

module Rbpptx
  # A single shape (+<p:sp>+) on a slide.
  #
  # Two layers of text access are exposed:
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
  # In addition, {#placeholder_type} reads the shape's placeholder role from
  # +<p:nvSpPr>/<p:nvPr>/<p:ph type="...">+ so callers can target shapes by
  # semantic role (e.g. +:ctrTitle+, +:subTitle+, +:body+) instead of by
  # index.
  class Shape
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main".freeze
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze

    # @param node [Nokogiri::XML::Element] the +<p:sp>+ element
    # @param slide [Slide, nil] back-reference used to mark the slide dirty
    #   when this shape is mutated
    def initialize(node:, slide: nil)
      @node = node
      @slide = slide
    end

    # @return [String, nil] value of +<p:cNvPr name="...">+, the human-readable
    #   shape name set by the authoring tool
    def name
      @node.at_xpath("./p:nvSpPr/p:cNvPr/@name", "p" => P_NS)&.value
    end

    # @return [Symbol, nil] the placeholder type as a Symbol (e.g.
    #   +:ctrTitle+, +:subTitle+, +:title+, +:body+, +:obj+, +:ftr+,
    #   +:dt+, +:sldNum+, +:pic+), or +nil+ if the shape is not a
    #   placeholder. Returns +:body+ when the shape has a +<p:ph>+ element
    #   with no +type+ attribute, matching the OOXML default.
    def placeholder_type
      ph = @node.at_xpath("./p:nvSpPr/p:nvPr/p:ph", "p" => P_NS)
      return nil unless ph

      type = ph["type"]
      type ? type.to_sym : :body
    end

    # @return [Array<Paragraph>] paragraphs in the shape's text body, in
    #   document order. Returns +[]+ if the shape has no +<p:txBody>+ (e.g.
    #   picture or group shapes).
    def paragraphs
      txbody = @node.at_xpath("./p:txBody", "p" => P_NS)
      return [] unless txbody

      txbody.xpath("./a:p", "a" => A_NS).map { |p| Paragraph.new(node: p, slide: @slide) }
    end

    # @return [Array<Run>] all runs in the shape's text body, flattened
    #   across paragraphs in document order
    def runs
      paragraphs.flat_map(&:runs)
    end

    # Concatenated text of the shape's text body. Paragraphs are joined with
    # +"\n"+; soft line breaks (+<a:br>+) inside a paragraph are also
    # rendered as +"\n"+.
    #
    # @return [String]
    def text
      paragraphs.map(&:text).join("\n")
    end

    # Replaces the shape's text body. Lines (separated by +"\n"+) become
    # individual +<a:p>+ paragraphs. Run properties from the first existing
    # +<a:rPr>+ (font, size, color, language) are cloned onto each new run
    # so template formatting carries through.
    #
    # This is the bulldozer-level setter; for edits that need to preserve
    # individual run formatting (e.g. replacing a placeholder inside a
    # multi-run paragraph), use {#runs} or {#paragraphs} and {Run#text=}
    # instead.
    #
    # Shapes without a +<p:txBody>+ (e.g. picture or group shapes) cannot
    # accept text and will raise.
    #
    # @param new_text [String]
    # @return [String] the value that was set
    # @raise [Rbpptx::Error] if the shape has no text body
    def text=(new_text)
      txbody = @node.at_xpath("./p:txBody", "p" => P_NS)
      raise Error, "shape has no <p:txBody>; cannot set text" unless txbody

      template_rpr = txbody.at_xpath(".//a:r/a:rPr", "a" => A_NS)
      txbody.xpath("./a:p", "a" => A_NS).each(&:remove)

      lines = new_text.to_s.empty? ? [""] : new_text.to_s.split("\n", -1)
      txbody.add_child(build_paragraphs_xml(lines))

      if template_rpr
        txbody.xpath("./a:p/a:r", "a" => A_NS).each do |run|
          first_child = run.children.first
          first_child ? first_child.add_previous_sibling(template_rpr.dup) : run.add_child(template_rpr.dup)
        end
      end

      @slide&.mark_dirty!
      new_text
    end

    private

    def build_paragraphs_xml(lines)
      lines.map do |line|
        if line.empty?
          "<a:p/>"
        else
          "<a:p><a:r><a:t>#{CGI.escapeHTML(line)}</a:t></a:r></a:p>"
        end
      end.join
    end
  end
end
