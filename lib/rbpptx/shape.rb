require "cgi"

module Rbpptx
  # A single shape (+<p:sp>+) on a slide.
  #
  # Currently exposes the shape's display name, its concatenated text, and a
  # text setter that rewrites the shape's text body while preserving the run
  # properties (+<a:rPr>+) of the first existing run, so font/size/color from
  # the template carry through to the replacement.
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

    # Concatenated text of the shape's text body. Paragraphs are joined with
    # +"\n"+; soft line breaks (+<a:br>+) inside a paragraph are also rendered
    # as +"\n"+.
    #
    # @return [String]
    def text
      paragraphs = @node.xpath("./p:txBody/a:p", "p" => P_NS, "a" => A_NS)
      paragraphs.map { |p| paragraph_text(p) }.join("\n")
    end

    # Replaces the shape's text body with +new_text+. Lines (separated by
    # +"\n"+) become individual +<a:p>+ paragraphs. Run properties from the
    # first existing +<a:rPr>+ (font, size, color, language) are cloned onto
    # each new run so that template formatting survives the rewrite.
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

    def paragraph_text(paragraph)
      buf = +""
      paragraph.children.each do |child|
        next unless child.namespace&.href == A_NS

        case child.name
        when "r"
          t = child.at_xpath("./a:t", "a" => A_NS)
          buf << t.text if t
        when "br"
          buf << "\n"
        end
      end
      buf
    end

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
