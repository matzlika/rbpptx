require "cgi"

module Rbpptx
  # Text access shared by everything that owns a DrawingML text body:
  # {Shape} (+<p:txBody>+) and {TableCell} (+<a:txBody>+).
  #
  # Including classes provide a private +text_body_node+ returning the text
  # body element (or +nil+), a private +text_body_label+ used in error
  # messages, and an +@slide+ owner that responds to +mark_dirty!+.
  module TextBody
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze

    # @return [Array<Paragraph>] paragraphs in the text body, in document
    #   order. Returns +[]+ when there is no text body.
    def paragraphs
      txbody = text_body_node
      return [] unless txbody

      txbody.xpath("./a:p", "a" => A_NS).map { |p| Paragraph.new(node: p, slide: @slide) }
    end

    # @return [Array<Run>] all runs in the text body, flattened across
    #   paragraphs in document order
    def runs
      paragraphs.flat_map(&:runs)
    end

    # Concatenated text of the text body. Paragraphs are joined with
    # +"\n"+; soft line breaks (+<a:br>+) inside a paragraph are also
    # rendered as +"\n"+.
    #
    # @return [String]
    def text
      paragraphs.map(&:text).join("\n")
    end

    # Replaces the text body. Lines (separated by +"\n"+) become individual
    # +<a:p>+ paragraphs. Template formatting carries through:
    #
    # * run properties from the first existing +<a:rPr>+ (font, size, color,
    #   language) are cloned onto each new run. When the body has no runs —
    #   an empty form field or table cell — the first +<a:endParaRPr>+ is
    #   used instead, since that is where PowerPoint keeps the formatting
    #   of an empty paragraph;
    # * the first paragraph's +<a:pPr>+ (alignment, indent, bullets) and
    #   +<a:endParaRPr>+ are cloned onto each new paragraph.
    #
    # This is the bulldozer-level setter; to preserve individual run
    # formatting use {#runs} and {Run#text=} instead.
    #
    # @param new_text [String]
    # @return [String] the value that was set
    # @raise [Rbpptx::Error] if there is no text body
    def text=(new_text)
      txbody = text_body_node
      raise Error, "#{text_body_label}; cannot set text" unless txbody

      template_rpr = txbody.at_xpath(".//a:r/a:rPr", "a" => A_NS)
      first_p = txbody.at_xpath("./a:p", "a" => A_NS)
      template_ppr = first_p&.at_xpath("./a:pPr", "a" => A_NS)
      template_end = first_p&.at_xpath("./a:endParaRPr", "a" => A_NS)
      template_rpr ||= txbody.at_xpath(".//a:p/a:endParaRPr", "a" => A_NS)&.dup&.tap { |n| n.name = "rPr" }
      txbody.xpath("./a:p", "a" => A_NS).each(&:remove)

      lines = new_text.to_s.empty? ? [""] : new_text.to_s.split("\n", -1)
      txbody.add_child(build_paragraphs_xml(lines))

      txbody.xpath("./a:p", "a" => A_NS).each do |para|
        para.prepend_child(template_ppr.dup) if template_ppr
        para.add_child(template_end.dup) if template_end
        next unless template_rpr

        para.xpath("./a:r", "a" => A_NS).each { |run| run.prepend_child(template_rpr.dup) }
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
