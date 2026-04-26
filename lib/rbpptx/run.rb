require "cgi"

module Rbpptx
  # A single run (+<a:r>+) inside a paragraph.
  #
  # Runs are the smallest unit of text that shares a single set of run
  # properties (+<a:rPr>+: font, size, color, language, hyperlink, ...). The
  # text setter replaces only the content of the run's +<a:t>+ element, so
  # every other property of the run survives the edit untouched — the
  # surgical counterpart to the bulldozer-style {Shape#text=} setter.
  class Run
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze

    # @param node [Nokogiri::XML::Element] the +<a:r>+ element
    # @param slide [Slide, nil] back-reference used to mark the slide dirty
    #   when this run is mutated
    def initialize(node:, slide: nil)
      @node = node
      @slide = slide
    end

    # @return [String] the run's text content (the inner text of +<a:t>+),
    #   or +""+ if the run has no +<a:t>+ element
    def text
      t = @node.at_xpath("./a:t", "a" => A_NS)
      t ? t.text : ""
    end

    # Replaces the run's text content. Only the +<a:t>+ element's character
    # data is rewritten; +<a:rPr>+ and any other children of the run are left
    # in place. If the run has no +<a:t>+ element, one is appended.
    #
    # @param new_text [String]
    # @return [String]
    def text=(new_text)
      value = new_text.to_s
      t = @node.at_xpath("./a:t", "a" => A_NS)
      if t
        t.content = value
      else
        @node.add_child("<a:t>#{CGI.escapeHTML(value)}</a:t>")
      end
      @slide&.mark_dirty!
      value
    end
  end
end
