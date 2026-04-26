module Rbpptx
  # A paragraph (+<a:p>+) inside a shape's text body.
  #
  # Paragraphs hold a sequence of {Run}s separated optionally by soft line
  # breaks (+<a:br>+). The {#text} accessor mirrors {Shape#text} restricted
  # to this paragraph: runs joined as-is, breaks rendered as +"\n"+. Since
  # paragraphs map 1:1 to +<a:p>+ elements, paragraph boundaries are *not*
  # rendered as +"\n"+ here — that is {Shape#text}'s job.
  class Paragraph
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze

    # @param node [Nokogiri::XML::Element] the +<a:p>+ element
    # @param slide [Slide, nil] back-reference propagated to runs so their
    #   setters can mark the slide dirty
    def initialize(node:, slide: nil)
      @node = node
      @slide = slide
    end

    # @return [Array<Run>] +<a:r>+ children in document order. +<a:fld>+
    #   (field) and +<a:br>+ (line break) elements are not included.
    def runs
      @node.xpath("./a:r", "a" => A_NS).map { |r| Run.new(node: r, slide: @slide) }
    end

    # @return [String] concatenated text. Run text appears verbatim; +<a:br>+
    #   between runs is rendered as +"\n"+.
    def text
      buf = +""
      @node.children.each do |child|
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
  end
end
