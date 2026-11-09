module Rbpptx
  # A table: a +<p:graphicFrame>+ whose graphic data is an +<a:tbl>+.
  #
  # Only cell text is exposed. Grid geometry, table styles, borders and fills
  # are left untouched; adding or removing rows and columns is out of scope.
  #
  #   table = slide.tables.first
  #   table.cell(0, 1).text            # => "Q2"
  #   table.cell(1, 1).text = "42"
  #   table.rows.map { |r| r.cells.map(&:text) }
  class Table
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main".freeze
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze

    # @param node [Nokogiri::XML::Element] the +<p:graphicFrame>+ element
    # @param slide [Slide, NotesSlide, nil] owning part
    def initialize(node:, slide: nil)
      @node = node
      @slide = slide
      @tbl = node.at_xpath("./a:graphic/a:graphicData/a:tbl", "a" => A_NS)
    end

    # @return [String, nil] value of +<p:nvGraphicFramePr>/<p:cNvPr name>+
    def name
      @node.at_xpath("./p:nvGraphicFramePr/p:cNvPr/@name", "p" => P_NS)&.value
    end

    # @return [Array<TableRow>]
    def rows
      @tbl.xpath("./a:tr", "a" => A_NS).map { |tr| TableRow.new(node: tr, slide: @slide) }
    end

    # @return [Integer]
    def row_count
      @tbl.xpath("./a:tr", "a" => A_NS).size
    end

    # @return [Integer] number of grid columns (+<a:tblGrid>/<a:gridCol>+)
    def column_count
      @tbl.xpath("./a:tblGrid/a:gridCol", "a" => A_NS).size
    end

    # @param row [Integer] zero-based row index
    # @param col [Integer] zero-based column index
    # @return [TableCell, nil] +nil+ when out of range
    def cell(row, col)
      rows[row]&.cells&.[](col)
    end
  end

  # A table row (+<a:tr>+).
  class TableRow
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze

    def initialize(node:, slide: nil)
      @node = node
      @slide = slide
    end

    # @return [Array<TableCell>] one per grid column, including cells that
    #   are covered by a horizontal or vertical merge (see
    #   {TableCell#spanned?})
    def cells
      @node.xpath("./a:tc", "a" => A_NS).map { |tc| TableCell.new(node: tc, slide: @slide) }
    end
  end

  # A table cell (+<a:tc>+). Text access is identical to {Shape}; see
  # {TextBody}.
  class TableCell
    include TextBody

    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze

    def initialize(node:, slide: nil)
      @node = node
      @slide = slide
    end

    # @return [Boolean] whether this cell is the top-left origin of a merged
    #   range (+gridSpan+ or +rowSpan+ greater than 1)
    def merge_origin?
      @node["gridSpan"].to_i > 1 || @node["rowSpan"].to_i > 1
    end

    # @return [Boolean] whether this cell is hidden under a merged range
    #   (+hMerge+ or +vMerge+ set). PowerPoint does not render its text.
    def spanned?
      truthy?(@node["hMerge"]) || truthy?(@node["vMerge"])
    end

    private

    def truthy?(value)
      value == "1" || value == "true"
    end

    def text_body_node
      @node.at_xpath("./a:txBody", "a" => A_NS)
    end

    def text_body_label
      "table cell has no <a:txBody>"
    end
  end
end
