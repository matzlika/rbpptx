module Rbpptx
  # Shape-tree navigation shared by {Slide}, {NotesSlide} and {GroupShape}.
  #
  # Including classes provide a private +shape_tree_node+ returning the
  # element whose children are the shapes (+<p:spTree>+ or +<p:grpSp>+),
  # and a private +owner_part+ returning the part to mark dirty on edits.
  #
  # Only the element kinds rbpptx understands are surfaced: +<p:sp>+,
  # +<p:grpSp>+, and +<p:graphicFrame>+ holding an +<a:tbl>+. Pictures,
  # connectors, charts, SmartArt and +<mc:AlternateContent>+ wrappers are
  # skipped and left untouched.
  module ShapeContainer
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main".freeze
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze

    # @return [Array<Shape>] direct +<p:sp>+ children, in document order
    def shapes
      direct_children("sp").map { |n| Shape.new(node: n, slide: owner_part) }
    end

    # @return [Array<GroupShape>] direct +<p:grpSp>+ children, in document
    #   order
    def groups
      direct_children("grpSp").map { |n| GroupShape.new(node: n, slide: owner_part) }
    end

    # @return [Array<Table>] direct +<p:graphicFrame>+ children that hold a
    #   table, in document order
    def tables
      direct_children("graphicFrame").filter_map { |n| build_table(n) }
    end

    # @return [Array<Shape>] every +<p:sp>+ in the tree, descending into
    #   groups (nested to any depth), in document order
    def all_shapes
      walk_tree(shape_tree_node, [], [])[0]
    end

    # @return [Array<Table>] every table in the tree, descending into groups,
    #   in document order
    def all_tables
      walk_tree(shape_tree_node, [], [])[1]
    end

    private

    def direct_children(local_name)
      node = shape_tree_node
      return [] unless node

      node.xpath("./p:#{local_name}", "p" => P_NS)
    end

    def build_table(frame)
      return nil unless frame.at_xpath("./a:graphic/a:graphicData/a:tbl", "a" => A_NS)

      Table.new(node: frame, slide: owner_part)
    end

    def walk_tree(node, shapes, tables)
      return [shapes, tables] unless node

      node.element_children.each do |child|
        next unless child.namespace&.href == P_NS

        case child.name
        when "sp" then shapes << Shape.new(node: child, slide: owner_part)
        when "grpSp" then walk_tree(child, shapes, tables)
        when "graphicFrame"
          table = build_table(child)
          tables << table if table
        end
      end
      [shapes, tables]
    end
  end

  # A group shape (+<p:grpSp>+). Groups nest; use {#shapes} for direct
  # children or {#all_shapes} to flatten the whole subtree.
  class GroupShape
    include ShapeContainer

    # @param node [Nokogiri::XML::Element] the +<p:grpSp>+ element
    # @param slide [Slide, NotesSlide, nil] owning part
    def initialize(node:, slide: nil)
      @node = node
      @slide = slide
    end

    # @return [String, nil] value of +<p:nvGrpSpPr>/<p:cNvPr name="...">+
    def name
      @node.at_xpath("./p:nvGrpSpPr/p:cNvPr/@name", "p" => P_NS)&.value
    end

    # @return [String] text of every shape in the group (see {#all_shapes}),
    #   one shape per line, empty shapes skipped. Read-only convenience.
    def text
      all_shapes.map(&:text).reject(&:empty?).join("\n")
    end

    private

    def shape_tree_node
      @node
    end

    def owner_part
      @slide
    end
  end

  # The placeholders of a slide, addressable by type or by index.
  #
  #   slide.placeholders[:title]   # first placeholder of type :title
  #   slide.placeholders[1]        # placeholder with <p:ph idx="1">
  #
  # Integer keys follow python-pptx: they match the placeholder's +idx+
  # attribute (see {Shape#placeholder_idx}), *not* its position in the list.
  class Placeholders
    include Enumerable

    # @param shapes [Array<Shape>]
    def initialize(shapes)
      @shapes = shapes.select(&:placeholder?)
    end

    # @yieldparam shape [Shape]
    def each(&block)
      return enum_for(:each) { size } unless block

      @shapes.each(&block)
      self
    end

    # @param key [Symbol, Integer] placeholder type, or placeholder +idx+
    # @return [Shape, nil]
    def [](key)
      case key
      when Symbol then @shapes.find { |s| s.placeholder_type == key }
      when Integer then @shapes.find { |s| s.placeholder_idx == key }
      else raise TypeError, "placeholder key must be a Symbol (type) or Integer (idx), got #{key.class}"
      end
    end

    # @return [Integer]
    def size
      @shapes.size
    end
    alias length size

    # @return [Boolean]
    def empty?
      @shapes.empty?
    end

    # @return [Array<Shape>]
    def to_a
      @shapes.dup
    end
  end
end
