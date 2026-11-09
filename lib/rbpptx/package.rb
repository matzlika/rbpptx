require "pathname"
require "set"

module Rbpptx
  # The OPC (Open Packaging Conventions) container underneath a presentation.
  #
  # +Package+ owns the open ZIP and is the single place that knows how parts
  # are read, mutated, added, removed and written back. Parts are parsed on
  # demand and cached; once a caller signals a mutation via {#touch}, the
  # cached document is serialized in place of the original bytes on save.
  # Everything else is streamed through with +copy_raw_entry+, so parts the
  # gem never looked at round-trip byte-for-byte.
  #
  # @api private
  class Package
    PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships".freeze
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types".freeze
    CONTENT_TYPES_PART = "[Content_Types].xml".freeze
    SAVE_OPTIONS = Nokogiri::XML::Node::SaveOptions::AS_XML

    # A single +<Relationship>+ entry, with its target already resolved to an
    # absolute part name (+nil+ for external targets such as hyperlinks).
    Relationship = Struct.new(:id, :type, :target, :part, :external, keyword_init: true)

    attr_reader :path

    def initialize(path)
      @path = path
      @zip = Zip::File.open(path)
      @docs = {}
      @modified = Set.new
      @deleted = Set.new
    end

    def close
      @zip&.close
      @zip = nil
    end

    # @return [Boolean] whether +name+ exists in the package as it would be
    #   saved right now
    def part?(name)
      return false if @deleted.include?(name)

      @docs.key?(name) || !@zip.find_entry(name).nil?
    end

    # @return [Array<String>] every part name as it would be saved right now
    def part_names
      names = @zip.entries.reject(&:directory?).map(&:name).reject { |n| @deleted.include?(n) }
      names | @modified.to_a
    end

    # Parses (and caches) the XML part +name+.
    #
    # @raise [error_class] if the part is missing or not well-formed
    def xml(name, error_class: PresentationFormatError)
      @docs[name] ||= begin
        entry = @deleted.include?(name) ? nil : @zip.find_entry(name)
        raise error_class, "#{@path}: missing part #{name}" unless entry

        doc = Nokogiri::XML(entry.get_input_stream.read)
        raise error_class, "#{@path}: #{name}: #{doc.errors.first}" unless doc.errors.empty?

        doc
      end
    end

    # Marks +name+ as mutated so its cached document is written on save.
    def touch(name)
      @modified << name
      true
    end

    # Adds a brand-new part (or replaces a deleted one).
    def add(name, doc)
      @deleted.delete(name)
      @docs[name] = doc
      @modified << name
      doc
    end

    # Removes +name+, its relationships part, and its content-type override.
    def delete(name)
      [name, rels_part_for(name)].each do |n|
        @docs.delete(n)
        @modified.delete(n)
        @deleted << n
      end
      remove_content_type_override(name)
    end

    # ---------- relationships ----------

    # @return [String] the rels part name for +part+, e.g.
    #   +ppt/slides/_rels/slide1.xml.rels+
    def rels_part_for(part)
      dir = File.dirname(part)
      base = File.basename(part)
      dir == "." ? "_rels/#{base}.rels" : "#{dir}/_rels/#{base}.rels"
    end

    # @return [Array<Relationship>] relationships whose source is +part+
    def relationships(part)
      rels_part = rels_part_for(part)
      return [] unless part?(rels_part)

      xml(rels_part).xpath("/pkg:Relationships/pkg:Relationship", "pkg" => PKG_REL_NS).map do |r|
        external = r["TargetMode"] == "External"
        Relationship.new(
          id: r["Id"], type: r["Type"], target: r["Target"],
          part: external ? nil : resolve(part, r["Target"].to_s), external: external
        )
      end
    end

    # @return [Array<String>] internal targets of +part+'s relationships of
    #   type +rel_type+
    def related(part, rel_type)
      relationships(part).select { |r| r.type == rel_type && r.part }.map(&:part)
    end

    # @return [Array<String>] parts that hold an internal relationship
    #   pointing at +target_part+
    def referrers_of(target_part)
      part_names.grep(%r{(\A|/)_rels/[^/]+\.rels\z}).filter_map do |rels_part|
        source = source_part_for(rels_part)
        source if relationships(source).any? { |r| r.part == target_part }
      end
    end

    # Adds a relationship from +source+ to +target_part+, creating the rels
    # part if needed.
    #
    # @return [String] the new relationship id
    def add_relationship(source, rel_type, target_part)
      rels_part = rels_part_for(source)
      unless part?(rels_part)
        add(rels_part, Nokogiri::XML(<<~XML))
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="#{PKG_REL_NS}"/>
        XML
      end

      doc = xml(rels_part)
      root = doc.root
      taken = root.xpath("./pkg:Relationship/@Id", "pkg" => PKG_REL_NS).map(&:value)
      n = 1
      n += 1 while taken.include?("rId#{n}")
      rid = "rId#{n}"

      rel = doc.create_element("Relationship", "Id" => rid, "Type" => rel_type,
                                               "Target" => relative_target(source, target_part))
      root.add_child(rel)
      touch(rels_part)
      rid
    end

    def remove_relationship(source, rid)
      rels_part = rels_part_for(source)
      return unless part?(rels_part)

      doc = xml(rels_part)
      doc.xpath("/pkg:Relationships/pkg:Relationship[@Id=$rid]", { "pkg" => PKG_REL_NS }, { "rid" => rid }).each(&:remove)
      touch(rels_part)
    end

    # Resolves a relationship +target+ relative to the directory of +source+.
    def resolve(source, target)
      return target.sub(%r{\A/}, "") if target.start_with?("/")

      base_dir = File.dirname(source)
      base_dir = "" if base_dir == "."
      File.expand_path(target, "/#{base_dir}").sub(%r{\A/}, "")
    end

    def relative_target(source, target_part)
      base_dir = File.dirname("/#{source}")
      Pathname.new("/#{target_part}").relative_path_from(Pathname.new(base_dir)).to_s
    end

    # ---------- content types ----------

    def add_content_type_override(part, content_type)
      doc = xml(CONTENT_TYPES_PART)
      el = doc.create_element("Override", "PartName" => "/#{part}", "ContentType" => content_type)
      doc.root.add_child(el)
      touch(CONTENT_TYPES_PART)
    end

    def remove_content_type_override(part)
      doc = xml(CONTENT_TYPES_PART)
      nodes = doc.xpath("/ct:Types/ct:Override[@PartName=$name]", { "ct" => CT_NS }, { "name" => "/#{part}" })
      return if nodes.empty?

      nodes.each(&:remove)
      touch(CONTENT_TYPES_PART)
    end

    # ---------- save ----------

    # Writes the package to +out_path+ via temp file + atomic rename.
    def save(out_path)
      tmp_path = "#{out_path}.rbpptx-tmp.#{Process.pid}.#{rand(1 << 32).to_s(16)}"
      begin
        Zip::OutputStream.open(tmp_path) do |out|
          written = Set.new
          @zip.each do |entry|
            next if entry.directory? || @deleted.include?(entry.name)

            if @modified.include?(entry.name)
              write_doc(out, entry.name)
            else
              copy_raw(out, entry)
            end
            written << entry.name
          end
          (@modified.to_a - written.to_a).each { |name| write_doc(out, name) }
        end
        File.rename(tmp_path, out_path)
      rescue StandardError
        File.unlink(tmp_path) if File.exist?(tmp_path)
        raise
      end

      # The open Zip::File re-reads entries from @path by name, so after
      # overwriting our own source the old central directory no longer
      # matches the file on disk. Reopen so later lazy reads stay valid.
      reopen if File.expand_path(out_path) == File.expand_path(@path)
      out_path
    end

    private

    DATA_DESCRIPTOR_FLAG = 0b1000
    private_constant :DATA_DESCRIPTOR_FLAG

    # Streams +entry+'s compressed bytes through unchanged.
    #
    # Entries written by streaming zippers (Google Slides exports, among
    # others) set general-purpose bit 3: "CRC and sizes follow the data in
    # a data descriptor". +copy_raw_entry+ writes the real CRC and sizes
    # into the local header but keeps that bit and emits no descriptor,
    # which PowerPoint rejects as corrupt. Clearing the bit on a copy makes
    # the header self-describing again; the payload is not recompressed.
    def copy_raw(out, entry)
      if entry.gp_flags & DATA_DESCRIPTOR_FLAG != 0
        entry = entry.dup
        entry.gp_flags &= ~DATA_DESCRIPTOR_FLAG
      end
      out.copy_raw_entry(entry)
    end

    def write_doc(out, name)
      out.put_next_entry(name)
      out.write(@docs.fetch(name).to_xml(save_with: SAVE_OPTIONS))
    end

    def reopen
      @zip.close
      @zip = Zip::File.open(@path)
      @deleted.clear
    end

    def source_part_for(rels_part)
      return "" if rels_part == "_rels/.rels"

      dir = File.dirname(rels_part)
      base = File.basename(rels_part, ".rels")
      parent = File.dirname(dir)
      parent == "." ? base : "#{parent}/#{base}"
    end
  end
end
