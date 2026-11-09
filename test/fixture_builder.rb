# frozen_string_literal: true

require "zip"
require "zlib"

module Rbpptx
  module Test
    # Builds minimal +.pptx+ fixtures from scratch using +rubyzip+ and
    # hand-coded OOXML, so the test suite never depends on a binary fixture
    # whose provenance we cannot fully account for. The structure is the
    # smallest one rbpptx will open: +[Content_Types].xml+, +_rels/.rels+,
    # +ppt/presentation.xml+, its rels, and one +ppt/slides/slideN.xml+ per
    # slide.
    #
    # The default fixture exercises every read/edit code path the gem
    # exposes today:
    #
    # * three slides in presentation order (so rels resolution is non-trivial)
    # * a slide with two placeholders (+:ctrTitle+, +:subTitle+)
    # * a text shape with two paragraphs
    # * a paragraph with two runs separated by an +<a:br>+ soft line break
    # * a run with a non-trivial +<a:rPr>+ (font, size, color, language)
    # * a shape without a +<p:txBody>+ (mimics a picture or group shape)
    # * a placeholder of type +:body+
    # * a notes master, and a notes page on slide 1 only (so both reading
    #   existing notes and creating a notes page are exercised)
    # * a group shape with a nested group on slide 3
    # * a 2x2 table on slide 3 whose second row is a horizontal merge; the
    #   covered cell is empty and keeps its formatting only in +<a:pPr>+ and
    #   +<a:endParaRPr>+, like an empty form field
    # * a "jump to slide 2" hyperlink on slide 3 (blocks naive deletion)
    # * a p14 section list referencing every slide id
    #
    # Inspired by the test-case taxonomy in python-pptx's +tests/unit/+
    # directory (MIT) — same structural concerns, our own bytes.
    module FixtureBuilder
      MAIN_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
      DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
      REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
      PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
      OFFICE_DOC_REL_TYPE = "#{REL_NS}/officeDocument"
      SLIDE_REL_TYPE = "#{REL_NS}/slide"
      NOTES_SLIDE_REL_TYPE = "#{REL_NS}/notesSlide"
      NOTES_MASTER_REL_TYPE = "#{REL_NS}/notesMaster"
      P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"

      module_function

      # Writes a minimal three-slide +.pptx+ to +path+.
      def build(path)
        Zip::OutputStream.open(path) do |out|
          write_entry(out, "[Content_Types].xml", content_types_xml)
          write_entry(out, "_rels/.rels", root_rels_xml)
          write_entry(out, "ppt/presentation.xml", presentation_xml)
          write_entry(out, "ppt/_rels/presentation.xml.rels", presentation_rels_xml)
          write_entry(out, "ppt/slides/slide1.xml", slide1_xml)
          write_entry(out, "ppt/slides/slide2.xml", slide2_xml)
          write_entry(out, "ppt/slides/slide3.xml", slide3_xml)
          write_entry(out, "ppt/slides/_rels/slide1.xml.rels", slide1_rels_xml)
          write_entry(out, "ppt/slides/_rels/slide3.xml.rels", slide3_rels_xml)
          write_entry(out, "ppt/notesMasters/notesMaster1.xml", notes_master_xml)
          write_entry(out, "ppt/notesSlides/notesSlide1.xml", notes_slide1_xml)
          write_entry(out, "ppt/notesSlides/_rels/notesSlide1.xml.rels", notes_slide1_rels_xml)
        end
        path
      end

      # Re-packs the ZIP at +src+ the way streaming zippers (e.g. Google
      # Slides exports) do: general-purpose bit 3 set, CRC and sizes zeroed
      # in each local header, and a data descriptor after each payload.
      def build_streaming_copy(src, dst)
        entries = Zip::File.open(src) do |zf|
          zf.entries.reject(&:directory?).map { |e| [e.name, e.get_input_stream.read] }
        end

        local = +"".b
        central = +"".b
        entries.each do |name, content|
          name = name.b
          data = Zlib::Deflate.new(Zlib::DEFAULT_COMPRESSION, -Zlib::MAX_WBITS).deflate(content, Zlib::FINISH)
          crc = Zlib.crc32(content)
          flags = 0x0808 # bit 3: data descriptor, bit 11: UTF-8 names
          offset = local.bytesize
          local << [0x04034b50, 20, flags, 8, 0, 0x21, 0, 0, 0, name.bytesize, 0].pack("VvvvvvVVVvv") << name << data
          local << [0x08074b50, crc, data.bytesize, content.bytesize].pack("VVVV")
          central << [0x02014b50, 20, 20, flags, 8, 0, 0x21, crc, data.bytesize, content.bytesize,
                      name.bytesize, 0, 0, 0, 0, 0, offset].pack("VvvvvvvVVVvvvvvVV") << name
        end
        eocd = [0x06054b50, 0, 0, entries.size, entries.size, central.bytesize, local.bytesize, 0].pack("VvvvvVVv")
        File.binwrite(dst, local + central + eocd)
        dst
      end

      def write_entry(out, name, body)
        out.put_next_entry(name)
        out.write(body)
      end

      def content_types_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Types xmlns="#{CONTENT_TYPES_NS}">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
            <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
            <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
            <Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
            <Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>
            <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
          </Types>
        XML
      end

      def root_rels_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="#{PACKAGE_REL_NS}">
            <Relationship Id="rId1" Type="#{OFFICE_DOC_REL_TYPE}" Target="ppt/presentation.xml"/>
          </Relationships>
        XML
      end

      def presentation_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <p:presentation xmlns:a="#{DRAW_NS}" xmlns:r="#{REL_NS}" xmlns:p="#{MAIN_NS}">
            <p:notesMasterIdLst><p:notesMasterId r:id="rId10"/></p:notesMasterIdLst>
            <p:sldIdLst>
              <p:sldId id="256" r:id="rId1"/>
              <p:sldId id="257" r:id="rId2"/>
              <p:sldId id="258" r:id="rId3"/>
            </p:sldIdLst>
            <p:sldSz cx="9144000" cy="6858000"/>
            <p:notesSz cx="6858000" cy="9144000"/>
            <p:extLst>
              <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
                <p14:sectionLst xmlns:p14="#{P14_NS}">
                  <p14:section name="Default Section" id="{00000000-0000-0000-0000-000000000001}">
                    <p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/><p14:sldId id="258"/></p14:sldIdLst>
                  </p14:section>
                </p14:sectionLst>
              </p:ext>
            </p:extLst>
          </p:presentation>
        XML
      end

      def presentation_rels_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="#{PACKAGE_REL_NS}">
            <Relationship Id="rId1" Type="#{SLIDE_REL_TYPE}" Target="slides/slide1.xml"/>
            <Relationship Id="rId2" Type="#{SLIDE_REL_TYPE}" Target="slides/slide2.xml"/>
            <Relationship Id="rId3" Type="#{SLIDE_REL_TYPE}" Target="slides/slide3.xml"/>
            <Relationship Id="rId10" Type="#{NOTES_MASTER_REL_TYPE}" Target="notesMasters/notesMaster1.xml"/>
          </Relationships>
        XML
      end

      # Slide 1: a title slide with two placeholders (ctrTitle + subTitle),
      # each holding a single-run paragraph.
      def slide1_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <p:sld xmlns:a="#{DRAW_NS}" xmlns:r="#{REL_NS}" xmlns:p="#{MAIN_NS}">
            <p:cSld>
              <p:spTree>
                <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                <p:sp>
                  <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1"/>
                    <p:cNvSpPr txBox="1"/>
                    <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
                  </p:nvSpPr>
                  <p:spPr/>
                  <p:txBody>
                    <a:bodyPr/>
                    <a:p><a:r><a:rPr lang="en-US"/><a:t>Hello rbpptx</a:t></a:r></a:p>
                  </p:txBody>
                </p:sp>
                <p:sp>
                  <p:nvSpPr>
                    <p:cNvPr id="3" name="Subtitle 2"/>
                    <p:cNvSpPr txBox="1"/>
                    <p:nvPr><p:ph idx="1" type="subTitle"/></p:nvPr>
                  </p:nvSpPr>
                  <p:spPr/>
                  <p:txBody>
                    <a:bodyPr/>
                    <a:p><a:r><a:rPr lang="en-US"/><a:t>fixture subtitle</a:t></a:r></a:p>
                  </p:txBody>
                </p:sp>
              </p:spTree>
            </p:cSld>
          </p:sld>
        XML
      end

      # Slide 2: a body placeholder with two paragraphs. The first paragraph
      # has two runs separated by a soft line break <a:br>; the first run
      # carries a non-trivial <a:rPr> with font/size/color so reload tests
      # can assert that surgical edits preserve those properties.
      def slide2_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <p:sld xmlns:a="#{DRAW_NS}" xmlns:r="#{REL_NS}" xmlns:p="#{MAIN_NS}">
            <p:cSld>
              <p:spTree>
                <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                <p:sp>
                  <p:nvSpPr>
                    <p:cNvPr id="4" name="Body 1"/>
                    <p:cNvSpPr txBox="1"/>
                    <p:nvPr><p:ph type="body"/></p:nvPr>
                  </p:nvSpPr>
                  <p:spPr/>
                  <p:txBody>
                    <a:bodyPr/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="2400" b="1">
                          <a:solidFill><a:srgbClr val="2E74B5"/></a:solidFill>
                          <a:latin typeface="Verdana"/>
                        </a:rPr>
                        <a:t>line one</a:t>
                      </a:r>
                      <a:br><a:rPr lang="en-US"/></a:br>
                      <a:r><a:rPr lang="en-US"/><a:t>line two</a:t></a:r>
                    </a:p>
                    <a:p>
                      <a:r><a:rPr lang="en-US"/><a:t>second paragraph</a:t></a:r>
                    </a:p>
                  </p:txBody>
                </p:sp>
              </p:spTree>
            </p:cSld>
          </p:sld>
        XML
      end

      # Slide 3: one shape without <p:txBody> (mimics a picture / group),
      # plus one regular text shape, to cover the no-text-body branch in
      # readers and setters.
      def slide3_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <p:sld xmlns:a="#{DRAW_NS}" xmlns:r="#{REL_NS}" xmlns:p="#{MAIN_NS}">
            <p:cSld>
              <p:spTree>
                <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                <p:sp>
                  <p:nvSpPr><p:cNvPr id="5" name="NoTextShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                </p:sp>
                <p:sp>
                  <p:nvSpPr><p:cNvPr id="6" name="Plain Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                  <p:spPr/>
                  <p:txBody>
                    <a:bodyPr/>
                    <a:p><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId1" action="ppaction://hlinksldjump"/></a:rPr><a:t>plain shape, no placeholder</a:t></a:r></a:p>
                  </p:txBody>
                </p:sp>
                <p:grpSp>
                  <p:nvGrpSpPr><p:cNvPr id="7" name="Outer Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                  <p:grpSpPr/>
                  <p:sp>
                    <p:nvSpPr><p:cNvPr id="8" name="Grouped Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                    <p:spPr/>
                    <p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>in group</a:t></a:r></a:p></p:txBody>
                  </p:sp>
                  <p:grpSp>
                    <p:nvGrpSpPr><p:cNvPr id="9" name="Inner Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                    <p:grpSpPr/>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="10" name="Nested Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                      <p:spPr/>
                      <p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>in nested group</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                  </p:grpSp>
                </p:grpSp>
                <p:graphicFrame>
                  <p:nvGraphicFramePr><p:cNvPr id="11" name="Table 1"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>
                  <p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="800000"/></p:xfrm>
                  <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                      <a:tbl>
                        <a:tblPr firstRow="1"/>
                        <a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>
                        <a:tr h="400000">
                          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Name</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Value</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                        </a:tr>
                        <a:tr h="400000">
                          <a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>merged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          <a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="r"/><a:endParaRPr lang="ja-JP" sz="1200"><a:latin typeface="Meiryo"/></a:endParaRPr></a:p></a:txBody><a:tcPr/></a:tc>
                        </a:tr>
                      </a:tbl>
                    </a:graphicData>
                  </a:graphic>
                </p:graphicFrame>
              </p:spTree>
            </p:cSld>
          </p:sld>
        XML
      end
      def slide1_rels_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="#{PACKAGE_REL_NS}">
            <Relationship Id="rId1" Type="#{NOTES_SLIDE_REL_TYPE}" Target="../notesSlides/notesSlide1.xml"/>
          </Relationships>
        XML
      end

      # Slide 3 links to slide 2 ("jump to slide"), like an agenda slide.
      def slide3_rels_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="#{PACKAGE_REL_NS}">
            <Relationship Id="rId1" Type="#{SLIDE_REL_TYPE}" Target="slide2.xml"/>
          </Relationships>
        XML
      end

      def notes_master_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <p:notesMaster xmlns:a="#{DRAW_NS}" xmlns:r="#{REL_NS}" xmlns:p="#{MAIN_NS}">
            <p:cSld>
              <p:spTree>
                <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr/>
              </p:spTree>
            </p:cSld>
            <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
          </p:notesMaster>
        XML
      end

      def notes_slide1_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <p:notes xmlns:a="#{DRAW_NS}" xmlns:r="#{REL_NS}" xmlns:p="#{MAIN_NS}">
            <p:cSld>
              <p:spTree>
                <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr/>
                <p:sp>
                  <p:nvSpPr><p:cNvPr id="2" name="Slide Image"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr>
                  <p:spPr/>
                </p:sp>
                <p:sp>
                  <p:nvSpPr><p:cNvPr id="3" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                  <p:spPr/>
                  <p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>speaker notes for slide 1</a:t></a:r></a:p></p:txBody>
                </p:sp>
              </p:spTree>
            </p:cSld>
          </p:notes>
        XML
      end

      def notes_slide1_rels_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="#{PACKAGE_REL_NS}">
            <Relationship Id="rId1" Type="#{NOTES_MASTER_REL_TYPE}" Target="../notesMasters/notesMaster1.xml"/>
            <Relationship Id="rId2" Type="#{SLIDE_REL_TYPE}" Target="../slides/slide1.xml"/>
          </Relationships>
        XML
      end
    end
  end
end
