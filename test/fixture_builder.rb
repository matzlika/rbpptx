# frozen_string_literal: true

require "zip"

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
        end
        path
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
            <p:sldIdLst>
              <p:sldId id="256" r:id="rId1"/>
              <p:sldId id="257" r:id="rId2"/>
              <p:sldId id="258" r:id="rId3"/>
            </p:sldIdLst>
            <p:sldSz cx="9144000" cy="6858000"/>
            <p:notesSz cx="6858000" cy="9144000"/>
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
                    <a:p><a:r><a:rPr lang="en-US"/><a:t>plain shape, no placeholder</a:t></a:r></a:p>
                  </p:txBody>
                </p:sp>
              </p:spTree>
            </p:cSld>
          </p:sld>
        XML
      end
    end
  end
end
