require "nokogiri"
require "zip"

require_relative "rbpptx/errors"
require_relative "rbpptx/run"
require_relative "rbpptx/paragraph"
require_relative "rbpptx/shape"
require_relative "rbpptx/slide"
require_relative "rbpptx/presentation"
require_relative "rbpptx/version"

# PowerPoint (+.pptx+) reader/writer for Ruby, modeled after +python-pptx+.
#
# Rbpptx exposes a single, unified +Presentation+ object that can be loaded
# from disk, mutated in memory, and written back out:
#
#   require "rbpptx"
#
#   pres = Rbpptx.open("deck.pptx")
#   pres.slides.each { |slide| puts slide.shapes.map(&:text).join(" / ") }
#   pres.close
#
# Block form auto-closes the underlying ZIP:
#
#   Rbpptx.open("deck.pptx") do |pres|
#     pres.slides.first.shapes.each { |s| puts s.text }
#   end
module Rbpptx
  class << self
    # Opens an existing +.pptx+ presentation.
    #
    # When a block is given, the presentation is yielded and automatically
    # closed when the block returns (or raises), mirroring the +File.open+ and
    # +Zip::File.open+ idiom.
    #
    # @param path [String, #to_path] filesystem path to a +.pptx+ file
    # @yieldparam pres [Rbpptx::Presentation] opened presentation; auto-closed
    #   when the block returns
    # @return [Rbpptx::Presentation, Object] the presentation when no block is
    #   given, otherwise the block's return value
    def open(path)
      pres = Presentation.open(path)
      return pres unless block_given?

      begin
        yield pres
      ensure
        pres.close
      end
    end

    # Creates a new empty presentation. Not yet implemented; tracked for
    # future write-from-scratch support.
    #
    # @return [Rbpptx::Presentation]
    def new
      raise NotImplementedError, "Rbpptx.new (blank presentation creation) is not implemented yet"
    end
  end
end
