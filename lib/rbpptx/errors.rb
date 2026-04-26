module Rbpptx
  # Base class for all errors raised by Rbpptx. Rescue this class to catch any
  # library-specific failure without catching unrelated +StandardError+
  # subclasses from the caller's code.
  class Error < StandardError; end

  # Raised by {Rbpptx.open} when the file is not a valid +.pptx+ container.
  # Most commonly fires on legacy +.ppt+ (BIFF/CFB) files — the message names
  # the detected format and suggests a conversion path rather than letting the
  # underlying ZIP parser surface an opaque error.
  class UnsupportedFormatError < Error; end

  # Raised when presentation-level XML is malformed or internally inconsistent,
  # for example when +ppt/presentation.xml+ cannot be parsed or references a
  # missing relationship target.
  class PresentationFormatError < Error; end

  # Raised when a slide XML entry cannot be parsed.
  class SlideFormatError < Error; end

  # Raised by {Rbpptx::Presentation#slide} when the requested slide index is
  # out of range.
  class SlideNotFoundError < Error; end

  # Raised when an operation is attempted against a presentation whose
  # underlying ZIP resources have already been released via +close+.
  class ClosedPresentationError < Error; end
end
