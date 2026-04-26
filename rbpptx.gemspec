require_relative "lib/rbpptx/version"

Gem::Specification.new do |spec|
  spec.name = "rbpptx"
  spec.version = Rbpptx::VERSION
  spec.authors = ["Taro KOBAYASHI"]
  spec.email = ["taro@matzlika.co.jp"]
  spec.summary = "PowerPoint (PPTX) processing for Ruby."
  spec.description = "rbpptx is a Ruby gem for reading and writing PowerPoint .pptx files, modeled after python-pptx."
  spec.homepage = "https://github.com/matzlika/rbpptx"
  spec.license = "MIT"
  spec.required_ruby_version = ">= 3.2"
  spec.metadata = {
    "homepage_uri" => spec.homepage,
    "source_code_uri" => "https://github.com/matzlika/rbpptx",
    "bug_tracker_uri" => "https://github.com/matzlika/rbpptx/issues",
    "changelog_uri" => "https://github.com/matzlika/rbpptx/releases"
  }

  spec.files = Dir["lib/**/*.rb"] + Dir["sig/**/*.rbs"] + %w[CHANGELOG.md LICENSE.txt README.md Rakefile]
  spec.require_paths = ["lib"]

  spec.add_dependency "rubyzip", "~> 2.3"
  spec.add_dependency "nokogiri", ">= 1.19", "< 2.0"
end
