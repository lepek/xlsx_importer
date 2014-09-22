# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'xlsx_importer/version'

Gem::Specification.new do |spec|
  spec.name          = "xlsx_importer"
  spec.version       = XlsxImporter::VERSION
  spec.authors       = ["Martin Bianculli"]
  spec.email         = ["mbianculli@gmail.com"]
  spec.description   = %q{Import a xlxs file}
  spec.summary       = %q{Import a xlxs file}
  spec.homepage      = ""
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"

  spec.add_dependency "simple_xlsx_reader", "~> 1.0.1"
end
