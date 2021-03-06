# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'rxl/version'

Gem::Specification.new do |spec|
  spec.name          = "rxl"
  spec.version       = Rxl::VERSION
  spec.authors       = ["Ian McWilliams"]
  spec.email         = ["ian.mcwilliams@f3mmedia.co.uk"]

  spec.summary       = "A ruby spreadsheet interface"
  spec.description   = <<~DESC
    Implements functionality written with Excel users in mind for straight reading and writing of sheets.
    Row and column scope values are written to only the cells used; cells are specified by their Excel (A1) ID.
  DESC
  spec.homepage      = "https://github.com/f3mmedia/rxl"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.12"
  spec.add_development_dependency "rake", "~> 10.0"
  spec.add_development_dependency "rspec", "~> 3.0"
end
