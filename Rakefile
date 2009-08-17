$:.unshift('lib')
require 'rubygems'
require 'rake/clean'
require 'rake/gempackagetask'
require 'rake/rdoctask'
require 'spec/rake/spectask'
require 'lib/poiyer/version'

NAME = "poiyer"
AUTHOR = "Brian Tatnall"
VERS = Poiyer::VERSION
PKG = "#{NAME}-#{VERS}"
PKG_FILES = FileList[
 "[A-Z]*",
 "lib/**/*.rb",
 "vendor/*.jar"
]
DEVELOPMENT_DEPENDENCIES = { 'rspec' => '>= 1.1.4' }
CLEAN.include ['*.gem']

rd = Rake::RDocTask.new do |rdoc|
  rdoc.rdoc_dir = File.dirname(__FILE__) + '/doc'
  rdoc.options << 'README'
  rdoc.rdoc_files.include('README', 'lib/**/*.rb')
end

SPEC = Gem::Specification.new do |s|
  s.name = NAME
  s.version = VERS
  s.platform = Gem::Platform::CURRENT
  s.summary = "JRuby wrapper for the Apache POI Project."
  s.description = s.summary
  s.author = AUTHOR
  s.email = "btatnall@gmail.com"
  s.files = PKG_FILES
  s.has_rdoc = true
  s.rdoc_options = rd.options
  DEVELOPMENT_DEPENDENCIES.each do |gem, requirements|
    s.add_development_dependency(gem, *requirements)
  end
end

task :package => [:clean] do
  Gem::Builder.new(SPEC).build
end

task :default => :spec

task :gem_debug do
  puts SPEC.to_ruby
end

desc "Run all specs"
Spec::Rake::SpecTask.new do |t|
  t.libs << "lib"
  t.spec_files = FileList['spec/**/*_spec.rb']
  t.spec_opts = ['--options', 'spec/spec.opts']
end
