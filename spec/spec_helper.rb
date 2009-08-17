require File.join( File.expand_path(File.dirname(__FILE__)), "..", "lib", "poiyer")
require 'spec'

module Poiyer::Helper
  def filename
    "test.xls"
  end

  def cleanup_test_file
    File.delete(filename) if File.exists?(filename)
  end
end
