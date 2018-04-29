require 'rxl/version'
require 'rubyXL'
require_relative 'workbook'

module Rxl

  def self.write_file(filepath, hash_workbook)
    rubyxl_workbook = Workbook.hash_workbook_to_rubyxl_workbook(hash_workbook)
    rubyxl_workbook.write(filepath)
  end

  def self.read_file(filepath)
    rubyxl_workbook = RubyXL::Parser.parse(filepath)
    Workbook.rubyxl_to_hash(rubyxl_workbook)
  end

  def self.read_file_as_tables(filepath)
    hash_workbook = read_file(filepath)
    Workbook.hash_workbook_to_hash_tables(hash_workbook)
  end

end
