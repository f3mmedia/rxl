require 'rxl/version'
require 'rubyXL'
require_relative 'workbook'

module Rxl

  def self.write_file(filepath, hash_workbook)
    begin
      rubyxl_workbook = Workbook.hash_workbook_to_rubyxl_workbook(hash_workbook)
    rescue => e
      return e
    end
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

  def self.read_files(filepaths_hash, read_style = nil)
    return_hash = {}
    filepaths_hash.each do |key, value|
      if read_style == :as_tables
        return_hash[key] = read_file_as_tables(value)
      else
        return_hash[key] = read_file(value)
      end
    end
    return_hash
  end

end
