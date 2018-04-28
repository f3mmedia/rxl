require 'rubyXL'
require 'workbook'
require 'worksheet'
require 'cell'

class Excel
  attr_accessor :hash_workbook

  # class generates based on input type:
  # pass in a string to read from that file location
  # pass in a hash to set @hash_workbook to that hash
  # pass in nil for @hash_workbook to be an empty hash
  # pass in an array of strings to generate empty worksheets at initialisation where the strings are tab names
  def initialize(source=nil)
    prepare_hash_workbook(source)
  end

  def prepare_hash_workbook(source)
    if source.is_a?(Hash)
      @hash_workbook = source
    elsif source.is_a?(String)
      read_file(source)
    elsif source.nil?
      @hash_workbook = {}
    elsif source.is_a?(Array)
      @hash_workbook = {}
      source.each { |sheet_name| @hash_workbook.update({sheet_name => {}}) }
    else
      raise("source argument of class '#{source.class}' not handled by the Excel class")
    end
  end

  def save_file(filepath)
    Workbook.validate_hash_workbook
    Workbook.hash_workbook_to_rubyxl_workbook.write(filepath)
  end

  def read_file(path)
    rubyxl_workbook = RubyXL::Parser.parse(path)
    @hash_workbook = Workbook.rubyxl_workbook_to_hash_workbook(rubyxl_workbook)
  end

end
