require 'rxl/version'
require_relative 'excel'
require_relative 'worksheet'

module Rxl

  def self.write_file(filepath, file_hash)
    xl = Excel.new(file_hash)
    xl.save_file(filepath)
  end

  def self.read_file(filepath)
    xl = Excel.new(filepath)
    xl.hash_workbook
  end

  def self.read_file_as_tables(filepath)
    raw_hash = read_file(filepath)
    processed_hash = {}
    raw_hash.each do |k, v|
      processed_hash[k] = Worksheet.hash_worksheet_to_hash_table(v)
    end
    processed_hash
  end

end
