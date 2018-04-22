require 'rxl/version'
require 'excel'

module Rxl

  def self.write_file(filepath, file_hash)
    xl = Excel.new(source: file_hash)
    xl.save_file(filepath)
  end

  def self.read_file(filepath)
    xl = Excel.new(source: filepath)
    xl.hash_workbook
  end

  def self.read_file_as_tables(filepath)
    raw_hash = read_file(filepath)
    raw_hash.each_with_object({}) do |tables, worksheet_name, raw_worksheet_hash|
      tables[worksheet_name] = Excel.hash_worksheet_to_hash_table(raw_worksheet_hash)
    end

  end

end
