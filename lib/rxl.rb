require 'rxl/version'
require 'excel'

module Rxl

  def self.write_file(filepath, file_hash)
    xl = Excel.new(file_hash)
    xl.save_file(filepath)
  end

  def self.read_file(filepath)
    xl = Excel.new(filepath)
    xl.hash_workbook
  end

end
