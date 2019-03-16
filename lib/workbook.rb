require_relative 'worksheet'

module Workbook

  def self.rubyxl_to_hash(rubyxl_workbook)
    hash_workbook = {}
    rubyxl_workbook.each do |rubyxl_worksheet|
      hash_workbook[rubyxl_worksheet.sheet_name] = Worksheet.rubyxl_to_hash(rubyxl_worksheet)
    end
    hash_workbook
  end

  def self.hash_workbook_to_rubyxl_workbook(hash_workbook)
    validate_hash_workbook(hash_workbook)
    rubyxl_workbook = RubyXL::Workbook.new
    first_worksheet = true
    hash_workbook.each do |hash_key, hash_value|
      if first_worksheet
        rubyxl_workbook.worksheets[0].sheet_name = hash_key
        first_worksheet = false
      else
        rubyxl_workbook.add_worksheet(hash_key)
      end
      Worksheet.hash_worksheet_to_rubyxl_worksheet(hash_value, rubyxl_workbook[hash_key])
    end
    rubyxl_workbook
  end

  def self.hashes_to_hash_workbook(hash_tables, write_headers: true)
    hash_workbook = {}
    hash_tables.each do |k, v|
      hash_workbook[k] = Worksheet.hashes_to_hash_worksheet(v[:rows], v[:columns], v[:formats] || {}, write_headers: write_headers)
    end
    hash_workbook
  end

  def self.hash_workbook_to_hash_tables(hash_workbook)
    hash_workbook.keys.each_with_object({}) do |key, hash_tables|
      hash_tables[key] = Worksheet.hash_worksheet_to_hash_table(hash_workbook[key])
    end
  end

  def self.validate_hash_workbook(hash_workbook)
    raise('workbook must be a Hash') unless hash_workbook.is_a?(Hash)
    hash_workbook.each do |hash_worksheet_name, hash_worksheet|
      Worksheet.validate_hash_worksheet(hash_worksheet_name, hash_worksheet)
    end
  end

end