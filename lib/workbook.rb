require_relative 'worksheet'

module Workbook

  def self.rubyxl_workbook_to_hash_workbook(rubyxl_workbook)
    hash_workbook = {}
    rubyxl_workbook.each do |rubyxl_worksheet|
      hash_worksheet = {row_count: rubyxl_worksheet.count, column_count: 1, rows: {}, columns: {}, cells: {}}
      Worksheet.rubyxl_worksheet_to_hash_worksheet(rubyxl_worksheet, hash_worksheet)
      Worksheet.process_sheet_to_populated_block(hash_worksheet)
      hash_workbook[rubyxl_worksheet.sheet_name] = hash_worksheet
    end
    hash_workbook
  end

  def self.hash_workbook_to_rubyxl_workbook
    rubyxl_workbook = RubyXL::Workbook.new
    first_worksheet = true
    @hash_workbook.each do |hash_key, hash_value|
      if first_worksheet
        rubyxl_workbook.worksheets[0].sheet_name = hash_key
        first_worksheet = false
      else
        rubyxl_workbook.add_worksheet(hash_key)
      end
      Worksheet.set_hash_worksheet_defaults(hash_value)
      Worksheet.hash_worksheet_to_rubyxl_worksheet(hash_value, rubyxl_workbook[hash_key])
    end
    rubyxl_workbook
  end

  def self.validate_hash_workbook
    raise("@hash_workbook is class '#{@hash_workbook.class}', should be a Hash") unless @hash_workbook.is_a?(Hash)
    @validation = []
    @hash_workbook.each { |hash_worksheet_name, hash_worksheet| Worksheet.validate_hash_worksheet(hash_worksheet_name, hash_worksheet) }
    raise(@validation.join("\n")) unless @validation.empty?
  end

end