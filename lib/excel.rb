require 'rubyXL'

class Excel
  attr_accessor :hash_workbook

  # class generates based on input type:
  # pass in a string to read from that file location
  # pass in a hash to set @hash_workbook to that hash
  # pass in nil for @hash_workbook to be an empty hash
  # pass in an array of strings to generate empty worksheets at initialisation where the strings are tab names
  def initialize(source: nil)
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

  def self.hash_worksheet_to_hash_table(raw_hash)
    cells = raw_hash[:cells]
    columns = cells.keys.map { |key| key[/\D+/] }.uniq
    cells.keys.map { |key| key[/\d+/] }.uniq.map do |row_number|
      columns[1..-1].each_with_object({}) do |column_letter, this_hash|
        this_hash[cells["#{column_letter}1"]][:value] = cells["#{column_letter}#{row_number}"][:value]
      end
    end
  end

  def validate_hash_workbook
    raise("@hash_workbook is class '#{@hash_workbook.class}', should be a Hash") unless @hash_workbook.is_a?(Hash)
    @validation = []
    @hash_workbook.each { |hash_worksheet_name, hash_worksheet| validate_hash_worksheet(hash_worksheet_name, hash_worksheet) }
    raise(@validation.join("\n")) unless @validation.empty?
  end

  def validate_hash_worksheet(hash_worksheet_name, hash_worksheet)
    @validation << "hash_worksheet_name is class '#{hash_worksheet_name.class}', should be a String" unless hash_worksheet_name.is_a?(String)
    # other validation...
  end

  def save_file(filepath)
    validate_hash_workbook
    hash_workbook_to_rubyxl_workbook.write(filepath)
  end

  def hash_workbook_to_rubyxl_workbook
    rubyxl_workbook = RubyXL::Workbook.new
    first_worksheet = true
    @hash_workbook.each do |hash_key, hash_value|
      if first_worksheet
        rubyxl_workbook.worksheets[0].sheet_name = hash_key
        first_worksheet = false
      else
        rubyxl_workbook.add_worksheet(hash_key)
      end
      set_hash_worksheet_defaults(hash_value)
      hash_worksheet_to_rubyxl_worksheet(hash_value, rubyxl_workbook[hash_key])
    end
    rubyxl_workbook
  end

  def set_hash_worksheet_defaults(hash_worksheet)
    %i[worksheet columns rows].each do |key|
      hash_worksheet[key] = {} unless hash_worksheet.has_key?(key)
    end
  end

  def set_hash_worksheet_extents(hash_worksheet)
    row_keys = (hash_worksheet[:rows] || {}).keys.map { |item| "A#{item}" }
    column_keys = (hash_worksheet[:columns] || {}).keys.map { |item| "#{item}1" }
    hash_worksheet[:row_count] = 0
    hash_worksheet[:column_count] = 0
    ((hash_worksheet[:cells] || {}).keys + row_keys + column_keys).each do |hash_cell_key|
      row_index, column_index = RubyXL::Reference.ref2ind(hash_cell_key)
      hash_worksheet[:row_count] = row_index + 1 if row_index >= hash_worksheet[:row_count]
      hash_worksheet[:column_count] = column_index + 1 if column_index >= hash_worksheet[:column_count]
    end
  end

  def hash_worksheet_to_rubyxl_worksheet(hash_worksheet, rubyxl_worksheet)
    process_sheet_to_populated_block(hash_worksheet)
    (hash_worksheet[:cells] || {}).sort.each do |hash_cell_key, hash_cell|
      combined_hash_cell = get_combined_hash_cell(hash_worksheet, hash_cell_key, hash_cell)
      row_index, column_index = RubyXL::Reference.ref2ind(hash_cell_key)
      add_rubyxl_cells(combined_hash_cell, rubyxl_worksheet, row_index, column_index)
      hash_cell_to_rubyxl_cell(combined_hash_cell, rubyxl_worksheet, row_index, column_index)
    end
  end

  def get_combined_hash_cell(hash_worksheet, hash_cell_key, hash_cell)
    # first get data from the matching column if it's specified
    column_keys = hash_worksheet[:columns].keys.select { |key| hash_cell_key =~ /^#{key}\d+$/ }
    column_keys.empty? ? hash_column = {} : hash_column = hash_worksheet[:columns][column_keys[0]]
    combined_hash_cell = hash_column.merge(hash_cell)
    # then get data from the matching row if it's specified
    row_keys = hash_worksheet[:rows].keys.select { |key| hash_cell_key =~ /^\D+#{key}$/ }
    row_keys.empty? ? hash_row = {} : hash_row = hash_worksheet[:rows][row_keys[0]]
    combined_hash_cell = hash_row.merge(combined_hash_cell)
    hash_worksheet[:worksheet].merge(combined_hash_cell)
  end

  def add_rubyxl_cells(combined_hash_cell, rubyxl_worksheet, row_index, column_index)
    if combined_hash_cell[:formula]
      rubyxl_worksheet.add_cell(row_index, column_index, '', combined_hash_cell[:formula]).set_number_format combined_hash_cell[:dp_2]
    else
      rubyxl_worksheet.add_cell(row_index, column_index, combined_hash_cell[:value])
    end
  end

  def hash_cell_to_rubyxl_cell(combined_hash_cell, rubyxl_worksheet, row_index, column_index)
    merge_row_index, merge_column_index = RubyXL::Reference.ref2ind(combined_hash_cell[:merge])

    rubyxl_worksheet.merge_cells(row_index, column_index, merge_column_index, merge_row_index) if combined_hash_cell[:merge]
    rubyxl_worksheet.change_column_width(column_index, combined_hash_cell[:width])  if combined_hash_cell[:width]

    rubyxl_worksheet[row_index][column_index].change_font_name(combined_hash_cell[:font_style]) if combined_hash_cell[:font_style]
    rubyxl_worksheet[row_index][column_index].change_font_size(combined_hash_cell[:font_size]) if combined_hash_cell[:font_size]
    rubyxl_worksheet[row_index][column_index].change_fill(combined_hash_cell[:fill]) if combined_hash_cell[:fill]
    rubyxl_worksheet[row_index][column_index].change_horizontal_alignment(combined_hash_cell[:align]) if combined_hash_cell[:align]
    rubyxl_worksheet[row_index][column_index].change_font_bold(combined_hash_cell[:bold]) if combined_hash_cell[:bold]

    if combined_hash_cell[:border_all]
      rubyxl_worksheet[row_index][column_index].change_border('top' , combined_hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('bottom' , combined_hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('left' , combined_hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('right' , combined_hash_cell[:border_all])
    end
  end

  def read_file(path)
    rubyxl_workbook = RubyXL::Parser.parse(path)
    @hash_workbook = rubyxl_workbook_to_hash_workbook(rubyxl_workbook)
  end

  def rubyxl_workbook_to_hash_workbook(rubyxl_workbook)
    hash_workbook = {}
    rubyxl_workbook.each do |rubyxl_worksheet|
      hash_worksheet = {row_count: rubyxl_worksheet.count, column_count: 1, rows: {}, columns: {}, cells: {}}
      rubyxl_worksheet_to_hash_worksheet(rubyxl_worksheet, hash_worksheet)
      process_sheet_to_populated_block(hash_worksheet)
      hash_workbook[rubyxl_worksheet.sheet_name] = hash_worksheet
    end
    hash_workbook
  end

  def rubyxl_worksheet_to_hash_worksheet(rubyxl_worksheet, hash_worksheet)
    rubyxl_worksheet.each_with_index do |rubyxl_row, rubyxl_row_index|
      rubyxl_row_cells = rubyxl_row&.cells
      if rubyxl_row_cells.nil?
        hash_cell_key = RubyXL::Reference.ind2ref(rubyxl_row_index, 0)
        hash_worksheet[:rows][hash_cell_key[/\D+/]]
      else
        rubyxl_row_cells.each_with_index do |rubyxl_cell, rubyxl_column_index|
          hash_cell_key = RubyXL::Reference.ind2ref(rubyxl_row_index, rubyxl_column_index)
          hash_worksheet[:cells][hash_cell_key] = rubyxl_cell_to_hash_cell(rubyxl_cell)
          hash_worksheet[:column_count] = rubyxl_column_index + 1 if rubyxl_column_index >= hash_worksheet[:column_count]
        end
      end
    end
  end

  def rubyxl_cell_to_hash_cell(rubyxl_cell)
    hash_cell = {}
    hash_cell[:value] = rubyxl_cell.value
    hash_cell[:format] = {
        string: 'text',
        fixnum: 'number',
        float: 'number',
        datetime: 'date'
    }[rubyxl_cell.value.class.to_s.downcase.to_sym]
    if rubyxl_cell.value.is_a?(Float)
      hash_cell[:number] = "0.#{'0' * rubyxl_cell.value.to_s[rubyxl_cell.value.to_s.index('.') + 1..-1].length}"
    end
    hash_cell
  end

  def process_sheet_to_populated_block(hash_worksheet)
    set_hash_worksheet_extents(hash_worksheet)
    hash_worksheet[:row_count].times do |row_index|
      hash_worksheet[:column_count].times do |column_index|
        cell_key = RubyXL::Reference.ind2ref(row_index, column_index)
        hash_worksheet[:cells][cell_key] = {} unless hash_worksheet[:cells][cell_key]
      end
    end
  end
end
