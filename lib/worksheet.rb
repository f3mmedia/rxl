require 'rubyXL'
require_relative 'cell'

module Worksheet

  def self.rubyxl_worksheet_to_hash_worksheet(rubyxl_worksheet, hash_worksheet)
    rubyxl_worksheet.each_with_index do |rubyxl_row, rubyxl_row_index|
      rubyxl_row_cells = rubyxl_row&.cells
      if rubyxl_row_cells.nil?
        hash_cell_key = RubyXL::Reference.ind2ref(rubyxl_row_index, 0)
        hash_worksheet[:rows][hash_cell_key[/\D+/]]
      else
        rubyxl_row_cells.each_with_index do |rubyxl_cell, rubyxl_column_index|
          hash_cell_key = RubyXL::Reference.ind2ref(rubyxl_row_index, rubyxl_column_index)
          hash_worksheet[:cells][hash_cell_key] = Cell.rubyxl_cell_to_hash_cell(rubyxl_cell)
          hash_worksheet[:column_count] = rubyxl_column_index + 1 if rubyxl_column_index >= hash_worksheet[:column_count]
        end
      end
    end
  end

  def self.hash_worksheet_to_rubyxl_worksheet(hash_worksheet, rubyxl_worksheet)
    Worksheet.process_sheet_to_populated_block(hash_worksheet)
    (hash_worksheet[:cells] || {}).sort.each do |hash_cell_key, hash_cell|
      combined_hash_cell = Cell.get_combined_hash_cell(hash_worksheet, hash_cell_key, hash_cell)
      row_index, column_index = RubyXL::Reference.ref2ind(hash_cell_key)
      Cell.add_rubyxl_cells(combined_hash_cell, rubyxl_worksheet, row_index, column_index)
      Cell.hash_cell_to_rubyxl_cell(combined_hash_cell, rubyxl_worksheet, row_index, column_index)
    end
  end

  def self.process_sheet_to_populated_block(hash_worksheet)
    Worksheet.set_hash_worksheet_extents(hash_worksheet)
    hash_worksheet[:row_count].times do |row_index|
      hash_worksheet[:column_count].times do |column_index|
        cell_key = RubyXL::Reference.ind2ref(row_index, column_index)
        hash_worksheet[:cells][cell_key] = {} unless hash_worksheet[:cells][cell_key]
      end
    end
  end

  def self.set_hash_worksheet_extents(hash_worksheet)
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

  def self.set_hash_worksheet_defaults(hash_worksheet)
    %i[worksheet columns rows].each do |key|
      hash_worksheet[key] = {} unless hash_worksheet.has_key?(key)
    end
  end

  def self.hash_worksheet_to_hash_table(raw_hash)
    cells = raw_hash[:cells]
    columns = cells.keys.map { |key| key[/\D+/] }.uniq
    cells.keys.map { |key| key[/\d+/] }.uniq[1..-1].map do |row_number|
      columns.each_with_object({}) do |column_letter, this_hash|
        this_hash[cells["#{column_letter}1"][:value]] = cells["#{column_letter}#{row_number}"][:value]
      end
    end
  end

  def self.validate_hash_worksheet(hash_worksheet_name, hash_worksheet)
    @validation << "hash_worksheet_name is class '#{hash_worksheet_name.class}', should be a String" unless hash_worksheet_name.is_a?(String)
    # other validation...
  end

end