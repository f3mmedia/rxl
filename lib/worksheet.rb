require 'rubyXL'
require 'mitrush'
require_relative 'cell'
require_relative 'cells'

module Worksheet

  ########################################################
  ###     GET HASH WORKSHEET FROM RUBYXL WORKSHEET     ###
  ########################################################

  def self.rubyxl_to_hash(rubyxl_worksheet)
    hash_worksheet = hash_worksheet_template(rubyxl_worksheet.count)
    rubyxl_rows = rubyxl_worksheet.each_with_index.map do |rubyxl_row, rubyxl_row_index|
      {rubyxl_row: rubyxl_row, rubyxl_row_index: rubyxl_row_index}
    end
    Cells.rubyxl_to_hash(rubyxl_rows, hash_worksheet)
    process_sheet_to_populated_block(hash_worksheet)
    Mitrush.delete_keys(hash_worksheet, %i[row_count column_count])
    hash_worksheet
  end

  def self.hash_worksheet_template(rubyxl_worksheet_row_count)
    {
        row_count: rubyxl_worksheet_row_count,
        column_count: 1,
        rows: {},
        columns: {},
        cells: {}
    }
  end


  ########################################################
  ###     GET RUBYXL WORKSHEET FROM HASH WORKSHEET     ###
  ########################################################

  def self.hash_worksheet_to_rubyxl_worksheet(hash_worksheet, rubyxl_worksheet)
    Worksheet.process_sheet_to_populated_block(hash_worksheet)
    (hash_worksheet[:cells] || {}).sort.each do |hash_cell_key, hash_cell|
      combined_hash_cell = Cell.get_combined_hash_cell(hash_worksheet, hash_cell_key, hash_cell)
      row_index, column_index = RubyXL::Reference.ref2ind(hash_cell_key)
      Cell.add_rubyxl_cells(combined_hash_cell, rubyxl_worksheet, row_index, column_index)
      Cell.hash_cell_to_rubyxl_cell(combined_hash_cell, rubyxl_worksheet, row_index, column_index)
    end
  end


  ##############################
  ###     SHARED METHODS     ###
  ##############################

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


  ####################################
  ###     OTHER PUBLIC METHODS     ###
  ####################################

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

  def self.validate_hash_worksheet(hash_worksheet_name, hash_worksheet, trace)
    unless hash_worksheet_name.is_a?(String)
      raise 'hash_worksheet key must be a String'
    end
    unless hash_worksheet.is_a?(Hash)
      raise "hash_worksheet value at path #{trace} must be a Hash"
    end
    unauthorised_keys = Mitrush.delete_keys(hash_worksheet.dup, %i[cells columns rows])
    unless unauthorised_keys.empty?
      raise "hash_worksheet at path #{trace} contains unauthorised key(s): #{unauthorised_keys.join(', ')}"
    end
    hash_worksheet.each do |type, cells_hash|
      unless cells_hash.is_a?(Hash)
        raise "value at path #{trace + [type]} must be a Hash"
      end
      cells_hash.each do |cell_id, hash_cell|
        Cell.validate_hash_cell(type, cell_id, hash_cell, trace + [type])
      end
    end
  end

end