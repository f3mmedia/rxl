require 'rubyXL'
require 'mitrush'
require_relative 'cell'
require_relative 'cells'

module Worksheet

  ########################################################
  ###     GET HASH WORKSHEET FROM RUBYXL WORKSHEET     ###
  ########################################################

  def self.rubyxl_to_hash(rubyxl_worksheet)
    rubyxl_rows = rubyxl_worksheet.map do |rubyxl_row|
      { rubyxl_row: rubyxl_row, rubyxl_row_index: rubyxl_row ? rubyxl_row.r - 1 : nil }
    end
    hash_worksheet = Cells.rubyxl_to_hash(rubyxl_rows)
    process_sheet_to_populated_block(hash_worksheet)
    Mitrush.delete_keys(hash_worksheet, %i[row_count column_count])
    hash_worksheet
  end


  ########################################################
  ###     GET RUBYXL WORKSHEET FROM HASH WORKSHEET     ###
  ########################################################

  def self.hash_worksheet_to_rubyxl_worksheet(hash_worksheet, rubyxl_worksheet)
    process_sheet_to_populated_block(hash_worksheet)
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
    extent = hash_worksheet_extents(hash_worksheet)
    extent[:row_count].times do |row_index|
      extent[:column_count].times do |column_index|
        cell_key = RubyXL::Reference.ind2ref(row_index, column_index)
        hash_worksheet[cell_key] = Cell.rubyxl_cell_to_hash_cell unless hash_worksheet[cell_key]
      end
    end
  end

  def self.hash_worksheet_extents(hash_worksheet)
    extents = { row_count: 0, column_count: 0 }
    (hash_worksheet || {}).keys.each do |hash_cell_key|
      row_index, column_index = RubyXL::Reference.ref2ind(hash_cell_key)
      extents[:row_count] = row_index + 1 if row_index >= extents[:row_count]
      extents[:column_count] = column_index + 1 if column_index >= extents[:column_count]
    end
    extents
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
    cells = Mitrush.deep_copy(raw_hash)
    columns = cells.keys.map { |key| key[/\D+/] }.uniq
    columns.delete_if { |item| cells["#{item}1"][:value].nil? }
    row_nums = cells.keys.map { |key| key[/\d+/].to_i }.uniq[1..-1] || []
    row_nums.sort! unless row_nums.empty?
    row_nums.map do |row_number|
      columns.each_with_object({}) do |column_letter, h|
        h[cells["#{column_letter}1"][:value]] = cells["#{column_letter}#{row_number}"][:value]
      end
    end
  end

  def self.validate_hash_worksheet(hash_worksheet_name, hash_worksheet)
    unless hash_worksheet_name.is_a?(String)
      raise('worksheet name must be a String')
    end
    raise('worksheet name must not be an empty String') if hash_worksheet_name.empty?
    unless hash_worksheet.is_a?(Hash)
      raise(%(worksheet value at path ["#{hash_worksheet_name}"] must be a Hash))
    end
    hash_worksheet.each do |hash_cell_key, hash_cell|
      Cell.validate_hash_cell(hash_cell_key, hash_cell, [hash_worksheet_name])
    end
  end

end