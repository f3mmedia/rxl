require 'rubyXL'

module Cell

  ##############################################
  ###     GET HASH CELL FROM RUBYXL CELL     ###
  ##############################################

  def self.rubyxl_cell_to_hash_cell(rubyxl_cell=nil)
    rubyxl_cell_value = rubyxl_cell.nil? ? RubyXL::Cell.new.value : rubyxl_cell.value
    {
        value: rubyxl_cell_value,
        format: hash_cell_format(rubyxl_cell_value)
    }
  end

  def self.hash_cell_format(rubyxl_cell_value)
    format = {
        nilclass: 'general',
        string: 'text',
        fixnum: 'number',
        float: 'number',
        datetime: 'date',
    }[rubyxl_cell_value.class.to_s.downcase.to_sym]
    format[:number] = hash_cell_float_format(rubyxl_cell_value) if rubyxl_cell_value.is_a?(Float)
    format
  end

  def self.hash_cell_float_format(rubyxl_cell_value)
    decimal_point_index = rubyxl_cell_value.to_s.index('.') + 1
    "0.#{'0' * rubyxl_cell_value.to_s[decimal_point_index..-1].length}"
  end


  ##############################################
  ###     GET RUBYXL CELL FROM HASH CELL     ###
  ##############################################

  def self.hash_cell_to_rubyxl_cell(combined_hash_cell, rubyxl_worksheet, row_index, column_index)
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

  def self.add_rubyxl_cells(combined_hash_cell, rubyxl_worksheet, row_index, column_index)
    if combined_hash_cell[:formula]
      rubyxl_worksheet.add_cell(row_index, column_index, '', combined_hash_cell[:formula]).set_number_format combined_hash_cell[:dp_2]
    else
      rubyxl_worksheet.add_cell(row_index, column_index, combined_hash_cell[:value])
    end
  end

  def self.get_combined_hash_cell(hash_worksheet, hash_cell_key, hash_cell)
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


  ##################################
  ###     VALIDATE HASH CELL     ###
  ##################################

  def self.validate_hash_cell(type, cell_id, hash_cell, trace)
    validate_cell_id(type, cell_id, trace)
    unless hash_cell.is_a?(Hash)
      raise("hash_cell at path [#{trace + [cell_id]}] is class #{hash_cell.class}, must be a Hash")
    end
    # TODO: add validation for hash_cell specification
  end

  def self.validate_cell_id(type, cell_id, trace)
    case type
      when :cells
        unless cell_id[/^\D+\d+$/]
          raise "hash_cell key at path [#{trace}] of type cell has invalid key: #{cell_id}, must be capitalised alpha(s) and numeric (eg AB123)"
        end
      when :columns
        unless cell_id[/^\D+$/]
          raise "hash_cell key at path [#{trace}] of type column has invalid key: #{cell_id}, must be capitalised alpha only (eg AB)"
        end
      when :rows
        unless cell_id[/^\d+$/]
          raise "hash_cell key at path [#{trace}] of type row has invalid key: #{cell_id}, must be stringified integer only (eg 123)"
        end
    end
  end

end