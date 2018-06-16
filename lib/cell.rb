require 'rubyXL'

module Cell

  ##############################################
  ###     GET HASH CELL FROM RUBYXL CELL     ###
  ##############################################

  def self.rubyxl_cell_to_hash_cell(rubyxl_cell = nil)
    rubyxl_cell_value = rubyxl_cell.nil? ? RubyXL::Cell.new.value : rubyxl_cell.value
    {
        value: rubyxl_cell_value,
        format: hash_cell_format(rubyxl_cell_value),
        formula: rubyxl_cell_formula(rubyxl_cell),
        h_align: rubyxl_cell_horizontal_alignment(rubyxl_cell),
        v_align: rubyxl_cell_vertical_alignment(rubyxl_cell)
    }
  end

  def self.rubyxl_cell_formula(rubyxl_cell)
    return nil if rubyxl_cell.nil? || rubyxl_cell.formula.nil? || rubyxl_cell.formula.expression.empty?
    rubyxl_cell.formula.expression
  end

  def self.rubyxl_cell_horizontal_alignment(rubyxl_cell)
    return nil if rubyxl_cell.nil? || rubyxl_cell.horizontal_alignment.nil?
    rubyxl_cell.horizontal_alignment.to_sym
  end

  def self.rubyxl_cell_vertical_alignment(rubyxl_cell)
    return :bottom if rubyxl_cell.nil? || rubyxl_cell.vertical_alignment.nil?
    rubyxl_cell.vertical_alignment.to_sym
  end

  def self.hash_cell_format(rubyxl_cell_value)
    format = {
        nilclass: :general,
        string: :text,
        fixnum: :number,
        float: :number,
        datetime: :date,
    }[rubyxl_cell_value.class.to_s.downcase.to_sym]
    format == :date && rubyxl_cell_value.strftime('%Y%m%d') == '18991231' ? :time : format
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

  def self.validate_hash_cell(hash_cell_key, hash_cell, trace)
    unless validate_cell_key(hash_cell_key)
      raise(%[invalid cell key at path #{trace}, must be String and in Excel format (eg "A1")])
    end
    unless hash_cell.is_a?(Hash)
      raise("cell value at path #{trace + [hash_cell_key]} must be a Hash")
    end
    unless hash_cell.keys.reject { |key| key.is_a?(Symbol) }.empty?
      raise("cell key at path #{trace + [hash_cell_key]} must be a Symbol")
    end
    unless hash_cell.keys.delete_if { |key| valid_cell_keys.include?(key) }.empty?
      valid_cell_keys_string = ":#{valid_cell_keys.join(', :')}"
      raise(%(invalid cell hash key at path #{trace + [hash_cell_key]}, valid keys are: [#{valid_cell_keys_string}]))
    end
    # TODO: add validation for hash_cell specification
  end

  def self.validate_cell_key(cell_key)
    return false unless cell_key.is_a?(String)
    return false unless cell_key[/^[A-Z]{1,3}[0-9]{1,7}$/]
    cell_index = RubyXL::Reference.ref2ind(cell_key)
    return false unless cell_index[0].between?(0, 1_048_575)
    return false unless cell_index[0].between?(0, 16383)
    true
  end

  def self.valid_cell_keys
    %i[
      value
      number
      formula
    ]
  end

end