require 'rubyXL'

module Cell

  ##############################################
  ###     GET HASH CELL FROM RUBYXL CELL     ###
  ##############################################

  def self.rubyxl_cell_to_hash_cell(rubyxl_cell = nil)
    rubyxl_cell_value = rubyxl_cell.nil? ? RubyXL::Cell.new.value : rubyxl_cell.value
    rubyxl_cell_value = resolve_date_ms(rubyxl_cell_value) if rubyxl_cell_value.is_a?(Date)
    format = hash_cell_format(rubyxl_cell_value)
    {
      value: rubyxl_cell_value,
      format: format,
      formula: rubyxl_cell_formula(rubyxl_cell),
      decimals: format == :number ? decimals(rubyxl_cell) : nil,
      h_align: rubyxl_cell_h_align(rubyxl_cell),
      v_align: rubyxl_cell_v_align(rubyxl_cell),
      bold: rubyxl_cell.nil? ? false : rubyxl_cell.is_bolded,
      fill: rubyxl_cell.nil? ? 'ffffff' : rubyxl_cell.fill_color,
      font_name: rubyxl_cell.nil? ? 'Calibri' : rubyxl_cell.font_name,
      font_size: rubyxl_cell.nil? ? 12 : rubyxl_cell.font_size.to_i,
      border: rubyxl_cell_to_border_hash(rubyxl_cell)
    }
  end

  def self.resolve_date_ms(value)
    value_ut = value.strftime('%s')
    ut = value.strftime('%L').to_i > 499 ? "#{value_ut.to_i + 1}" : value_ut
    DateTime.strptime(ut, '%s')
  end

  def self.decimals(rubyxl_cell)
    number_format = rubyxl_cell.number_format
    return nil unless number_format
    format_code = number_format.format_code
    i = format_code.reverse.index('.')
    format_code[0 - i..-1].length if i
  end

  def self.rubyxl_cell_to_border_hash(rubyxl_cell)
    {
      top: rubyxl_cell.nil? ? nil : rubyxl_cell.get_border(:top),
      bottom: rubyxl_cell.nil? ? nil : rubyxl_cell.get_border(:bottom),
      left: rubyxl_cell.nil? ? nil : rubyxl_cell.get_border(:left),
      right: rubyxl_cell.nil? ? nil : rubyxl_cell.get_border(:right)
    }
  end

  def self.rubyxl_cell_formula(rubyxl_cell)
    return nil if rubyxl_cell.nil? || rubyxl_cell.formula.nil? || rubyxl_cell.formula.expression.empty?
    rubyxl_cell.formula.expression
  end

  def self.rubyxl_cell_h_align(rubyxl_cell)
    return :left if rubyxl_cell.nil? || rubyxl_cell.horizontal_alignment.nil?
    rubyxl_cell.horizontal_alignment.to_sym
  end

  def self.rubyxl_cell_v_align(rubyxl_cell)
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
    rubyxl_worksheet[row_index][column_index].change_horizontal_alignment(combined_hash_cell[:h_align]) if combined_hash_cell[:h_align]
    rubyxl_worksheet[row_index][column_index].change_font_bold(combined_hash_cell[:bold]) if combined_hash_cell[:bold]

    if combined_hash_cell[:border_all]
      rubyxl_worksheet[row_index][column_index].change_border('top' , combined_hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('bottom' , combined_hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('left' , combined_hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('right' , combined_hash_cell[:border_all])
    end
  end

  def self.add_rubyxl_cells(hash_cell, rubyxl_worksheet, row_index, column_index)
    number_format = write_format(hash_cell)
    if hash_cell[:formula]
      if number_format
        rubyxl_worksheet.add_cell(row_index, column_index, '', hash_cell[:formula])
          .set_number_format(number_format)
      else
        rubyxl_worksheet.add_cell(row_index, column_index, '', hash_cell[:formula])
      end
    else
      if number_format
        cell_value = hash_cell[:value].is_a?(Date) ? date_to_num(hash_cell[:value]) : hash_cell[:value]
        rubyxl_worksheet.add_cell(row_index, column_index, cell_value)
          .set_number_format(number_format)
      else
        rubyxl_worksheet.add_cell(row_index, column_index, hash_cell[:value])
      end
    end
  end

  def self.write_format(hash_cell)
    case hash_cell[:format]
    when :number
      hash_cell[:decimals] ? "0.#{ '0' * hash_cell[:decimals] }" : '0'
    when :date
      hash_cell[:date_format] ? hash_cell[:date_format] : 'dd/mm/yyyy'
    when :time
      hash_cell[:date_format] ? hash_cell[:date_format] : 'hh:mm:ss'
    when :percentage
      hash_cell[:decimals] ? "0.#{ '0' * hash_cell[:decimals] }%" : '0%'
    else
      case hash_cell[:value].class.to_s
      when 'Fixnum'
        '0'
      when 'Float'
        value = hash_cell[:value].to_s
        decimals = value[value.index('.') + 1..-1].length
        "0.#{ '0' * decimals }"
      when 'DateTime'
        return 'dd/mm/yyyy hh:mm:ss'
      else
        nil
      end
    end
  end

  def self.date_to_num(date)
    workbook = RubyXL::Workbook.new
    workbook.date_to_num(date)
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
    invalid_keys = hash_cell.keys.delete_if { |key| valid_cell_keys.include?(key) }
    unless invalid_keys.empty?
      raise(%(invalid cell hash key(s) #{invalid_keys} at path #{trace + [hash_cell_key]}))
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
      format
      date_format
      formula
      decimals
      bold
      h_align
      v_align
      border
      fill
      font_name
      font_size
    ]
  end

end