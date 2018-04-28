require 'rubyXL'

module Cell

  def self.rubyxl_cell_to_hash_cell(rubyxl_cell=nil)
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

end