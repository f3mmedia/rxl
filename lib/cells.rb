module Cells

  def self.rubyxl_to_hash(rubyxl_rows, hash_worksheet)
    rubyxl_rows.each do |rubyxl_row_hash|
      rubyxl_row = rubyxl_row_hash[:rubyxl_row]
      rubyxl_row_index = rubyxl_row_hash[:rubyxl_row_index]
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

end