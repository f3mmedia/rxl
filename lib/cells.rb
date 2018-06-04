module Cells

  def self.rubyxl_to_hash(rubyxl_rows)
    hash_cells = {}
    rubyxl_rows.each do |rubyxl_row_hash|
      rubyxl_row = rubyxl_row_hash[:rubyxl_row]
      rubyxl_row_index = rubyxl_row_hash[:rubyxl_row_index]
      rubyxl_row_cells = rubyxl_row&.cells
      unless rubyxl_row_cells.nil?
        rubyxl_row_cells.each_with_index do |rubyxl_cell, rubyxl_column_index|
          hash_cell_key = RubyXL::Reference.ind2ref(rubyxl_row_index, rubyxl_column_index)
          hash_cells[hash_cell_key] = Cell.rubyxl_cell_to_hash_cell(rubyxl_cell)
        end
      end
    end
    hash_cells
  end

end