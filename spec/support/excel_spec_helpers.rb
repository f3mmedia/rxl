module ExcelSpecHelpers

  def create_temp_xlsx_dir_unless_exists
    path = Pathname.new(ENV['TEMP_XLSX_PATH'])
    FileUtils.mkdir(path.to_s) unless path.exist?
  end

  def destroy_temp_xlsx_dir_if_exists
    path = Pathname.new(ENV['TEMP_XLSX_PATH'])
    FileUtils.rmtree(path.to_s) if path.exist?
  end

  def generate_test_excel_file_to_spec(spec)
    filepath = "#{ENV['TEMP_XLSX_PATH']}/#{spec_to_filename(spec)}"
    file = Excel.new(source: spec_to_write_hash(spec))
    file.save_file(filepath)
    path = Pathname.new(filepath)
    expect(path.exist?)
  end

  def verify_read_hash_matches_expected(spec)
    filepath = "#{ENV['TEMP_XLSX_PATH']}/#{spec_to_filename(spec)}"
    file = Excel.new(source: filepath)
    expect(file.hash_workbook).to eq(spec_to_test_hash(spec))
  end

  def spec_to_filename(spec)
    {
        empty_xlsx: 'empty_file_test.xlsx',
        sheet_names_xlsx: 'sheet_names_test.xlsx'
    }[spec]
  end

  def spec_to_write_hash(spec)
    {
        empty_xlsx: {},
        sheet_names_xlsx: {'test_a' => {}, 'test_b' => {}}
    }[spec]
  end

  def spec_to_test_hash(spec)
    {
        empty_xlsx: {'Sheet1'=>{row_count: 0, column_count: 0, rows: {}, columns: {}, cells: {}}},
        sheet_names_xlsx: {
            'test_a'=>{row_count: 0, column_count: 0, rows: {}, columns: {}, cells: {}},
            'test_b'=>{row_count: 0, column_count: 0, rows: {}, columns: {}, cells: {}}
        }
    }[spec]
  end

end
