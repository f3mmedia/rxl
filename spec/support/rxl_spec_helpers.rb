module RxlSpecHelpers

  def self.create_temp_xlsx_dir_unless_exists
    path = Pathname.new(ENV['TEMP_XLSX_PATH'])
    FileUtils.mkdir(path.to_s) unless path.exist?
  end

  def self.destroy_temp_xlsx_dir_if_exists
    path = Pathname.new(ENV['TEMP_XLSX_PATH'])
    FileUtils.rmtree(path.to_s) if path.exist?
  end

  def self.generate_test_excel_file(test, key)
    filepath = "#{ENV['TEMP_XLSX_PATH']}/#{test_filenames(key)}"
    Rxl.write_file(filepath, write_hash(key))
    path = Pathname.new(filepath)
    test.expect(path.exist?)
  end

  def self.verify_read_hash_matches_expected(test, key)
    filepath = "#{ENV['TEMP_XLSX_PATH']}/#{test_filenames(key)}"
    read_hash = Rxl.read_file(filepath)
    test.expect(read_hash).to test.eq(expected_hash(key))
  end

  def self.test_filenames(key)
    {
        empty_xlsx: 'empty_file_test.xlsx',
        sheet_names_xlsx: 'sheet_names_test.xlsx'
    }[key]
  end

  def self.write_hash(key)
    {
        empty_xlsx: {},
        sheet_names_xlsx: {'test_a' => {}, 'test_b' => {}}
    }[key]
  end

  def self.expected_hash(key)
    {
        empty_xlsx: {'Sheet1'=>{row_count: 0, column_count: 0, rows: {}, columns: {}, cells: {}}},
        sheet_names_xlsx: {
            'test_a'=>{row_count: 0, column_count: 0, rows: {}, columns: {}, cells: {}},
            'test_b'=>{row_count: 0, column_count: 0, rows: {}, columns: {}, cells: {}}
        }
    }[key]
  end

end
