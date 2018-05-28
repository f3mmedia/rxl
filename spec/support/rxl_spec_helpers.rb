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
    filepath = test_data(:filepath, key)
    Rxl.write_file(filepath, test_data(:write_hash, key))
    path = Pathname.new(filepath)
    test.expect(path.exist?)
  end

  def self.test_data(type, key)
    temp_xlsx_path = ENV['TEMP_XLSX_PATH']
    return_value = {
      filepath: derived_filepath(key, temp_xlsx_path),
      write_hash: {
        empty_file: {},
        worksheet_names: {'test_a' => {}, 'test_b' => {}}
      }[key],
      expected_hash: {
        empty_file: {'Sheet1' => {rows: {}, columns: {}, cells: {}}},
        worksheet_names: {
          'test_a' => {rows: {}, columns: {}, cells: {}},
          'test_b' => {rows: {}, columns: {}, cells: {}}
        },
        cell_raw_string_read: {
          'B2' => {value: 'abcde', format: :text, formula: nil},
          'B3' => {value: 'abcde', format: :text, formula: nil},
          'B4' => {value: 'abcde', format: :text, formula: nil},
          'B5' => {value: 'abcde', format: :text, formula: nil},
          'B6' => {value: 'abcde', format: :text, formula: nil},
          'B7' => {value: 'abcde', format: :text, formula: nil}
        },
        cell_raw_number_read: {
          'B2' => {value: 12345, format: :number, formula: nil},
          'B3' => {value: '12345', format: :text, formula: nil},
          'B4' => {value: 12345, format: :number, formula: nil}
        },
        cell_raw_float_read: {
          'B3' => {value: '123.45', format: :text, formula: nil},
          'B4' => {value: 123.45, format: :number, formula: nil}
        },
        cell_raw_date_read: {
          'B3' => {value: '01/01/2000', format: :text, formula: nil},
          'B5' => {value: DateTime.parse('01/01/2000'), format: :date, formula: nil},
          'B6' => {value: DateTime.parse('01/01/2000'), format: :date, formula: nil},
          'B7' => {value: '01/01/2000%', format: :text, formula: nil}
        },
        cell_raw_time_read: {
          'B3' => {value: '10:15:30', format: :text, formula: nil},
          'B6' => {value: DateTime.parse('31/12/1899 10:15:30'), format: :time, formula: nil},
          'B7' => {value: '10:15:30%', format: :text, formula: nil}
        },
        cell_raw_percentage_read: {
          'B3' => {value: '100%', format: :text, formula: nil},
          'B7' => {value: 1, format: :number, formula: nil}
        },
        cell_raw_percentage_float_read: {
          'B3' => {value: '123.45%', format: :text, formula: nil},
          'B7' => {value: 1.2345, format: :number, formula: nil}
        },
        cell_raw_empty_read: {
          'B2' => {value: nil, format: :general, formula: nil},
          'B3' => {value: nil, format: :general, formula: nil},
          'B4' => {value: nil, format: :general, formula: nil},
          'B5' => {value: nil, format: :general, formula: nil},
          'B6' => {value: nil, format: :general, formula: nil},
          'B7' => {value: nil, format: :general, formula: nil}
        },
        cell_formula_string_read: {
          'C2' => {value: 'abcde', format: :text, formula: 'CONCATENATE("abc","de")'},
          'C3' => {value: 'abcde', format: :text, formula: 'CONCATENATE("abc","de")'},
          'C4' => {value: 'abcde', format: :text, formula: nil},
          'C5' => {value: 'abcde', format: :text, formula: nil},
          'C6' => {value: 'abcde', format: :text, formula: 'CONCATENATE("abc","de")'},
          'C7' => {value: 'abcde', format: :text, formula: nil}
        },
        cell_formula_number_read: {
          'C2' => {value: 12345, format: :number, formula: '12340+5'},
          'C3' => {value: '=12340+5', format: :text, formula: nil},
          'C4' => {value: 12345, format: :number, formula: '12340+5'}
        },
        cell_formula_float_read: {
          'C3' => {value: '=123.41+0.04', format: :text, formula: nil},
          'C4' => {value: 123.45, format: :number, formula: '123.41+0.04'}
        },
        cell_formula_date_read: {
          'C3' => {value: '=DATE(2000,1,1)', format: :text, formula: nil},
          'C5' => {value: DateTime.parse('01/01/2000'), format: :date, formula: 'DATE(2000,1,1)'},
          'C6' => {value: DateTime.parse('01/01/2000'), format: :date, formula: 'DATE(2000,1,1)'}
        },
        cell_formula_time_read: {
          'C3' => {value: '=TIME(10,15,30)', format: :text, formula: nil},
          'C6' => {value: DateTime.parse('31/12/1899 10:15:30'), format: :time, formula: 'TIME(10,15,30)'}
        },
        cell_formula_percentage_read: {
          'C3' => {value: '=50%+50%', format: :text, formula: nil},
          'C7' => {value: 1, format: :number, formula: '50%+50%'}
        },
        cell_formula_percentage_float_read: {
          'C3' => {value: '=123.41%+0.04%', format: :text, formula: nil},
          'C7' => {value: 1.2345, format: :number, formula: '123.41%+0.04%'}
        },
      }[key],
      validation: {
        non_hash_workbook: 'workbook must be a Hash',
        non_string_worksheet_name: 'worksheet name must be a String',
        non_hash_worksheet: 'worksheet value at path ["worksheet_a"] must be a Hash',
        invalid_worksheet_keys: 'worksheet at path ["worksheet_a"] contains unauthorised key(s)'
      }[key]
    }[type]
    raise "no value found for type :#{type} and key :#{key}" unless return_value
    return_value
  end

  def self.derived_filepath(key, xlsx_path)
    filepath = {
      cell_values_and_formats: 'spec/support/static_test_files/cell_values_and_formats.xlsx'
    }[key]
    filepath || "#{xlsx_path}/#{key}.xlsx"
  end

end
