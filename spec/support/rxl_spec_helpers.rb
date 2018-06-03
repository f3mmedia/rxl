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

  def self.test_data(type, key, args = {})
    temp_xlsx_path = ENV['TEMP_XLSX_PATH']
    return_value = {
      filepath: derived_filepath(key, temp_xlsx_path),
      write_hash: {
        empty_file: {},
        worksheet_names: { 'test_a' => {}, 'test_b' => {} }
      }[key],
      expected_hash: {
        empty_file: { 'Sheet1' => {rows: {}, columns: {}, cells: {} } },
        worksheet_names: {
          'test_a' => { rows: {}, columns: {}, cells: {} },
          'test_b' => { rows: {}, columns: {}, cells: {} }
        },
        cell_raw_string_read: {
          'B2' => { value: 'abcde', format: :text },
          'B3' => { value: 'abcde', format: :text },
          'B4' => { value: 'abcde', format: :text },
          'B5' => { value: 'abcde', format: :text },
          'B6' => { value: 'abcde', format: :text },
          'B7' => { value: 'abcde', format: :text }
        },
        cell_raw_number_read: {
          'B2' => { value: 12345, format: :number },
          'B3' => { value: '12345', format: :text },
          'B4' => { value: 12345, format: :number }
        },
        cell_raw_float_read: {
          'B3' => { value: '123.45', format: :text },
          'B4' => { value: 123.45, format: :number }
        },
        cell_raw_date_read: {
          'B3' => { value: '01/01/2000', format: :text },
          'B5' => { value: DateTime.parse('01/01/2000'), format: :date },
          'B6' => { value: DateTime.parse('01/01/2000'), format: :date },
          'B7' => { value: '01/01/2000%', format: :text }
        },
        cell_raw_time_read: {
          'B3' => { value: '10:15:30', format: :text },
          'B6' => { value: DateTime.parse('31/12/1899 10:15:30'), format: :time },
          'B7' => { value: '10:15:30%', format: :text }
        },
        cell_raw_percentage_read: {
          'B3' => { value: '100%', format: :text },
          'B7' => { value: 1, format: :number }
        },
        cell_raw_percentage_float_read: {
          'B3' => { value: '123.45%', format: :text },
          'B7' => { value: 1.2345, format: :number }
        },
        cell_raw_empty_read: {
          'B2' => { format: :general },
          'B3' => { format: :general },
          'B4' => { format: :general },
          'B5' => { format: :general },
          'B6' => { format: :general },
          'B7' => { format: :general }
        },
        cell_formula_string_read: {
          'C2' => { value: 'abcde', format: :text, formula: 'CONCATENATE("abc","de")' },
          'C3' => { value: 'abcde', format: :text, formula: 'CONCATENATE("abc","de")' },
          'C4' => { value: 'abcde', format: :text },
          'C5' => { value: 'abcde', format: :text },
          'C6' => { value: 'abcde', format: :text, formula: 'CONCATENATE("abc","de")' },
          'C7' => { value: 'abcde', format: :text }
        },
        cell_formula_number_read: {
          'C2' => { value: 12345, format: :number, formula: '12340+5' },
          'C3' => { value: '=12340+5', format: :text },
          'C4' => { value: 12345, format: :number, formula: '12340+5' }
        },
        cell_formula_float_read: {
          'C3' => { value: '=123.41+0.04', format: :text },
          'C4' => { value: 123.45, format: :number, formula: '123.41+0.04' }
        },
        cell_formula_date_read: {
          'C3' => { value: '=DATE(2000,1,1)', format: :text },
          'C5' => { value: DateTime.parse('01/01/2000'), format: :date, formula: 'DATE(2000,1,1)' },
          'C6' => { value: DateTime.parse('01/01/2000'), format: :date, formula: 'DATE(2000,1,1)' }
        },
        cell_formula_time_read: {
          'C3' => { value: '=TIME(10,15,30)', format: :text },
          'C6' => { value: DateTime.parse('31/12/1899 10:15:30'), format: :time, formula: 'TIME(10,15,30)' }
        },
        cell_formula_percentage_read: {
          'C3' => { value: '=50%+50%', format: :text },
          'C7' => { value: 1, format: :number, formula: '50%+50%' }
        },
        cell_formula_percentage_float_read: {
          'C3' => { value: '=123.41%+0.04%', format: :text },
          'C7' => { value: 1.2345, format: :number, formula: '123.41%+0.04%' }
        }
      }[key],
      validation: {
        non_hash_workbook: 'workbook must be a Hash',
        non_string_worksheet_name: 'worksheet name must be a String',
        empty_string_worksheet_name: 'worksheet name must not be an empty String',
        non_hash_worksheet: "worksheet value at path #{args[:path]} must be a Hash",
        invalid_cell_key: %[invalid cell key at path #{args[:path]}, must be String and in Excel format (eg "A1")],
        non_hash_cell_value: "cell value at path #{args[:path]} must be a Hash",
        non_symbol_cell_hash_key: "cell key at path #{args[:path]} must be a Symbol",
        invalid_cell_hash_key: %(invalid cell hash key at path #{args[:path]}, valid keys are: [#{args[:valid_cell_keys_string]}])
      }[key]
    }[type]
    raise("no value found for type :#{type} and key :#{key}") unless return_value
    return_value
  end

  def self.derived_filepath(key, xlsx_path)
    filepath = {
      cell_values_and_formats: 'spec/support/static_test_files/cell_values_and_formats.xlsx'
    }[key]
    filepath || "#{xlsx_path}/#{key}.xlsx"
  end

  def self.raw_cell_value_test_data_hash
    {
      cell_raw_string_read: {
        worksheet_name: 'string',
        cell_range: /^B[2-7]$/,
        description: 'with string input as String :text regardless of number format'
      },
      cell_raw_number_read: {
        worksheet_name: 'number',
        cell_range: /^B[2-4]$/,
        description: 'with whole number input as String :text for text number format, as FixNum :number for general/number number formats'
      },
      cell_raw_float_read: {
        worksheet_name: 'float',
        cell_range: /^B[3-4]$/,
        description: 'with float input as String :text for text number format, as FixNum :number for number number format'
      },
      cell_raw_date_read: {
        worksheet_name: 'date',
        cell_range: /^B(3|[5-7])$/,
        description: 'with date input as String :text for text/percentage format, as DateTime :date for time/date number formats'
      },
      cell_raw_time_read: {
        worksheet_name: 'time',
        cell_range: /^B(3|[6-7])$/,
        description: 'with time input as String :text for text/percentage format, as DateTime :time for time number format'
      },
      cell_raw_percentage_read: {
        worksheet_name: 'percentage',
        cell_range: /^B(3|7)$/,
        description: 'with percentage input as String :text for text format, as FixNum :number for percentage number format'
      },
      cell_raw_percentage_float_read: {
        worksheet_name: 'percentage_float',
        cell_range: /^B(3|7)$/,
        description: 'with percentage float input as String :text for text format, as Float :number for percentage number format'
      },
      cell_raw_empty_read: {
        worksheet_name: 'empty',
        cell_range: /^B[2-7]$/,
        description: 'with empty input as NilClass :general regardless of number format'
      }
    }
  end

  def self.formula_cell_value_test_data_hash
    {
      cell_formula_string_read: {
        worksheet_name: 'string',
        cell_range: /^C[2-7]$/,
        description: 'with string result as String :text regardless of number format, and collects formula'
      },
      cell_formula_number_read: {
        worksheet_name: 'number',
        cell_range: /^C[2-4]$/,
        description: 'with whole number result as String :text for text number format, as FixNum :number for general/number number formats, and collects formula'
      },
      cell_formula_float_read: {
        worksheet_name: 'float',
        cell_range: /^C[3-4]$/,
        description: 'with float result as String :text for text number format, as FixNum :number for number number format, and collects formula'
      },
      cell_formula_date_read: {
        worksheet_name: 'date',
        cell_range: /^C(3|5|6)$/,
        description: 'with date result as String :text for text number format, as DateTime :date for time/date number formats, and collects formula'
      },
      cell_formula_time_read: {
        worksheet_name: 'time',
        cell_range: /^C(3|6)$/,
        description: 'with time result as String :text for text format, as DateTime :time for time number format, and collects formula'
      },
      cell_formula_percentage_read: {
        worksheet_name: 'percentage',
        cell_range: /^C(3|7)$/,
        description: 'with percentage result as String :text for text format, as FixNum :number for percentage number format, and collects formula'
      },
      cell_formula_percentage_float_read: {
        worksheet_name: 'percentage_float',
        cell_range: /^C(3|7)$/,
        description: 'with percentage float result as String :text for text format, as Float :number for percentage number format, and collects formula'
      }
    }
  end

  def self.read_and_test_cell_values(test, expected_key, worksheet_name, cell_range)
    read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
    cell_range = read_hash[worksheet_name][:cells].select { |key, _| key[cell_range] }
    test.expect(cell_range).to test.eq(RxlSpecHelpers.test_data(:expected_hash, expected_key))
  end

  def self.non_string_key_arrays
    [
      [:worksheet_a],
      ['worksheet_a', :worksheet_b],
      [0, 'worksheet_b'],
      ['worksheet_a', nil],
      [[], 'worksheet_b'],
      ['worksheet_a', {}],
      [true, 'worksheet_b'],
      ['worksheet_a', false]
    ]
  end

  def self.empty_string_key_arrays
    [
      [''],
      ['worksheet_a', ''],
      ['', 'worksheet_a'],
      ['worksheet_a', '', 'worksheet_b']
    ]
  end

  def self.example_class_values
    [
      nil,
      true,
      false,
      '',
      'abc',
      :cells,
      0,
      [],
      %w[a b c],
      [1, 2, 3],
      {},
      { a: 1, b: 2 },
      {}.to_json
    ]
  end

  def self.invalid_string_cell_keys
    keys = %w[
      !
      A!
      !1
      1A
      aaa
      A
      ZZZ
      0
      1234
      123a
      123A
      1a2
      1A2
      a1a
      A1A
      AAAA1
      A11111111
    ]
    keys + %i[invalid A1]
  end

  def self.invalid_cell_hash_key_arrays
    [
      %i[a b c],
      %i[value number formula cell]
    ]
  end

end
