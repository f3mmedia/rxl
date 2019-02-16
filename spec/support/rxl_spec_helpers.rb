module RxlSpecHelpers

  def self.hash_cell_template
    {
      value: nil,
      format: :general,
      formula: nil,
      h_align: nil,
      v_align: :bottom
    }
  end

  def self.create_temp_xlsx_dir_unless_exists
    path = Pathname.new(ENV['TEMP_XLSX_PATH'])
    FileUtils.mkdir(path.to_s) unless path.exist?
  end

  def self.destroy_temp_xlsx_dir_if_exists
    path = Pathname.new(ENV['TEMP_XLSX_PATH'])
    FileUtils.rmtree(path.to_s) if path.exist?
  end

  def self.generate_test_excel_file(test, key, path)
    filepath = test_data(:filepath, key, path: path)
    Rxl.write_file(filepath, test_data(:write_hash, key))
    path = Pathname.new(filepath)
    test.expect(path.exist?)
  end

  def self.test_data(type, key = nil, args = {})
    return_value = {
      filepath: "#{args[:path]}/#{key}.xlsx",
      write_hash: {
        empty_file: {},
        save_as_table: save_as_table_hashes,
        save_as_table_with_formatting: save_as_table_with_formatting_hashes,
        save_with_content: save_with_content_hash,
        test_file: {},
        worksheet_names: { 'test_a' => {}, 'test_b' => {} }
      }[key],
      expected_hash: {
        empty_file: { 'Sheet1' => {} },
        test_file: { 'Sheet1' => {} },
        test_table_file: { 'Sheet1' => [] },
        worksheet_names: {
          'test_a' => {},
          'test_b' => {}
        },
        cell_raw_string_read: {
          'B2' => hash_cell_template.merge(value: 'abcde', format: :text),
          'B3' => hash_cell_template.merge(value: 'abcde', format: :text),
          'B4' => hash_cell_template.merge(value: 'abcde', format: :text),
          'B5' => hash_cell_template.merge(value: 'abcde', format: :text),
          'B6' => hash_cell_template.merge(value: 'abcde', format: :text),
          'B7' => hash_cell_template.merge(value: 'abcde', format: :text)
        },
        cell_raw_number_read: {
          'B2' => hash_cell_template.merge(value: 12345, format: :number),
          'B3' => hash_cell_template.merge(value: '12345', format: :text),
          'B4' => hash_cell_template.merge(value: 12345, format: :number)
        },
        cell_raw_float_read: {
          'B3' => hash_cell_template.merge(value: '123.45', format: :text),
          'B4' => hash_cell_template.merge(value: 123.45, format: :number)
        },
        cell_raw_date_read: {
          'B3' => hash_cell_template.merge(value: '01/01/2000', format: :text),
          'B5' => hash_cell_template.merge(value: DateTime.parse('01/01/2000'), format: :date),
          'B6' => hash_cell_template.merge(value: DateTime.parse('01/01/2000'), format: :date),
          'B7' => hash_cell_template.merge(value: '01/01/2000%', format: :text)
        },
        cell_raw_time_read: {
          'B3' => hash_cell_template.merge(value: '10:15:30', format: :text),
          'B6' => hash_cell_template.merge(value: DateTime.parse('31/12/1899 10:15:30'), format: :time),
          'B7' => hash_cell_template.merge(value: '10:15:30%', format: :text)
        },
        cell_raw_percentage_read: {
          'B3' => hash_cell_template.merge(value: '100%', format: :text),
          'B7' => hash_cell_template.merge(value: 1, format: :number)
        },
        cell_raw_percentage_float_read: {
          'B3' => hash_cell_template.merge(value: '123.45%', format: :text),
          'B7' => hash_cell_template.merge(value: 1.2345, format: :number)
        },
        cell_raw_empty_read: {
          'B2' => hash_cell_template.merge(format: :general),
          'B3' => hash_cell_template.merge(format: :general),
          'B4' => hash_cell_template.merge(format: :general),
          'B5' => hash_cell_template.merge(format: :general),
          'B6' => hash_cell_template.merge(format: :general),
          'B7' => hash_cell_template.merge(format: :general)
        },
        cell_formula_string_read: {
          'C2' => hash_cell_template.merge(value: 'abcde', format: :text, formula: 'CONCATENATE("abc","de")'),
          'C3' => hash_cell_template.merge(value: 'abcde', format: :text, formula: 'CONCATENATE("abc","de")'),
          'C4' => hash_cell_template.merge(value: 'abcde', format: :text),
          'C5' => hash_cell_template.merge(value: 'abcde', format: :text),
          'C6' => hash_cell_template.merge(value: 'abcde', format: :text, formula: 'CONCATENATE("abc","de")'),
          'C7' => hash_cell_template.merge(value: 'abcde', format: :text)
        },
        cell_formula_number_read: {
          'C2' => hash_cell_template.merge(value: 12345, format: :number, formula: '12340+5'),
          'C3' => hash_cell_template.merge(value: '=12340+5', format: :text),
          'C4' => hash_cell_template.merge(value: 12345, format: :number, formula: '12340+5')
        },
        cell_formula_float_read: {
          'C3' => hash_cell_template.merge(value: '=123.41+0.04', format: :text),
          'C4' => hash_cell_template.merge(value: 123.45, format: :number, formula: '123.41+0.04')
        },
        cell_formula_date_read: {
          'C3' => hash_cell_template.merge(value: '=DATE(2000,1,1)', format: :text),
          'C5' => hash_cell_template.merge(value: DateTime.parse('01/01/2000'), format: :date, formula: 'DATE(2000,1,1)'),
          'C6' => hash_cell_template.merge(value: DateTime.parse('01/01/2000'), format: :date, formula: 'DATE(2000,1,1)')
        },
        cell_formula_time_read: {
          'C3' => hash_cell_template.merge(value: '=TIME(10,15,30)', format: :text),
          'C6' => hash_cell_template.merge(value: DateTime.parse('31/12/1899 10:15:30'), format: :time, formula: 'TIME(10,15,30)')
        },
        cell_formula_percentage_read: {
          'C3' => hash_cell_template.merge(value: '=50%+50%', format: :text),
          'C7' => hash_cell_template.merge(value: 1, format: :number, formula: '50%+50%')
        },
        cell_formula_percentage_float_read: {
          'C3' => hash_cell_template.merge(value: '=123.41%+0.04%', format: :text),
          'C7' => hash_cell_template.merge(value: 1.2345, format: :number, formula: '123.41%+0.04%')
        },
        horizontal_and_vertical_alignment: horizontal_and_vertical_alignment_expected_hash,
        as_tables: as_tables_expected_hash,
        tables_ignore_no_header_columns: tables_ignore_no_header_columns_expected_hash
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

  def self.save_with_content_hash
    {
      'first_sheet' => {
        'A1' => { value: 'cell_a1' },
        'A2' => { value: 'cell_a2' }
      }
    }
  end

  def self.horizontal_and_vertical_alignment_expected_hash
    {
      'A1' => {
        value: 'abc',
        format: :text,
        formula: nil,
        h_align: :left,
        v_align: :top
      },
      'B1' => {
        value: 'abc',
        format: :text,
        formula: nil,
        h_align: :center,
        v_align: :top
      },
      'C1' => {
        value: 'abc',
        format: :text,
        formula: nil,
        h_align: :right,
        v_align: :top
      },
      'A2' => {
        value: 'abc',
        format: :text,
        formula: nil,
        h_align: :left,
        v_align: :center
      },
      'B2' => {
        value: 'abc',
        format: :text,
        formula: nil,
        h_align: :center,
        v_align: :center
      },
      'C2' => {
        value: 'abc',
        format: :text,
        formula: nil,
        h_align: :right,
        v_align: :center
      },
      'A3' => {
        value: 'abc',
        format: :text,
        formula: nil,
        h_align: :left,
        v_align: :bottom
      },
      'B3' => {
        value: 'abc',
        format: :text,
        formula: nil,
        h_align: :center,
        v_align: :bottom
      },
      'C3' => {
        value: 'abc',
        format: :text,
        formula: nil,
        h_align: :right,
        v_align: :bottom
      },
      'A4' => {
        value: 'abc',
        format: :text,
        formula: nil,
        h_align: nil,
        v_align: :bottom
      },
      'B4' => {
        value: nil,
        format: :general,
        formula: nil,
        h_align: nil,
        v_align: :bottom
      },
      'C4' => {
        value: nil,
        format: :general,
        formula: nil,
        h_align: nil,
        v_align: :bottom
      }
    }
  end

  def self.as_tables_expected_hash
    {
      'Sheet1' => [
        {
          'header_1' => 'row_1_a',
          'header_2' => 'row_1_b',
          'header_3' => 'row_1_c'
        },
        {
          'header_1' => 'row_2_a',
          'header_2' => 'row_2_b',
          'header_3' => 'row_2_c'
        },
        {
          'header_1' => nil,
          'header_2' => nil,
          'header_3' => nil
        },
        {
          'header_1' => 'sum_1',
          'header_2' => 'sum_2',
          'header_3' => 'sum_3'
        }
      ],
      'Sheet2' => [
        {
          'header_1' => 'row_1_a',
          'header_2' => 'row_1_b'
        },
        {
          'header_1' => 'row_2_a',
          'header_2' => 'row_2_b'
        },
        {
          'header_1' => 'row_3_a',
          'header_2' => 'row_3_b'
        },
        {
          'header_1' => 'row_4_a',
          'header_2' => 'row_4_b'
        },
        {
          'header_1' => 'row_5_a',
          'header_2' => 'row_5_b'
        },
        {
          'header_1' => 'row_6_a',
          'header_2' => 'row_6_b'
        },
        {
          'header_1' => 'row_7_a',
          'header_2' => 'row_7_b'
        },
        {
          'header_1' => 'row_8_a',
          'header_2' => 'row_8_b'
        },
        {
          'header_1' => nil,
          'header_2' => nil
        },
        {
          'header_1' => nil,
          'header_2' => nil
        },
        {
          'header_1' => 'sum_1',
          'header_2' => 'sum_2'
        }
      ]
    }
  end

  def self.tables_ignore_no_header_columns_expected_hash
    {
      'Sheet1' => [
        {
          'header_1' => 'row_1_b',
          'header_3' => 'row_1_d'
        },
        {
          'header_1' => 'row_2_b',
          'header_3' => 'row_2_d'
        }
      ]
    }
  end

  def self.save_as_table_hashes
    {
      'test' => [
        {
          'col_1' => 'r1c1',
          'col_2' => 'r1c2'
        },
        {
          'col_1' => 'r2c1',
          'col_2' => 'r2c2'
        }
      ]
    }
  end

  def self.save_as_table_with_formatting_hashes
    {
      formats: {
        'Sheet1' => {
          headers: {
            bold: true,
            h_align: 'center'
          },
          'B' => {
            fill: 'feb302'
          }
        }
      },
      'Sheet1' => [
        {
          'col_1' => 'some_value',
          'col_2' => 'other_value'
        },
        {
          'col_1' => 'some_value',
          'col_2' => 'other_value'
        }
      ]
    }
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
    path = 'spec/support/static_test_files'
    read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats, path: path))
    cell_range = read_hash[worksheet_name].select { |key, _| key[cell_range] }
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
