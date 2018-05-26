describe Rxl do

  before(:all) do
    RxlSpecHelpers.destroy_temp_xlsx_dir_if_exists
    RxlSpecHelpers.create_temp_xlsx_dir_unless_exists
  end

  after(:all) do
    RxlSpecHelpers.destroy_temp_xlsx_dir_if_exists
  end

  context 'when reading in an excel file' do

    it 'returns a hash_workbook with one empty hash_worksheet from an empty file' do
      RxlSpecHelpers.generate_test_excel_file(self, :empty_file)
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :empty_file))
      expect(read_hash).to eq(RxlSpecHelpers.test_data(:expected_hash, :empty_file))
    end

    it 'reads in worksheet names' do
      RxlSpecHelpers.generate_test_excel_file(self, :worksheet_names)
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :worksheet_names))
      expect(read_hash).to eq(RxlSpecHelpers.test_data(:expected_hash, :worksheet_names))
    end

  end

  # NB "number format" means the selection from the excel number dropdown

  context 'reads cell raw values' do

    it 'with string input as String :text regardless of number format' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['string'][:cells].select { |key, _| key[/^B[2-7]$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_raw_string_read))
    end

    it 'with whole number input as String :text for text number format, as FixNum :number for general/number number formats' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['number'][:cells].select { |key, _| key[/^B[2-4]$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_raw_number_read))
    end

    it 'with float input as String :text for text number format, as FixNum :number for number number format' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['float'][:cells].select { |key, _| key[/^B[3-4]$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_raw_float_read))
    end

    it 'with date input as String :text for text/percentage format, as DateTime :date for time/date number formats' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['date'][:cells].select { |key, _| key[/^B(3|[5-7])$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_raw_date_read))
    end

    it 'with time input as String :text for text/percentage format, as DateTime :time for time number format' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['time'][:cells].select { |key, _| key[/^B(3|[6-7])$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_raw_time_read))
    end

    it 'with percentage input as String :text for text format, as FixNum :number for percentage number format' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['percentage'][:cells].select { |key, _| key[/^B(3|7)$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_raw_percentage_read))
    end

    it 'with percentage float input as String :text for text format, as Float :number for percentage number format' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['percentage_float'][:cells].select { |key, _| key[/^B(3|7)$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_raw_percentage_float_read))
    end

    it 'with empty input as NilClass :general regardless of number format' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['empty'][:cells].select { |key, _| key[/^B[2-7]$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_raw_empty_read))
    end

  end

  context 'reads cell formula values' do

    it 'with string result as String :text regardless of number format, and collects formula' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['string'][:cells].select { |key, _| key[/^C[2-7]$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_formula_string_read))
    end

    it 'with whole number result as String :text for text number format, as FixNum :number for general/number number formats, and collects formula' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['number'][:cells].select { |key, _| key[/^C[2-4]$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_formula_number_read))
    end

    it 'with float result as String :text for text number format, as FixNum :number for number number format, and collects formula' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['float'][:cells].select { |key, _| key[/^C[3-4]$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_formula_float_read))
    end

    it 'with date result as String :text for text number format, as DateTime :date for time/date number formats, and collects formula' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['date'][:cells].select { |key, _| key[/^C(3|5|6)$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_formula_date_read))
    end

    it 'with time result as String :text for text format, as DateTime :time for time number format, and collects formula' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['time'][:cells].select { |key, _| key[/^C(3|6)$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_formula_time_read))
    end

    it 'with percentage result as String :text for text format, as FixNum :number for percentage number format, and collects formula' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['percentage'][:cells].select { |key, _| key[/^C(3|7)$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_formula_percentage_read))
    end

    it 'with percentage float result as String :text for text format, as Float :number for percentage number format, and collects formula' do
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :cell_values_and_formats))
      cell_range = read_hash['percentage_float'][:cells].select { |key, _| key[/^C(3|7)$/] }
      expect(cell_range).to eq(RxlSpecHelpers.test_data(:expected_hash, :cell_formula_percentage_float_read))
    end

  end

  context 'when writing an excel file' do

    it 'saves an empty hash as a file with the specified file name and a single empty sheet as "Sheet1"' do
      Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :empty_file), {})
      expect(Pathname(RxlSpecHelpers.test_data(:filepath, :empty_file)).exist?).to eq(true)
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :empty_file))
      expect(read_hash.keys.length).to eq(1)
      expect(read_hash.keys[0]).to eq('Sheet1')
      expect(read_hash['Sheet1']).to eq({rows: {}, columns: {}, cells: {}})
    end

    it 'saves one or more sheets with the name specified and removes "Sheet1"' do
      worksheet_name_arrays = [
          ['worksheet_a'],
          %w[a b c].map { |id| "worksheet_#{id}" },
          ('a'..'dd').map { |id| "worksheet_#{id}" }
      ]
      worksheet_name_arrays.each do |worksheet_name_array|
        hash_workbook_input = worksheet_name_array.each_with_object({}) do |worksheet_name, hash|
          hash[worksheet_name] = {}
        end
        Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :empty_file), hash_workbook_input)
        read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :empty_file))
        expect(read_hash.keys.length).to eq(worksheet_name_array.length)
        expect(read_hash.keys).to eq(worksheet_name_array)
        worksheet_name_array.each do |worksheet_name|
          expect(read_hash[worksheet_name]).to eq({rows: {}, columns: {}, cells: {}})
        end
      end
    end

    # MANUAL TESTS FOR SUCCESSFUL WRITE
    # see the manual_tests directory for the raw ruby
    # hashes and expected outcomes for writing each with Rxl

    context 'it returns an exception where' do

      context 'the hash_workbook is not a hash' do
        hash_workbook_inputs = [nil, true, false, '', 'abc', 0, [], ['a', 'b', 'c'], [1, 2, 3], {}.to_json]
        hash_workbook_inputs.each_with_index do |hash_workbook_input, i|
          it "[example ##{i}]" do
            exception = Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :hash_validation), hash_workbook_input)
            expect(exception.message).to eq(RxlSpecHelpers.test_data(:validation, :non_hash_workbook))
          end
        end
      end

      context 'the hash_workbook contains non-string keys' do
        key_arrays = [
            [:worksheet_a],
            ['worksheet_a', :worksheet_b],
            [0, 'worksheet_b'],
            ['worksheet_a', nil],
            [[], 'worksheet_b'],
            ['worksheet_a', {}],
            [true, 'worksheet_b'],
            ['worksheet_a', false]
        ]
        key_arrays.each_with_index do |key_array, i|
          it "[example ##{i}]" do
            hash_workbook_input = key_array.each_with_object({}) { |key, hash| hash[key] = {} }
            exception = Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :hash_validation), hash_workbook_input)
            expect(exception.message).to eq(RxlSpecHelpers.test_data(:validation, :non_string_worksheet_name))
          end
        end
      end

    end

  end

end
