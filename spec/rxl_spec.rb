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
      path = ENV['TEMP_XLSX_PATH']
      RxlSpecHelpers.generate_test_excel_file(self, :empty_file, path)
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :empty_file, path: path))
      expect(read_hash).to eq(RxlSpecHelpers.test_data(:expected_hash, :empty_file))
    end

    it 'reads in worksheet names' do
      path = ENV['TEMP_XLSX_PATH']
      RxlSpecHelpers.generate_test_excel_file(self, :worksheet_names, path)
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :worksheet_names, path: path))
      expect(read_hash).to eq(RxlSpecHelpers.test_data(:expected_hash, :worksheet_names))
    end

    context 'reads cell raw values' do
      RxlSpecHelpers.raw_cell_value_test_data_hash.each do |expected_key, value_hash|
        it value_hash[:description] do
          RxlSpecHelpers.read_and_test_cell_values(self, expected_key, value_hash[:worksheet_name], value_hash[:cell_range])
        end
      end
    end

    context 'reads cell formula values' do
      RxlSpecHelpers.formula_cell_value_test_data_hash.each do |expected_key, value_hash|
        it value_hash[:description] do
          RxlSpecHelpers.read_and_test_cell_values(self, expected_key, value_hash[:worksheet_name], value_hash[:cell_range])
        end
      end
    end

    it 'reads horizontal and vertical cell alignment' do
      path = ENV['TEST_XLSX_FILES']
      read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :horizontal_and_vertical_alignment, path: path))
      expect(read_hash['values']).to eq(RxlSpecHelpers.test_data(:expected_hash, :horizontal_and_vertical_alignment))
    end

    context 'when reading workbook as tables' do

      it 'reads in sheets as tables' do
        path = ENV['TEST_XLSX_FILES']
        read_hash = Rxl.read_file_as_tables(RxlSpecHelpers.test_data(:filepath, :as_tables, path: path))
        expected_hash = RxlSpecHelpers.test_data(:expected_hash, :as_tables)
        expect(read_hash.keys).to eq(expected_hash.keys)
        expected_hash.keys.each do |key|
          expect(read_hash[key]).to be_a(Array)
          expect(read_hash[key].length).to eq(expected_hash[key].length)
          expected_hash[key].each do |item|
            expect(read_hash[key]).to include(item)
          end
        end
      end

      it 'does not include columns with no header' do
        path = ENV['TEST_XLSX_FILES']
        read_hash = Rxl.read_file_as_tables(RxlSpecHelpers.test_data(:filepath, :tables_ignore_no_header_columns, path: path))
        expected_hash = RxlSpecHelpers.test_data(:expected_hash, :tables_ignore_no_header_columns)
        expect(read_hash['Sheet1']).to be_a(Array)
        expect(read_hash['Sheet1'].length).to eq(expected_hash['Sheet1'].length)
        expected_hash['Sheet1'].each do |item|
          expect(read_hash['Sheet1']).to include(item)
        end
      end

    end

    context 'when reading multiple files' do

      it 'returns multiple files as a hash when given a hash of filepaths' do
        path = ENV['TEMP_XLSX_PATH']
        RxlSpecHelpers.generate_test_excel_file(self, :test_file, path)
        input = {
          first_file: RxlSpecHelpers.test_data(:filepath, :test_file, path: path),
          second_file: RxlSpecHelpers.test_data(:filepath, :test_file, path: path)
        }
        read_hash = Rxl.read_files(input)
        expect(read_hash).to be_a(Hash)
        expect(read_hash.keys).to eq(%i[first_file second_file])
        %i[first_file second_file].each do |key|
          expect(read_hash[key]).to eq(RxlSpecHelpers.test_data(:expected_hash, :test_file))
        end
      end

      it 'returns multiple files as tables as a hash when given a hash of filepaths' do
        path = ENV['TEMP_XLSX_PATH']
        RxlSpecHelpers.generate_test_excel_file(self, :test_file, path)
        input = {
          first_file: RxlSpecHelpers.test_data(:filepath, :test_file, path: path),
          second_file: RxlSpecHelpers.test_data(:filepath, :test_file, path: path)
        }
        read_hash = Rxl.read_files(input, :as_tables)
        expect(read_hash).to be_a(Hash)
        expect(read_hash.keys).to eq(%i[first_file second_file])
        %i[first_file second_file].each do |key|
          expect(read_hash[key]).to eq(RxlSpecHelpers.test_data(:expected_hash, :test_table_file))
        end
      end
    end
  end

  # MANUAL TESTS FOR SUCCESSFUL WRITE
  # see the manual_tests directory for the raw ruby
  # hashes and expected outcomes for writing each with Rxl

  context 'when writing an excel file' do
    context 'it saves successfully' do
      it 'saves an empty hash as a file with the specified file name and a single empty sheet as "Sheet1"' do
        path = ENV['TEMP_XLSX_PATH']
        Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :empty_file, path: path), {})
        expect(Pathname(RxlSpecHelpers.test_data(:filepath, :empty_file, path: path)).exist?).to eq(true)
        read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :empty_file, path: path))
        expect(read_hash.keys.length).to eq(1)
        expect(read_hash.keys[0]).to eq('Sheet1')
        expect(read_hash['Sheet1']).to eq({})
      end

      context 'saves one or more sheets with the name specified and removes "Sheet1"' do
        worksheet_name_arrays = [
          ['worksheet_a'],
          %w[a b c].map { |id| "worksheet_#{id}" },
          ('a'..'dd').map { |id| "worksheet_#{id}" }
        ]
        worksheet_name_arrays.each_with_index do |worksheet_name_array, i|
          it "[example ##{i + 1}]" do
            hash_workbook_input = worksheet_name_array.each_with_object({}) do |worksheet_name, hash|
              hash[worksheet_name] = {}
            end
            path = ENV['TEMP_XLSX_PATH']
            Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :empty_file, path: path), hash_workbook_input)
            read_hash = Rxl.read_file(RxlSpecHelpers.test_data(:filepath, :empty_file, path: path))
            expect(read_hash.keys.length).to eq(worksheet_name_array.length)
            expect(read_hash.keys).to eq(worksheet_name_array)
            worksheet_name_array.each do |worksheet_name|
              expect(read_hash[worksheet_name]).to eq({})
            end
          end
        end
      end

      it 'saves a file with worksheet content' do
        path = ENV['TEMP_XLSX_PATH']
        save_hash = RxlSpecHelpers.test_data(:write_hash, :save_with_content)
        save_filepath = RxlSpecHelpers.test_data(:filepath, :save_with_content, path: path)
        Rxl.write_file(save_filepath, save_hash)
        read_hash = Rxl.read_file(save_filepath)
        expect(read_hash).to be_a(Hash)
        expect(read_hash.keys).to eq(['first_sheet'])
        expect(read_hash['first_sheet'].keys).to eq(%w[A1 A2])
        expect(read_hash['first_sheet']['A1'][:value]).to eq('cell_a1')
      end

      it 'saves an array of hashes as a table' do
        path = ENV['TEMP_XLSX_PATH']
        save_hash = RxlSpecHelpers.test_data(:write_hash, :save_as_table)
        save_filepath = RxlSpecHelpers.test_data(:filepath, :save_as_table, path: path)
        order = %w[col_1 col_2]
        error = Rxl.write_file_as_tables(save_filepath, save_hash, order)
        expect(error).to be_nil
        read_hash = Rxl.read_file_as_tables(save_filepath)
        expect(read_hash).to be_a(Hash)
        expect(read_hash.keys).to eq(['test'])
        expect(read_hash['test']).to be_a(Array)
        expect(read_hash).to eq(save_hash)
      end

      it 'saves an array of hashes as a table without headers' do
        path = ENV['TEMP_XLSX_PATH']
        save_hash = RxlSpecHelpers.test_data(:write_hash, :save_as_table)
        save_filepath = RxlSpecHelpers.test_data(:filepath, :save_as_table, path: path)
        order = %w[col_1 col_2]
        error = Rxl.write_file_as_tables(save_filepath, save_hash, order, write_headers: false)
        expect(error).to be_nil
        read_hash = Rxl.read_file(save_filepath)
        expect(read_hash).to be_a(Hash)
        expect(read_hash.keys).to eq(['test'])
        expect(read_hash['test']).to be_a(Hash)
        expect(read_hash['test']['A1'][:value]).to eq('r1c1')
      end
    end

    context 'it returns an exception where' do
      context 'the workbook is not a hash' do
        non_hash_values = RxlSpecHelpers.example_class_values.delete_if { |value| value.is_a?(Hash) }
        non_hash_values.each_with_index do |hash_workbook_input, i|
          it "[example ##{i + 1}]" do
            exception = Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :hash_validation), hash_workbook_input)
            expect(exception.class).to eq(RuntimeError)
            expect(exception.message).to eq(RxlSpecHelpers.test_data(:validation, :non_hash_workbook))
          end
        end
      end

      context 'the workbook contains non-string keys' do
        RxlSpecHelpers.non_string_key_arrays.each_with_index do |key_array, i|
          it "[example ##{i + 1}]" do
            hash_workbook_input = key_array.each_with_object({}) { |key, hash| hash[key] = {} }
            exception = Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :hash_validation), hash_workbook_input)
            expect(exception.class).to eq(RuntimeError)
            expect(exception.message).to eq(RxlSpecHelpers.test_data(:validation, :non_string_worksheet_name))
          end
        end
      end

      context 'the workbook contains an empty string key' do
        RxlSpecHelpers.empty_string_key_arrays.each_with_index do |key_array, i|
          it "[example ##{i + 1}]" do
            hash_workbook_input = key_array.each_with_object({}) { |key, hash| hash[key] = {} }
            exception = Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :hash_validation), hash_workbook_input)
            expect(exception.class).to eq(RuntimeError)
            expect(exception.message).to eq(RxlSpecHelpers.test_data(:validation, :empty_string_worksheet_name))
          end
        end
      end

      context 'the worksheet is not a hash' do
        non_hash_values = RxlSpecHelpers.example_class_values.delete_if { |value| value.is_a?(Hash) }
        non_hash_values.each_with_index do |worksheet_input, i|
          it "[example ##{i + 1}]" do
            filepath = RxlSpecHelpers.test_data(:filepath, :hash_validation)
            exception = Rxl.write_file(filepath, {'worksheet_a' => worksheet_input})
            expect(exception.class).to eq(RuntimeError)
            expected_message = RxlSpecHelpers.test_data(:validation, :non_hash_worksheet, path: ['worksheet_a'])
            expect(exception.message).to eq(expected_message)
          end
        end
      end

      context 'the worksheet contains keys that are not valid excel cell keys' do
        RxlSpecHelpers.invalid_string_cell_keys.each_with_index do |key, i|
          it "[example ##{i + 1}]" do
            hash_workbook_input = { 'worksheet_a' => { key => {} } }
            exception = Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :hash_validation), hash_workbook_input)
            expect(exception.class).to eq(RuntimeError)
            expected_message = RxlSpecHelpers.test_data(:validation, :invalid_cell_key, path: ['worksheet_a'])
            expect(exception.message).to eq(expected_message)
          end
        end
      end

      context 'the worksheet contains values that are not hashes' do
        non_hash_values = RxlSpecHelpers.example_class_values.delete_if { |value| value.is_a?(Hash) }
        non_hash_values.each_with_index do |cell_value, i|
          it "[example ##{i + 1}]" do
            hash_workbook_input = { 'worksheet_a' => { 'A1' => cell_value } }
            exception = Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :hash_validation), hash_workbook_input)
            expect(exception.class).to eq(RuntimeError)
            expected_message = RxlSpecHelpers.test_data(:validation, :non_hash_cell_value, path: ['worksheet_a', 'A1'])
            expect(exception.message).to eq(expected_message)
          end
        end
      end

      context 'the cell hash contains non-symbol keys' do
        non_symbol_values = RxlSpecHelpers.example_class_values.reject { |value| value.is_a?(Symbol) }
        non_symbol_values.each_with_index do |key, i|
          it "[example ##{i + 1}]" do
            hash_workbook_input = { 'worksheet_a' => { 'A1' => { key => nil } } }
            exception = Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :hash_validation), hash_workbook_input)
            expect(exception.class).to eq(RuntimeError)
            expected_message = RxlSpecHelpers.test_data(:validation, :non_symbol_cell_hash_key, path: ['worksheet_a', 'A1'])
            expect(exception.message).to eq(expected_message)
          end
        end
      end

      context 'the cell hash contains invalid keys' do
        RxlSpecHelpers.invalid_cell_hash_key_arrays.each_with_index do |key_array, i|
          it "[example ##{i + 1}]" do
            cell_hash = key_array.each_with_object({}) { |key, hash| hash[key] = nil }
            hash_workbook_input = { 'worksheet_a' => { 'A1' => cell_hash } }
            exception = Rxl.write_file(RxlSpecHelpers.test_data(:filepath, :hash_validation), hash_workbook_input)
            expect(exception.class).to eq(RuntimeError)
            args = { path: ['worksheet_a', 'A1'], valid_cell_keys_string: ':value, :number, :formula' }
            expected_message = RxlSpecHelpers.test_data(:validation, :invalid_cell_hash_key, args)
            expect(exception.message).to eq(expected_message)
          end
        end
      end
    end
  end
end
