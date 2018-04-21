require_relative 'support/env'

include ExcelSpecHelpers


describe 'Read excel file' do

  before(:all) do
    destroy_temp_xlsx_dir_if_exists
    create_temp_xlsx_dir_unless_exists
  end

  after(:all) do
    destroy_temp_xlsx_dir_if_exists
  end

  it 'should open and read an empty file' do
    spec = :empty_xlsx
    generate_test_excel_file_to_spec(spec)
    verify_read_hash_matches_expected(spec)
  end

  it 'should read in sheet names' do
    spec = :sheet_names_xlsx
    generate_test_excel_file_to_spec(spec)
    verify_read_hash_matches_expected(spec)
  end

end
