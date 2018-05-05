describe Rxl do

  before(:all) do
    RxlSpecHelpers.destroy_temp_xlsx_dir_if_exists
    RxlSpecHelpers.create_temp_xlsx_dir_unless_exists
  end

  after(:all) do
    RxlSpecHelpers.destroy_temp_xlsx_dir_if_exists
  end

  it 'reads an empty file and returns a hash_workbook with one empty hash_worksheet' do
    RxlSpecHelpers.generate_test_excel_file(self, :empty_xlsx)
    RxlSpecHelpers.verify_read_hash_matches_expected(self, :empty_xlsx)
  end

end
