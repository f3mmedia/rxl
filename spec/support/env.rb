require 'rspec'
require 'fileutils'
require 'pathname'
require_relative 'excel_spec_helpers'
require_relative '../../lib/excel'

ENV['TEMP_XLSX_PATH'] = 'spec/temp_xlsx'
