$LOAD_PATH.unshift File.expand_path('../../lib', __FILE__)
require 'fileutils'
require 'json'
require 'pathname'
require 'rxl'
require 'rspec'
require_relative 'support/rxl_spec_helpers'

ENV['TEMP_XLSX_PATH'] = 'spec/temp_xlsx'
ENV['TEST_XLSX_FILES'] = 'spec/support/static_test_files'
