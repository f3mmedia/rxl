$LOAD_PATH.unshift File.expand_path('../../lib', __FILE__)
require 'rxl'
require 'rspec'
require 'fileutils'
require 'json'
require 'pathname'
require_relative 'support/rxl_spec_helpers'

ENV['TEMP_XLSX_PATH'] = 'spec/temp_xlsx'

