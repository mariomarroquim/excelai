#!/usr/bin/ruby

require "ai4r"
require "nokogiri"
require "zip"
require "rubyXL"

excel = RubyXL::Parser.parse ARGV.first
real_lines = excel.worksheets[0].extract_data

data_set = Ai4r::Data::DataSet.new data_items: real_lines[1..(real_lines.size - 1)],
                                   data_labels: real_lines.first

id3 = Ai4r::Classifiers::ID3.new.build data_set

test_worksheet = excel.worksheets[1]

test_lines = test_worksheet.extract_data

test_lines[1..(test_lines.size - 1)].each_with_index do |columns, index|
  new_value = id3.eval(columns[0..(columns.size - 2)])
  test_worksheet[index + 1][columns.size - 1].change_contents new_value
end

excel.write ARGV.first
