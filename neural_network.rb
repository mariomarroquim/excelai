#!/usr/bin/ruby

require "ai4r"
require "nokogiri"
require "zip"
require "rubyXL"

excel = RubyXL::Parser.parse ARGV.first
real_lines = excel.worksheets[0].extract_data
real_columns = real_lines.first.size - 1

net = Ai4r::NeuralNetwork::Backpropagation.new [real_columns, real_columns, 1]

(real_lines.size*100).times do
  real_lines[1..(real_lines.size - 1)].each_with_index do |columns, index|
    net.train columns[0..(columns.size - 2)].collect{|i| i.to_i}, [columns.last.to_i]
  end
end

test_worksheet = excel.worksheets[1]

test_lines = test_worksheet.extract_data

test_lines[1..(test_lines.size - 1)].each_with_index do |columns, index|
  result = net.eval(columns[0..(columns.size - 2)].collect{|i| i.to_i}).first
  test_worksheet[index + 1][columns.size - 1].change_contents result
end

excel.write ARGV.first