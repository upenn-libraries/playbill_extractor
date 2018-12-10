require 'json'
require './extractor'

path = ARGV.first

puts JSON.pretty_generate PlaybillExtractor.new(path).output_data
