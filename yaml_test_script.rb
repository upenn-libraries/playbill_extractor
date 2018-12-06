require 'xlsx_to_pqc_xml'
require 'yaml'
require 'pry'
require 'pp'
require 'json'

require './script'

PE = PlaybillExtractor.new('test/fixtures/Playbills_9970657833503681.xlsx')

pp PE.extract_event
=begin
  config = YAML.load open('data/performance.yml').read
  config[:sheet_position] = 2
  xlsx_data = XlsxToPqcXml::XlsxData.new xlsx_path: '/Users/ggord/Desktop/Playbills_9970657833503681.xlsx', config: config
  pp xlsx_data.data

  # puts xlsx_data.errors # => {} empty errors hash

  p ARGV.first


  contributors = xlsx_data.data[1..-1]

  puts xlsx_data.data.first.merge({contributors: contributors}).to_json
=end
