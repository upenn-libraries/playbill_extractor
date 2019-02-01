require 'xlsx_to_pqc_xml'
require 'yaml'
require 'json'
require 'pry'

# ------------ ^ for now ^ ---------------

require_relative '../lib/furness/furness_extractor'

XLSX_PATH   = '../lib/furness/doc/Furness_image_collection_data_good.xlsx'

#XlsxToPqcXml::XlsxData.new(xlsx_path: XLSX_PATH, config: YAML.load(open(CONFIG_PATH))).data.each do |d|
#  puts JSON.pretty_generate(FurnessExtractor.new(nil).clean_hash(FurnessExtractor.new(nil).combine_data(d)))
#end

um = FurnessExtractor.new(XLSX_PATH).get_result

# um = FurnessExtractor.new(nil).grand_combine(XlsxToPqcXml::XlsxData.new(xlsx_path: XLSX_PATH, config: YAML.load(open(CONFIG_PATH))).data)
puts um.to_json

=begin
  require_relative '../lib/furness/furness_extractor'

  # file_path = ARGV.first

  # abort 'No file given' unless file_path
  # abort "No file found at #{file_path}" unless File.exist?(folder_path)

  File.open('write_path', 'w+'){ |f| f.puts FurnessExtractor.new(file_path).get_result}
=end