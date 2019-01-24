require 'xlsx_to_pqc_xml'
require 'yaml'
require 'json'
require 'pry'

XLSX_PATH   = '../lib/furness/doc/Furness_image_collection_data_good.xlsx'
CONFIG_PATH = '../lib/furness/furness_config.yml'

pp XlsxToPqcXml::XlsxData.new(xlsx_path: XLSX_PATH, config: YAML.load(open(CONFIG_PATH)))
