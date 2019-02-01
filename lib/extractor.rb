require 'xlsx_to_pqc_xml'
require 'yaml'
require 'json'

class XlsxDataExtractor
  def initialize(xlsx_path)
    @errors = []
    @path = xlsx_path
  end

  def add_error(sheet_name, msg)
    error_string =  "  -- [#{sheet_name}]#{msg}"
    @errors << error_string
  end

  def clean_hash(hash)
    for_deletion = []
    hash.each do |key, val|
      case val
      when Hash
        for_deletion << key if clean_hash(val).empty?
      when Arrays
        for_deletion << key if clean_array(val).empty?
      when String
        for_deletion << key if val.strip.empty?
      when NilClass
        for_deletion << key
      end
    end
    for_deletion.each{ |k| hash.delete(k)}
    hash
  end

  def clean_array(array)
    for_deletion = []
    array.each do |entry|
      case entry
      when Hash
        for_deletion << entry if clean_hash(entry).empty?
      when Array
        for_deletion << entry if clean_array(entry).empty?
      when String
        for_deletion << entry if entry.strip.empty?
      end
    end
    for_deletion.each{ |e| array.delete(e)}
    array.compact
  end
end


ExtractorResult = Struct.new(:result) do
  def to_h
    result
  end

  def to_json
    JSON.pretty_generate(result)
  end
end