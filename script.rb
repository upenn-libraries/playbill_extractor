
# ==================================================================================

# sheets = RubyXL::Parser.parse(path).worksheets

# shmeets = {}

# lables = {
#   event:      'Show', # [1]
#   attraction: 'Coming attraction 1', # [-3]
#   additional_text: 'Additional Text' # [-2]
# }

# lables.keys.each do |k|
#   shmeets[k] = sheets.find{ |s| s.sheet_name == lables[k] }
# end

# shmeets[:performances] = sheets.select{ |s| s.sheet_name.include?('Performance') }


# ==================================================================================



# {
#   'BibID'               => :metadata
#   'Show'                => :event
#   /Performance [1-9]/   => :performance
#   'Coming Attraction 1' => :coming_attraction
#   'Additional Text'     => :additional_text
# }

# BASE_SHEET_POSITIONS = {

#   performance: # 3
#   attraction: 3
#   additional_text: 4
# }

# sheets.each do |sheet|
#   config = YAML.load open("data/#{sheet}.yml").read

# # unless %w(BibID Show).include?(sheet.sheet_name)
#   unless config[:sheet_position]

#     config[:sheet_position] =
#   end

#   sheet_data = XlsxToPqcXml::XlsxData.new(xlsx_path: xlsx_path, config: config)
# end



# require 'rubyXL'

require 'xlsx_to_pqc_xml'
require 'yaml'
require 'json'

path = ARGV.first

# PE = PlaybillExtractor.new(path)
# PE.print_output

class PlaybillExtractor
  def initialize(path)
    @path = path
    sheetnames = RubyXL::Parser.parse(@path).worksheets.map(&:sheet_name)
    @performance_count = sheetnames.count{ |n| /Performance [1-9]/ === n }
  end

  def load_config(file_name)
    YAML.load(open("data/#{file_name}.yml").read)
  end

  def get_sheet_data(config)
    XlsxToPqcXml::XlsxData.new(xlsx_path: @path, config: config).data
  end

# ====================================================================

  def extract_hash(hash, syms = [])
    hash.select{ |k| syms.include?(k) }
  end

  def extract_metadata
    config = load_config('metadata')
    data   = get_sheet_data(config)
    data.first
  end

  def extract_event
    config = load_config('event')
    data   = get_sheet_data(config)

    details = data.first
    tickets = data[1..-1]

    formatted_details = {
      date: details.select{ |k| %i(date_standard date_as_written).include?(k)},
      venue: details.select{ |k| %i(@id venue_name venue_location).include?(k)},
      occasion: details.select{ |k| k == :occasion_type },
      organization: details[:organization],
      manager: details.select{ |k| k == :manager_name }
    }

    ticketing = data.map{ |record| extract_hash(record, %i(price price_as_written ticket_location)) }
    formatted_details.merge(ticketing: ticketing, performances: assemble_performances)
  end

  def extract_performance(performance_number)
    config = load_config('performance')
    config[:sheet_position] = 1 + performance_number
    data   = get_sheet_data(config)

    contributors = data[1..-1]
    data.first.merge({contributor: contributors})
  end

  def extract_coming_attraction
    config = load_config('attraction')
    config[:sheet_position] = 2 + @performance_count
    data   = get_sheet_data(config)
    {title: data.first[:title], date: data.first.tap{ |h| h.delete(:title) }}
  end

  def extract_additional_text
    config = load_config('additional_text')
    config[:sheet_position] = 3 + @performance_count
    data   = get_sheet_data(config)
  end

# ====================================================================

  def assemble_performances
    (1...@performance_count).map do |num|
      extract_performance(num)
    end
  end

  def combine_information
    {
      metadata: extract_metadata,
      playbill: {
        event: extract_event,
        coming_attraction: extract_coming_attraction,
        additional_text: extract_additional_text
      }
    }
  end

  def clean_hash(h)
    h.keys.each do |k|
      return h.delete(k) if val.empty?
      case val.class
      when Hash
        clean_hash(val)
      when Array
        val.each{ |e| clean_hash(e) if e.is_a?(Hash) }
      end
    end
  end

  def print_output
    clean_hash(combine_information).to_json
  end
end











