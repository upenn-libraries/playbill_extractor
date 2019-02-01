require_relative '../extractor'

class PlaybillExtractor < XlsxDataExtractor
  def initialize(xlsx_path)
    # @errors = []
    # @path = xlsx_path
    super(xlsx_path)
    sheetnames = RubyXL::Parser.parse(@path).worksheets.map(&:sheet_name)
    @performance_count = sheetnames.grep(/Performance [1-9]/).size
  end

  def read_data(config_filename, sheet_position = nil)
    config = YAML.load(open("../lib/playbills/data/#{config_filename}.yml").read)
    config[:sheet_position] = sheet_position if sheet_position
    get_sheet_data(config)
  end

  def get_sheet_data(config)
    xlsx_data = XlsxToPqcXml::XlsxData.new(xlsx_path: @path, config: config)
    return xlsx_data.data if xlsx_data.valid?

    xlsx_data.errors.each do |key, list|
      sheet_name = "Performance #{config[:sheet_position] - 1}" if config[:sheet_name] == 'Performance'
      sheet_name ||= config[:sheet_name]
      list.each do |struct|
        msg =
          case key
          when :required_header_missing then ": required header missing #{struct.text[/\(.*\)/]}"
          when :required_value_missing then "(#{struct.address}): required value missing #{struct.text[/\(.*\)/]}"
          else "(#{struct.address}) #{key} #{struct.text}"
          end
        add_error(sheet_name, msg)
      end
    end
    nil
  end

# ====================================================================

  def extract_subhash(hash, syms = [])
    unless hash
    #  puts "\ngot empty sheet"
      return {}
    end
    sh1 = hash.select{ |k| syms.include?(k) }
    sh2 = sh1.inject({}) do |result, pair|
      pair[0] = :@id if pair.first.to_s =~ /^@id/
      result.update([pair].to_h)
    end
    sh2.delete(:@id) unless sh2[:@id] =~ /^https?:\/\//i
    sh2.sort_by{ |k,v| k[0] == ?@ ? 0 : 1 }.to_h
  end

  def extract_metadata
    data = read_data('metadata')
    return {} unless data
    data.first
  end

  def extract_event
    data = read_data('event')
    return {} unless data
    details = data.first

    formatted_details = {
      date: extract_subhash(details, %i(date_standard date_as_written)),
    # venue: extract_subhash(details, %i(@id venue_name venue_identified_name venue_location)),
      occasion: details.select{ |k| k == :occasion_type },
      organization: details[:organization]
    }

    # add_error('Show', ': required value missing (Venue Name)' )   unless formatted_details[:venue][:venue_name]
    # add_error('Show', ': required value missing (Venue Address or Location)') unless formatted_details[:venue][:venue_location]

    venue     = data.map{ |record| extract_subhash(record, %i(@id venue_name venue_identified_name venue_location)) }
    add_error('Show', ': required value missing (Venue Name)' )   unless venue.first[:venue_name]
    add_error('Show', ': required value missing (Venue Address or Location)') unless venue.first[:venue_location]
    manager   = data.map{ |record| record[:manager_name] }.compact
    ticketing = data.map{ |record| extract_subhash(record, %i(price price_as_written ticket_location)) }
    formatted_details.merge(venue: venue, manager: manager, ticketing: ticketing, performance: assemble_performances)
  end

  def assemble_performances
    (1..@performance_count).map do |num|
      extract_performance(num)
    end
  end

  def extract_performance(performance_number)
    data = read_data('performance', 1 + performance_number)
    return {} unless data

    details = extract_subhash(data.first, %i(@id title performance_description attractions performance_other))

    contributors = data.map{ |record| extract_subhash(record, %i(@id_contrib contributor_type contributor_name character headliner)) }
    details.merge({contributor: contributors})
  end

  def extract_coming_attraction
    data = read_data('attraction', 2 + @performance_count)
    return {} unless data

    data.map do |record|
      details = extract_subhash(record, %i(title headliner))
      date    = extract_subhash(record, %i(date_standard date_as_written))
      details.merge(date: date)
    end
  end

  def extract_additional_text
    data = read_data('additional_text', 3 + @performance_count)
    return {} unless data

    %i(advertisement printer_publisher additional_other).map do |sym|
      [sym, data.map{ |record| record[sym] }.compact]
    end.to_h
  end

# ====================================================================

  def combine_data
    {
      metadata: extract_metadata,
      playbill: {
        event: extract_event,
        coming_attraction: extract_coming_attraction,
        additional_text: extract_additional_text
      }
    }
  end

  def get_result
    data = combine_data
    @errors.empty? ? ExtractorResult.new(clean_hash(data)) : @errors
  end
end

# ExtractorResult = Struct.new(:result) do
#   def to_h
#     result
#   end

#   def to_json
#     JSON.pretty_generate(result)
#   end
# end











