require 'xlsx_to_pqc_xml'
require 'yaml'

class PlaybillExtractor
  def initialize(path)
    @path = path
    sheetnames = RubyXL::Parser.parse(@path).worksheets.map(&:sheet_name)
    @performance_count = sheetnames.grep(/Performance [1-9]/).size
  end

  def load_config(file_name)
    YAML.load(open("data/#{file_name}.yml").read)
  end

  def get_sheet_data(config)
    xlsx_data = XlsxToPqcXml::XlsxData.new(xlsx_path: @path, config: config)
    return xlsx_data.data if xlsx_data.valid?

    xlsx_data.errors.each do |key,list|
      list.each do |struct|
        msg = "ERROR -- #{key} "
        msg += "location #{struct.address}; " if struct.address
        msg += struct.text
        $stderr.puts msg
      end
    end
    abort 'Invalid spreadsheet; see errors above'
  end

# ====================================================================

  def extract_subhash(hash, syms = [])
    sh1 = hash.select{ |k| syms.include?(k) }
    sh2 = sh1.inject({}) do |result, pair|
      pair[0] = :@id if pair.first.to_s =~ /^\@id/
      result.update([pair].to_h)
    end
    sh2.delete(:@id) unless sh2[:@id] =~ /^http:\/\//
    sh2.sort_by{ |k,v| k[0] == ?@ ? 0 : 1 }.to_h
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

    formatted_details = {
      date: extract_subhash(details, %i(date_standard date_as_written)),
      venue: extract_subhash(details, %i(@id venue_name venue_location)),
      occasion: details.select{ |k| k == :occasion_type },
      organization: details[:organization]
    }
    manager   = data.map{ |record| record[:manager_name] }.compact
    ticketing = data.map{ |record| extract_subhash(record, %i(price price_as_written ticket_location)) }
    formatted_details.merge(manager: manager, ticketing: ticketing, performance: assemble_performances)
  end

  def assemble_performances
    (1..@performance_count).map do |num|
      extract_performance(num)
    end
  end

  def extract_performance(performance_number)
    config = load_config('performance')
    config[:sheet_position] = 1 + performance_number
    data   = get_sheet_data(config)

    details = extract_subhash(data.first, %i(@id title performance_description attractions performance_other))
    contributors = data.map{ |record| extract_subhash(record, %i(@id_contrib contributor_type contributor_name character headliner)) }

    details.merge({contributor: contributors})
  end

  def extract_coming_attraction
    config = load_config('attraction')
    config[:sheet_position] = 2 + @performance_count
    data   = get_sheet_data(config)

    data.map do |record|
      details = extract_subhash(record, %i(title headliner))
      date    = extract_subhash(record, %i(date_standard date_as_written))
      details.merge(date: date)
    end
  end

  def extract_additional_text
    config = load_config('additional_text')
    config[:sheet_position] = 3 + @performance_count
    data   = get_sheet_data(config)

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

  def clean_hash(h)
    h.keys.each do |k|
      val = h[k]
      next (h.delete(k)) if val.empty?

      case val
      when Hash
        clean_hash(val)
      when Array
        val.each{ |e| clean_hash(e) if e.is_a?(Hash) }
      when String
        val.strip.empty? && h.delete(k)
      end
    end
    h
  end

  def output_data
    clean_hash(combine_data)
  end
end


