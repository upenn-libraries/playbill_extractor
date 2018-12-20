require 'xlsx_to_pqc_xml'
require 'yaml'
require 'json'

class PlaybillExtractor
  def initialize(path)
    @errors = []
    @path = path
    sheetnames = RubyXL::Parser.parse(@path).worksheets.map(&:sheet_name)
    @performance_count = sheetnames.grep(/Performance [1-9]/).size
  end

  def read_data(config_filename, sheet_position = nil)
    config = load_config(config_filename)
    config[:sheet_position] = sheet_position if sheet_position
    get_sheet_data(config)
  end

  def load_config(file_name)
    YAML.load(open("data/#{file_name}.yml").read)
  end

  def get_sheet_data(config)
    xlsx_data = XlsxToPqcXml::XlsxData.new(xlsx_path: @path, config: config)
    return xlsx_data.data if xlsx_data.valid?

    xlsx_data.errors.each do |key,list|

      sheet_name = "Performance #{config[:sheet_position] - 1}" if config[:sheet_name] == 'Performance'
      sheet_name ||= config[:sheet_name]

      # list.each do |struct|
      #   msg = "ERROR -- [#{sheet_name}]: #{key} "
      #   msg += "location #{struct.address}; " if struct.address
      #   msg += struct.text
      #   $stderr.puts msg
      # end

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

  def add_error(sheet_name, msg)
    error_string =  "ERROR -- [#{sheet_name}]#{msg}"
    @errors << error_string
  end

# ====================================================================

  def extract_subhash(hash, syms = [])
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
      venue: extract_subhash(details, %i(@id venue_name venue_identified_name venue_location)),
      occasion: details.select{ |k| k == :occasion_type },
      organization: details[:organization]
    }

    add_error('Show', ': required value missing (Venue Name)' )   unless formatted_details[:venue][:venue_name]
    add_error('Show', ': required value missing (Venue Address or Location)') unless formatted_details[:venue][:venue_location]

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
    data = read_data('performance', 1 + performance_number)
    return {} unless data
    details = extract_subhash(data.first, %i(@id title performance_description attractions performance_other))

    # add_error("Performance #{performance_number}", ': required value missing (Title)') unless details[:title]

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

  def clean_hash(h)
    h.keys.each do |k|
      val = h[k]
      next (h.delete(k)) if  val.nil? || val.empty?

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
    data = combine_data
    abort(@errors.join(?\n)) unless @errors.empty?
    clean_hash(data)
  end

  def write_json_to_file
    File.open "#{@path[0..-6]}.json", 'w+' do |f|
      f.puts JSON.pretty_generate(output_data)
    end
  end
end

# File.open 'file_name.json', 'w+' do |f|
#   f.puts the_data.to_json
# end

# folder_path = path
# xlsx_files  = Dir.entries(folder_path).select{ |e| e =~ /^[^\.~].*\.xlsx/ }

# xlsx_files.each do |file|
#   puts "working on #{file}"
#   file_path = folder_path + file
#   pe = PlaybillExtractor.new(file_path)
#   pe.write_json_to_file
# end








