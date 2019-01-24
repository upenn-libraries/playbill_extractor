require 'xlsx_to_pqc_xml'
require 'yaml'
require 'json'

class FurnessExtractor
  def initialize(xlsx_path)
    @xlsx_path = xlsx_path
  end

  def get_sheet_data
    xlsx_data = XlsxToPqcXml::XlsxData.new(xlsx_path: @xlsx_path, config: YAML.load(open(CONFIG_PATH)))
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

  def get_result
    data = combine_data
    @errors.empty? ? ExtractorResult.new(clean_hash(data)) : @errors
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


{
  image: {
    digital_facsimile: {
      media_URL: media_URL,
      ssid: ssid,
      filename: filename,
      id_number: id_number
    },
    collection_object: {
      creator: {'@id_creator_URI': creator_URI},
      title: {'@id_title_URI': title_URI},
      description: {'@id_description_URI': description_URI},
      actor_character: actor_character,
      '@id_actor_character_URI': actor_character_URI,
      related_agent: {'@id_related_agent_URI': related_agent_URI},
      play: play,
      date: date,
      date_range: date_range,
      earliest_date: earliest_date,
      latest_date: latest_date,
      style_period: style_period,
      materials_techniques: materials_techniques,
      measurements: measurements,
      artstor_clasification: artstor_clasification,
      work_type: work_type,
      repository: repository,
      creation_discovery_site: creation_discovery_site,
      country: country,
      subject: subject,
      relationships: relationships,
      source: source,
      legacy_bibliography_number: legacy_bibliography_number,
      donor_statement: donor_statement,
      image_view_description: image_view_description
    }
  }
}