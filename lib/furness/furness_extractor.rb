require_relative '../extractor'

class FurnessExtractor < XlsxDataExtractor
  def initialize(xlsx_path)
    super(xlsx_path)
  end

  CONFIG_PATH = '../lib/furness/furness_config.yml'

  def get_sheet_data
    xlsx_data = XlsxToPqcXml::XlsxData.new(xlsx_path: @path, config: YAML.load(open(CONFIG_PATH)))
    return xlsx_data.data if xlsx_data.valid?

    puts xlsx_data.errors
  end

  def purge_non_URI(data)
    %i(creator_URI title_URI description_URI actor_character_URI related_agent_URI).each do |key|
      case data[key]
      when Array
        data[key] = data[key].map{ |s| s[0] == ?h ? s : nil }
      when String
        data[key] = nil unless data[key][0] == ?h
      when nil
        nil
      end
    end
    data
  end

  def grand_combine(records)
    {furness_image_dataset: records.map{ |record| clean_hash(combine_data(purge_non_URI(record))) }}
  end

  def combine_data(data)
    {
      image: {
        digital_facsimile: {
          media_URL: data[:media_URL],
          ssid: data[:ssid],
          filename: data[:filename],
          id_number: data[:id_number]
        },
        collection_object: {

          creator: pair_off(data[:creator], data[:creator_URI]),

          title: {
            '@id': data[:title_URI],
            title_text: data[:title]
          },

          description: pair_description(data[:description],data[:description_URI]),

          actor_character: {
            '@id': data[:actor_character_URI],
            name: data[:actor_character]
          },

          related_agent: pair_off(data[:related_agent], data[:related_agent_URI]),

          play: data[:play],
          date: data[:date],
          date_range: data[:date_range],
          earliest_date: data[:earliest_date],
          latest_date: data[:latest_date],
          style_period: data[:style_period],
          materials_techniques: data[:materials_techniques],
          measurements: data[:measurements],
          artstor_clasification: data[:artstor_clasification],
          work_type: data[:work_type],
          repository: data[:repository],
          creation_discovery_site: data[:creation_discovery_site],
          country: data[:country],
          subject: data[:subject],
          relationships: data[:relationships],
          source: data[:source],
          legacy_bibliography_number: data[:legacy_bibliography_number],
          donor_statement: data[:donor_statement],
          image_view_description: data[:image_view_description]
        }
      }
    }
  end

  def pair_off(names, ids)
    pair = [names, ids].map{ |e| e ? e : [] }

    (0...pair.map(&:length).max).map do |index|
      {'@id': pair.last[index], name:  pair.first[index]}
    end
  end

  def pair_description(descrip, ids)
    (0...(ids ? ids.length : 0)).map do |index|
      next({}) if (ids.length > 1) && ids[index].nil?
      {'@id': ids && ids[index], description_text: descrip}
    end
  end

  def get_result
    data = grand_combine(get_sheet_data)
    ExtractorResult.new(data)
  end
end


