require 'json'
require './extractor'

path = ARGV.first

abort 'No filepath given'        unless path
abort "No file found at #{path}" unless File.exist?(path)

puts JSON.pretty_generate(PlaybillExtractor.new(path).output_data)




# ........................
require './extractor'

folder_path = ARGV.first

abort 'No folder given'                   unless folder_path
abort "No folder found at #{folder_path}" unless Dir.exist?(folder_path)

xlsx_files  = Dir.entries(folder_path).select{ |e| e =~ /^[^\.~].*\.xlsx$/ }

xlsx_files.each do |file|
  puts "working on #{file}"
  file_path = folder_path + file
  p_e = PlaybillExtractor.new(file_path)


  File.open "#{file_path[0..-6]}.json", 'w+' do |f|
    f.puts JSON.pretty_generate(p_e.result.to_json)
  end

end


####

