# require 'json'
# require './extractor'

# path = ARGV.first

# abort 'No filepath given'        unless path
# abort "No file found at #{path}" unless File.exist?(path)

# puts JSON.pretty_generate(PlaybillExtractor.new(path).output_data)




# ........................
require './extractor'

folder_path = ARGV.first

abort 'No folder given'                   unless folder_path
abort "No folder found at #{folder_path}" unless Dir.exist?(folder_path)

xlsx_files  = Dir.entries(folder_path).select{ |e| e =~ /^[^\.~].*\.xlsx$/ }

total = xlsx_files.size 

xlsx_files.each_with_index do |file, i|
  print "extracting from #{file} (#{i + 1}/#{total})..."
  file_path = folder_path + file
  p_e = PlaybillExtractor.new(file_path)

  File.open "#{file_path[0..-6]}.json", 'w+' do |f|
    f.puts p_e.get_result.to_json
  end
  # percent = ((i + 1).to_f / total * 100).round
  puts "	saved JSON"# (#{percent}% done)"
end


####

