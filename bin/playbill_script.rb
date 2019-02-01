require 'pry' #########
require 'json' #
require_relative '../lib/playbills/playbill_extractor'

folder_path = ARGV.first

abort 'No folder given'                   unless folder_path
abort "No folder found at #{folder_path}" unless Dir.exist?(folder_path)

xlsx_files = Dir.entries(folder_path).select{ |e| e =~ /^[^\.~].*\.xlsx$/ }
total = xlsx_files.size
errors = []

stuffs = [] ######

xlsx_files.each_with_index do |file, i|
  print "extracting from #{file} (#{i + 1}/#{total})..."
  print ' ' * [(40 - file.length), 0].max

  file_path = folder_path + file;  # next(print ?\n) if File.exist?("#{file_path[0..-6]}.json") ########
  result    = PlaybillExtractor.new(file_path).get_result

  if result.is_a?(ExtractorResult)
  	# File.open("#{file_path[0..-6]}.json", 'w+'){ |f| f.puts result.to_json }
  	# puts "	saved JSON"
    stuffs << result.to_h
    puts 'extracted data'
  else
  	errors << [file, result]
  	puts " FOUND #{result.length} ERRORS"
  end
end

File.open("#{folder_path}/result.json", 'w+') do |f|
  f.puts JSON.pretty_generate({stuffs: stuffs}) ######
end

if errors.any?
	puts ''
	puts '======================= ERRORS ======================='
	errors.each do |group|
		puts ?\n + group.first
		group.last.each{ |e| puts e }
	end
end


