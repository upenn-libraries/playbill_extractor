require './extractor'

folder_path = ARGV.first

abort 'No folder given'                   unless folder_path
abort "No folder found at #{folder_path}" unless Dir.exist?(folder_path)

xlsx_files  = Dir.entries(folder_path).select{ |e| e =~ /^[^\.~].*\.xlsx$/ }
total = xlsx_files.size 

errors = []

xlsx_files.each_with_index do |file, i|
  print "extracting from #{file} (#{i + 1}/#{total})..."
  print ' ' * [(40 - file.length), 0].max
  
  file_path = folder_path + file
  result    = PlaybillExtractor.new(file_path).get_result

  if result.is_a?(ExtractorResult)
  	 File.open("#{file_path[0..-6]}.json", 'w+'){ |f| f.puts result.to_json }
  	 puts "	saved JSON"
  else
  	errors << [file, result]
  	puts " FOUND #{result.length} ERRORS"
  end
end

if errors.any?
	puts ''
	puts '======================= ERRORS ======================='
	errors.each do |group|
		puts ?\n + group.first
		group.last.each{ |e| puts e }
	end
end


