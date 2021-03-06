require_relative '../lib/playbills/playbill_extractor'

folder_path = ARGV.first
folder_path += ?/ unless /[\/\\]/ === folder_path[-1]

abort 'No folder given'                   unless folder_path
abort "No folder found at #{folder_path}" unless Dir.exist?(folder_path)

xlsx_files = Dir.entries(folder_path).select{ |e| e =~ /^[^\.~].*\.xlsx$/ }
total = xlsx_files.size
errors  = []
records = []

xlsx_files.each_with_index do |file, i|
  print "extracting from #{file} (#{i + 1}/#{total})..."
  print ' ' * [(40 - file.length), 0].max

  file_path = folder_path + file
  result    = PlaybillExtractor.new(file_path).get_result

  if result.is_a?(ExtractorResult)
    records << result.to_h
    puts 'extracted data'
  else
  	errors << [file, result]
    puts "FOUND #{result.length} ERRORS"
  end
end

outfile = File.join(folder_path, 'extractor_result.json')

File.open(outfile, 'w+') do |f|
  f.puts JSON.pretty_generate({playbills_dataset: records})
end

if errors.any?
	puts ''
	puts '======================= ERRORS ======================='
	errors.each do |group|
		puts $/ + group.first
		group.last.each{ |e| puts e }
	end
end


