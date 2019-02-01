require_relative '../lib/furness/furness_extractor'

file_path = ARGV.first

abort 'No file given' unless file_path
abort "No file found at #{file_path}" unless File.exist?(file_path)

out_file  = File.join(File.dirname(file_path), "#{File.basename(file_path, '.xlsx')}.json")

File.open(out_file, 'w+'){ |f| f.puts FurnessExtractor.new(file_path).get_result.to_json }
