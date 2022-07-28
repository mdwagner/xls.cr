require "../src/xls"

# Example: show information about xls
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  puts s
end
