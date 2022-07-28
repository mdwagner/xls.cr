require "../src/xls"

# Example: list xls formats per worksheet
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  s.formats.each do |f|
    puts f
  end
end
