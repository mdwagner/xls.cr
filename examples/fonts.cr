require "../src/xls"

# Example: list xls fonts per worksheet
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  s.worksheets.each do |ws|
    s.fonts.each do |f|
      puts "#{f.name}"
    end
  end
end
