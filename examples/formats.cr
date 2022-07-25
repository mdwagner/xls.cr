require "../src/xls"

# Example: list xls formats per worksheet
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  s.worksheets.each do |ws|
    s.formats.each do |f|
      puts "#{f.index} => #{f.value}"
    end
  end
end
