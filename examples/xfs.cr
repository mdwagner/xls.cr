require "../src/xls"

# Example: list xls extended formats (XF) per worksheet
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  s.worksheets.each do |ws|
    s.xfs.each_with_index do |xf, index|
      puts "Index: #{index}"
      puts
      puts xf
      puts
    end
  end
end
