require "../src/xls"

# Example: list xls columns per worksheet
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  s.worksheets.each do |ws|
    ws.columns_info.each do |info|
      puts info
    end
  end
end
