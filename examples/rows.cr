require "../src/xls"

# Example: list xls rows per worksheet
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  s.worksheets.each do |ws|
    ws.rows.each do |row|
      puts row
    end
  end
end
