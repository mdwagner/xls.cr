require "../src/xls"

# Example: list xls cells per worksheet
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  s.worksheets.each do |ws|
    ws.rows.each do |row|
      row.cells.each do |cell|
        puts cell
      end
    end
  end
end
