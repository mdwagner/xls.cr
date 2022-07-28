require "../src/xls"

# Example: convert xls into csv
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  s.worksheets.each do |ws|
    ws.rows.each do |row|
      row.cells.each_with_index do |cell, index|
        value = cell.value
        if str = value.as_s?
          if str.includes?(",")
            print "\"#{str}\""
          else
            print str
          end
        else
          print cell.value.raw
        end
        unless index == row.cells.size - 1
          print ","
        end
      end
      print "\n"
    end
  end
end
