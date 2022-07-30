require "../src/xls"
require "html"

# Example: convert xls table into html
#
# Usage: [program] <xls file path>
Xls::Spreadsheet.open(Path.new(ARGV[0])) do |s|
  io = IO::Memory.new

  s.worksheets.each do |ws|
    io.clear
    # set styles from workbook
    io << %(<style type="text/css">\n#{s.css}</style>\n)
    ws.rows.each do |row|
      io << %(<table border="0" cellspacing="0" cellpadding="2">)
      io << %(<tr>)
      row.cells.each do |cell|
        next if cell.is_hidden?

        io << %(<td)
        io << %( colspan="#{cell.col_span}") if cell.col_span != 0
        io << %( rowspan="#{cell.row_span}") if cell.row_span != 0

        io << %( class="xf#{cell.xf_index}">)

        HTML.escape(cell.value.raw.to_s, io)

        io << %(</td>)
      end
      io << %(</table>)
    end
    puts io.to_s
  end
end
