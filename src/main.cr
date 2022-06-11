require "./xls"

alias LibXls = Xls::LibXls

def main
  puts String.new(LibXls.version)

  wb = LibXls.open_file(ARGV[0], "UTF-8", out error)
  unless wb.null?
    sheet = LibXls.get_worksheet(wb, 0)
    status = LibXls.parse_worksheet(sheet)
    row = LibXls.row(sheet, 0)
    cell = row.value.cells.cell[0]
    puts String.new(cell.str)
  end
ensure
  LibXls.close_workbook(wb) if wb
end

main
