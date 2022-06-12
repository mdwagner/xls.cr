require "./xls"

alias LibXls = Xls::LibXls

def main
  puts Xls::Spreadsheet.xls_version
  xls = Xls::Spreadsheet.new(Path.new(ARGV[0]))
  wb = xls.workbook_ptr

  # Make CSV
  wb.value.sheets.count.times do |sheet_index|
    work_sheet = LibXls.get_worksheet(wb, sheet_index)

    break if work_sheet.null?

    status = LibXls.parse_worksheet(work_sheet)

    unless status.libxls_ok?
      puts String.new(LibXls.error(status))
      puts
    end

    work_sheet.value.rows.lastrow.times do |row_index|
      row = LibXls.row(work_sheet, row_index)
      last_col = work_sheet.value.rows.lastcol

      row.value.cells.cell.to_slice(last_col).each_with_index do |cell, cell_index|
        if cell.str.null?
          print ""
        else
          print %("#{String.new(cell.str)}")
        end
        print ","
      end
      puts
    end
  end
end

main
