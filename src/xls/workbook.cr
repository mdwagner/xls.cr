# A `Xls::Workbook` contains metadata about the `Xls::Spreadsheet`.
class Xls::Workbook
  protected def initialize(@workbook : LibXls::XlsWorkBook*)
  end

  # Returns sheets for the workbook
  def sheets : Sheets
    Sheets.new(@workbook)
  end
end
