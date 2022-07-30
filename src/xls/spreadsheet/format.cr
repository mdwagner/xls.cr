class Xls::Spreadsheet
  # See http://sc.openoffice.org/excelfileformat.pdf, page 174
  record Format,
    index : UInt16,
    value : String
end
