class Xls::Workbook
  protected def initialize(@workbook : LibXls::XlsWorkBook*)
  end

  def sheets : Sheets
    Sheets.new(@workbook)
  end

  def charset : String
    if ptr = @workbook.value.charset
      String.new(ptr)
    else
      ""
    end
  end

  def summary : String
    if ptr = @workbook.value.summary
      String.new(ptr)
    else
      ""
    end
  end

  def doc_summary : String
    if ptr = @workbook.value.docSummary
      String.new(ptr)
    else
      ""
    end
  end
end
