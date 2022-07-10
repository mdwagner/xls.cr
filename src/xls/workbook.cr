# A `Xls::Workbook` contains metadata about the `Xls::Spreadsheet`.
class Xls::Workbook
  protected def initialize(@workbook : LibXls::XlsWorkBook*)
  end

  # Returns sheets for the workbook
  def sheets : Sheets
    Sheets.new(@workbook)
  end

  # Returns the encoding of the workbook
  def charset : String
    Xls::Utils.ptr_to_str(@workbook.value.charset)
  end

  # Returns the Summary of the workbook
  def summary : String
    Xls::Utils.ptr_to_str(@workbook.value.summary)
  end

  # Returns the Document Summary of the workbook
  def doc_summary : String
    Xls::Utils.ptr_to_str(@workbook.value.docSummary)
  end

  # Returns the index to the active (displayed) worksheet
  def active_worksheet_index : UInt16
    @workbook.value.activeSheetIdx
  end

  # :nodoc:
  def raw_ole_stream
    @workbook.value.olestr
  end

  # :ditto:
  def raw_filepos
    @workbook.value.filepos
  end

  # :ditto:
  def raw_is5ver
    @workbook.value.is5ver
  end

  # :ditto:
  def raw_is1904
    @workbook.value.is1904
  end

  # :ditto:
  def raw_type
    @workbook.value.type
  end

  # :ditto:
  def raw_codepage
    @workbook.value.codepage
  end

  # :ditto:
  def raw_sst
    @workbook.value.sst
  end

  # :ditto:
  def raw_xfs
    @workbook.value.xfs
  end

  # :ditto:
  def raw_fonts
    @workbook.value.fonts
  end

  # :ditto:
  def raw_formats
    @workbook.value.formats
  end

  # :ditto:
  def raw_converter : Void*
    @workbook.value.converter
  end

  # :ditto:
  def raw_utf16_converter : Void*
    @workbook.value.utf16_converter
  end

  # :ditto:
  def raw_utf8_locale : Void*
    @workbook.value.utf8_locale
  end
end
