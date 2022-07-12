# A `Xls::Spreadsheet` represents the document and contains metadata.
class Xls::Spreadsheet
  # Returns `libxls` version
  def self.xls_version : String
    Xls::Utils.ptr_to_str(LibXls.version)
  end

  # Toggles debug mode for `libxls`
  #
  # *value* controls different types of debugging information.
  def self.debugging(enable = true, value : UInt32 = 1) : Nil
    LibXls.xls(enable ? value : 0)
  end

  # Creates a new `Xls::Spreadsheet` by opening a file
  #
  # *charset* can be an encoding other than UTF-8.
  #
  # Throws `Xls::FileNotFound` if the filepath cannot be found.
  def self.open_file(path : Path, charset : String = "UTF-8")
    raise FileNotFound.new unless File.file?(path)
    wb = LibXls.open_file(path.to_s, charset, out error)
    new(wb, error)
  end

  # Yields a new `Xls::Spreadsheet` by providing a filepath or string or IO
  #
  # Calls `#validate!` on the newly created `Xls::Spreadsheet`.
  #
  # Always invokes `#close` after yielding.
  def self.open(file_or_content, & : Spreadsheet ->)
    begin
      instance = new file_or_content
      instance.validate!
      yield instance
    ensure
      instance.try &.close
    end
  end

  # Creates a new `Xls::Spreadsheet` by providing a filepath
  #
  # Throws `Xls::FileNotFound` if the filepath cannot be found.
  def self.new(path : Path)
    raise FileNotFound.new unless File.file?(path)
    new File.read(path)
  end

  # Creates a new `Xls::Spreadsheet` by providing a string
  def self.new(content : String)
    new IO::Memory.new(content)
  end

  # Creates a new `Xls::Spreadsheet` by providing an IO
  def self.new(io : IO)
    wb = LibXls.open_buffer(io.to_s, io.size, io.encoding, out error)
    new(wb, error)
  end

  private def initialize(
    @workbook : LibXls::XlsWorkBook*,
    @workbook_err : LibXls::XlsError
  )
  end

  # Validates the spreadsheet
  #
  # Can only be called once.
  #
  # Throws `Xls::Error` if spreadsheet is invalid.
  # Throws `Xls::WorkbookParserError` if spreadsheet failed to parse.
  def validate! : Nil
    @validated ||= begin
      raise Error.new(@workbook_err) unless @workbook
      status = LibXls.parse_workbook(@workbook)
      raise WorkbookParserError.new(status) unless status.libxls_ok?
      true
    end
  end

  # Closes the `Xls::Spreadsheet` and any `Xls::Worksheet`'s
  #
  # Once a `Xls::Spreadsheet` is closed, it cannot be reopened in the same instance. You must create a new `Xls::Spreadsheet` instance to reopen it.
  def close : Nil
    @closed ||= begin
      @worksheets.each { |ws| LibXls.close_worksheet(ws) }
      LibXls.close_workbook(@workbook)
      true
    end
  end

  # Checks if the `Xls::Spreadsheet` is closed
  def closed? : Bool
    @closed ? true : false
  end

  # Returns worksheets for the spreadsheet
  #
  # Throws `Xls::WorksheetParserError` if a worksheet failed to parse.
  def worksheets : Array(Worksheet)
    @worksheets ||= begin
      raw_sheets = @workbook.value.sheets
      sheets = raw_sheets.sheet.to_slice(raw_sheets.count)
      sheets.each.map_with_index do |sheet, index|
        raw_worksheet = LibXls.get_worksheet(@workbook, index)
        status = LibXls.parse_worksheet(raw_worksheet)
        raise WorksheetParserError.new(status) unless status.libxls_ok?
        Worksheet.new(
          raw_worksheet: raw_worksheet,
          sheet_name: Xls::Utils.ptr_to_str(sheet.name),
          raw_visibility: sheet.visibility,
          raw_type: sheet.type,
          raw_filepos: sheet.filepos
        )
      end.to_a
    end
  end

  # Returns the encoding of the spreadsheet
  def charset : String
    @charset ||= Xls::Utils.ptr_to_str(@workbook.value.charset)
  end

  # Returns the *Summary* of the spreadsheet
  def summary : String
    @summary ||= Xls::Utils.ptr_to_str(@workbook.value.summary)
  end

  # Returns the *Document Summary* of the spreadsheet
  def doc_summary : String
    @doc_summary ||= Xls::Utils.ptr_to_str(@workbook.value.docSummary)
  end

  # Returns the index to the active (displayed) worksheet
  def active_worksheet_index : UInt16
    @workbook.value.activeSheetIdx
  end

  # :nodoc:
  def raw_workbook
    @workbook
  end

  # :ditto:
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
