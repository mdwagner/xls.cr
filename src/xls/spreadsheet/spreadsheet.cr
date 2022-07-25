require "./font"
require "./format"
require "./xf"
require "./worksheet"

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

  @validated : Bool?
  @closed : Bool?
  @worksheets : Array(Worksheet)?
  @fonts : Array(Font)?
  @formats : Array(Format)?
  @xfs : Array(Xf)?

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
  #
  # Throws `Xls::WorkbookParserException` if spreadsheet failed to parse.
  def validate! : Nil
    @validated ||= begin
      raise Error.new(@workbook_err) unless @workbook
      status = LibXls.parse_workbook(@workbook)
      raise WorkbookParserException.new(status) unless status.libxls_ok?
      true
    end
  end

  # Closes the `Xls::Spreadsheet` and any `Xls::Worksheet`'s
  #
  # Once a `Xls::Spreadsheet` is closed, it cannot be reopened in the same instance.
  # You must create a new `Xls::Spreadsheet` instance to reopen it.
  def close : Nil
    @closed ||= begin
      worksheets.each { |ws| LibXls.close_worksheet(ws) }
      LibXls.close_workbook(@workbook)
      true
    end
  end

  # Checks if the `Xls::Spreadsheet` is closed
  def closed? : Bool
    @closed == true
  end

  # Returns worksheets for the spreadsheet
  #
  # Throws `Xls::WorksheetParserException` if a worksheet failed to parse.
  def worksheets : Array(Worksheet)
    @worksheets ||= begin
      raw_sheets = @workbook.value.sheets
      sheets = raw_sheets.sheet.to_slice(raw_sheets.count)
      sheets.each.map_with_index do |sheet, index|
        raw_worksheet = LibXls.get_worksheet(@workbook, index)
        status = LibXls.parse_worksheet(raw_worksheet)
        raise WorksheetParserException.new(status) unless status.libxls_ok?
        Worksheet.new(
          worksheet: raw_worksheet,
          sheet: sheet
        )
      end.to_a
    end
  end

  # Returns the active (displayed) worksheet
  def active_worksheet : Worksheet?
    begin
      worksheets[@workbook.value.activeSheetIdx]?
    rescue WorksheetParserException
      nil
    end
  end

  # Returns the encoding of the spreadsheet
  def charset : String
    Xls::Utils.ptr_to_str(@workbook.value.charset)
  end

  # Returns the *Summary* of the spreadsheet
  def summary : String
    Xls::Utils.ptr_to_str(@workbook.value.summary)
  end

  # Returns the *Document Summary* of the spreadsheet
  def doc_summary : String
    Xls::Utils.ptr_to_str(@workbook.value.docSummary)
  end

  def to_unsafe
    @workbook
  end

  # Returns the text encoding used to write byte strings, stored as MS Windows code page identifier
  #
  # For more information see https://en.wikipedia.org/wiki/Character_encoding.
  def codepage : UInt16
    @workbook.value.codepage
  end

  # Returns information about a used font, including character formatting
  #
  # All FONT records occur together in a sequential list.
  # Other records referencing a FONT record contain an index into this list.
  #
  # The font with index 4 is omitted in all BIFF versions.
  # This means the first four fonts have zero-based indexes, and the fifth font and all following fonts are refereced with one-based indexes.
  def fonts : Array(Font)
    @fonts ||= begin
      raw_fonts = @workbook.value.fonts
      raw_fonts_slice = raw_fonts.font.to_slice(raw_fonts.count)
      raw_fonts_slice.each.map_with_index do |font, index|
        real_index : UInt32 = index.to_u32
        if index >= 4
          real_index += 1
        end
        Font.new(font: font, real_index: real_index)
      end.to_a
    end
  end

  # Returns information about a number format
  #
  # All FORMAT records occur together in a sequential list.
  def formats : Array(Format)
    @formats ||= begin
      raw_formats = @workbook.value.formats
      raw_formats.format.to_slice(raw_formats.count).each.map do |format|
        Format.new(
          index: format.index,
          value: Xls::Utils.ptr_to_str(format.value)
        )
      end.to_a
    end
  end

  # Returns formatting information for cells, rows, columns or styles
  def xfs : Array(Xf)
    @xfs ||= begin
      raw_xfs = @workbook.value.xfs
      raw_xfs.xf.to_slice(raw_xfs.count).each.map do |xf|
        Xf.new(xf)
      end.to_a
    end
  end

  def get_xf?(value : Worksheet::ColumnInfo) : Xf?
    xfs[value.xf]?
  end

  def get_xf?(value : Worksheet::Cell) : Xf?
    xfs[value.xf]?
  end

  def get_xf?(value : Worksheet::Row) : Xf?
    xfs[value.xf]?
  end

  def get_font?(value : Xf) : Font?
    fonts.find { |font| font.real_index == value.font }
  end

  def get_format?(value : Xf) : Format?
    formats[value.format]?
  end
end
