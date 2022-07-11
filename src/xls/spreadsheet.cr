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

  def self.open_file(path : Path, charset : String = "UTF-8")
    raise "File not found" unless File.file?(path)
    wb = LibXls.open_file(path.to_s, charset, out error)
    new(wb, error)
  end

  def self.open(file_or_content, & : Spreadsheet ->)
    begin
      instance = new file_or_content
      instance.validate!
      yield instance
    ensure
      instance.try &.close
    end
  end

  def self.new(path : Path)
    raise "File not found" unless File.file?(path)
    new File.read(path)
  end

  def self.new(content : String)
    new IO::Memory.new(content)
  end

  def self.new(io : IO)
    wb = LibXls.open_buffer(io.to_s, io.size, io.encoding, out error)
    new(wb, error)
  end

  ###############
  # begin
  #   Spreadsheet.open("path") do |s|
  #     # workbook_ptr
  #     # check for invalid workbook_ptr
  #     s.summary # metadata
  #     s.worksheets.each do |worksheet| # (valid) worksheets
  #       # parses each worksheet first
  #       worksheet.name
  #       worksheet.each_row { ... }
  #     end
  #   end
  # rescue Spreadsheet::Error # workbook_ptr is invalid
  #   exit 1
  # end
  ###############

  ###############
  # s = Spreadsheet.new("path")
  # s.validate! : Nil # throws
  # s.valid? : Bool
  # s.summary
  # s.worksheets.each { ... }
  # s.raw_worksheets # Worksheet.new(..., parse = false)
  # s.close!
  ###############

  ###############
  # def worksheets : Array(Worksheet)
  #   # TODO: consider memoization
  #   raw_sheets = @workbook.value.sheets
  #   sheets = raw_sheets.sheet.to_slice(raw_sheets.count)
  #   sheets.map_with_index do |sheet, index|
  #     sheet_name = Xls::Utils.ptr_to_str(sheet.name)
  #     raw_visibility = sheet.visibility
  #     raw_type = sheet.type
  #     raw_filepos = sheet.filepos
  #     raw_worksheet = LibXls.get_worksheet(@workbook, index)
  #     LibXls.parse_worksheet(raw_worksheet) # TODO: erorr handling for invalid worksheet
  #     Worksheet.new(raw_worksheet, sheet_name, raw_visibility, raw_type, raw_filepos)
  #   end.to_a
  # end
  ###############

  private def initialize(@workbook : LibXls::XlsWorkBook*, @workbook_err : LibXls::XlsError)
  end

  # Validates the Spreadsheet
  #
  # Can throw if invalid.
  def validate! : Nil
    unless @workbook
      message = Xls::Utils.ptr_to_str(LibXls.error(@workbook_err))
      raise message # TODO: make custom exception
    end
  end

  # Validates the Spreadsheet *safely*
  def valid? : Bool
    begin
      validate!
      true
    rescue
      false
    end
  end

  # Closes the Spreadsheet
  #
  # Once a Spreadsheet is closed, it cannot be reopened in the same instance. You must create a new Spreadsheet instance to reopen it.
  def close : Nil
    @closed ||= begin
      LibXls.close_workbook(@workbook)
      true
    end
  end

  # Checks if the Spreadsheet is closed
  def closed? : Bool
    @closed ? true : false
  end

  # Returns worksheets for the Spreadsheet
  #
  # TODO: not implemented yet
  def worksheets : Array(Worksheet)
    ###############
    # def worksheets : Array(Worksheet)
    #   # TODO: consider memoization
    #   raw_sheets = @workbook.value.sheets
    #   sheets = raw_sheets.sheet.to_slice(raw_sheets.count)
    #   sheets.map_with_index do |sheet, index|
    #     sheet_name = Xls::Utils.ptr_to_str(sheet.name)
    #     raw_visibility = sheet.visibility
    #     raw_type = sheet.type
    #     raw_filepos = sheet.filepos
    #     raw_worksheet = LibXls.get_worksheet(@workbook, index)
    #     LibXls.parse_worksheet(raw_worksheet) # TODO: erorr handling for invalid worksheet
    #     Worksheet.new(raw_worksheet, sheet_name, raw_visibility, raw_type, raw_filepos)
    #   end.to_a
    # end
    ###############

    [] of Worksheet
  end

  # Returns the encoding of the Spreadsheet
  def charset : String
    Xls::Utils.ptr_to_str(@workbook.value.charset)
  end

  # Returns the Summary of the Spreadsheet
  def summary : String
    Xls::Utils.ptr_to_str(@workbook.value.summary)
  end

  # Returns the Document Summary of the Spreadsheet
  def doc_summary : String
    Xls::Utils.ptr_to_str(@workbook.value.docSummary)
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
