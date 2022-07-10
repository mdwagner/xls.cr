class Xls::Spreadsheet
  # Retrieve libxls version
  def self.xls_version
    String.new(LibXls.version)
  end

  # Enable debug mode for libxls
  def self.debugging(enable = true, value = 1) : Nil
    LibXls.xls(enable ? value : 0)
  end

  def self.open_file(path : Path, charset : String = "UTF-8")
    raise "File not found" unless File.file?(path)
    wb = LibXls.open_file(path.to_s, charset, out error)
    new(wb, error)
  end

  def self.open(file_or_content)
    begin
      instance = new file_or_content
      yield instance.workbook
    ensure
      instance.try &.close!
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

  # Closes libxls Workbook
  def close! : Nil
    LibXls.close_workbook(@workbook)
  end

  # Raw Pointer to libxls Workbook for Advanced Usage
  def workbook! : LibXls::XlsWorkBook*
    @workbook
  end

  def workbook : Workbook
    unless @workbook
      message = String.new(LibXls.error(@workbook_err))
      raise "Error reading file: #{message}"
    end
    Workbook.new(@workbook)
  end
end
