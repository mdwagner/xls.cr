class Xls::Spreadsheet
  # Retrieve libxls version
  def self.xls_version
    String.new(LibXls.version)
  end

  # Enable debug mode for libxls
  def self.debugging(enable : Bool = true) : Nil
    LibXls.xls(enable ? 1 : 0)
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
