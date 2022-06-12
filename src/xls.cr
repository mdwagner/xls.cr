require "./version"
require "./lib_xls"

module Xls
  class Spreadsheet
    alias XlsWorkBook = LibXls::XlsWorkBook
    alias XlsError = LibXls::XlsError

    def self.xls_version
      String.new(LibXls.version)
    end

    def self.new(path : Path, charset : String = "UTF-8")
      raise "File not found" unless File.file?(path)
      wb = LibXls.open_file(path.to_s, charset, out error)
      new(wb, error)
    end

    def self.new(content : String)
      new IO::Memory.new(content)
    end

    def self.new(io : IO)
      wb = LibXls.open_buffer(io.to_s, io.size, io.encoding, out error)
      new(wb, error)
    end

    private def initialize(@workbook_ptr : XlsWorkBook*, @workbook_error : XlsError)
      at_exit { LibXls.close_workbook(@workbook_ptr) }
    end

    def workbook_ptr : XlsWorkBook*
      if @workbook_ptr.null?
        raise "Error reading file: #{String.new(LibXls.error(@workbook_error))}"
      end
      @workbook_ptr
    end
  end
end
