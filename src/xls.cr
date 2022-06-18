require "./version"
require "./lib_xls"

module Xls
  class Spreadsheet
    def self.xls_version
      String.new(LibXls.version)
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

    def close! : Nil
      LibXls.close_workbook(@workbook)
    end

    # Raw Pointer to Workbook for Advanced Usage
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

  class Workbook
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

  class Worksheet
    class ParserError < Exception
      def initialize(message = "Unknown")
        super(message)
      end
    end

    protected def initialize(@worksheet : LibXls::XlsWorkSheet*)
      status = LibXls.parse_worksheet(@worksheet)
      unless status.libxls_ok?
        message = String.new(LibXls.error(status))
        raise ParserError.new(message)
      end
    end

    def close! : Nil
      LibXls.close_worksheet(@worksheet)
    end
  end

  class Sheets
    include Enumerable(Worksheet)

    @sheets : Slice(LibXls::StSheetData)

    protected def initialize(@workbook : LibXls::XlsWorkBook*)
      raw_sheet = @workbook.value.sheets
      @sheets = raw_sheet.sheet.to_slice(raw_sheet.count)
    end

    def names : Array(String)
      @sheets.map { |sheet| String.new(sheet.name) }.to_a
    end

    def count
      @sheets.size
    end

    def each
      @sheets.each_with_index do |_, index|
        begin
          raw_worksheet = LibXls.get_worksheet(@workbook, index)
          yield Worksheet.new(raw_worksheet)
        rescue ex
          case ex
          when Worksheet::ParserError
            puts ex.message
            next
          else
            break
          end
        end
      end
    end
  end
end
