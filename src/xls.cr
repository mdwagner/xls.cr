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
    end

    def parse! : Nil
      status = LibXls.parse_worksheet(@worksheet)
      unless status.libxls_ok?
        message = String.new(LibXls.error(status))
        raise ParserError.new(message)
      end
    end

    def close! : Nil
      LibXls.close_worksheet(@worksheet)
    end

    def row_count
      @worksheet.value.rows.lastrow.to_i
    end

    def col_count
      @worksheet.value.rows.lastcol.to_i
    end

    def each_row(**kwargs, &)
      headers = get_headers
      row_count.times do |row|
        next if row == 0

        hash = {} of String => String

        if kwargs.empty?
          headers.each do |key, col|
            cell = LibXls.cell(@worksheet, row, col)
            value = cell_value(cell.value).to_s
            hash[key] = value

            x = {
              "id" => match_id(cell.value.id),
              "str" => cell.value.str ? String.new(cell.value.str) : cell.value.str
            }
            pp! x
          end
        else
          kwargs.each do |hash_key, matching_header|
            col = headers[matching_header]
            cell = LibXls.cell(@worksheet, row, col)
            value = cell_value(cell.value).to_s
            hash["#{hash_key}"] = value
          end
        end

        yield hash
      end
    end

    private def get_headers
      headers = {} of String => Int32
      col_count.times do |col|
        cell = LibXls.cell(@worksheet, 0, col)
        value = cell_value(cell.value).to_s
        headers[value] = col
      end
      headers
    end

    private def cell_value(cell)
      # pp! match_id(cell.id)
      if cell.id == LibXls::XLS_RECORD_BLANK
        nil
      elsif cell.id == LibXls::XLS_RECORD_NUMBER
        cell.d
      else
        if cell.str
          String.new(cell.str)
        else
          nil
        end
      end
    end

    def match_id(id)
      {% begin %}
      {%
        constants = LibXls.constants.select do |constant|
          constant.starts_with?("XLS_RECORD")
        end
      %}
        case id
        {% for constant in constants %}
        when LibXls::{{constant.id}}
          {{constant.id.stringify}}
        {% end %}
        else
          id
        end
      {% end %}
    end
  end

  private class WorksheetIterator
    include Iterator(Worksheet)
    alias SheetData = ::Slice(LibXls::StSheetData)

    def initialize(@sheets : SheetData, @workbook : LibXls::XlsWorkBook*)
      @index = 0
    end

    def next
      if @index < @sheets.size
        index = @index
        @index += 1
        raw_worksheet = LibXls.get_worksheet(@workbook, index)
        Worksheet.new(raw_worksheet)
      else
        stop
      end
    end
  end

  class Sheets
    include Enumerable(Worksheet)
    include Iterable(Worksheet)

    @sheets : Slice(LibXls::StSheetData)

    protected def initialize(@workbook : LibXls::XlsWorkBook*)
      raw_sheet = @workbook.value.sheets
      @sheets = raw_sheet.sheet.to_slice(raw_sheet.count)
    end

    def names : Array(String)
      @sheets.each.map { |sheet| String.new(sheet.name) }.to_a
    end

    def count
      @sheets.size
    end

    def each(& : Worksheet ->) : Nil
      @sheets.each_with_index do |_, index|
        raw_worksheet = LibXls.get_worksheet(@workbook, index)
        yield Worksheet.new(raw_worksheet)
      end
    end

    def each!(& : Worksheet ->) : Nil
      each do |ws|
        begin
          ws.parse!
          yield ws
        ensure
          ws.close!
        end
      end
    end

    def each
      WorksheetIterator.new(@sheets, @workbook)
    end
  end
end
