module Xls
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
end
