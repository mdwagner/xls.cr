class Xls::Worksheet
  # class Cell
  # protected def initialize(@cell : LibXls::StCellData)
  # end
  # end

  class Row
    protected def initialize(@row : LibXls::StRowData)
    end

    # Returns the index of this `Xls::Worksheet::Row` in the `Xls::Worksheet`
    def index
      @row.index
    end

    def cells : Array(Cell)
      @cells ||= begin
        raw_cells = @row.cells
        raw_cells.cell.to_slice(raw_cells.count).each.map do |cell|
          # Cell.new(cell)
        end.to_a
      end
    end

    # Returns the first `Xls::Worksheet::Cell`'s index
    #
    # NOTE: Untested
    def fcell
      @row.fcell
    end

    # Returns the last `Xls::Worksheet::Cell`'s index
    #
    # NOTE: Untested
    def lcell
      @row.lcell
    end

    # Returns the height of this `Xls::Worksheet::Row`
    #
    # NOTE: Untested
    def height
      @row.height
    end

    # :nodoc:
    def raw_flags
      @row.flags
    end

    # :ditto:
    def raw_xf
      @row.xf
    end

    # :ditto:
    def raw_xfflags
      @row.xfflags
    end
  end

  class ColumnInfo
    protected def initialize(@colinfo : LibXls::StColInfoData)
    end

    # Returns the index to the first column in the range
    def first
      @colinfo.first
    end

    # Returns the index to the last column in the range
    def last
      @colinfo.last
    end

    # Returns the width of the columns in 1/256 of the width of the zero character, using default font (the first FONT record in the file)
    def width
      @colinfo.width
    end

    # Returns the index to the XF record for default column formatting
    def xf
      @colinfo.xf
    end

    record ColumnInfoFlags,
      columns_hidden : Bool,
      columns_outline_level : UInt16,
      columns_collapsed : Bool

    # Returns any option flags
    #
    # See table below.
    #
    # ```markdown
    # | Bits | Mask | Contents                                      |
    # |------|------|-----------------------------------------------|
    # | 0    | 0001 | 1 = Columns are hidden                        |
    # | 10-8 | 0700 | Outline level of the columns (0 = no outline) |
    # | 12   | 1000 | 1 = Columns are collapsed                     |
    # ```
    def flags : ColumnInfoFlags
      ColumnInfoFlags.new(
        columns_hidden: @colinfo.flags.bit(0) == 1 ? true : false,
        columns_outline_level: @colinfo.flags.bits(8..10),
        columns_collapsed: @colinfo.flags.bit(12) == 1 ? true : false
      )
    end
  end

  protected def initialize(
    raw_worksheet @worksheet : LibXls::XlsWorkSheet*,
    @sheet_name : String,
    raw_visibility @sheet_visibility : UInt8,
    raw_type @sheet_type : UInt8,
    raw_filepos @sheet_filepos : UInt32
  )
  end

  # Returns the name of the worksheet
  def name
    @sheet_name
  end

  def columns : Array(ColumnInfo)
    @columns ||= begin
      raw_colinfo = @worksheet.value.colinfo
      raw_colinfo.col.to_slice(raw_colinfo.count).each.map do |info|
        ColumnInfo.new(info)
      end.to_a
    end
  end

  def rows : Array(Row)
    @rows ||= begin
      raw_rows = @worksheet.value.rows
      raw_rows.row.to_slice(raw_rows.lastrow).each.map do |row|
        Row.new(row)
      end.to_a
    end
  end

  enum SheetState : UInt8
    Visibile
    Hidden
    VeryHidden
  end

  def sheet_visibility : SheetState
    SheetState.from_value(@sheet_visibility)
  end

  enum SheetType : UInt8
    Worksheet
    Chart             = 2
    VisualBasicModule = 6
  end

  def sheet_type : SheetType
    SheetType.from_value(@sheet_type)
  end

  # Returns the absolute stream position of the BOF record of the sheet represented by this record
  def sheet_filepos
    @sheet_filepos
  end

  # :ditto:
  def worksheet_filepos
    @worksheet.value.filepos
  end

  # Returns the default column width for columns that do not have a specific width already set
  #
  # Column width in characters, using the width of the zero character from default font (first FONT record in the file).
  def defcolwidth
    @worksheet.value.defcolwidth
  end

  #
  #
  #

  # def row_count
  # @worksheet.value.rows.lastrow.to_i
  # end

  # def col_count
  # @worksheet.value.rows.lastcol.to_i
  # end

  # def each_row(**kwargs, &)
  # headers = get_headers
  # row_count.times do |row|
  # next if row == 0

  # hash = {} of String => String

  # if kwargs.empty?
  # headers.each do |key, col|
  # cell = LibXls.cell(@worksheet, row, col)
  # value = cell_value(cell.value).to_s
  # hash[key] = value
  # end
  # else
  # kwargs.each do |hash_key, matching_header|
  # col = headers[matching_header]
  # cell = LibXls.cell(@worksheet, row, col)
  # value = cell_value(cell.value).to_s
  # hash["#{hash_key}"] = value
  # end
  # end

  # yield hash
  # end
  # end

  # private def get_headers
  # headers = {} of String => Int32
  # col_count.times do |col|
  # cell = LibXls.cell(@worksheet, 0, col)
  # value = cell_value(cell.value).to_s
  # headers[value] = col
  # end
  # headers
  # end

  # private def ptr_to_s(ptr) : String
  # if ptr
  # String.new(ptr)
  # else
  # ""
  # end
  # end

  # private def cell_value(cell) : Cell::Any
  # if id = XlsRecord.from_value?(cell.id)
  # case id
  # when .record_boolerr?
  # cell_boolerr(cell)
  # when .record_number?, .record_rk?
  # Cell::Any.new(cell.d.to_f64)
  # when .record_labelsst?, .record_label?, .record_rstring?
  # Cell::Any.new(ptr_to_s(cell.str))
  # else
  # Cell::Any.new(nil)
  # end
  # else
  # Cell::Any.new(nil)
  # end
  # end

  # private def cell_boolerr(cell, str_fallback = false) : Cell::Any
  # str = ptr_to_s(cell.str)
  # case str
  # when "bool"
  # Cell::Any.new(cell.d.to_f64 > 0)
  # when "error"
  # Cell::Any.new(Cell::Error.new)
  # else
  # if str_fallback
  # Cell::Any.new(str)
  # else
  # Cell::Any.new(nil)
  # end
  # end
  # end
end
