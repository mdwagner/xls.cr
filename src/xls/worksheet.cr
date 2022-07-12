class Xls::Worksheet
  class Cell
    struct Error
      def to_s(io)
        io << "Cell::Error"
      end

      def inspect(io)
        io << "Cell::Error"
      end
    end

    struct Any
      alias Type = Nil | Bool | Float64 | String | Error

      getter raw : Type

      def initialize(@raw)
      end

      # Checks that the underlying value is `Nil`, and returns `nil`
      # Raises otherwise.
      def as_nil : Nil
        @raw.as(Nil)
      end

      # Checks that the underlying value is `Bool`, and returns its value
      # Raises otherwise.
      def as_bool : Bool
        @raw.as(Bool)
      end

      # Checks that the underlying value is `Bool`, and returns its value
      # Returns `nil` otherwise.
      def as_bool? : Bool?
        as_bool if @raw.is_a?(Bool)
      end

      # Checks that the underlying value is `Float64`, and returns its value
      # Raises otherwise.
      def as_f : Float64
        @raw.as(Float64)
      end

      # Checks that the underlying value is `Float64`, and returns its value
      # Returns `nil` otherwise.
      def as_f? : Float64?
        as_f if @raw.is_a?(Float64)
      end

      # Checks that the underlying value is `String`, and returns its value
      # Raises otherwise.
      def as_s : String
        @raw.as(String)
      end

      # Checks that the underlying value is `String`, and returns its value
      # Returns `nil` otherwise.
      def as_s? : String?
        as_s if @raw.is_a?(String)
      end

      # Checks that the underlying value is `Xls::Worksheet::Cell::Error`, and returns its value
      # Raises otherwise.
      def as_error : Error
        @raw.as(Error)
      end

      # Checks that the underlying value is `Xls::Worksheet::Cell::Error`, and returns its value
      # Returns `nil` otherwise.
      def as_error? : Error?
        as_error if @raw.is_a?(Error)
      end
    end

    protected def initialize(@cell : LibXls::StCellData)
    end

    def id : XlsRecord
      XlsRecord.from_value(@cell.id)
    end

    def row : UInt16
      @cell.row
    end

    def col : UInt16
      @cell.col
    end

    # See `Xls::Worksheet::ColumnInfo#xf`
    def xf : UInt16
      @cell.xf
    end

    # See `Xls::Worksheet#defcolwidth`
    def width : UInt16
      @cell.width
    end

    # Returns the span of columns
    #
    # NOTE: Untested
    def colspan : UInt16
      @cell.colspan
    end

    # Returns the span of rows
    #
    # NOTE: Untested
    def rowspan : UInt16
      @cell.rowspan
    end

    # Returns whether this cell is hidden
    def is_hidden? : Bool
      @cell.isHidden == 1
    end

    private def raw_string
      Xls::Utils.ptr_to_str(@cell.str)
    end

    private def raw_double
      @cell.d.to_f64
    end

    private def raw_long
      @cell.l.to_i32
    end

    # Returns the value of this cell as `Xls::Worksheet::Cell::Any`
    #
    # You must invoke `Xls::Worksheet::Cell::Any#raw` to get the raw value.
    def value : Any
      @value ||= begin
        case id
        when .record_boolerr?
          case raw_string
          when "bool"
            return Any.new(raw_double > 0)
          when "error"
            return Any.new(Error.new)
          end
        when .record_number?, .record_rk?
          return Any.new(raw_double)
        when .record_labelsst?, .record_label?, .record_rstring?
          return Any.new(raw_string)
        end

        Any.new(nil)
      end
    end
  end

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
          Cell.new(cell)
        end.to_a
      end
    end

    # Returns the index to the column of the first cell which is described by a cell record
    def fcell : UInt16
      @row.fcell
    end

    # Returns the index to the column of the last cell which is described by a cell record, increased by 1
    def lcell : UInt16
      @row.lcell
    end

    record HeightMetadata,
      height : UInt16,
      default_height : Bool do
      def is_custom_height? : Bool
        !default_height
      end

      def is_default_height? : Bool
        default_height
      end
    end

    # Returns the height of the row, in twips = 1/20 of a point
    def height : HeightMetadata
      HeightMetadata.new(
        height: @row.height.bits(0..14),
        default_height: @row.height.bit(15) == 1
      )
    end

    # :nodoc:
    def flags
      @row.flags
    end

    # :ditto:
    def xf
      @row.xf
    end

    # :ditto:
    def xfflags
      @row.xfflags
    end
  end

  class ColumnInfo
    protected def initialize(@colinfo : LibXls::StColInfoData)
    end

    # Returns the index to the first column in the range
    def first : UInt16
      @colinfo.first
    end

    # Returns the index to the last column in the range
    def last : UInt16
      @colinfo.last
    end

    # Returns the width of the columns in 1/256 of the width of the zero character, using default font (the first FONT record in the file)
    def width : UInt16
      @colinfo.width
    end

    # Returns the index to the XF record for default column formatting
    def xf : UInt16
      @colinfo.xf
    end

    record OptionFlags,
      columns_hidden : Bool,
      columns_outline_level : UInt16,
      columns_collapsed : Bool do
      def columns_hidden? : Bool
        columns_hidden
      end

      def columns_collapsed? : Bool
        columns_collapsed
      end

      def no_outline? : Bool
        columns_outline_level == 0
      end
    end

    # Returns option flags
    def flags : OptionFlags
      OptionFlags.new(
        columns_hidden: @colinfo.flags.bit(0) == 1,
        columns_outline_level: @colinfo.flags.bits(8..10),
        columns_collapsed: @colinfo.flags.bit(12) == 1
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
  def name : String
    @sheet_name
  end

  def columns_info : Array(ColumnInfo)
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

  # Returns the worksheet visibility
  def sheet_visibility : SheetState
    SheetState.from_value(@sheet_visibility)
  end

  enum SheetType : UInt8
    Worksheet
    Chart             = 2
    VisualBasicModule = 6
  end

  # Returns the worksheet type
  def sheet_type : SheetType
    SheetType.from_value(@sheet_type)
  end

  # Returns the absolute stream position of the BOF record of the sheet represented by this record
  def sheet_filepos : UInt32
    sheet_filepos = @sheet_filepos
    worksheet_filepos = @worksheet.value.filepos
    if sheet_filepos != worksheet_filepos
      Log.warn &.emit("Worksheet filepos mismatched",
        sheet_filepos: sheet_filepos,
        worksheet_filepos: worksheet_filepos,
        worksheet_name: name)
    end
    sheet_filepos
  end

  # Returns the default column width for columns that do not have a specific width already set
  #
  # Column width in characters, using the width of the zero character from default font (first FONT record in the file).
  def defcolwidth : UInt16
    @worksheet.value.defcolwidth
  end
end
