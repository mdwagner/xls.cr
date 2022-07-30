class Xls::Worksheet
  class Cell
    # See http://sc.openoffice.org/excelfileformat.pdf, page 21
    enum CellError : UInt8
      # Intersection of two cell ranges is empty
      Null
      # Division by zero
      DivZero = 0x07
      # Wrong type of operand
      Value = 0x0F
      # Illegal or deleted cell reference
      Ref = 0x17
      # Wrong function or range name
      Name = 0x1D
      # Value range overflow
      Num = 0x24
      # Argument or function not available
      NA = 0x2A
    end

    # See http://sc.openoffice.org/excelfileformat.pdf, page 91
    struct Any
      include InspectableMethods

      alias Type = Nil | Bool | Float64 | String | CellError

      protected def initialize(@raw : Type)
      end

      @[Inspectable]
      protected def type
        @raw.class
      end

      @[Inspectable]
      def raw
        @raw
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

      # Checks that the underlying value is `Xls::Worksheet::Cell::CellError`, and returns its value
      # Raises otherwise.
      def as_cell_error : CellError
        @raw.as(CellError)
      end

      # Checks that the underlying value is `Xls::Worksheet::Cell::CellError`, and returns its value
      # Returns `nil` otherwise.
      def as_cell_error? : CellError?
        as_cell_error if @raw.is_a?(CellError)
      end
    end

    include InspectableMethods

    @value : Any?

    protected def initialize(@cell : LibXls::StCellData)
    end

    @[Inspectable]
    def id : XlsRecord
      XlsRecord.from_value(@cell.id)
    end

    @[Inspectable]
    def row : UInt16
      @cell.row
    end

    @[Inspectable]
    def col : UInt16
      @cell.col
    end

    def xf?(spreadsheet : Spreadsheet) : Spreadsheet::Xf?
      spreadsheet.xfs[@cell.xf]?
    end

    # See `Xls::Worksheet#defcolwidth`
    @[Inspectable]
    def width : UInt16
      @cell.width
    end

    # Returns whether this cell is hidden
    @[Inspectable]
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
    @[Inspectable]
    def value : Any
      @value ||= begin
        case id
        when .record_boolerr?
          case raw_string
          when "bool"
            return Any.new(raw_double > 0)
          when "error"
            return Any.new(CellError.from_value(raw_double.to_u8))
          end
        when .record_number?, .record_rk?
          return Any.new(raw_double)
        when .record_labelsst?, .record_label?, .record_rstring?
          return Any.new(raw_string)
        end

        Any.new(nil)
      end
    end

    def to_unsafe
      pointerof(@cell)
    end
  end
end
