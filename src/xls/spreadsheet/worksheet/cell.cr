class Xls::Worksheet
  class Cell
    struct Error
      def to_s(io : IO) : Nil
        io << self.class.name
      end

      def inspect(io : IO) : Nil
        to_s(io)
      end
    end

    struct Any
      alias Type = Nil | Bool | Float64 | String | Error

      getter raw : Type

      def initialize(@raw)
      end

      def to_s(io : IO) : Nil
        io << self.class.name
        io << "("

        io << "raw: "
        raw.inspect(io)

        io << ")"
      end

      def inspect(io : IO) : Nil
        to_s(io)
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

    @value : Any?

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

    def to_s(io : IO) : Nil
      io << self.class.name
      io << "("

      io << "id: "
      id.inspect(io)
      io << ", "

      io << "row: "
      row.inspect(io)
      io << ", "

      io << "col: "
      col.inspect(io)
      io << ", "

      io << "xf: "
      xf.inspect(io)
      io << ", "

      io << "width: "
      width.inspect(io)
      io << ", "

      io << "is_hidden?: "
      is_hidden?.inspect(io)
      io << ", "

      io << "value: "
      value.inspect(io)

      io << ")"
    end

    def inspect(io : IO) : Nil
      to_s(io)
    end

    def to_unsafe
      pointerof(@cell)
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
end
