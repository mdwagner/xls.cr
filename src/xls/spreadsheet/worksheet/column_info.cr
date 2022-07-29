class Xls::Worksheet
  class ColumnInfo
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

      def to_s(io : IO) : Nil
        io << self.class.name
        io << "("

        io << "columns_hidden?: "
        columns_hidden?.inspect(io)
        io << ", "

        io << "columns_collapsed?: "
        columns_collapsed?.inspect(io)
        io << ", "

        io << "columns_outline_level: "
        columns_outline_level.inspect(io)
        io << ", "

        io << "no_outline?: "
        no_outline?.inspect(io)

        io << ")"
      end

      def inspect(io : IO) : Nil
        to_s(io)
      end
    end

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

    # Returns the XF record for default column formatting
    def xf?(spreadsheet : Spreadsheet) : Spreadsheet::Xf?
      spreadsheet.xfs[@colinfo.xf]?
    end

    # Returns option flags
    def flags : OptionFlags
      OptionFlags.new(
        columns_hidden: @colinfo.flags.bit(0) == 1,
        columns_outline_level: @colinfo.flags.bits(8..10),
        columns_collapsed: @colinfo.flags.bit(12) == 1
      )
    end

    def to_s(io : IO) : Nil
      io << self.class.name
      io << "("

      io << "first: "
      first.inspect(io)
      io << ", "

      io << "last: "
      last.inspect(io)
      io << ", "

      io << "width: "
      width.inspect(io)
      io << ", "

      io << "flags: "
      flags.inspect(io)

      io << ")"
    end

    def inspect(io : IO) : Nil
      to_s(io)
    end

    def to_unsafe
      pointerof(@colinfo)
    end
  end
end
