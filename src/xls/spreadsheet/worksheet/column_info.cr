class Xls::Worksheet
  # See http://sc.openoffice.org/excelfileformat.pdf, page 146
  class ColumnInfo
    struct OptionFlags
      include InspectableMethods

      protected def initialize(@flags : UInt16)
      end

      @[Inspectable]
      def columns_hidden? : Bool
        @flags.bit(0) == 1
      end

      @[Inspectable]
      def columns_collapsed? : Bool
        @flags.bit(12) == 1
      end

      @[Inspectable]
      def columns_outline_level : UInt16
        @flags.bits(8..10)
      end

      @[Inspectable]
      def no_outline? : Bool
        columns_outline_level == 0
      end
    end

    include InspectableMethods

    protected def initialize(@colinfo : LibXls::StColInfoData)
    end

    # Returns the index to the first column in the range
    @[Inspectable]
    def first : UInt16
      @colinfo.first
    end

    # Returns the index to the last column in the range
    @[Inspectable]
    def last : UInt16
      @colinfo.last
    end

    # Returns the width of the columns in 1/256 of the width of the zero character, using default font (the first FONT record in the file)
    @[Inspectable]
    def width : UInt16
      @colinfo.width
    end

    # Returns the `Xls::Spreadsheet::Xf` for default column formatting
    def xf(spreadsheet : Spreadsheet) : Spreadsheet::Xf
      spreadsheet.xfs[@colinfo.xf]
    end

    # Returns option flags
    @[Inspectable]
    def flags : OptionFlags
      OptionFlags.new(@colinfo.flags)
    end

    def to_unsafe
      pointerof(@colinfo)
    end
  end
end
