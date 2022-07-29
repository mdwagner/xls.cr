class Xls::Worksheet
  class Row
    record Height,
      height : UInt16,
      default_height : Bool do
      def is_custom_height? : Bool
        !default_height
      end
    end

    include InspectableMethods

    @cells : Array(Cell)?

    protected def initialize(@row : LibXls::StRowData)
    end

    # Returns the index of this `Xls::Worksheet::Row` in the `Xls::Worksheet`
    @[Inspectable]
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
    @[Inspectable]
    def fcell : UInt16
      @row.fcell
    end

    # Returns the index to the column of the last cell which is described by a cell record, increased by 1
    @[Inspectable]
    def lcell : UInt16
      @row.lcell
    end

    # Returns the height of the row (represented as `Xls::Worksheet::Row::Height`), in twips = 1/20 of a point
    @[Inspectable]
    def height : Height
      Height.new(
        height: @row.height.bits(0..14),
        default_height: @row.height.bit(15) == 1
      )
    end

    def xf?(spreadsheet : Spreadsheet) : Spreadsheet::Xf?
      spreadsheet.xfs[@row.xf]?
    end

    def to_unsafe
      pointerof(@row)
    end
  end
end
