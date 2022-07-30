class Xls::Spreadsheet
  # See http://sc.openoffice.org/excelfileformat.pdf, page 224
  class Xf
    enum Type : UInt16
      Cell
      Style
    end

    struct TypeProtection
      include InspectableMethods

      protected def initialize(@type : UInt16)
      end

      @[Inspectable]
      def locked? : Bool
        cell_protection.bit(0) == 1
      end

      @[Inspectable]
      def formula_hidden? : Bool
        cell_protection.bit(1) == 1
      end

      @[Inspectable]
      def type : Type
        Type.from_value(cell_protection.bit(2))
      end

      @[Inspectable(base: 16, precision: 3)]
      def parent_xf_index : UInt16
        @type.bits(4..15)
      end

      private def cell_protection
        @type.bits(0..2)
      end
    end

    enum HorizontalAlignment : UInt8
      General
      Left
      Centered
      Right
      Filled
      Justified
      CenteredAcrossSelection
      Distributed
    end

    enum VerticalAlignment : UInt8
      Top
      Centered
      Bottom
      Justified
      Distributed
    end

    struct Alignment
      include InspectableMethods

      protected def initialize(@align : UInt8)
      end

      @[Inspectable]
      def horizontal : HorizontalAlignment
        HorizontalAlignment.from_value(@align.bits(0..2))
      end

      @[Inspectable]
      def vertical : VerticalAlignment
        VerticalAlignment.from_value(@align.bits(4..6))
      end

      # Text is wrapped at right border
      @[Inspectable]
      def text_wrap_right? : Bool
        @align.bit(3) == 1
      end

      # Justify last line in either `Justified` or `Distributed` texts
      @[Inspectable]
      def justify_last_line? : Bool
        @align.bit(7) == 1
      end
    end

    enum RotationType
      # Not rotated
      NotRotated
      # 1 to 90 degrees counterclockwise
      CCW
      # 1 to 90 degrees clockwise
      CW
      # Letters are stacked top-to-bottom, but not rotated
      TopToBottom
      # Not in spec
      Unknown
    end

    # See http://sc.openoffice.org/excelfileformat.pdf, page 220
    record Rotation, type : RotationType, value : UInt8

    enum TextDirection : UInt8
      # Uses `LeftToRight` or `RightToLeft` depending on text from script
      AccordingToContext
      LeftToRight
      RightToLeft
    end

    struct Indentation
      include InspectableMethods

      protected def initialize(@indent : UInt8)
      end

      # Indent level
      @[Inspectable]
      def level : UInt8
        @indent.bits(0..3)
      end

      # Shrink content to fit into cell
      @[Inspectable]
      def shrink_to_fit? : Bool
        @indent.bit(4) == 1
      end

      @[Inspectable]
      def text_direction
        TextDirection.from_value(@indent.bits(6..7))
      end
    end

    record UsedAttributesValidity,
      # Flag for number format
      number_format : Bool,
      # Flag for font
      font : Bool,
      # Flag for horizontal and vertical alignment, text wrap, indentation, orientation, rotation, and text direction
      alignment : Bool,
      # Flag for border lines
      border_lines : Bool,
      # Flag for background area style
      background_style : Bool,
      # Flag for cell protection (cell locked and formula hidden)
      cell_protection : Bool

    enum LineStyle : UInt8
      NoLine
      Thin
      Medium
      Dashed
      Dotted
      Thick
      Double
      Hair
      MediumDashed
      ThinDashDotted
      MediumDashDotted
      ThinDashDotDotted
      MediumDashDotDotted
      SlantedMediumDashDotted
    end

    struct BorderLineBackground
      include InspectableMethods

      protected def initialize(
        @line_style : UInt32,
        @line_color : UInt32,
        @background_color : UInt16
      )
      end

      @[Inspectable]
      def left_line_style : LineStyle
        LineStyle.from_value(@line_style.bits(0..3))
      end

      @[Inspectable]
      def right_line_style : LineStyle
        LineStyle.from_value(@line_style.bits(4..7))
      end

      @[Inspectable]
      def top_line_style : LineStyle
        LineStyle.from_value(@line_style.bits(8..11))
      end

      @[Inspectable]
      def bottom_line_style : LineStyle
        LineStyle.from_value(@line_style.bits(12..15))
      end

      # See http://sc.openoffice.org/excelfileformat.pdf, page 196 for color index mapping
      def left_line_color_index
        @line_style.bits(16..22)
      end

      @[Inspectable(base: 16, precision: 6)]
      def left_line_color(default_color_index : UInt16 = 0) : UInt32
        LibXls.get_color(left_line_color_index, default_color_index)
      end

      # See http://sc.openoffice.org/excelfileformat.pdf, page 196 for color index mapping
      def right_line_color_index
        @line_style.bits(23..29)
      end

      # Diagonal line from top left to right bottom
      @[Inspectable]
      def diag_from_top_left?
        @line_style.bit(30) == 1
      end

      # Diagonal line from bottom left to right top
      @[Inspectable]
      def diag_from_bottom_left?
        @line_style.bit(31) == 1
      end

      # See http://sc.openoffice.org/excelfileformat.pdf, page 196 for color index mapping
      def top_line_color_index
        @line_color.bits(0..6)
      end

      @[Inspectable(base: 16, precision: 6)]
      def top_line_color(default_color_index : UInt16 = 0) : UInt32
        LibXls.get_color(top_line_color_index, default_color_index)
      end

      # See http://sc.openoffice.org/excelfileformat.pdf, page 196 for color index mapping
      def bottom_line_color_index
        @line_color.bits(7..13)
      end

      @[Inspectable(base: 16, precision: 6)]
      def bottom_line_color(default_color_index : UInt16 = 0) : UInt32
        LibXls.get_color(bottom_line_color_index, default_color_index)
      end

      # See http://sc.openoffice.org/excelfileformat.pdf, page 196 for color index mapping
      def diag_line_color_index
        @line_color.bits(14..20)
      end

      @[Inspectable(base: 16, precision: 6)]
      def diag_line_color(default_color_index : UInt16 = 0) : UInt32
        LibXls.get_color(diag_line_color_index, default_color_index)
      end

      # Diagonal line style
      @[Inspectable]
      def diag_line_style : LineStyle
        LineStyle.from_value(@line_color.bits(21..24))
      end

      @[Inspectable]
      def fill_pattern : LineStyle
        LineStyle.from_value(@line_color.bits(26..31))
      end

      # See http://sc.openoffice.org/excelfileformat.pdf, page 196 for color index mapping
      def pattern_color_index
        @background_color.bits(0..6)
      end

      @[Inspectable(base: 16, precision: 6)]
      def pattern_color(default_color_index : UInt16 = 0) : UInt32
        LibXls.get_color(pattern_color_index, default_color_index)
      end

      # See http://sc.openoffice.org/excelfileformat.pdf, page 196 for color index mapping
      def pattern_background_color_index
        @background_color.bits(7..13)
      end

      @[Inspectable(base: 16, precision: 6)]
      def pattern_background_color(default_color_index : UInt16 = 0) : UInt32
        LibXls.get_color(pattern_background_color_index, default_color_index)
      end
    end

    include InspectableMethods

    protected def initialize(@xf : LibXls::StXfData, @spreadsheet : Spreadsheet)
    end

    # Returns associated `Xls::Spreadsheet::Font` or nil
    @[Inspectable]
    def font? : Font?
      @spreadsheet.fonts.find { |font| font.real_index == @xf.font }
    end

    # Returns associated `Xls::Spreadsheet::Format` or nil
    @[Inspectable]
    def format? : Format?
      @spreadsheet.formats[@xf.format]?
    end

    # Returns XF type, cell protection, and parent style XF
    @[Inspectable]
    def type : TypeProtection
      TypeProtection.new(@xf.type)
    end

    # Returns alignment and text break
    @[Inspectable]
    def alignment : Alignment
      Alignment.new(@xf.align)
    end

    # Returns text rotation angle
    @[Inspectable]
    def rotation : Rotation
      case value = @xf.rotation
      when 0
        Rotation.new(:not_rotated, value)
      when 1..90
        Rotation.new(:ccw, value)
      when 91..180
        Rotation.new(:cw, value)
      when 255
        Rotation.new(:top_to_bottom, value)
      else
        Rotation.new(:unknown, value)
      end
    end

    # Returns indentation, shrink to cell size, and text direction
    @[Inspectable]
    def indent : Indentation
      Indentation.new(@xf.ident)
    end

    # Returns flags for used attribute groups
    #
    # Checks for validity on specific groups on this XF record
    @[Inspectable]
    def used_attrs : UsedAttributesValidity
      current_used_attrs = get_used_attrs_validity

      case type.type
      in .cell?
        parent_used_attrs = @spreadsheet.xfs[type.parent_xf_index].get_used_attrs_validity
        UsedAttributesValidity.new(
          number_format: current_used_attrs.number_format ? parent_used_attrs.number_format : true,
          font: current_used_attrs.font ? parent_used_attrs.font : true,
          alignment: current_used_attrs.alignment ? parent_used_attrs.alignment : true,
          border_lines: current_used_attrs.border_lines ? parent_used_attrs.border_lines : true,
          background_style: current_used_attrs.background_style ? parent_used_attrs.background_style : true,
          cell_protection: current_used_attrs.cell_protection ? parent_used_attrs.cell_protection : true
        )
      in .style?
        current_used_attrs
      end
    end

    protected def get_used_attrs_validity : UsedAttributesValidity
      UsedAttributesValidity.new(
        number_format: @xf.usedattr.bits(2..7).bit(0) == 0,
        font: @xf.usedattr.bits(2..7).bit(1) == 0,
        alignment: @xf.usedattr.bits(2..7).bit(2) == 0,
        border_lines: @xf.usedattr.bits(2..7).bit(3) == 0,
        background_style: @xf.usedattr.bits(2..7).bit(4) == 0,
        cell_protection: @xf.usedattr.bits(2..7).bit(5) == 0,
      )
    end

    # Returns information on border line styles, line colors, and background colors
    @[Inspectable]
    def border_line_background : BorderLineBackground
      BorderLineBackground.new(@xf.linestyle, @xf.linecolor, @xf.groundcolor)
    end

    def to_unsafe
      pointerof(@xf)
    end
  end
end
