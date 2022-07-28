class Xls::Spreadsheet
  # TODO: Incomplete
  #
  # See http://sc.openoffice.org/excelfileformat.pdf, page 224
  class Xf
    def initialize(@xf : LibXls::StXfData, @spreadsheet : Spreadsheet)
    end

    # Returns index to FONT record
    def font : UInt16
      @xf.font
    end

    # Returns index to FORMAT record
    def format : UInt16
      @xf.format
    end

    enum Type : UInt16
      Cell
      Style
    end

    struct TypeProtection
      def initialize(@type : UInt16)
      end

      def locked? : Bool
        cell_protection.bit(0) == 1
      end

      def formula_hidden? : Bool
        cell_protection.bit(1) == 1
      end

      def type : Type
        Type.from_value(cell_protection.bit(2))
      end

      def parent_xf : UInt16
        @type.bits(4..15)
      end

      private def cell_protection
        @type.bits(0..2)
      end

      def to_s(io : IO) : Nil
        io << self.class.name
        io << "("

        io << "locked?: "
        locked?.inspect(io)
        io << ", "

        io << "formula_hidden?: "
        formula_hidden?.inspect(io)
        io << ", "

        io << "type: "
        type.inspect(io)
        io << ", "

        io << "parent_xf: "
        parent_xf.inspect(io)

        io << ")"
      end

      def inspect(io : IO) : Nil
        to_s(io)
      end
    end

    # Returns XF type, cell protection, and parent style XF
    def type : TypeProtection
      TypeProtection.new(@xf.type)
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
      def initialize(@align : UInt8)
      end

      def horizontal : HorizontalAlignment
        HorizontalAlignment.from_value(@align.bits(0..2))
      end

      def vertical : VerticalAlignment
        VerticalAlignment.from_value(@align.bits(4..6))
      end

      # Text is wrapped at right border
      def text_wrap_right? : Bool
        @align.bit(3) == 1
      end

      # Justify last line in either `Justified` or `Distributed` texts
      def justify_last_line? : Bool
        @align.bit(7) == 1
      end

      def to_s(io : IO) : Nil
        io << self.class.name
        io << "("

        io << "horizontal: "
        horizontal.inspect(io)
        io << ", "

        io << "vertical: "
        vertical.inspect(io)
        io << ", "

        io << "text_wrap_right?: "
        text_wrap_right?.inspect(io)
        io << ", "

        io << "justify_last_line?: "
        justify_last_line?.inspect(io)

        io << ")"
      end

      def inspect(io : IO) : Nil
        to_s(io)
      end
    end

    # Returns alignment and text break
    def align : Alignment
      Alignment.new(@xf.align)
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

    record Rotation, type : RotationType, value : UInt8

    # Returns text rotation angle
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

    enum TextDirection : UInt8
      # Uses `LeftToRight` or `RightToLeft` depending on text from script
      AccordingToContext
      LeftToRight
      RightToLeft
    end

    struct Indentation
      def initialize(@indent : UInt8)
      end

      # Indent level
      def level : UInt8
        @indent.bits(0..3)
      end

      # Shrink content to fit into cell
      def shrink_to_fit? : Bool
        @indent.bit(4) == 1
      end

      def text_direction
        TextDirection.from_value(@indent.bits(6..7))
      end

      def to_s(io : IO) : Nil
        io << self.class.name
        io << "("

        io << "level: "
        level.inspect(io)
        io << ", "

        io << "shrink_to_fit?: "
        shrink_to_fit?.inspect(io)
        io << ", "

        io << "text_direction: "
        text_direction.inspect(io)

        io << ")"
      end

      def inspect(io : IO) : Nil
        to_s(io)
      end
    end

    # Returns indentation, shrink to cell size, and text direction
    def indent : Indentation
      Indentation.new(@xf.ident)
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

    # Returns flags for used attribute groups
    def used_attrs : UsedAttributesValidity
      current_used_attrs = get_used_attrs_validity

      case type.type
      in .cell?
        parent_used_attrs = @spreadsheet.xfs[type.parent_xf].get_used_attrs_validity
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

    # Returns line style
    def line_style
      # linestyle, uint32
      {% raise "not implemented" %}
    end

    # Returns line color
    def line_color
      # linecolor, uint32
      {% raise "not implemented" %}
    end

    # Returns background color
    def background_color
      # groundcolor, uint16
      {% raise "not implemented" %}
    end

    # TODO: show not implemented methods when implemented
    def to_s(io : IO) : Nil
      io << self.class.name
      io << "("

      io << "font: "
      font.inspect(io)
      io << ", "

      io << "format: "
      format.inspect(io)
      io << ", "

      io << "type: "
      type.inspect(io)
      io << ", "

      io << "align: "
      align.inspect(io)
      io << ", "

      io << "rotation: "
      rotation.inspect(io)
      io << ", "

      io << "indent: "
      indent.inspect(io)
      io << ", "

      io << "used_attrs: "
      used_attrs.inspect(io)

      io << ")"
    end

    def inspect(io : IO) : Nil
      to_s(io)
    end

    def to_unsafe
      pointerof(@xf)
    end
  end
end
