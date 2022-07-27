class Xls::Spreadsheet
  # TODO: Incomplete
  #
  # See http://sc.openoffice.org/excelfileformat.pdf, page 224
  class Xf
    def initialize(@xf : LibXls::StXfData)
    end

    # Returns index to FONT record
    def font : UInt16
      @xf.font
    end

    # Returns index to FORMAT record
    def format : UInt16
      @xf.format
    end

    # Returns XF type, cell protection, and parent style XF
    def type
      {% raise "not implemented" %}
    end

    # Returns alignment and text break
    def align
      {% raise "not implemented" %}
    end

    # Returns indentation, shrink to cell size, and text direction
    def indent
      {% raise "not implemented" %}
    end

    # Returns flags for used attribute groups
    def used_attrs
      {% raise "not implemented" %}
    end

    # Returns line style
    def line_style
      {% raise "not implemented" %}
    end

    # Returns line color
    def line_color
      {% raise "not implemented" %}
    end

    # Returns background color
    def background_color
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
