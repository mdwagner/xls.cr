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
      raise NotImplementedError.new("Xf#type")
    end

    # Returns alignment and text break
    def align
      raise NotImplementedError.new("Xf#align")
    end

    # Returns indentation, shrink to cell size, and text direction
    def indent
      raise NotImplementedError.new("Xf#indent")
    end

    # Returns flags for used attribute groups
    def used_attrs
      raise NotImplementedError.new("Xf#used_attrs")
    end

    # Returns line style
    def line_style
      raise NotImplementedError.new("Xf#line_style")
    end

    # Returns line color
    def line_color
      raise NotImplementedError.new("Xf#line_color")
    end

    # Returns background color
    def background_color
      raise NotImplementedError.new("Xf#background_color")
    end

    def to_unsafe
      pointerof(@xf)
    end
  end
end
