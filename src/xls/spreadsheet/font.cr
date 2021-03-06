class Xls::Spreadsheet
  class Font
    protected def initialize(
      @font : LibXls::StFontData,
      @real_index : UInt32
    )
    end

    def name : String
      Xls::Utils.ptr_to_str(@font.name)
    end

    # Returns the height of the font (in twips = 1/20 of a point)
    def height : UInt16
      @font.height
    end

    record OptionFlags,
      bold : Bool,
      italic : Bool,
      underlined : Bool,
      struck_out : Bool,
      outlined : Bool,
      shadowed : Bool,
      condensed : Bool,
      extended : Bool

    def flag : OptionFlags
      OptionFlags.new(
        bold: @font.flag.bit(0) == 1,
        italic: @font.flag.bit(1) == 1,
        underlined: @font.flag.bit(2) == 1,
        struck_out: @font.flag.bit(3) == 1,
        outlined: @font.flag.bit(4) == 1,
        shadowed: @font.flag.bit(5) == 1,
        condensed: @font.flag.bit(6) == 1,
        extended: @font.flag.bit(7) == 1,
      )
    end

    enum FontColor : UInt16
      EgaBlack
      EgaWhite
      EgaRed
      EgaGreen
      EgaBlue
      EgaYellow
      EgaMagenta
      EgaCyan
      Automatic  = 0x7FFF
    end

    def color : FontColor
      FontColor.from_value(@font.color)
    end

    def bold : UInt16
      @font.bold
    end

    def standard_font_weight? : Bool
      bold == 400
    end

    def bold_font_weight? : Bool
      bold == 700
    end

    enum FontEscapement : UInt16
      None
      Superscript
      Subscript
    end

    def escapement : FontEscapement
      FontEscapement.from_value(@font.escapement)
    end

    enum FontUnderline : UInt8
      None
      Single
      Double
      SingleAccounting = 0x21
      DoubleAccounting = 0x22
    end

    def underline : FontUnderline
      FontUnderline.from_value(@font.underline)
    end

    enum FontFamily : UInt8
      None       # unknown
      Roman      # variable width, serif
      Swiss      # variable width, sans-serif
      Modern     # fixed width, serif or sans-serif
      Script     # cursive
      Decorative # specialised
    end

    def family : FontFamily
      FontFamily.from_value(@font.family)
    end

    enum FontCharset : UInt8
      AnsiLatin
      SystemDefault
      Symbol
      AppleRoman                 = 0x4D
      AnsiJapaneseShiftJis       = 0x80
      AnsiKoreanHangul           = 0x81
      AnsiKoreanJohab            = 0x82
      AnsiChineseSimplifedGbk    = 0x86
      AnsiChineseTraditionalBig5 = 0x88
      AnsiGreek                  = 0xA1
      AnsiTurkish                = 0xA2
      AnsiVietnamese             = 0xA3
      AnsiHebrew                 = 0xB1
      AnsiArabic                 = 0xB2
      AnsiBaltic                 = 0xBA
      AnsiCyrillic               = 0xCC
      AnsiThai                   = 0xDE
      AnsiLatin2                 = 0xEE
      OemLatin1                  = 0xFF
    end

    def charset : FontCharset
      FontCharset.from_value(@font.charset)
    end

    def to_unsafe
      pointerof(@font)
    end
  end
end
