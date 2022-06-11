module Xls
  # https://github.com/libxls/libxls
  @[Link("xlsreader")]
  lib LibXls
    alias Byte = LibC::UInt8T
    alias Word = LibC::UInt16T
    alias DWord = LibC::UInt32T

    enum XlsError
      LIBXLS_OK
      LIBXLS_ERROR_OPEN
      LIBXLS_ERROR_SEEK
      LIBXLS_ERROR_READ
      LIBXLS_ERROR_PARSE
      LIBXLS_ERROR_MALLOC
    end
    type XlsErrorT = XlsError

    struct StSheetData
      filepos : DWord
      visibility, type : Byte
      name : LibC::Char*
    end

    struct StSheet
      count : DWord
      sheet : StSheetData*
    end

    struct StFontData
      height, flag, color, bold, escapement : Word
      underline, family, charset : Byte
      name : LibC::Char*
    end

    struct StFont
      count : DWord
      font : StFontData*
    end

    struct StFormatData
      index : Word
      value : LibC::Char*
    end

    struct StFormat
      count : DWord
      format : StFormatData*
    end

    struct StXfData
      font, format, type : Word
      align, rotation, ident, usedattr : Byte
      linestyle, linecolor : DWord
      groundcolor : Word
    end

    struct StXf
      count : DWord
      xf : StXfData*
    end

    struct StrSSTString
      str : LibC::Char*
    end

    struct StSST
      count, lastid, continued, lastln, lastrt, lastsz : DWord
      string : StrSSTString*
    end

    struct StCellData
      id, row, col, xf : Word
      str : LibC::Char*
      d : LibC::Double
      l : LibC::Int32T
      width, colspan, rowspan : Word
      isHidden : Byte
    end

    struct StCell
      count : DWord
      cell : StCellData*
    end

    struct StRowData
      index, fcell, lcell, height, flags, xf : Word
      xfflags : Byte
      cells : StCell
    end

    struct StRow
      lastcol, lastrow : Word
      row : StRowData*
    end

    struct StColInfoData
      first, last, width, xf, flags : Word
    end

    struct StColInfo
      count : DWord
      col : StColInfoData*
    end

    struct XlsWorkBook
      # file : Void* # FILE*
      # olestr : Void* # OLE2Stream*
      filepos : LibC::Int32T

      # From Header (BIFF)
      is5ver, is1904 : Byte
      type, activeSheetIdx : Word

      # Other data
      codepage : Word
      charset : LibC::Char*
      sheets : StSheet
      sst : StSST
      xfs : StXf
      fonts : StFont
      formats : StFormat

      summary, docSummary : LibC::Char*

      converter, utf16_converter, utf8_locale : Void*
    end

    struct XlsWorkSheet
      filepos : DWord
      defcolwidth : Word
      rows : StRow
      workbook : XlsWorkBook*
      colinfo : StColInfo
    end

    alias XlsCell = StCellData
    alias XlsRow = StRowData

    struct XlsSummaryInfo
      title, subject, author, keywords, comment, lastAuthor, appName, category, manager, company : Byte*
    end

    alias XlsFormulaHandler = Word, Word, Byte* -> Void

    fun version = xls_getVersion : LibC::Char*
    fun error = xls_getError(code : XlsErrorT) : LibC::Char*

    # Set debug. Force library to load?
    fun xls(debug : Int32) : Int32
    fun set_formula_handler = xls_set_formula_hander(handler : XlsFormulaHandler) : Void

    fun parse_workbook = xls_parseWorkBook(workbook : XlsWorkBook*) : XlsErrorT
    fun parse_worksheet = xls_parseWorkSheet(worksheet : XlsWorkSheet*) : XlsErrorT

    # # Preferred API
    # charset - convert 16bit strings within the spread sheet to this 8bit encoding (UTF-8 default)
    fun open_file = xls_open_file(file : LibC::Char*, charset : LibC::Char*, out_error : XlsErrorT*) : XlsWorkBook*
    fun open_buffer = xls_open_buffer(data : LibC::Char*, data_len : LibC::SizeT, charset : LibC::Char*, out_error : XlsErrorT*) : XlsWorkBook*
    fun close_workbook = xls_close_WB(workbook : XlsWorkBook*) : Void

    # # Historical API
    fun open = xls_open(file : LibC::Char*, charset : LibC::Char*) : XlsWorkBook*

    fun get_worksheet = xls_getWorkSheet(workbook : XlsWorkBook*, num : Int32) : XlsWorkSheet*
    fun close_worksheet = xls_close_WS(worksheet : XlsWorkSheet*) : Void

    fun summary_info = xls_summaryInfo(workbook : XlsWorkBook*) : XlsSummaryInfo*
    fun close_summary_info = xls_close_summaryInfo(summary : XlsSummaryInfo*) : Void

    # # utility function
    fun row = xls_row(worksheet : XlsWorkSheet*, cell_row : Word) : XlsRow*
    fun cell = xls_cell(worksheet : XlsWorkSheet*, cell_row : Word, cell_col : Word) : XlsCell*
  end
end
