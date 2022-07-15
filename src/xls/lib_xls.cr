module Xls
  # https://github.com/libxls/libxls
  @[Link("xlsreader")]
  lib LibXls
    XLS_RECORD_EOF              = 0x000A
    XLS_RECORD_DEFINEDNAME      = 0x0018
    XLS_RECORD_NOTE             = 0x001C
    XLS_RECORD_1904             = 0x0022
    XLS_RECORD_CONTINUE         = 0x003C
    XLS_RECORD_WINDOW1          = 0x003D
    XLS_RECORD_CODEPAGE         = 0x0042
    XLS_RECORD_OBJ              = 0x005D
    XLS_RECORD_MERGEDCELLS      = 0x00E5
    XLS_RECORD_DEFCOLWIDTH      = 0x0055
    XLS_RECORD_COLINFO          = 0x007D
    XLS_RECORD_BOUNDSHEET       = 0x0085
    XLS_RECORD_PALETTE          = 0x0092
    XLS_RECORD_MULRK            = 0x00BD
    XLS_RECORD_MULBLANK         = 0x00BE
    XLS_RECORD_RSTRING          = 0x00D6
    XLS_RECORD_DBCELL           = 0x00D7
    XLS_RECORD_XF               = 0x00E0
    XLS_RECORD_MSODRAWINGGROUP  = 0x00EB
    XLS_RECORD_MSODRAWING       = 0x00EC
    XLS_RECORD_SST              = 0x00FC
    XLS_RECORD_LABELSST         = 0x00FD
    XLS_RECORD_EXTSST           = 0x00FF
    XLS_RECORD_TXO              = 0x01B6
    XLS_RECORD_HYPERREF         = 0x01B8
    XLS_RECORD_BLANK            = 0x0201
    XLS_RECORD_NUMBER           = 0x0203
    XLS_RECORD_LABEL            = 0x0204
    XLS_RECORD_BOOLERR          = 0x0205
    XLS_RECORD_STRING           = 0x0207 # only follows a formula
    XLS_RECORD_ROW              = 0x0208
    XLS_RECORD_INDEX            = 0x020B
    XLS_RECORD_ARRAY            = 0x0221 # Array-entered formula
    XLS_RECORD_DEFAULTROWHEIGHT = 0x0225
    XLS_RECORD_FONT             = 0x0031 # spec says 0x0231 but Excel expects 0x0031
    XLS_RECORD_FONT_ALT         = 0x0231
    XLS_RECORD_WINDOW2          = 0x023E
    XLS_RECORD_RK               = 0x027E
    XLS_RECORD_STYLE            = 0x0293
    XLS_RECORD_FORMULA          = 0x0006
    XLS_RECORD_FORMULA_ALT      = 0x0406 # Apple Numbers bug
    XLS_RECORD_FORMAT           = 0x041E
    XLS_RECORD_BOF              = 0x0809

    enum XlsError
      LIBXLS_OK
      LIBXLS_ERROR_OPEN
      LIBXLS_ERROR_SEEK
      LIBXLS_ERROR_READ
      LIBXLS_ERROR_PARSE
      LIBXLS_ERROR_MALLOC
    end

    type OLE2StreamT = Void*

    struct StSheetData
      filepos : LibC::UInt32T
      visibility, type : LibC::UInt8T
      name : LibC::Char*
    end

    struct StSheet
      count : LibC::UInt32T
      sheet : StSheetData*
    end

    struct StFontData
      height, flag, color, bold, escapement : LibC::UInt16T
      underline, family, charset : LibC::UInt8T
      name : LibC::Char*
    end

    struct StFont
      count : LibC::UInt32T
      font : StFontData*
    end

    struct StFormatData
      index : LibC::UInt16T
      value : LibC::Char*
    end

    struct StFormat
      count : LibC::UInt32T
      format : StFormatData*
    end

    struct StXfData
      font, format, type : LibC::UInt16T
      align, rotation, ident, usedattr : LibC::UInt8T
      linestyle, linecolor : LibC::UInt32T
      groundcolor : LibC::UInt16T
    end

    struct StXf
      count : LibC::UInt32T
      xf : StXfData*
    end

    struct StrSSTString
      str : LibC::Char*
    end

    struct StSST
      count, lastid, continued, lastln, lastrt, lastsz : LibC::UInt32T
      string : StrSSTString*
    end

    struct StCellData
      id, row, col, xf : LibC::UInt16T
      str : LibC::Char*
      d : LibC::Double
      l : LibC::Int32T
      width, colspan, rowspan : LibC::UInt16T
      isHidden : LibC::UInt8T
    end

    struct StCell
      count : LibC::UInt32T
      cell : StCellData*
    end

    struct StRowData
      index, fcell, lcell, height, flags, xf : LibC::UInt16T
      xfflags : LibC::UInt8T
      cells : StCell
    end

    struct StRow
      lastcol, lastrow : LibC::UInt16T
      row : StRowData*
    end

    struct StColInfoData
      first, last, width, xf, flags : LibC::UInt16T
    end

    struct StColInfo
      count : LibC::UInt32T
      col : StColInfoData*
    end

    struct XlsWorkBook
      olestr : OLE2StreamT*
      filepos : LibC::Int32T
      is5ver, is1904 : LibC::UInt8T
      type, activeSheetIdx : LibC::UInt16T
      codepage : LibC::UInt16T
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
      filepos : LibC::UInt32T
      defcolwidth : LibC::UInt16T
      rows : StRow
      workbook : XlsWorkBook*
      colinfo : StColInfo
    end

    alias XlsCell = StCellData
    alias XlsRow = StRowData

    struct XlsSummaryInfo
      title, subject, author, keywords, comment, lastAuthor, appName, category, manager, company : LibC::UInt8T*
    end

    alias XlsFormulaHandler = LibC::UInt16T, LibC::UInt16T, LibC::UInt8T* -> Void

    fun version = xls_getVersion : LibC::Char*
    fun error = xls_getError(code : XlsError) : LibC::Char*

    # Set debug. Force library to load?
    fun xls(debug : Int32) : Int32
    fun set_formula_handler = xls_set_formula_hander(handler : XlsFormulaHandler) : Void

    fun parse_workbook = xls_parseWorkBook(workbook : XlsWorkBook*) : XlsError
    fun parse_worksheet = xls_parseWorkSheet(worksheet : XlsWorkSheet*) : XlsError

    # Preferred API
    # charset - convert 16bit strings within the spread sheet to this 8bit encoding (UTF-8 default)
    fun open_file = xls_open_file(file : LibC::Char*, charset : LibC::Char*, out_error : XlsError*) : XlsWorkBook*
    fun open_buffer = xls_open_buffer(data : LibC::Char*, data_len : LibC::SizeT, charset : LibC::Char*, out_error : XlsError*) : XlsWorkBook*
    fun close_workbook = xls_close_WB(workbook : XlsWorkBook*) : Void

    # Historical API
    fun open = xls_open(file : LibC::Char*, charset : LibC::Char*) : XlsWorkBook*

    @[Raises]
    fun get_worksheet = xls_getWorkSheet(workbook : XlsWorkBook*, num : Int32) : XlsWorkSheet*
    fun close_worksheet = xls_close_WS(worksheet : XlsWorkSheet*) : Void

    fun summary_info = xls_summaryInfo(workbook : XlsWorkBook*) : XlsSummaryInfo*
    fun close_summary_info = xls_close_summaryInfo(summary : XlsSummaryInfo*) : Void

    # utility function
    fun row = xls_row(worksheet : XlsWorkSheet*, cell_row : LibC::UInt16T) : XlsRow*
    fun cell = xls_cell(worksheet : XlsWorkSheet*, cell_row : LibC::UInt16T, cell_col : LibC::UInt16T) : XlsCell*
  end
end
