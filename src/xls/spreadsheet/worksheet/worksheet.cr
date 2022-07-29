require "./cell"
require "./column_info"
require "./row"

class Xls::Worksheet
  enum SheetState : UInt8
    Visibile
    Hidden
    VeryHidden
  end

  enum SheetType : UInt8
    Worksheet
    Chart             = 2
    VisualBasicModule = 6
  end

  include InspectableMethods

  @columns : Array(ColumnInfo)?
  @rows : Array(Row)?

  protected def initialize(
    @worksheet : LibXls::XlsWorkSheet*,
    @sheet : LibXls::StSheetData
  )
  end

  # Returns the name of the worksheet
  @[Inspectable]
  def name : String
    Xls::Utils.ptr_to_str(@sheet.name)
  end

  def columns_info : Array(ColumnInfo)
    @columns ||= begin
      raw_colinfo = @worksheet.value.colinfo
      raw_colinfo.col.to_slice(raw_colinfo.count).each.map do |info|
        ColumnInfo.new(info)
      end.to_a
    end
  end

  def rows : Array(Row)
    @rows ||= begin
      raw_rows = @worksheet.value.rows
      raw_rows.row.to_slice(raw_rows.lastrow).each.map do |row|
        Row.new(row)
      end.to_a
    end
  end

  # Returns the worksheet visibility
  @[Inspectable]
  def sheet_visibility : SheetState
    SheetState.from_value(@sheet.visibility)
  end

  # Returns the worksheet type
  @[Inspectable]
  def sheet_type : SheetType
    SheetType.from_value(@sheet.type)
  end

  # Returns the absolute stream position of the BOF record of the sheet represented by this record
  @[Inspectable]
  def sheet_filepos : UInt32
    sheet_filepos = @sheet.filepos
    worksheet_filepos = @worksheet.value.filepos
    if sheet_filepos != worksheet_filepos
      Log.warn &.emit("Worksheet filepos mismatched",
        sheet_filepos: sheet_filepos,
        worksheet_filepos: worksheet_filepos,
        worksheet_name: name)
    end
    sheet_filepos
  end

  # Returns the default column width for columns that do not have a specific width already set
  #
  # Column width in characters, using the width of the zero character from default font (first FONT record in the file).
  @[Inspectable]
  def defcolwidth : UInt16
    @worksheet.value.defcolwidth
  end

  def to_unsafe
    @worksheet
  end
end
