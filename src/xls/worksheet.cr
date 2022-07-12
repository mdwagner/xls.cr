class Xls::Worksheet
  protected def initialize(
    raw_worksheet @worksheet : LibXls::XlsWorkSheet*,
    @sheet_name : String,
    raw_visibility @sheet_visibility : UInt8,
    raw_type @sheet_type : UInt8,
    raw_filepos @sheet_filepos : UInt32
  )
  end

  # Returns the name of the worksheet
  def name
    @sheet_name
  end

  # :nodoc:
  def raw_visibility
    @sheet_visibility
  end

  # :ditto:
  def raw_type
    @sheet_type
  end

  # :ditto:
  def raw_filepos
    @sheet_filepos
  end

  def row_count
    @worksheet.value.rows.lastrow.to_i
  end

  def col_count
    @worksheet.value.rows.lastcol.to_i
  end

  def each_row(**kwargs, &)
    headers = get_headers
    row_count.times do |row|
      next if row == 0

      hash = {} of String => String

      if kwargs.empty?
        headers.each do |key, col|
          cell = LibXls.cell(@worksheet, row, col)
          value = cell_value(cell.value).to_s
          hash[key] = value
        end
      else
        kwargs.each do |hash_key, matching_header|
          col = headers[matching_header]
          cell = LibXls.cell(@worksheet, row, col)
          value = cell_value(cell.value).to_s
          hash["#{hash_key}"] = value
        end
      end

      yield hash
    end
  end

  private def get_headers
    headers = {} of String => Int32
    col_count.times do |col|
      cell = LibXls.cell(@worksheet, 0, col)
      value = cell_value(cell.value).to_s
      headers[value] = col
    end
    headers
  end

  private def ptr_to_s(ptr) : String
    if ptr
      String.new(ptr)
    else
      ""
    end
  end

  private def cell_value(cell) : Cell::Any
    if id = XlsRecord.from_value?(cell.id)
      case id
      when .record_boolerr?
        cell_boolerr(cell)
      when .record_number?, .record_rk?
        Cell::Any.new(cell.d.to_f64)
      when .record_labelsst?, .record_label?, .record_rstring?
        Cell::Any.new(ptr_to_s(cell.str))
      else
        Cell::Any.new(nil)
      end
    else
      Cell::Any.new(nil)
    end
  end

  private def cell_boolerr(cell, str_fallback = false) : Cell::Any
    str = ptr_to_s(cell.str)
    case str
    when "bool"
      Cell::Any.new(cell.d.to_f64 > 0)
    when "error"
      Cell::Any.new(Cell::Error.new)
    else
      if str_fallback
        Cell::Any.new(str)
      else
        Cell::Any.new(nil)
      end
    end
  end
end
