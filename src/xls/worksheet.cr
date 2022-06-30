class Xls::Worksheet
  class ParserError < Exception
    def initialize(message = "Unknown")
      super(message)
    end
  end

  class CellError < Exception
    def to_s(io)
      io << self.class.name
    end
  end

  protected def initialize(@worksheet : LibXls::XlsWorkSheet*)
  end

  def parse! : Nil
    status = LibXls.parse_worksheet(@worksheet)
    unless status.libxls_ok?
      message = String.new(LibXls.error(status))
      raise ParserError.new(message)
    end
  end

  def close! : Nil
    LibXls.close_worksheet(@worksheet)
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

  private def cell_value(cell)
    if id = XlsRecord.from_value?(cell.id)
      case id
      when .record_boolerr?
        cell_boolerr(cell)
      when .record_number?, .record_rk?
        cell.d.to_f64
      when .record_labelsst?, .record_label?, .record_rstring?
        ptr_to_s(cell.str)
      else
        nil
      end
    else
      nil
    end
  end

  private def cell_boolerr(cell, str_fallback = false)
    str = ptr_to_s(cell.str)
    case str
    when "bool"
      cell.d.to_f64 > 0
    when "error"
      CellError.new
    else
      if str_fallback
        str
      else
        nil
      end
    end
  end
end
