class Xls::Spreadsheet
  record Format,
    index : UInt16,
    value : String do
    def to_s(io : IO) : Nil
      io << self.class.name
      io << "("
      io << "index: "
      index.inspect(io)
      io << ", "
      io << "value: "
      value.inspect(io)
      io << ")"
    end

    def inspect(io : IO) : Nil
      to_s(io)
    end
  end
end
