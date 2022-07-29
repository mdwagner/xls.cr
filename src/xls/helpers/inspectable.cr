module Xls
  annotation Inspectable
  end

  module InspectableMethods
    def to_s(io : IO) : Nil
      io << self.class.name
      io << "("
      {% begin %}
      {% i_methods = @type.methods.select &.annotation(Inspectable) %}
      {% found_one = false %}
      {% for method in i_methods %}
        {% anno = method.annotation(Inspectable) %}
        {% if found_one %}
          io << ", "
        {% end %}
        io << {{method.name.stringify}} << ": "
        {% if anno[:base] %}
          {% if anno[:base] == 16 %}
            io << "0x"
          {% elsif anno[:base] == 8 %}
            io << "0o"
          {% elsif anno[:base] == 2 %}
            io << "0b"
          {% end %}
          {{method.name.id}}.to_s(
            io: io,
            base: {{anno[:base].id}},
            precision: {{anno[:precision] ? anno[:precision].id : 2}},
            upcase: true
          )
        {% else %}
          {{method.name.id}}.inspect(io)
        {% end %}
        {% found_one = true %}
      {% end %}
    {% end %}
      io << ")"
    end

    def inspect(io : IO) : Nil
      to_s(io)
    end
  end
end
