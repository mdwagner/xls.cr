module Xls
  {% begin %}
  {%
    constants = LibXls.constants.select do |constant|
      constant.starts_with?("XLS_RECORD")
    end
  %}
    enum XlsRecord
      {% for constant in constants %}
      {{ constant.gsub(/^XLS\_/, "").capitalize.camelcase.id }} = LibXls::{{ constant.id }}
      {% end %}
    end
  {% end %}
end
