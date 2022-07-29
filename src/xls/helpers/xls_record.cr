module Xls
  {% begin %}
    {% constants = LibXls.constants.select &.starts_with?("XLS_RECORD") %}
    enum XlsRecord
      {% for c in constants %}
      {{ c.gsub(/^XLS\_/, "").capitalize.camelcase.id }} = LibXls::{{ c.id }}
      {% end %}
    end
  {% end %}
end
