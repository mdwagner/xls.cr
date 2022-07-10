class Xls::Worksheet
  module Cell
    struct Error
    end

    struct Any
      alias Type = Nil | Bool | Float64 | String | Error

      getter raw : Type

      def initialize(@raw)
      end

      # TODO: add methods to access underlying type safe/unsafe
    end
  end
end
