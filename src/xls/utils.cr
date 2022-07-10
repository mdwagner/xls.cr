module Xls::Utils
  extend self

  def ptr_to_str(ptr) : String
    if ptr
      String.new(ptr)
    else
      ""
    end
  end
end
