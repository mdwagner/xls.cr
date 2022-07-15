module Xls::Utils
  extend self

  def ptr_to_str(ptr) : String
    if ptr
      String.new(ptr)
    else
      ""
    end
  end

  def internal_err_to_str(err : LibXls::XlsError) : String
    ptr = LibXls.error(err)
    ptr_to_str(ptr)
  end
end
