module Xls
  class FileNotFound < Exception
  end

  class WorksheetParserException < Exception
    def initialize(err : LibXls::XlsError)
      super(Xls::Utils.internal_err_to_str(err))
    end
  end

  class Error < Exception
    def initialize(err : LibXls::XlsError)
      super(Xls::Utils.internal_err_to_str(err))
    end
  end

  class WorkbookParserException < Exception
    def initialize(err : LibXls::XlsError)
      super(Xls::Utils.internal_err_to_str(err))
    end
  end
end
