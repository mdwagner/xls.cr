module Xls
  class FileNotFound < Exception
  end

  class WorksheetParserExcpetion < Exception
    def initialize(err : LibXls::XlsError)
      super(Xls::Utils.internal_err_to_str(err))
    end
  end

  class Error < Exception
    def initialize(err : LibXls::XlsError)
      super(Xls::Utils.internal_err_to_str(err))
    end
  end

  class WorkbookParserExcpetion < Exception
    def initialize(err : LibXls::XlsError)
      super(Xls::Utils.internal_err_to_str(err))
    end
  end
end
