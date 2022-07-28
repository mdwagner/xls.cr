module Xls
  class FileNotFound < Exception
  end

  abstract class BaseException < Exception
    def initialize(err : LibXls::XlsError)
      super(Xls::Utils.internal_err_to_str(err))
    end
  end

  class WorksheetParserException < BaseException
  end

  class Error < BaseException
  end

  class WorkbookParserException < BaseException
  end
end
