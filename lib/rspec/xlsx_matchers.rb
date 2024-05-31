# frozen_string_literal: true

require_relative "xlsx_matchers/version"
require_relative "xlsx_matchers/sheets"
require_relative "xlsx_matchers/columns"

begin
  require "roo"
rescue LoadError
  # optional dependency
end

begin
  require "caxlsx"
rescue LoadError
  # optional dependency
end
module RSpec
  module XlsxMatchers
    class Error < StandardError; end

    def have_excel_sheets(sheet_names)
      Sheets.new(sheet_names)
    end

    def have_excel_columns(column_names)
      Columns.new(column_names)
    end
  end
end
