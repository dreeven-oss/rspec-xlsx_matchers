# frozen_string_literal: true

require_relative "xlsx_matchers/version"
require_relative "xlsx_matchers/exact_match"
require_relative "xlsx_matchers/in_row"
require_relative "xlsx_matchers/utils"
require_relative "xlsx_matchers/base_sheet"

require_relative "xlsx_matchers/sheets"
require_relative "xlsx_matchers/columns"
require_relative "xlsx_matchers/empty_row"
require_relative "xlsx_matchers/cells"


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
    # class Error < StandardError; end

    def have_excel_sheets(sheet_names)
      Sheets.new(sheet_names)
    end

    def have_excel_columns(column_names)
      Columns.new(column_names)
    end

    def have_excel_empty_row(index)
      EmptyRow.new(index)
    end

    def have_excel_cells(cells)
      Cells.new(cells)
    end
  end
end
