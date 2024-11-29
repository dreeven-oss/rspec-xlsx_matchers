# frozen_string_literal: true

require_relative "xlsx_matchers/version"
require_relative "xlsx_matchers/exact_match"
require_relative "xlsx_matchers/in_row"
require_relative "xlsx_matchers/in_column"
require_relative "xlsx_matchers/utils"
require_relative "xlsx_matchers/base_sheet"

require_relative "xlsx_matchers/sheets"
require_relative "xlsx_matchers/columns"
require_relative "xlsx_matchers/empty_row"
require_relative "xlsx_matchers/cells"
require_relative "xlsx_matchers/cell_value"
require_relative "xlsx_matchers/rows_with_data"

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
  # # RSpec::XlsxMatchers adds the following matchers to rspec
  #  - have_excel_sheets
  #  - have_excel_columns
  #  - have_excel_column
  #  - have_excel_empty_row
  #  - have_excel_cells
  #  - have_excel_cell_value
  #  - have_excel_rows_with_data
  module XlsxMatchers
    # class Error < StandardError; end

    def have_excel_sheets(sheet_names)
      Sheets.new(sheet_names)
    end

    def have_excel_columns(column_names)
      Columns.new(column_names)
    end

    def have_excel_column(column_name)
      raise ArgumentError, "Column name should not be an Array" if column_name.is_a?(Array)

      Columns.new([column_name])
    end

    def have_excel_empty_row(index)
      EmptyRow.new(index)
    end

    def have_excel_cells(cells)
      Cells.new(cells)
    end

    def have_excel_cell_value(value)
      CellValue.new(value)
    end

    def have_excel_rows_with_data(count)
      RowsWithData.new(count)
    end
  end
end
