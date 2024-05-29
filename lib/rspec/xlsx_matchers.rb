# frozen_string_literal: true

require_relative "xlsx_matchers/version"
require_relative "xlsx_matchers/sheets"

module Rspec
  module XlsxMatchers
    class Error < StandardError; end

    def have_excel_sheets(sheet_names)
      Sheets.new(sheet_names)
    end
  end
end
