# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    class EmptyRow < BaseSheet
      attr_reader :row_index

      def initialize(row_index)
        super()
        @row_index = row_index
      end

      def failure_message_when_negated
        return sheet_failure_message if sheet.nil?

        "Row at index '#{row_index}' was expected to NOT be empty, but was empty."
      end

      def failure_message
        return sheet_failure_message if sheet.nil?

        "Row at index '#{row_index}' was expected to be empty, but was not empty."
      end

      private

      def process_axlsx_sheet(worksheet)
        row = worksheet.rows[row_index]
        true if row.nil?
      end

      def process_roo_sheet(worksheet)
        row = worksheet.row(row_index).compact
        true if row.empty?
      end
    end
  end
end
