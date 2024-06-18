# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    ## have_excel_cell_value
    class CellValue < BaseSheet
      include InRow
      include InColumn
      attr_reader :expected_value, :actual_value

      def initialize(expected_value)
        super()
        @expected_value = expected_value
      end

      def failure_message
        return sheet_failure_message if sheet.nil?
        return header_column_not_found_message if header_column_not_found?
        return row_not_found_message if row_not_found?
        return column_not_found_message if column_not_found?

        "Mismatch cell value in column #{failure_message_in_column} of row '#{row_index}': " \
          "expected: '#{expected_value}', received: '#{actual_value}'"
      end

      def failure_message_in_column
        if column_name.nil?
          "column with index #{column_index}"
        else
          "column '#{column_name}'"
        end
      end

      private

      def process_axlsx_sheet
        return false if header_column_not_found?
        return false if row.nil?
        return false if column_index.nil?

        @actual_value = row[column_index]&.value
        perform_match
      end

      def process_roo_sheet
        return false if header_column_not_found?
        return false if row.nil?
        return false if column_index.nil?

        @actual_value = row[column_index]
        perform_match
      end

      def perform_match
        convert_to_excel_value(expected_value) == actual_value
      end

      def convert_to_excel_value(value)
        case value
        when TrueClass
          "1"
        when FalseClass
          "0"
        else
          value
        end
      end
    end
  end
end
