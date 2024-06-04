# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    # common base row based matchers
    module InRow
      attr_reader :row_index

      def in_row(row_index)
        @row_index = row_index
        self
      end

      private

      def row
        return if sheet.nil?
        return if row_index.nil?

        @row ||= find_row
      end

      def find_row
        if defined?(Axlsx::Worksheet) && sheet.is_a?(Axlsx::Worksheet)
          sheet.rows[row_index]
        elsif defined?(Roo::Excelx) && sheet.is_a?(Roo::Excelx::Sheet)
          sheet.row(row_index + 1)
        end
      end

      def row_not_found_message
        "Row #{row_index} is empty" if row_not_found?
      end

      def row_not_found?
        row.nil? || row.compact.empty?
      end
    end
  end
end
