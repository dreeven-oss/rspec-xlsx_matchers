# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    # filled rows count matcher
    class FilledRows < BaseSheet
      attr_reader :expected_filled_row_count

      def initialize(count)
        super()
        @expected_filled_row_count = count
      end

      def matches?(subject)
        @subject = subject
        @sheet = find_sheet
        return false if sheet.nil?

        actual_filled_row_count == expected_filled_row_count
      end

      def failure_message
        return sheet_failure_message if sheet.nil?

        "Expected #{expected_filled_row_count} filled rows but found #{actual_filled_row_count}"
      end

      def failure_message_when_negated
        return sheet_failure_message if sheet.nil?

        "Expected NOT to find #{expected_filled_row_count} filled rows but did"
      end

      private

      def actual_filled_row_count
        @actual_filled_row_count ||= process_sheet
      end

      def process_axlsx_sheet
        sheet.rows.compact.count { |row| !row.nil? && !row.empty? }
      end

      def process_roo_sheet
        (1..sheet.last_row).count { |index| !sheet.row(index).compact.empty? }
      end
    end
  end
end
