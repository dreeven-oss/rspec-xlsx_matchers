# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    # have_excel_columns / have_excel_column
    class Columns < BaseSheet
      include ExactMatch
      attr_reader :extra_columns, :missing_columns, :expected_columns, :actual_columns

      def initialize(expected_columns)
        super()
        @expected_columns = force_array(expected_columns)
        @missing_columns = []
        @extra_columns = []
        @actual_columns = []
      end

      def failure_message
        return sheet_failure_message if sheet.nil?

        msg = failure_message_in_sheet("Columns mismatch")

        "#{msg}:\n\t#{column_failure_message}"
      end

      private

      def process_axlsx_sheet
        @actual_columns = sheet.rows[0]&.map(&:value)
        if exact_match
          perform_exact_match
        else
          perform_loose_match
        end
      end

      def process_roo_sheet
        @actual_columns = sheet.row(1)
        if exact_match
          perform_exact_match
        else
          perform_loose_match
        end
      end

      def perform_exact_match
        return true if expected_columns == actual_columns

        (expected_columns - actual_columns).each do |column|
          extra_columns << column
        end

        (actual_columns - expected_columns).each do |column|
          missing_columns << column
        end
        false
      end

      def perform_loose_match
        expected_columns.each do |column|
          missing_columns << column unless actual_columns&.include?(column)
        end
        missing_columns.empty?
      end

      def column_failure_message
        column_mismatch_failure_message || column_order_failure_message
      end

      def column_order_failure_message
        return unless exact_match

        "Expected: #{map_output(expected_columns)}\n\t" \
          "Received: #{map_output(actual_columns)}"
      end

      def column_mismatch_failure_message
        msg = [missing_column_failure_message, extra_columns_failure_message].compact.join("\n\t")
        if msg.empty?
          nil
        else
          msg
        end
      end

      def missing_column_failure_message
        return if missing_columns.empty?

        "Missing columns: #{map_output(missing_columns)}"
      end

      def extra_columns_failure_message
        return if extra_columns.empty?

        "Extra columns: #{map_output(extra_columns)}"
      end
    end
  end
end
