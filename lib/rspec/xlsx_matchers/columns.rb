# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    class Columns
      attr_reader :extra_columns, :missing_columns, :expected_columns, :sheet_name, :sheet, :exact_match,
                  :actual_columns

      def initialize(expected_columns)
        @expected_columns = expected_columns
        @missing_columns = []
        @extra_columns = []
        @exact_match = false
        @actual_columns = []
      end

      def matches?(subject)
        @sheet = find_sheet(subject)
        return false if sheet.nil?

        if defined?(Axlsx) && sheet.is_a?(Axlsx::Worksheet)
          match_axlsx_columns(sheet)
        elsif defined?(Roo::Excelx) && sheet.is_a?(Roo::Excelx::Sheet)
          match_roo_columns(sheet)
        else
          raise "Unsupported worksheet type: #{worksheet.class}"
        end
      end

      def in_sheet(name)
        @sheet_name = name
        self
      end

      def exactly
        @exact_match = true
        self
      end

      def failure_message
        return sheet_failure_message if sheet.nil?

        msg = if sheet_name
                "Columns mismatch in sheet #{sheet_name}"
              else
                "Columns mismatch"
              end

        "#{msg}:\n\t#{column_failure_message}"
      end

      private

      def find_sheet(subject)
        find_axlsx_sheet(subject) || find_roo_sheet(subject)
      rescue RangeError, ArgumentError
        nil
      end

      def find_axlsx_sheet(subject)
        return unless defined?(Axlsx)

        if subject.is_a?(Axlsx::Workbook)
          axlsx_sheet_from_workbook(subject)
        elsif subject.is_a?(Axlsx::Worksheet)
          subject
        end
      end

      def axlsx_sheet_from_workbook(worbook)
        if sheet_name.is_a?(String)
          worbook.sheet_by_name(sheet_name)
        elsif sheet_name.is_a?(Integer)
          worbook.worksheets(sheet_name)
        else
          raise ArgumentError, "Missing sheet name"
        end
      end

      def find_roo_sheet(subject)
        return unless defined?(Roo::Excelx)
        return subject if subject.is_a?(Roo::Excelx::Sheet)

        spreadsheet = if subject.is_a?(Roo::Excelx)
                        subject
                      elsif subject.is_a?(String)
                        Roo::Spreadsheet.open(subject)
                      end

        return unless spreadsheet
        return unless sheet_name

        spreadsheet.sheet_for(sheet_name)
      end

      def match_axlsx_columns(_worksheet)
        raise "Matcher not implemented yet"
      end

      def match_roo_columns(worksheet)
        @actual_columns = worksheet.row(1)
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
          missing_columns << column unless actual_columns.include?(column)
        end
        missing_columns.empty?
      end

      def sheet_failure_message
        if sheet_name
          "Could not find sheet #{sheet_name}"
        else
          "Sheet not provided"
        end
      end

      def column_failure_message
        column_mismatch_failure_message || column_order_failure_message
      end

      def column_order_failure_message
        return unless exact_match

        "Expected: #{expected_columns.map { |c| "'#{c}'" }.join(", ")}\n\t"\
        "Received: #{actual_columns.map { |c| "'#{c}'" }.join(", ")}"
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

        "Missing columns: #{(missing_columns.map { |c| "'#{c}'" }).join(", ")}"
      end

      def extra_columns_failure_message
        return if extra_columns.empty?

        "Extra columns: #{(extra_columns.map { |c| "'#{c}'" }).join(", ")}"
      end
    end
  end
end
