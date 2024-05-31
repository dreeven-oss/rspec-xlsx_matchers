# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    class Sheets
      attr_reader :errors, :expected_sheet_names

      def initialize(expected_sheet_names)
        @expected_sheet_names = if expected_sheet_names.is_a?(Array)
                                  expected_sheet_names
                                else
                                  [expected_sheet_names]
                                end
        @errors = []
      end

      def matches?(subject)
        if subject.is_a?(String)
          match_string(subject)
        elsif defined?(Roo::Excelx) && subject.is_a?(Roo::Excelx)
          match_roo_excelx(subject)
        elsif defined?(Axlsx)
          matches_axlsx?(subject)
        else
          invalid_file
        end
      end

      def failure_message
        "Xlsx file sheets not found: #{errors.map { |s| "'#{s}'" }.join(",")}"
      end

      private

      def matches_axlsx?(subject)
        if subject.is_a?(Axlsx::Package)
          match_caxlsx(subject.workbook)
        elsif defined?(Axlsx) && subject.is_a?(Axlsx::Workbook)
          match_caxlsx(subject)
        else
          invalid_file
        end
      end

      def invalid_file
        raise ArgumentError, "Could not evaluate the sheets existance, "\
            "the matcher expected an instance of Roo::Excelx, "\
            "Axlsx::Package or Axlsx::Workbook, got #{excel_file.class}."
      end

      def match_caxlsx(excel_file)
        return invalid_file unless defined?(Axlsx)

        workbook = if excel_file.is_a?(Axlsx::Package)
                     excel_file.workbook
                   elsif excel_file.is_a?(Axlsx::Workbook)
                     excel_file
                   end

        return invalid_file if workbook.nil?

        match_caxlsx_workbook(workbook)
      end

      def match_caxlsx_workbook(workbook)
        expected_sheet_names.each do |expected_sheet_name|
          found = if expected_sheet_name.is_a?(String)
                    workbook.sheet_by_name(expected_sheet_name)
                  elsif expected_sheet_name.is_a?(Integer)
                    workbook.worksheets[expected_sheet_name]
                  end

          errors << expected_sheet_name unless found.is_a?(Axlsx::Worksheet)
        end
        errors.empty?
      end

      def match_string(excel_file)
        if defined?(Roo::Spreadsheet)
          roo_file = Roo::Spreadsheet.open(excel_file)
          match_roo_excelx(roo_file)
        else
          raise ArgumentError, "Could not evaluate the sheets existance, "\
          "the matcher received a string, but Roo::Spreadsheet is not defined"
        end
      end

      def match_roo_excelx(excel_file)
        expected_sheet_names.each do |expected_sheet_name|
          excel_file.sheet(expected_sheet_name)
        rescue ArgumentError, RangeError
          errors << expected_sheet_name
        rescue TypeError => e
          errors << "#{expected_sheet_name} : #{e.message}"
        end
        errors.empty?
      end
    end
  end
end
