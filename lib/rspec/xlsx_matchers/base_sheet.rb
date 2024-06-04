# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    class BaseSheet
      include Utils
      attr_reader :sheet_name, :sheet

      def initialize; end

      def matches?(subject)
        @sheet = find_sheet(subject, sheet_name)
        return false if sheet.nil?

        process_sheet
      end

      def in_sheet(name)
        @sheet_name = name
        self
      end

      private

      def find_sheet(subject, sheet_name = nil)
        find_axlsx_sheet(subject, sheet_name) || find_roo_sheet(subject, sheet_name)
      rescue RangeError, ArgumentError
        nil
      end

      def find_axlsx_sheet(subject, sheet_name = nil)
        return unless defined?(Axlsx)
        return subject if subject.is_a?(Axlsx::Worksheet)

        subject = subject.workbook if subject.is_a?(Axlsx::Package)

        return unless subject.is_a?(Axlsx::Workbook)

        axlsx_sheet_from_workbook(subject, sheet_name)
      end

      def axlsx_sheet_from_workbook(workbook, sheet_name = nil)
        if sheet_name.is_a?(String)
          workbook.sheet_by_name(sheet_name)
        elsif sheet_name.is_a?(Integer)
          workbook.worksheets[sheet_name]
        else
          raise ArgumentError, "Missing sheet name"
        end
      end

      def find_roo_sheet(subject, sheet_name = nil)
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

      def process_sheet
        if defined?(Axlsx) && sheet.is_a?(Axlsx::Worksheet)
          process_axlsx_sheet
        elsif defined?(Roo::Excelx) && sheet.is_a?(Roo::Excelx::Sheet)
          process_roo_sheet
        else
          raise "Unsupported worksheet type: #{worksheet.class}"
        end
      end

      def sheet_failure_message
        if sheet_name
          "Could not find sheet #{sheet_name}"
        else
          "Sheet not provided"
        end
      end

      def failure_message_in_sheet(msg)
        if sheet_name
          "#{msg} in sheet #{sheet_name}"
        else
          msg
        end
      end
    end
  end
end
