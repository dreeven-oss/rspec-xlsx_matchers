# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    # Base class for sheet based machers
    class BaseSheet
      include Utils
      attr_reader :sheet_name, :sheet, :subject

      def matches?(subject)
        @subject = subject
        @sheet = find_sheet
        return false if sheet.nil?

        process_sheet.tap do
          finalize
        end
      end

      def in_sheet(name)
        @sheet_name = name
        self
      end

      private

      def find_sheet
        find_axlsx_sheet || find_roo_sheet
      rescue RangeError, ArgumentError
        nil
      end

      def find_axlsx_sheet
        return unless defined?(Axlsx)
        return subject if subject.is_a?(Axlsx::Worksheet)

        workbook = if subject.is_a?(Axlsx::Package)
                     subject.workbook
                   else
                     subject
                   end

        return unless workbook.is_a?(Axlsx::Workbook)

        axlsx_sheet_from_workbook(workbook)
      end

      def axlsx_sheet_from_workbook(workbook)
        if sheet_name.is_a?(String)
          workbook.sheet_by_name(sheet_name)
        elsif sheet_name.is_a?(Integer)
          workbook.worksheets[sheet_name]
        else
          raise ArgumentError, "Missing sheet name"
        end
      end

      def find_roo_sheet
        return unless defined?(Roo::Excelx)
        return subject if subject.is_a?(Roo::Excelx::Sheet)

        return unless roo_spreadsheet
        return unless sheet_name

        roo_spreadsheet.sheet_for(sheet_name)
      end

      def roo_spreadsheet
        return @roo_spreadsheet unless @roo_spreadsheet.nil?

        @roo_spreadsheet = if subject.is_a?(Roo::Excelx)
                             subject
                           elsif subject.is_a?(String) || subject.is_a?(File)
                             Roo::Spreadsheet.open(subject)
                           end
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

      def finalize
        roo_spreadsheet&.close
      rescue StandardError => e
        puts "Warning: error closing Roo Spreadsheet: #{e}"
      end
    end
  end
end
