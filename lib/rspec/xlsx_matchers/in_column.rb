# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    # common base for columns matchers
    module InColumn
      attr_reader :column_name

      def in_column(column_name)
        @column_name = column_name
        self
      end

      private

      def column_index
        return @column_index unless @column_index.nil?

        return if sheet.nil? || column_row.nil?

        find_column
        @column_index
      end

      def column_row
        return @column_row unless @column_row.nil?

        @column_row = axlsx_column_row
        @column_row = roo_column_row if @column_row.nil?
        @column_row
      end

      def axlsx_column_row
        sheet.rows[0]&.map(&:value) if defined?(Axlsx::Worksheet) && sheet.is_a?(Axlsx::Worksheet)
      end

      def roo_column_row
        sheet.row(1) if defined?(Roo::Excelx) && sheet.is_a?(Roo::Excelx::Sheet)
      rescue StandardError
        # do nothing
      end

      def find_column
        column_row.each_with_index do |c, i|
          break unless @column_index.nil?

          @column_index = i if c.to_s == column_name
        end
      end

      def header_column_not_found?
        column_row.nil? || column_row.compact.empty?
      end

      def header_column_not_found_message
        "First row not found" if header_column_not_found?
      end

      def column_not_found?
        column_index.nil?
      end

      def column_not_found_message
        "Column #{column_name} not found"
      end
    end
  end
end
