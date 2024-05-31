# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    class Rows < BaseSheet
      attr_reader :expected_cells, :actual_cells

      def initialize(expected_cells)
        super()
        @expected_cells = force_array(expected_cells)
        @actual_cells = []
      end

      private

      def process_axlsx_sheet(worksheet)
        @actual_columns = worksheet.rows[0]&.map(&:value)
        if exact_match
          perform_exact_match
        else
          perform_loose_match
        end
      end

      def process_roo_sheet(worksheet)
        @actual_columns = worksheet.row(1)
        if exact_match
          perform_exact_match
        else
          perform_loose_match
        end
      end
    end
  end
end
