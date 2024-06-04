# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    class Cells < BaseSheet
      include InRow
      attr_reader :expected_cells, :actual_cells, :mismatch_indexes

      EXPECTED = "Expected"
      RECEIVED = "Received"

      def initialize(expected_cells)
        super()
        @expected_cells = force_array(expected_cells)
        @actual_cells = []
        @mismatch_indexes = []
      end

      def failure_message
        return sheet_failure_message if sheet.nil?
        return row_not_found_message if row.nil?

        message = failure_message_header(String.new("Rows did not match : \n"))
        # Cells content
        expected_cells.each_with_index do |expected_cell, idx|
          sizeof_actual_cell = actual_cells[idx].to_s.size
          sizeof_expected_cell = expected_cell.to_s.size
          message << " #{idx}"
          message << (" " * (4 - idx.to_s.size))
          message << "| #{actual_cells[idx]}"
          message << (" " * (biggest_actual_cell_size - sizeof_actual_cell))
          message << " | #{expected_cell}"
          message << (" " * (biggest_expected_cell_size - sizeof_expected_cell))
          message << " | #{mismatch_indexes.include?(idx) ? "<----- Mismatch" : ""}"
          message << "\n"
        end
        message
      end

      private

      def process_axlsx_sheet
        return false if row.nil?

        @actual_cells = row.cells.map(&:value)
        perform_match
      end

      def process_roo_sheet
        return false if row.nil?

        @actual_cells = row
        perform_match
      end

      def perform_match
        expected_cells.each_with_index do |expected_cell, idx|
          actual_cell = actual_cells[idx]
          mismatch_indexes << idx if actual_cell.to_s != expected_cell.to_s
        end
        mismatch_indexes.empty?
      end

      def biggest_actual_cell_size
        @biggest_actual_cell_size ||= (actual_cells.map { |c| c.to_s.size } + [RECEIVED.size]).max
      end

      def biggest_expected_cell_size
        @biggest_expected_cell_size ||= (expected_cells.map { |c| c.to_s.size } + [EXPECTED.size]).max
      end

      def failure_message_header(message)
        # Table Header
        message << "     | #{RECEIVED}"
        message << (" " * (biggest_actual_cell_size - RECEIVED.size))
        message << " | #{EXPECTED}"
        message << (" " * (biggest_expected_cell_size - EXPECTED.size))
        message << " |\n"
        # Header / Content separator
        message << ("#{"-" * 5}|#{"-" * (2 + biggest_actual_cell_size)}|#{"-" * (2 + @biggest_expected_cell_size)}|\n")
        message
      end
    end
  end
end
