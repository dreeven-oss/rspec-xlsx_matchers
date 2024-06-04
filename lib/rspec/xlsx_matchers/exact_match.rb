# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    # common base for matchers with exact match option
    module ExactMatch
      attr_reader :exact_match

      def exactly
        @exact_match = true
        self
      end
    end
  end
end
