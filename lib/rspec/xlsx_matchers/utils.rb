# frozen_string_literal: true

module RSpec
  module XlsxMatchers
    # internal utilities
    module Utils
      def force_array(value)
        if value.is_a?(Array)
          value
        else
          [value]
        end
      end

      def map_output(array)
        array.map do |v|
          if v.is_a?(String)
            "'#{v}'"
          else
            v
          end
        end.join(", ")
      end
    end
  end
end
