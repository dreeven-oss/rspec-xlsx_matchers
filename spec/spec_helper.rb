# frozen_string_literal: true

require "rspec/xlsx_matchers"

RSPEC_ROOT = File.dirname __FILE__

RSpec.configure do |config|
  config.include RSpec::XlsxMatchers
  # Enable flags like --only-failures and --next-failure
  config.example_status_persistence_file_path = ".rspec_status"

  # Disable RSpec exposing methods globally on `Module` and `main`
  config.disable_monkey_patching!

  config.expect_with :rspec do |c|
    c.syntax = :expect
  end
end

Dir[("#{RSPEC_ROOT}/support/**/*.rb")].each { |file| require file }

def fixture_file_path(name)
  "#{RSPEC_ROOT}/fixtures/files/#{name}"
end
