# frozen_string_literal: true

require "spec_helper"
require "roo"

RSpec.describe Rspec::XlsxMatchers::Sheets do
  subject(:matcher) { described_class.new(expected_sheet_names) }

  let(:file_path) { fixture_file_path "file_example_XLSX_10.xlsx" }

  RSpec.shared_examples "sheet matcher" do
    context "with sheet name" do
      let(:expected_sheet_names) { ["Sheet1"] }

      it "is successful" do
        expect(matcher.match(matcher_argument)).to be(true)
      end
    end

    context "with invalid sheet name" do
      let(:expected_sheet_names) { ["Test"] }

      it "is successful" do
        expect(matcher.match(matcher_argument)).to be(false)
      end

      it "has_expected_message" do
        matcher.match(matcher_argument)
        expect(matcher.failure_message).to eq("Xlsx file sheets not found: [Test]")
      end
    end

    context "with index" do
      let(:expected_sheet_names) { [0] }

      it "is successful" do
        expect(matcher.match(matcher_argument)).to be(true)
      end
    end

    context "with invalid index" do
      let(:expected_sheet_names) { [3] }

      it "is successful" do
        expect(matcher.match(matcher_argument)).to be(false)
      end

      it "has_expected_message" do
        matcher.match(matcher_argument)
        expect(matcher.failure_message).to eq("Xlsx file sheets not found: [3]")
      end
    end
  end

  context "when providing a path" do
    let(:matcher_argument) { file_path }

    it_behaves_like "sheet matcher"
  end

  context "when providing a Roo::Excelx file" do
    let(:matcher_argument) { Roo::Spreadsheet.open(file_path) }

    it_behaves_like "sheet matcher"
  end

  context "when providing a Axlsx::Package" do
    include_context "with a simple caxlsx instance"
    let(:matcher_argument) { caxlsx_data }

    it_behaves_like "sheet matcher"
  end

  context "when providing a Axlsx::Workbook" do
    include_context "with a simple caxlsx instance"
    let(:matcher_argument) { caxlsx_data.workbook }

    it_behaves_like "sheet matcher"
  end
end
