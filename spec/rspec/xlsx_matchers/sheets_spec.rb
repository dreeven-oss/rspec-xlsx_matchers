# frozen_string_literal: true

require "spec_helper"

RSpec.describe RSpec::XlsxMatchers::Sheets do
  let(:file_path) { fixture_file_path "file_example_XLSX_10.xlsx" }

  RSpec.shared_examples "sheet matcher" do
    it "finds sheet by name" do
      expect(subject).to have_excel_sheets("Sheet1")
    end

    it "finds sheet by name with array" do
      expect(subject).to have_excel_sheets(["Sheet1"])
    end

    it "fails with invalid name" do
      expect do
        expect(subject).to have_excel_sheets("Test")
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Xlsx file sheets not found: 'Test'"
      )
    end

    it "fails with invalid name with array" do
      expect do
        expect(subject).to have_excel_sheets(["Test"])
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Xlsx file sheets not found: 'Test'"
      )
    end

    it "finds sheet by index" do
      expect(subject).to have_excel_sheets(0)
    end

    it "finds sheet by index with array" do
      expect(subject).to have_excel_sheets([0])
    end

    it "fails with invalid index" do
      expect do
        expect(subject).to have_excel_sheets(1)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Xlsx file sheets not found: '1'"
      )
    end

    it "fails with invalid index with array" do
      expect do
        expect(subject).to have_excel_sheets([1])
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Xlsx file sheets not found: '1'"
      )
    end
  end

  context "when providing a path" do
    let(:subject) { file_path }

    it_behaves_like "sheet matcher"
  end

  context "when providing a Roo::Excelx file" do
    let(:subject) { Roo::Spreadsheet.open(file_path) }

    it_behaves_like "sheet matcher"
  end

  context "when providing a Axlsx::Package" do
    include_context "with a simple caxlsx instance"
    let(:subject) { caxlsx_data }

    it_behaves_like "sheet matcher"
  end

  context "when providing a Axlsx::Workbook" do
    include_context "with a simple caxlsx instance"
    let(:subject) { caxlsx_data.workbook }

    it_behaves_like "sheet matcher"
  end
end
