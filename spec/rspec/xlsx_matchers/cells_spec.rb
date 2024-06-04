# frozen_string_literal: true

require "spec_helper"

RSpec.describe RSpec::XlsxMatchers::Cells do
  include_context "with simple example data"

  RSpec.shared_examples "have_excel_cells in_sheet" do |sheet|
    it "succeeds with proper data" do
      expect(subject).to have_excel_cells(raw_data[1]).in_row(1).in_sheet(sheet)
    end

    it "fails with missing sheet name" do
      expect do
        expect(subject).to have_excel_cells(raw_data[1]).in_row(1)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Sheet not provided"
      )
    end

    it "fails with invalid row" do
      expect do
        expect(subject).to have_excel_cells(raw_data[1]).in_row(2).in_sheet(sheet)
      end.to raise_error(RSpec::Expectations::ExpectationNotMetError)
    end
  end

  RSpec.shared_examples "have_excel_cells without sheet" do
    it "succeeds with proper data" do
      expect(subject).to have_excel_cells(raw_data[1]).in_row(1)
    end

    it "fails with invalid row" do
      expect do
        expect(subject).to have_excel_cells(raw_data[1]).in_row(22)
      end.to raise_error(RSpec::Expectations::ExpectationNotMetError)
    end
  end

  RSpec.shared_examples "have_excel_cells with sheet" do
    describe "by name" do
      include_examples "have_excel_cells in_sheet", "Sheet1"
    end

    describe "by index" do
      include_examples "have_excel_cells in_sheet", 0
    end
  end

  context "when providing a path" do
    subject(:subject) { file_path }

    it_behaves_like "have_excel_cells with sheet"
  end

  context "when providing a Roo::Excelx" do
    subject(:subject) { Roo::Spreadsheet.open(file_path) }

    it_behaves_like "have_excel_cells with sheet"
  end

  context "when providing a Roo::Excelx::Sheet" do
    subject(:subject) { Roo::Spreadsheet.open(file_path).sheet_for("Sheet1") }

    it_behaves_like "have_excel_cells without sheet"
  end

  context "when providing a Axlsx::Package" do
    subject(:subject) { caxlsx_data }

    it_behaves_like "have_excel_cells with sheet"
  end

  context "when providing a Axlsx::Workbook" do
    subject(:subject) { caxlsx_data.workbook }

    it_behaves_like "have_excel_cells with sheet"
  end

  context "when providing a Axlsx::Worksheet" do
    subject(:subject) { caxlsx_data.workbook.worksheets[0] }

    it_behaves_like "have_excel_cells without sheet"
  end
end
