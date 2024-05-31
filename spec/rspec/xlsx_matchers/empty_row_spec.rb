# frozen_string_literal: true

require "spec_helper"

RSpec.describe RSpec::XlsxMatchers::EmptyRow do
  let(:file_path) { fixture_file_path "file_example_XLSX_10.xlsx" }

  RSpec.shared_examples "empty_row matcher without sheet" do
    it "fails when the row is not empty" do
      expect do
        expect(subject).to have_excel_empty_row(1)
      end.to raise_error(RSpec::Expectations::ExpectationNotMetError, "Row at index '1' was expected to be empty, but was not empty.")
    end

    it "succeeds if the row is empty" do
      expect(subject).to have_excel_empty_row(11)
    end

    it "succeeds when we pass a very large number clearly out of the file" do
      expect(subject).to have_excel_empty_row(99_999)
    end

    describe "negated" do
      it "succeeds the row is present" do
        expect(subject).not_to have_excel_empty_row(1)
      end

      it "fails when the row is not present" do
        expect do
          expect(subject).not_to have_excel_empty_row(999)
        end.to raise_error(RSpec::Expectations::ExpectationNotMetError, "Row at index '999' was expected to NOT be empty, but was empty.")
      end
    end
  end

  RSpec.shared_examples "empty_row matcher in sheet" do |sheet_name|
    it "fails without the sheet" do
      expect do
        expect(subject).to have_excel_empty_row(1)
      end.to raise_error(RSpec::Expectations::ExpectationNotMetError, "Sheet not provided")
    end

    it "fails with an invalid sheet" do
      invalid_sheet_name = if sheet_name.is_a?(String)
                             "Invalid"
                           else
                             8
                           end
      expect do
        expect(subject).to have_excel_empty_row(3).in_sheet(invalid_sheet_name)
      end.to raise_error(RSpec::Expectations::ExpectationNotMetError, "Could not find sheet #{invalid_sheet_name}")
    end

    it "fails when the row is not empty" do
      expect do
        expect(subject).to have_excel_empty_row(1).in_sheet(sheet_name)
      end.to raise_error(RSpec::Expectations::ExpectationNotMetError, "Row at index '1' was expected to be empty, but was not empty.")
    end

    it "succeeds if the row is empty" do
      expect(subject).to have_excel_empty_row(11).in_sheet(sheet_name)
    end

    it "succeeds when we pass a very large number clearly out of the file" do
      expect(subject).to have_excel_empty_row(99_999).in_sheet(sheet_name)
    end

    describe "negated" do
      it "succeeds the row is present" do
        expect(subject).not_to have_excel_empty_row(1).in_sheet(sheet_name)
      end

      it "fails when the row is not present" do
        expect do
          expect(subject).not_to have_excel_empty_row(999).in_sheet(sheet_name)
        end.to raise_error(RSpec::Expectations::ExpectationNotMetError, "Row at index '999' was expected to NOT be empty, but was empty.")
      end
    end
  end

  RSpec.shared_examples "empty_row matcher with sheet" do
    describe "by name" do
      include_examples "empty_row matcher in sheet", "Sheet1"
    end

    describe "by index" do
      include_examples "empty_row matcher in sheet", 0
    end
  end

  context "when providing a path" do
    subject(:subject) { file_path }

    it_behaves_like "empty_row matcher with sheet"
  end

  context "when providing a Roo::Excelx" do
    subject(:subject) { Roo::Spreadsheet.open(file_path) }

    it_behaves_like "empty_row matcher with sheet"
  end

  context "when providing a Roo::Excelx::Sheet" do
    subject(:subject) { Roo::Spreadsheet.open(file_path).sheet_for("Sheet1") }

    it_behaves_like "empty_row matcher without sheet"
  end

  context "when providing a Axlsx::Package" do
    subject(:subject) { caxlsx_data }

    include_context "with a simple caxlsx instance"

    it_behaves_like "empty_row matcher with sheet"
  end

  context "when providing a Axlsx::Workbook" do
    subject(:subject) { caxlsx_data.workbook }

    include_context "with a simple caxlsx instance"

    it_behaves_like "empty_row matcher with sheet"
  end

  context "when providing a Axlsx::Worksheet" do
    subject(:subject) { caxlsx_data.workbook.worksheets[0] }

    include_context "with a simple caxlsx instance"

    it_behaves_like "empty_row matcher without sheet"
  end
end
