# frozen_string_literal: true

require "spec_helper"

RSpec.describe RSpec::XlsxMatchers::FilledRows do
  include_context "with simple example data"

  RSpec.shared_examples "filled_rows matcher without sheet" do
    it "succeeds if checking for 10 rows" do
      expect(subject).to have_excel_filled_rows(10)
    end

    it "fails if checking for more than 10 rows" do
      expect do
        expect(subject).to have_excel_filled_rows(11)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Expected 11 filled rows but found 10"
      )
    end

    it "fails if checking for less than 10 rows" do
      expect do
        expect(subject).to have_excel_filled_rows(9)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Expected 9 filled rows but found 10"
      )
    end

    describe "negated" do
      it "succeeds if puting another count number" do
        expect(subject).not_to have_excel_filled_rows(999)
      end

      it "fails when matching the count number" do
        expect do
          expect(subject).not_to have_excel_filled_rows(10)
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Expected NOT to find 10 filled rows but did"
        )
      end
    end
  end

  RSpec.shared_examples "filled_rows matcher in sheet" do |sheet_name|
    it "fails without the sheet" do
      expect do
        expect(subject).to have_excel_filled_rows(1)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Sheet not provided"
      )
    end

    it "fails with an invalid sheet" do
      invalid_sheet_name = if sheet_name.is_a?(String)
                             "Invalid"
                           else
                             8
                           end
      expect do
        expect(subject).to have_excel_filled_rows(1).in_sheet(invalid_sheet_name)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Could not find sheet #{invalid_sheet_name}"
      )
    end

    it "fails when checking for more than 10 rows" do
      expect do
        expect(subject).to have_excel_filled_rows(11).in_sheet(sheet_name)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Expected 11 filled rows but found 10"
      )
    end

    it "fails when checking for less than 10 rows" do
      expect do
        expect(subject).to have_excel_filled_rows(9).in_sheet(sheet_name)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Expected 9 filled rows but found 10"
      )
    end

    it "succeeds if checking for 10 rows" do
      expect(subject).to have_excel_filled_rows(10).in_sheet(sheet_name)
    end

    describe "negated" do
      it "succeeds if puting another count number" do
        expect(subject).not_to have_excel_filled_rows(999).in_sheet(sheet_name)
      end

      it "fails when matching the count number" do
        expect do
          expect(subject).not_to have_excel_filled_rows(10).in_sheet(sheet_name)
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Expected NOT to find 10 filled rows but did"
        )
      end
    end
  end

  RSpec.shared_examples "filled_rows matcher with sheet" do
    describe "by name" do
      include_examples "filled_rows matcher in sheet", "Sheet1"
    end

    describe "by index" do
      include_examples "filled_rows matcher in sheet", 0
    end
  end

  context "when providing a path" do
    subject(:subject) { file_path }

    it_behaves_like "filled_rows matcher with sheet"
  end

  context "when providing a File" do
    subject(:subject) { File.open(file_path) }

    it_behaves_like "filled_rows matcher with sheet"
  end

  context "when providing a Roo::Excelx" do
    subject(:subject) { Roo::Spreadsheet.open(file_path) }

    it_behaves_like "filled_rows matcher with sheet"
  end

  context "when providing a Roo::Excelx::Sheet" do
    subject(:subject) { Roo::Spreadsheet.open(file_path).sheet_for("Sheet1") }

    it_behaves_like "filled_rows matcher without sheet"
  end

  context "when providing a Axlsx::Package" do
    subject(:subject) { caxlsx_data }

    it_behaves_like "filled_rows matcher with sheet"
  end

  context "when providing a Axlsx::Workbook" do
    subject(:subject) { caxlsx_data.workbook }

    it_behaves_like "filled_rows matcher with sheet"
  end

  context "when providing a Axlsx::Worksheet" do
    subject(:subject) { caxlsx_data.workbook.worksheets[0] }

    it_behaves_like "filled_rows matcher without sheet"
  end
end
