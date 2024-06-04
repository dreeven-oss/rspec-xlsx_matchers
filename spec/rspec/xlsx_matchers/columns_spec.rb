# frozen_string_literal: true

require "spec_helper"

RSpec.describe RSpec::XlsxMatchers::Columns do
  include_context "with simple example data"
  let(:expected_columns) { raw_data[0] }
  let(:expected_column) { raw_data[0][2] }

  RSpec.shared_examples "columns matcher in sheet" do
    it "succeeds with all columns in sheet by name" do
      expect(subject).to have_excel_columns(expected_columns).in_sheet("Sheet1")
    end

    it "succeeds with all columns in sheet by index" do
      expect(subject).to have_excel_columns(expected_columns).in_sheet(0)
    end

    it "succeeds with some columns in sheet by name" do
      expect(subject).to have_excel_columns(["Gender", "First Name"]).in_sheet("Sheet1")
    end

    it "fails if sheet is absent" do
      expect do
        expect(subject).to have_excel_columns(expected_columns)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Sheet not provided"
      )
    end

    it "fails if sheet name is invalid" do
      expect do
        expect(subject).to have_excel_columns(expected_columns).in_sheet("Not A Sheet")
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Could not find sheet Not A Sheet"
      )
    end

    it "fails if sheet index is invalid" do
      expect do
        expect(subject).to have_excel_columns(expected_columns).in_sheet(2)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Could not find sheet 2"
      )
    end

    it "fails if one column is not found" do
      expect do
        expect(subject).to have_excel_columns(expected_columns + ["SSN"]).in_sheet(0)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Columns mismatch in sheet 0:\n\tMissing columns: 'SSN'"
      )
    end

    it "fails if many columns are not found" do
      expect do
        expect(subject).to have_excel_columns(expected_columns + %w[SSN State]).in_sheet(0)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Columns mismatch in sheet 0:\n\tMissing columns: 'SSN', 'State'"
      )
    end

    describe "exactly" do
      it "succeeds with all columns in sheet by name" do
        expect(subject).to have_excel_columns(expected_columns).in_sheet("Sheet1").exactly
      end

      it "fails when extra columns are found" do
        expect do
          expect(subject).to have_excel_columns(["First Name",	"Last Name",	"Gender"]).in_sheet("Sheet1").exactly
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Columns mismatch in sheet Sheet1:\n\tMissing columns: 0, 'Country', 'Age', 'Date', 'Id'"
        )
      end

      it "fails when columns are missing" do
        expect do
          expect(subject).to have_excel_columns(expected_columns + %w[SSN State])
            .in_sheet("Sheet1").exactly
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Columns mismatch in sheet Sheet1:\n\tExtra columns: 'SSN', 'State'"
        )
      end

      it "fails when columns are missing and there are extras" do
        expect do
          expect(subject).to have_excel_columns(expected_columns - ["First Name"] + %w[SSN State])
            .in_sheet("Sheet1").exactly
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Columns mismatch in sheet Sheet1:\n\tMissing columns: 'First Name'\n\tExtra columns: 'SSN', 'State'"
        )
      end

      it "fails when columns are not in the same order" do
        expected = expected_columns - ["First Name"] + ["First Name"]
        expect do
          expect(subject).to have_excel_columns(expected)
            .in_sheet("Sheet1").exactly
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Columns mismatch in sheet Sheet1:\n\t" \
          "Expected: 0, 'Last Name', 'Gender', 'Country', 'Age', 'Date', 'Id', 'First Name'\n\t" \
          "Received: 0, 'First Name', 'Last Name', 'Gender', 'Country', 'Age', 'Date', 'Id'"
        )
      end
    end
  end

  RSpec.shared_examples "column matcher in sheet" do
    it "succeeds with one column in sheet by name" do
      expect(subject).to have_excel_column(expected_column).in_sheet("Sheet1")
    end

    it "succeeds with one column (integer) in sheet by name" do
      expect(subject).to have_excel_column(0).in_sheet("Sheet1")
    end

    it "fails with an array of colums in sheet by name" do
      expect do
        expect(subject).to have_excel_column(expected_columns).in_sheet("Sheet1")
      end.to raise_error(
        ArgumentError,
        "Column name should not be an Array"
      )
    end

    it "succeeds with one column in sheet by index" do
      expect(subject).to have_excel_columns(expected_column).in_sheet(0)
    end

    it "fails if sheet is absent" do
      expect do
        expect(subject).to have_excel_column(expected_column)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Sheet not provided"
      )
    end

    it "fails if sheet name is invalid" do
      expect do
        expect(subject).to have_excel_columns(expected_column).in_sheet("Not A Sheet")
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Could not find sheet Not A Sheet"
      )
    end

    it "fails if sheet index is invalid" do
      expect do
        expect(subject).to have_excel_columns(expected_column).in_sheet(2)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Could not find sheet 2"
      )
    end

    it "fails if the column is not found" do
      expect do
        expect(subject).to have_excel_columns("SSN").in_sheet("Sheet1")
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Columns mismatch in sheet Sheet1:\n\tMissing columns: 'SSN'"
      )
    end
  end

  RSpec.shared_examples "column matcher without sheet" do
    it "succeeds with one column by name" do
      expect(subject).to have_excel_column("Gender")
    end

    it "succeeds with one column by name (integer)" do
      expect(subject).to have_excel_column(0)
    end

    it "fails if one column is not found" do
      expect do
        expect(subject).to have_excel_column("SSN")
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Columns mismatch:\n\tMissing columns: 'SSN'"
      )
    end
  end

  RSpec.shared_examples "columns matcher without sheet" do
    it "succeeds with some columns" do
      expect(subject).to have_excel_columns(["Gender", "First Name"])
    end

    it "succeeds with all columns" do
      expect(subject).to have_excel_columns(expected_columns)
    end

    it "fails if one column is not found" do
      expect do
        expect(subject).to have_excel_columns(expected_columns + ["SSN"])
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Columns mismatch:\n\tMissing columns: 'SSN'"
      )
    end

    it "fails if many columns are not found" do
      expect do
        expect(subject).to have_excel_columns(expected_columns + %w[SSN State])
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Columns mismatch:\n\tMissing columns: 'SSN', 'State'"
      )
    end

    describe "exactly" do
      it "succeeds with all columns in sheet by name" do
        expect(subject).to have_excel_columns(expected_columns).exactly
      end

      it "fails when extra columns are found" do
        expect do
          expect(subject).to have_excel_columns(["First Name",	"Last Name",	"Gender"]).exactly
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Columns mismatch:\n\tMissing columns: 0, 'Country', 'Age', 'Date', 'Id'"
        )
      end

      it "fails when columns are missing" do
        expect do
          expect(subject).to have_excel_columns(expected_columns + %w[SSN State])
            .exactly
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Columns mismatch:\n\tExtra columns: 'SSN', 'State'"
        )
      end

      it "fails when columns are missing and there are extras" do
        expect do
          expect(subject).to have_excel_columns(expected_columns - ["First Name"] + %w[SSN State])
            .exactly
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Columns mismatch:\n\tMissing columns: 'First Name'\n\tExtra columns: 'SSN', 'State'"
        )
      end

      it "fails when columns are not in the same order" do
        expected = expected_columns - ["First Name"] + ["First Name"]
        expect do
          expect(subject).to have_excel_columns(expected)
            .exactly
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Columns mismatch:\n\t" \
          "Expected: 0, 'Last Name', 'Gender', 'Country', 'Age', 'Date', 'Id', 'First Name'\n\t" \
          "Received: 0, 'First Name', 'Last Name', 'Gender', 'Country', 'Age', 'Date', 'Id'"
        )
      end
    end
  end

  context "when providing a path" do
    subject(:subject) { file_path }

    it_behaves_like "columns matcher in sheet"
    it_behaves_like "column matcher in sheet"
  end

  context "when providing a File" do
    subject(:subject) { File.open(file_path) }

    it_behaves_like "columns matcher in sheet"
    it_behaves_like "column matcher in sheet"
  end

  context "when providing a Roo::Excelx" do
    subject(:subject) { Roo::Spreadsheet.open(file_path) }

    it_behaves_like "columns matcher in sheet"
    it_behaves_like "column matcher in sheet"
  end

  context "when providing a Roo::Excelx::Sheet" do
    subject(:subject) { Roo::Spreadsheet.open(file_path).sheet_for("Sheet1") }

    it_behaves_like "columns matcher without sheet"
    it_behaves_like "column matcher without sheet"
  end

  context "when providing a Axlsx::Package" do
    subject(:subject) { caxlsx_data }

    it_behaves_like "columns matcher in sheet"
    it_behaves_like "column matcher in sheet"
  end

  context "when providing a Axlsx::Workbook" do
    subject(:subject) { caxlsx_data.workbook }

    it_behaves_like "columns matcher in sheet"
    it_behaves_like "column matcher in sheet"
  end

  context "when providing a Axlsx::Worksheet" do
    subject(:subject) { caxlsx_data.workbook.worksheets[0] }

    it_behaves_like "columns matcher without sheet"
    it_behaves_like "column matcher without sheet"
  end
end
