# frozen_string_literal: true

require "spec_helper"

RSpec.describe RSpec::XlsxMatchers::CellValue do
  RSpec.shared_examples "have_excel_cell_value in_sheet" do |sheet|
    it "succeeds with string" do
      expect(subject).to have_excel_cell_value("Hashimoto").in_row(2).in_sheet(sheet).in_column("Last Name")
    end

    it "succeeds with 1" do
      expect(subject).to have_excel_cell_value(1).in_row(1).in_sheet(sheet).in_column("0")
    end

    it "fails when sheet is not provided" do
      expect do
        expect(subject).to have_excel_cell_value("Hashimoto").in_row(2).in_column("Last Name")
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Sheet not provided"
      )
    end

    it "fails when provided row is empty" do
      expect do
        expect(subject).to have_excel_cell_value("Hashimoto").in_row(22).in_column("Last Name").in_sheet(sheet)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Row 22 is empty"
      )
    end

    it "fails when column is not found by name" do
      expect do
        expect(subject).to have_excel_cell_value("Hashimoto").in_row(2).in_column("SSN").in_sheet(sheet)
      end.to raise_error(
        RSpec::Expectations::ExpectationNotMetError,
        "Column SSN not found"
      )
    end
  end

  RSpec.shared_examples "have_excel_cell_value with sheet" do
    describe "by name" do
      include_examples "have_excel_cell_value in_sheet", "Sheet1"

      it "fails when provided sheet name is invalid" do
        expect do
          expect(subject).to have_excel_cell_value("Hashimoto")
            .in_row(2)
            .in_column("Last Name")
            .in_sheet("InvalidSheetName")
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Could not find sheet InvalidSheetName"
        )
      end
    end

    describe "by index" do
      include_examples "have_excel_cell_value in_sheet", 0

      it "fails when provided sheet index is invalid" do
        expect do
          expect(subject).to have_excel_cell_value("Hashimoto").in_row(2).in_column("Last Name").in_sheet(1)
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "Could not find sheet 1"
        )
      end
    end
  end

  RSpec.shared_examples "have_excel_cell_value without sheet" do
    it "succeeds with string" do
      expect(subject).to have_excel_cell_value("Hashimoto").in_row(2).in_column("Last Name")
    end

    it "succeeds with 1" do
      expect(subject).to have_excel_cell_value(1).in_row(1).in_column("0")
    end
  end

  context "with simple example data" do
    include_context "with simple example data"

    context "when providing a path" do
      subject(:subject) { file_path }

      it_behaves_like "have_excel_cell_value with sheet"
    end

    context "when providing a File" do
      subject(:subject) { File.open(file_path) }

      it_behaves_like "have_excel_cell_value with sheet"
    end

    context "when providing a Roo::Excelx" do
      subject(:subject) { Roo::Spreadsheet.open(file_path) }

      it_behaves_like "have_excel_cell_value with sheet"
    end

    context "when providing a Roo::Excelx::Sheet" do
      subject(:subject) { Roo::Spreadsheet.open(file_path).sheet_for("Sheet1") }

      it_behaves_like "have_excel_cell_value without sheet"
    end

    context "when providing a Axlsx::Package" do
      subject(:subject) { caxlsx_data }

      it_behaves_like "have_excel_cell_value with sheet"
    end

    context "when providing a Axlsx::Workbook" do
      subject(:subject) { caxlsx_data.workbook }

      it_behaves_like "have_excel_cell_value with sheet"
    end

    context "when providing a Axlsx::Worksheet" do
      subject(:subject) { caxlsx_data.workbook.worksheets[0] }

      it_behaves_like "have_excel_cell_value without sheet"
    end
  end

  RSpec.shared_examples "have_excel_cell_value empty data" do |with_sheet|
    if with_sheet
      it "fails when there is no header row with sheet name" do
        expect do
          expect(subject).to have_excel_cell_value("Hashimoto").in_row(2).in_column("Last Name").in_sheet("Sheet1")
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "First row not found"
        )
      end

      it "fails when there is no header row with sheet index" do
        expect do
          expect(subject).to have_excel_cell_value("Hashimoto").in_row(2).in_column("Last Name").in_sheet(0)
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "First row not found"
        )
      end
    else
      it "fails when there is no header row without sheet" do
        expect do
          expect(subject).to have_excel_cell_value("Hashimoto").in_row(2).in_column("Last Name")
        end.to raise_error(
          RSpec::Expectations::ExpectationNotMetError,
          "First row not found"
        )
      end
    end
  end

  context "with empty data" do
    include_context "with simple example empty data"

    context "when providing a path" do
      subject(:subject) { file_path }

      it_behaves_like "have_excel_cell_value empty data", true
    end

    context "when providing a File" do
      subject(:subject) { File.open(file_path) }

      it_behaves_like "have_excel_cell_value empty data", true
    end

    context "when providing a Roo::Excelx" do
      subject(:subject) { Roo::Spreadsheet.open(file_path) }

      it_behaves_like "have_excel_cell_value empty data", true
    end

    context "when providing a Roo::Excelx::Sheet" do
      subject(:subject) { Roo::Spreadsheet.open(file_path).sheet_for("Sheet1") }

      it_behaves_like "have_excel_cell_value empty data", false
    end

    context "when providing a Axlsx::Package" do
      subject(:subject) { caxlsx_data }

      it_behaves_like "have_excel_cell_value empty data", true
    end

    context "when providing a Axlsx::Workbook" do
      subject(:subject) { caxlsx_data.workbook }

      it_behaves_like "have_excel_cell_value empty data", true
    end

    context "when providing a Axlsx::Worksheet" do
      subject(:subject) { caxlsx_data.workbook.worksheets[0] }

      it_behaves_like "have_excel_cell_value empty data", false
    end
  end
end
