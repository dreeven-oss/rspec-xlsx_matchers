# frozen_string_literal: true

module DataSamples
  RSpec.shared_context "with simple example data" do
    let(:raw_data) do
      [
        [0,	"First Name",	"Last Name",	"Gender",	"Country",	"Age",	"Date",	"Id"],
        [1,	"Dulce",	"Abril",	"Female",	"United States", 32, "15/10/2017", 1562],
        [2, "Mara", "Hashimoto", "Female", "Great Britain", 25, "16/08/2016", 1582],
        [3, "Philip", "Gent", "Male", "France", 36, "21/05/2015", 2587],
        [4, "Kathleen", "Hanner", "Female", "United States", 25, "15/10/2017", 3549],
        [5, "Nereida", "Magwood", "Female", "United States", 58, "16/08/2016", 2468],
        [6, "Gaston", "Brumm", "Male", "United States", 24, "21/05/2015", 2554],
        [7, "Etta", "Hurn", "Female", "Great Britain", 56, "15/10/2017", 3598],
        [8, "Earlean", "Melgar", "Female", "United States", 27, "16/08/2016", 2456],
        [9, "Vincenza", "Weiland", "Female", "United States", 40, "21/05/2015", 6548]
      ]
    end
    let(:caxlsx_data) do
      Axlsx::Package.new do |p|
        p.workbook.add_worksheet(name: "Sheet1") do |sheet|
          raw_data.each do |d|
            sheet.add_row d
          end
        end
      end
    end

    let(:file_path) { fixture_file_path "file_example_XLSX_10.xlsx" }
  end
end
