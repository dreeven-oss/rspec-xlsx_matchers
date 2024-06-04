# RSpec::XlsxMatchers

Provides rspec matchers for testing xlsx data with `roo` or `caxlsx` gems.

[TOC]

## Optional dependencies

 - [roo](https://github.com/roo-rb/roo)
 - [caxlsx](https://github.com/caxlsx/caxlsx)

## Configure
  
add `config.include RSpec::XlsxMatchers` to your RSpec.configure declaration

```ruby
# spec_helper.rb
RSpec.configure do |config|
  config.include RSpec::XlsxMatchers
  # ...
end
```

## Matchers

### have_excel_sheets

Tests that xlsx data contains specific sheets (by name or zero-based index)

`subject` must be one of the following types:
  - `String`: file path to an excel file (requires roo gem)
  - `File`: file path to an excel file (requires roo gem)
  - `Roo::Excelx`: excel file loaded with roo (requires roo gem)
  - `Axlsx::Package`: (requires caxlsx gem)
  - `Axlsx::Workbook`: (requires caxlsx gem)


#### examples
```ruby
  it "has a sheet Sheet1" do
    expect(subject).to have_excel_sheets("Sheet1")
  end

  it "has sheets Sheet1 and Sheet2" do
    expect(subject).to have_excel_sheets(%w[Sheet1 Sheet2])
  end

  it "has at least 3 sheets" do
    expect(subject).to have_excel_sheets(2)
  end

  it "has at least 3 sheets and a sheet named Sheet1" do
    expect(subject).to have_excel_sheets([2, "Sheet1"])
  end
```

### have_excel_columns

Tests that a specific sheet contains some columns with specific value on the first row

`subject` must be one of the following types:
  - `String`: file path to an excel file (requires roo gem)
  - `File`: file path to an excel file (requires roo gem)
  - `Roo::Excelx`: excel file loaded with roo (requires roo gem)
  - `Roo::Excelx::Sheet`: (requires roo gem)
  - `Axlsx::Package`: (requires caxlsx gem)
  - `Axlsx::Workbook`: (requires caxlsx gem)
  - `Axlsx::Worksheet`: (requires caxlsx gem)

#### chaining

- `in_sheet(sheet_name)`: sheet where the data is expected. name (String) or index (Integer)
  - **required** unless subject is a `Roo::Excelx::Sheet` or a `Axlsx::Worksheet`
- `exactly`: Match columns in order and fail if extra columns are found

#### examples
```ruby
  it "has a Gender and First Name columns in sheet Sheet 1" do
    expect(subject).to have_excel_columns(["Gender", "First Name"]).in_sheet("Sheet1")
  end

  it "has column Last Name in first sheet" do
    expect(subject).to have_excel_columns("Last Name").in_sheet(0)
  end

  it "has columns First Name and Last Name in first sheet, matching exactly" do
    expect(subject).to have_excel_columns(["First Name", "Last Name"]).in_sheet(0).exactly
  end
```
### have_excel_column

Tests that a specific sheet contains columns with specific value on the first row

`subject` must be one of the following types:
  - `String`: file path to an excel file (requires roo gem)
  - `File`: file path to an excel file (requires roo gem)
  - `Roo::Excelx`: excel file loaded with roo (requires roo gem)
  - `Roo::Excelx::Sheet`: (requires roo gem)
  - `Axlsx::Package`: (requires caxlsx gem)
  - `Axlsx::Workbook`: (requires caxlsx gem)
  - `Axlsx::Worksheet`: (requires caxlsx gem)
  

#### chaining

- `in_sheet(sheet_name)`: sheet where the data is expected. name (String) or index (Integer)
  - **required** unless subject is a `Roo::Excelx::Sheet` or a `Axlsx::Worksheet`

#### examples
```ruby
  it "has a Gender column in sheet Sheet 1" do
    expect(subject).to have_excel_columns("Gender").in_sheet("Sheet1")
  end

  it "has a Last Name column in first sheet" do
    expect(subject).to have_excel_columns("Last Name").in_sheet(0)
  end
```

### have_excel_cells

Tests that a specific row in a specific sheet contains the expected cell values

`subject` must be one of the following types:
  - `String`: file path to an excel file (requires roo gem)
  - `File`: file path to an excel file (requires roo gem)
  - `Roo::Excelx`: excel file loaded with roo (requires roo gem)
  - `Roo::Excelx::Sheet`: (requires roo gem)
  - `Axlsx::Package`: (requires caxlsx gem)
  - `Axlsx::Workbook`: (requires caxlsx gem)
  - `Axlsx::Worksheet`: (requires caxlsx gem)

#### chaining

- `in_sheet(sheet_name)`: sheet where the data is expected. name (String) or index (Integer, zero based)
  - **required** unless subject is a `Roo::Excelx::Sheet` or a `Axlsx::Worksheet`
- `in_row(row_index)`: row where the data is expected (row_index: Integer, zero-based)

#### examples
```ruby
  it "has John Smith information on second row" do
    expect(subject).to have_excel_cells(["John", "Smith", 24]).in_row(1).in_sheet("Sheet1")
  end
```

### have_excel_empty_row

Tests that a specific row in a specific sheet is empty

`subject` must be one of the following types:
  - `String`: file path to an excel file (requires roo gem)
  - `File`: file path to an excel file (requires roo gem)
  - `Roo::Excelx`: excel file loaded with roo (requires roo gem)
  - `Roo::Excelx::Sheet`: (requires roo gem)
  - `Axlsx::Package`: (requires caxlsx gem)
  - `Axlsx::Workbook`: (requires caxlsx gem)
  - `Axlsx::Worksheet`: (requires caxlsx gem)

#### chaining

- `in_sheet(sheet_name)`: sheet where the data is expected. name (String) or index (Integer)
  - **required** unless subject is a `Roo::Excelx::Sheet` or a `Axlsx::Worksheet`


#### examples
```ruby
  it "has no data on third row" do
    expect(subject).to have_excel_empty_row(2)
  end
```

### have_excel_cell_value

Tests that a cell in a specific column of a specific row from a specific sheet has expected value

`subject` must be one of the following types:
  - `String`: file path to an excel file (requires roo gem)
  - `File`: file path to an excel file (requires roo gem)
  - `Roo::Excelx`: excel file loaded with roo (requires roo gem)
  - `Roo::Excelx::Sheet`: (requires roo gem)
  - `Axlsx::Package`: (requires caxlsx gem)
  - `Axlsx::Workbook`: (requires caxlsx gem)
  - `Axlsx::Worksheet`: (requires caxlsx gem)

#### chaining

- `in_sheet(sheet_name)`: sheet where the data is expected. name (String) or index (Integer, zero based)
  - **required** unless subject is a `Roo::Excelx::Sheet` or a `Axlsx::Worksheet`
- `in_row(row_index)`: row where the data is expected (row_index: Integer, zero-based)
- `column(first_row_value)`: value of the column in the first row


#### examples
```ruby
  it "has John in the First Name column of the third row of the first sheet" do
    expect(subject).to have_excel_cell_value("John").in_sheet(0).in_row(2).in_column("First Name")
  end
```
