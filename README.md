```markdown
# Desky Datatable Activities

**Desky Datatable Activities** is a custom library for UiPath designed to simplify and enhance data manipulation within DataTables. This library provides a variety of activities for converting data types, merging tables, splitting tables, summing values, and more.

## Features

- **Convert Columns to Numeric**: Converts specified columns in a DataTable to a numeric data type (double).
- **Convert Columns to Date**: Converts specified columns in a DataTable to the DateTime data type.
- **Convert DataTable to HTML**: Converts a DataTable into an HTML table format.
- **Merge DataTable**: Merges two DataTables into one.
- **Reorder DataTable Columns**: Reorders the columns of a DataTable based on specified order.
- **Split DataTable Vertically**: Splits a DataTable into two separate DataTables based on a specified column index.
- **Sum Column Values**: Calculates the sum of the values in a specified column of a DataTable.
- **Get Column Name By Index**: Retrieves the name of a column at a specified index.
- **Get Column Index By Name**: Retrieves the index of a column with a specified name.

## Installation

1. Download the latest release of the **Desky Datatable Activities** library.
2. In UiPath Studio, go to **Manage Packages** and select **nuget.org**.
3. Search *Desky Datatable Activities** .
4. Click **Install** to add it to your project.

## Usage

### Convert Columns to Numeric

```xml
<ConvertColumnToNumeric>
    <InputDataTable>Your DataTable</InputDataTable>
    <ColumnNames>
        <string>Column1</string>
        <string>Column2</string>
    </ColumnNames>
    <OutputDataTable>Result DataTable</OutputDataTable>
</ConvertColumnToNumeric>
```

### Get Column Index By Name

```xml
<GetColumnIndexByName>
    <InputDataTable>Your DataTable</InputDataTable>
    <ColumnName>Column1</ColumnName>
    <ColumnIndex>Index</ColumnIndex>
</GetColumnIndexByName>
```

## Activities Overview

### 1. Convert Columns to Numeric
- **Description**: Converts specified columns in a DataTable to a numeric data type (double).
- **Input**: DataTable, List of Column Names.
- **Output**: DataTable with converted columns.

### 2. Convert Columns to Date
- **Description**: Converts specified columns in a DataTable to DateTime.
- **Input**: DataTable, List of Column Names, Date Format.
- **Output**: DataTable with converted columns.

### 3. Convert DataTable to HTML
- **Description**: Converts a DataTable into HTML format.
- **Input**: DataTable.
- **Output**: HTML string representation.

### 4. Merge DataTable
- **Description**: Merges two DataTables.
- **Input**: Destination DataTable, Source DataTable.
- **Output**: Updated Destination DataTable.

### 5. Reorder DataTable Columns
- **Description**: Reorders DataTable columns based on specified order.
- **Input**: DataTable, List of Expected Column Names.
- **Output**: DataTable with reordered columns.

### 6. Split DataTable Vertically
- **Description**: Splits a DataTable into two based on a specified column index.
- **Input**: DataTable, Column Index.
- **Output**: Two DataTables.

### 7. Sum Column Values
- **Description**: Calculates the sum of values in a specified column.
- **Input**: DataTable, Column Name.
- **Output**: Sum of numeric values.

### 8. Get Column Name By Index
- **Description**: Retrieves the name of a column at a specified index.
- **Input**: DataTable, Column Index.
- **Output**: Column Name.

### 9. Get Column Index By Name
- **Description**: Retrieves the index of a column by its name.
- **Input**: DataTable, Column Name.
- **Output**: Column Index.

## Contributing

Contributions are welcome! If you have suggestions for improvements or new features, please fork the repository and submit a pull request.

## License

This project is licensed under the MIT License.

## Contact

For any inquiries or issues, please contact [your email or contact method].
```

### Notes
- **Customization**: Replace `[your email or contact method]` with your actual contact information.
- **Examples**: You may add more detailed examples or even screenshots of the activities in use, depending on your target audience.
- **License**: Adjust the license section based on your project's licensing approach.

Feel free to adjust the structure or content as needed! If you have any specific additions or modifications in mind, let me know!
