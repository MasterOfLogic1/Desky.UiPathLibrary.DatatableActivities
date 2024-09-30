using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Desky.Datatable
{
    internal static class ActionHelper
    {

        public static DataTable ConvertColumnToNumeric(DataTable InputDatatable, string columnName)
        {
            // Check if the DataTable is not empty
            if (InputDatatable != null && InputDatatable.Rows.Count > 0)
            {
                Console.WriteLine("Converting column to double for " + columnName);
                // Check if the column exists in the DataTable
                if (InputDatatable.Columns.Contains(columnName))
                {
                    // Get the ordinal position of the column
                    int pos = InputDatatable.Columns[columnName].Ordinal;
                    // Create a new column with "_new" appended to the column name
                    string newColumnName = columnName + "_new";
                    InputDatatable.Columns.Add(newColumnName, typeof(decimal)).SetOrdinal(pos + 1);
                    // Iterate through each row and update the new column with the parsed double
                    foreach (DataRow row in InputDatatable.Rows)
                    {
                        string valueString = row[columnName].ToString().Trim();
                        double parsedValue;
                        // Try to parse the string to double
                        if (double.TryParse(valueString, out parsedValue))
                        {
                            row[newColumnName] = parsedValue;
                        }
                        else
                        {
                            // Optionally handle cases where parsing fails (e.g., set to DBNull)
                            row[newColumnName] = DBNull.Value;
                        }
                    }
                    // Remove the old column and rename the new one to the original column name
                    InputDatatable.Columns.Remove(columnName);
                    InputDatatable.Columns[newColumnName].ColumnName = columnName;
                }
                else
                {
                    throw new Exception("the column named [" + columnName + "] does not exist in referenced table");
                }
            }
            else
            {
                throw new Exception("referenced table should not be null");
            }
            return InputDatatable;
        }

        public static DataTable ConvertColumnToDate(DataTable InputDatatable, string columnName, string expectedDateFormat)
        {

            // Check if the DataTable is not empty
            if (InputDatatable != null && InputDatatable.Rows.Count > 0)
            {
                Console.WriteLine("fixing date column for " + columnName);
                // Check if the column exists in the DataTable
                if (InputDatatable.Columns.Contains(columnName))
                {
                    // Get the ordinal position of the column
                    int pos = InputDatatable.Columns[columnName].Ordinal;
                    // Create a new column with "_new" appended to the column name
                    string newColumnName = columnName + "_new";
                    InputDatatable.Columns.Add(newColumnName, typeof(DateTime)).SetOrdinal(pos + 1);

                    // Iterate through each row and update the new column with the parsed date
                    foreach (DataRow row in InputDatatable.Rows)
                    {
                        string dateString = row[columnName].ToString().Trim().Replace(" ", string.Empty);
                        DateTime parsedDate;

                        // Try to parse the date string using the specified format
                        if (DateTime.TryParseExact(dateString, expectedDateFormat, null, System.Globalization.DateTimeStyles.None, out parsedDate))
                        {
                            row[newColumnName] = parsedDate;
                        }
                    }
                    // Remove the old column and rename the new one to the original column name
                    InputDatatable.Columns.Remove(columnName);
                    InputDatatable.Columns[newColumnName].ColumnName = columnName;
                }
                else
                {
                    throw new ArgumentException("the column named [" + columnName + "] does not exist in referenced table");
                }
            }
            else
            {
                throw new ArgumentException("referenced table should not be null");
            }

            return InputDatatable;
        }

        public static DataTable ConvertAllColumnsToString(DataTable InputDatatable)
        {
            DataTable dtFix = new DataTable();
            // Check if the DataTable is not empty
            if (InputDatatable != null && InputDatatable.Rows.Count > 0)
            {
                foreach (DataColumn col in InputDatatable.Columns)
                {
                    dtFix.Columns.Add(col.ColumnName, typeof(string));
                }

                foreach (DataRow row in InputDatatable.Rows)
                {
                    object[] newRow = new object[row.ItemArray.Length];
                    for (int i = 0; i < row.ItemArray.Length; i++)
                    {
                        newRow[i] = row[i]?.ToString();
                    }
                    dtFix.Rows.Add(newRow);
                }
            }
            else
            {
                throw new ArgumentException("referenced table should not be null");
            }
            return dtFix;
        }

        public static DataTable ReorderColumns(DataTable InputDatatable, List<string> expectedColumnOrder)
        {
            if (InputDatatable != null && InputDatatable.Columns.Count > 0)
            {
                int currentPosition = 0;

                foreach (string columnName in expectedColumnOrder)
                {
                    if (InputDatatable.Columns.Contains(columnName))
                    {
                        InputDatatable.Columns[columnName].SetOrdinal(currentPosition);
                        currentPosition++;
                    }
                }
                return InputDatatable;
            }
            else
            {
                throw new ArgumentException("referenced table should have one or more column");
            }
        }

        public static DataTable MergeDatatables(DataTable sourceDataTable, DataTable destinationDataTable)
        {
            if (sourceDataTable != null && sourceDataTable.Rows.Count > 0)
            {
                if (destinationDataTable == null)
                {
                    destinationDataTable = sourceDataTable.Clone(); // Clone the structure of sourceDataTable
                }
                destinationDataTable.Merge(sourceDataTable, false, MissingSchemaAction.Ignore); // Merge the sourceDataTable into destinationDataTable
            }
            return destinationDataTable;
        }

        public static double SumColumnValues(DataTable InputDatatable, string columnName)
        {
            if (InputDatatable == null || !InputDatatable.Columns.Contains(columnName))
            {
                throw new ArgumentException("the column named [" + columnName + "] does not exist in referenced table");
            }

            double sum = InputDatatable.AsEnumerable().Sum(row =>
            {
                string cellValue = row[columnName]?.ToString().Trim();
                if (string.IsNullOrEmpty(cellValue) || !double.TryParse(cellValue, out double numericValue))
                {
                    return 0; // Treat non-numeric or empty values as zero
                }
                return numericValue;
            });

            return sum;
        }

        public static string ConvertDataTableToHtml(DataTable InputDatatable)
        {
            StringBuilder builder = new StringBuilder();
            string tableOpeningTag = "<table border='1' style='border-collapse:collapse'>";

            Console.WriteLine("Setting HTML mail table header....");

            if (InputDatatable != null && InputDatatable.Rows.Count > 0)
            {
                builder.Append(tableOpeningTag);

                // Create table header
                var tableHeader = "<tr>" +
                    string.Join(Environment.NewLine, InputDatatable.Columns.Cast<DataColumn>()
                    .Select(c => $"<th style=\"text-align: center\">{c.ColumnName}</th>")) +
                    "</tr>";
                builder.Append(tableHeader);

                // Generate the body of the table
                Console.WriteLine("Now generating body of HTML table....");
                string tdTag = "<td style=\"text-align: center\">{0}</td>";

                var rows = InputDatatable.AsEnumerable().Select(r =>
                    "<tr>" +
                    string.Join(Environment.NewLine, r.ItemArray.Select(e => string.Format(tdTag, e.ToString()))) +
                    "</tr>");

                builder.Append(string.Join(Environment.NewLine, rows));

                // Add table closing tag
                builder.Append("</table>");
            }

            return builder.ToString();
        }

        public static void SplitTableVertically(DataTable InputDatatable, int columnToSplitFrom, out DataTable dt1, out DataTable dt2)
        {
            dt1 = new DataTable();
            dt2 = new DataTable();

            if (InputDatatable != null && InputDatatable.Columns.Count > 0)
            {
                // Add columns to the first table up to the split point
                for (int i = 0; i < columnToSplitFrom; i++)
                {
                    dt1.Columns.Add(InputDatatable.Columns[i].ColumnName, InputDatatable.Columns[i].DataType);
                }

                // Add remaining columns to the second table
                for (int i = columnToSplitFrom; i < InputDatatable.Columns.Count; i++)
                {
                    dt2.Columns.Add(InputDatatable.Columns[i].ColumnName, InputDatatable.Columns[i].DataType);
                }

                // Copy rows from the original table to the new tables
                foreach (DataRow row in InputDatatable.Rows)
                {
                    DataRow row1 = dt1.NewRow();
                    DataRow row2 = dt2.NewRow();

                    // Fill the first row with data up to the split point
                    for (int i = 0; i < columnToSplitFrom; i++)
                    {
                        row1[i] = row[i];
                    }

                    // Fill the second row with the remaining data
                    for (int i = columnToSplitFrom; i < InputDatatable.Columns.Count; i++)
                    {
                        row2[i - columnToSplitFrom] = row[i];
                    }

                    // Add rows to respective tables
                    dt1.Rows.Add(row1);
                    dt2.Rows.Add(row2);
                }
            }
        }

        public static string GetColumnNameByIndex(DataTable InputDatatable, int columnIndex)
        {
            // Check if the index is within the range of columns in the DataTable
            if (columnIndex < 0 || columnIndex >= InputDatatable.Columns.Count)
            {
                throw new IndexOutOfRangeException($"Column index {columnIndex} is out of range.");
            }
            // Return the column name
            return InputDatatable.Columns[columnIndex].ColumnName;
        }

        public static int GetColumnIndexByName(DataTable InputDatatable, string columnName)
        {
            // Check if the column name exists in the DataTable
            if (!InputDatatable.Columns.Contains(columnName))
            {
                throw new ArgumentException($"Column name '{columnName}' does not exist in the DataTable.");
            }

            // Return the index of the column with the specified name
            return InputDatatable.Columns[columnName].Ordinal;
        }
    }
}
