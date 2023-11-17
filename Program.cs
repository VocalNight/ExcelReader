using ExcelReader;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

string filePath = "C:\\Programming\\testing.xlsx";
var headers = new List<string>();

List<Dictionary<string, object>> excelData = ReadExcelToList(filePath);
var context = new ExcelContext();

using (context)
{
    context.Database.EnsureDeleted();
    context.Database.Migrate();

    // Dynamically create a table with columns based on the Excel headers
    CreateTable(context, "DynamicTables", headers);

    // Populate the table with data from the Excel sheet
    PopulateTable(context, "DynamicTables", headers, excelData);
}



foreach (var row in excelData)
    {
        foreach (var kvp in row)
        {
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");
        }
        Console.WriteLine();
    }

List<Dictionary<string, object>> ReadExcelToList( string filePath )
{
    List<Dictionary<string, object>> excelData = new List<Dictionary<string, object>>();

    using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
    {
        var worksheet = package.Workbook.Worksheets[0]; // Assuming data is on the first worksheet

        int rowCount = worksheet.Dimension.Rows;
        int colCount = worksheet.Dimension.Columns;

        // Read headers from the first row
        for (int col = 1; col <= colCount; col++)
        {
            headers.Add(worksheet.Cells[1, col].Text);
        }

        // Read data from subsequent rows
        for (int row = 2; row <= rowCount; row++)
        {
            var rowData = new Dictionary<string, object>();

            for (int col = 1; col <= colCount; col++)
            {
                string header = headers[col - 1];
                string cellValue = worksheet.Cells[row, col].Text;

                rowData[header] = cellValue;
            }

            excelData.Add(rowData);
        }
    }

    return excelData;
}

static void CreateTable( ExcelContext dbContext, string tableName, List<string> columns )
{
    // Build a SQL command to create the table dynamically
    string createTableSql = $"CREATE TABLE {tableName} (Id INT PRIMARY KEY IDENTITY(1,1), {string.Join(", ", columns.Select(c => $"{c} NVARCHAR(MAX)"))})";

    // Execute the SQL command to create the table
    dbContext.Database.ExecuteSqlRaw(createTableSql);
}

static void PopulateTable( ExcelContext dbContext, string tableName, List<string> columns, List<Dictionary<string, object>> data )
{
    foreach (var row in data)
    {
        // Create an instance of the DynamicTable model and set properties dynamically
        var dynamicTableInstance = new DynamicTable();

        foreach (var column in columns)
        {
            var value = row.ContainsKey(column) ? row[column]?.ToString() : null;
            dynamicTableInstance.GetType().GetProperty(column)?.SetValue(dynamicTableInstance, value);
        }

        // Add the instance to the DbSet and save changes
        dbContext.Add(dynamicTableInstance);
        dbContext.SaveChanges();
    }
}

