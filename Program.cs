using ExcelReader;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Migrations.Operations.Builders;
using OfficeOpenXml;
using System.Configuration;
using System.Data;
using System.Reflection.PortableExecutable;




var connectionString = ConfigurationManager.AppSettings.Get("ConnectionString");

// Your database name
string databaseName = "ArquiveReader";
var file = "C:\\Programming\\testing.xlsx";

// Drop the database if it exists
DropDatabaseIfExists(connectionString, databaseName);

DataTable columnHeaders = GetHeaders(file);

// Create the database
CreateDatabase(connectionString, databaseName);
CreateTable(connectionString, databaseName, columnHeaders);

Console.WriteLine("Database created");

// Load Excel data into a DataTable
PopulateTable(file, connectionString, columnHeaders, databaseName);

Console.WriteLine("Saving to the database.");

// Save data to the database
//SaveDataToDatabase(dt, connectionString);

// Perform other initialization steps if needed

Console.WriteLine("Database recreated successfully.");

static void PopulateTable(string file, string connectionString, DataTable columnHeaders, string databaseName)
{
    using (var package = new ExcelPackage(file))
    {

        var worksheet = package.Workbook.Worksheets[0];
        DataTable excelData = columnHeaders;

        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
        {
            var dataRow = excelData.NewRow();

            // Populate dataRow with values from Excel
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                // Use the column name from excelSchema to get the correct index
                

                var columnName = columnHeaders.Columns[col - 1].ColumnName;
                dataRow[columnName] = worksheet.Cells[row, col].Text;

            }

            excelData.Rows.Add(dataRow);
        }

    
        // Insert data into SQL Server
        using (var bulkCopy = new SqlBulkCopy(ConfigurationManager.AppSettings.Get("DatabaseConnectionString")))
        {
            bulkCopy.DestinationTableName = "FileRead";
            bulkCopy.WriteToServer(excelData);
        }
    }
}

static void CreateTable(string connectionString, string databaseName, DataTable excelSchema)
{
    using (var connection = new SqlConnection(connectionString))
    {
        connection.Open();

        using (var command = connection.CreateCommand())
        {
            command.CommandText = $"CREATE TABLE {databaseName}.dbo.FileRead (" +
                                  string.Join(", ", excelSchema.Columns.Cast<DataColumn>().Select(c => $"{ConvertColumnNames(c.ColumnName)} NVARCHAR(MAX)")) +
                                  ")";
           var response = command.ExecuteNonQuery();
            Console.WriteLine(response);
        }
    }
}

static DataTable GetHeaders(string filePath)
{
    DataTable excelSchema = new DataTable();

    using (var package = new ExcelPackage(filePath))
    {
        // Assume the first worksheet is your target
        var worksheet = package.Workbook.Worksheets[0];

        // Populate excelSchema with column names from Excel
        foreach (var header in worksheet.Cells["A1:Z1"])
        {
            excelSchema.Columns.Add(header.Text);
        }
    }

    return excelSchema;
}

static void DropDatabaseIfExists( string connectionString, string databaseName )
{
    using (SqlConnection connection = new SqlConnection(connectionString))
    {
        connection.Open();

        string dropDatabaseSql = $"IF EXISTS (SELECT 1 FROM sys.databases WHERE name = '{databaseName}') " +
                                $"DROP DATABASE [{databaseName}]";

        using (SqlCommand command = new SqlCommand(dropDatabaseSql, connection))
        {
            command.ExecuteNonQuery();
        }
    }
}

static string ConvertColumnNames(string names)
{
    return names.Replace(" ", "_").Trim();
}

static void SaveDataToDatabase( DataTable dt, string connectionString )
{
    using (SqlConnection connection = new SqlConnection(connectionString))
    {
        connection.Open();

        foreach (DataRow row in dt.Rows)
        {
            using (SqlCommand command = connection.CreateCommand())
            {
                // Assuming your table name is "YourTableName"
                command.CommandText = $"INSERT INTO ReadFile ({string.Join(", ", dt.Columns.Cast<DataColumn>().Select(c => ConvertColumnNames(c.ColumnName)))}) VALUES ({string.Join(", ", dt.Columns.Cast<DataColumn>().Select(c => $"@{ConvertColumnNames(c.ColumnName)}"))})";

                // Add parameters dynamically
                foreach (DataColumn col in dt.Columns)
                {
                    command.Parameters.AddWithValue($"@{col.ColumnName}", row[col]);
                }

                // Execute the query
                command.ExecuteNonQuery();
            }
        }
    }
}

static DataTable LoadExcelData( string filePath )
{
    // Load Excel data into a DataTable using EPPlus or another library
    // Example using EPPlus
    using (var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(filePath)))
    {
        var worksheet = package.Workbook.Worksheets[0];
        DataTable dt = new DataTable();

        // Add columns dynamically based on Excel headers
        foreach (var headerCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
        {
            dt.Columns.Add(headerCell.Text);
        }

        // Iterate through rows and populate data
        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
        {
            var dataRow = dt.Rows.Add();
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                dataRow[col - 1] = worksheet.Cells[row, col].Text;
            }
        }

        return dt;
    }
}

static void CreateDatabase( string connectionString, string databaseName )
{
    using (SqlConnection connection = new SqlConnection(connectionString))
    {
        connection.Open();

        string createDatabaseSql = $"CREATE DATABASE [{databaseName}]";

        using (SqlCommand command = new SqlCommand(createDatabaseSql, connection))
        {
            command.ExecuteNonQuery();
        }
    }
}

/*
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
            string header = worksheet.Cells[1, col].Text;
            headers.Add(header.Replace(" ", "_").Trim());
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
    Console.WriteLine(createTableSql);

    // Execute the SQL command to create the table
    dbContext.Database.ExecuteSqlRaw(createTableSql);
}

static void PopulateTable( ExcelContext dbContext, string tableName, List<string> columns, List<Dictionary<string, object>> data )
{
    foreach (var row in data)
    {
        // Create an instance of the DynamicTable model and set properties dynamically
        var dynamicTableInstance = new DynamicTable
        {
            DynamicProperties = new Dictionary<string, object>()
        };

        foreach (var column in columns)
        {
            var value = row.ContainsKey(column) ? row[column]?.ToString() : null;
            dynamicTableInstance.DynamicProperties[column] = value;
        }

        // Add the instance to the DbSet and save changes
        dbContext.Add(dynamicTableInstance);
        dbContext.SaveChanges();
    }
}

*/