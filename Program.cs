using OfficeOpenXml;

string filePath = "C:\\Programming\\testing.xlsx";

List<Dictionary<string, object>> excelData = ReadExcelToList(filePath);

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

        var headers = new List<string>();

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