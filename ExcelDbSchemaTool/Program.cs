using ExcelDbSchemaTool;
using ExcelDbSchemaTool.Models;
using Microsoft.Office.Interop.Excel;
using Npgsql;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;

string connectionString = "your database connection string goes over here";

List<string> schemaNames = ["public", "actor", "Arcane", "TheLastoFUs"]; // Your DB schemes 

List<SchemaModel> schemas = new();
foreach (string schema in schemaNames)
    schemas.Add(CollectTableDetailBySchema(schema));

CreateExcel(schemas);

SchemaModel CollectTableDetailBySchema(string schema)
{
    SchemaModel schemaModel = new();
    schemaModel.Name = schema;

    using var connection = new NpgsqlConnection(connectionString);
    connection.Open();

    string columnsQuery = $@"
                SELECT table_name, column_name, data_type,d.description
                FROM information_schema.columns c
                LEFT JOIN 
    pg_catalog.pg_class t
ON 
    t.relname = c.table_name
LEFT JOIN 
    pg_catalog.pg_description d
ON 
    d.objoid = t.oid AND d.objsubid = c.ordinal_position
                
                WHERE table_schema = '{schema}'
                ORDER BY table_name, ordinal_position";

    using var columnsCommand = new NpgsqlCommand(columnsQuery, connection);

    List<TableDBModel> tablesFromDB = new List<TableDBModel>();

    using var reader = columnsCommand.ExecuteReader();

    while (reader.Read())
    {
        TableDBModel tableDB = new();
        tableDB.TableName = reader["table_name"].ToString();
        tableDB.ColumnName = reader["column_name"].ToString();
        tableDB.ColumnType = reader["data_type"].ToString();
        tableDB.Comment = reader["description"]?.ToString();
        tablesFromDB.Add(tableDB);
    }
    connection.Close();

    schemaModel.Tables = tablesFromDB.ToTables();
    return schemaModel;
}

void CreateExcel(List<SchemaModel> schemas)
{
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    using var package = new ExcelPackage();

    foreach (var schema in schemas)
    {
        var worksheet = package.Workbook.Worksheets.Add(schema.Name.ToUpper());

        int rowCounter = 1;
        AddFrame(worksheet.Cells[$"A{rowCounter}:C{rowCounter}"]);
        foreach (var table in schema.Tables)
        {
            var cellRange = worksheet.Cells[$"A{rowCounter}:C{rowCounter}"];
            cellRange.Value = table.Name;

            AddStyleToCell(cellRange, true);
            rowCounter++;

            worksheet.Cells[$"A{rowCounter}:C{rowCounter}"].Merge = true;
            AddFrame(worksheet.Cells[$"A{rowCounter}:C{rowCounter}"]);
            rowCounter++;

            var cellA = worksheet.Cells[$"A{rowCounter}"];
            var cellB = worksheet.Cells[$"B{rowCounter}"];
            var cellC = worksheet.Cells[$"C{rowCounter}"];
            AddStyleToCell(cellA, false);
            AddStyleToCell(cellB, false);
            AddStyleToCell(cellC, false);

            cellA.Value = "Column Name";
            cellB.Value = "Column Type";
            cellC.Value = "Description";
            rowCounter++;

            foreach (var column in table.Columns)
            {
                worksheet.Cells[$"A{rowCounter}"].Value = column.Name;
                worksheet.Cells[$"B{rowCounter}"].Value = column.Type;
                worksheet.Cells[$"C{rowCounter}"].Value = column.Description;

                AddFrame(worksheet.Cells[$"A{rowCounter}:C{rowCounter}"]);
                rowCounter++;
            }

            rowCounter++;

        }
        AddFrame(worksheet.Cells[$"A{rowCounter}:C{rowCounter}"]);
        worksheet.Cells.AutoFitColumns();
    }


    Save(package, "Documentation");
}
void Save(ExcelPackage package, string projectName)
{
    var filePath = $"C:/Users/yourPcUserName/Desktop/{projectName}{Guid.NewGuid()}.xlsx";
    File.WriteAllBytes(filePath, package.GetAsByteArray());
    Console.WriteLine($"File has been created: {filePath}");
}
void AddStyleToCell(ExcelRange cellRange, bool isMerge)
{
    if (isMerge)
        cellRange.Merge = true;

    cellRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
    cellRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
    cellRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
    cellRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSteelBlue);
    cellRange.Style.Font.Bold = true;

    AddFrame(cellRange);

}
void AddFrame(ExcelRange cellRange)
{

    cellRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
    cellRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
    cellRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
    cellRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
}


