// See https://aka.ms/new-console-template for more information
using Aspose.Cells;
using ExportDBToolBasel3;

Console.WriteLine("Hello, World!");

var importFilePath = @"C:\Users\Windows\Downloads\Doi_chieu_ToolB3_ETL.xlsx";
var workbook = new Workbook(importFilePath);
var output = @"";
foreach (var worksheet in workbook.Worksheets)
{
   var listDataModel =  Process.GetListDataInSheet(worksheet);

    var listTable = listDataModel.Select(o => o.TableName).Distinct().ToList();
    foreach (var tableName in listTable)
    {
        var tables = listDataModel.Where(o => o.TableName == tableName).ToList();
        var nameTable = tables[0].TableName;
         output += Process.ConvertTableObjToContent(tables, nameTable);
    }
}

Process.WriteFile(@"D:\Export\ScriptOrcal\", "SCRIPT_DB_OUTPUT_BASEL3.sql", output);