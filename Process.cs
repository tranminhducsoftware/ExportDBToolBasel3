using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

namespace ExportDBToolBasel3
{
    public static class Process
    {

        public static List<DataModel> GetListDataInSheet(Worksheet worksheet)
        {
            var rowKey = 0;
            var maxRows = worksheet.Cells.MaxDataRow;
            var listData = new List<DataModel>();

            for (var row = rowKey; row <= maxRows; row++)
            {
                var tableName = worksheet.Cells[row, 0]?.Value?.ToString() ?? string.Empty;
                var fieldName = worksheet.Cells[row, 1]?.Value?.ToString() ?? string.Empty;
                var fieldType = worksheet.Cells[row, 2]?.Value?.ToString() ?? string.Empty;
                if (fieldType == "String" || fieldType == "Varchar")
                {
                    fieldType = "varchar2";
                }

                listData.Add(new DataModel()
                {
                    TableName = tableName,
                    FieldName = fieldName,
                    FieldType = fieldType,
                });
            }
            return listData;
        }

        public static string ConvertTableObjToContent(List<DataModel> tableObjs, string tableName)
        {
            var output = string.Empty;
            var tableScript = $"---- SCRIPT TẠO BẢNG: {tableName} --------------------------- \n";
            tableScript += $"DROP TABLE BASEL3.{tableName}; \n";
            tableScript += $"CREATE TABLE BASEL3.{tableName}( \n";
            foreach (var tableObj in tableObjs)
            {
 
                if (tableObj.FieldType.ToUpper() == "DATE")
                {
                    tableScript += $"\t{tableObj.FieldName} DATE, \n";

                }
                if (tableObj.FieldType.ToUpper() == "VARCHAR2")
                {
                    tableScript += $"\t{tableObj.FieldName} VARCHAR2(225),\n";

                }
                if (tableObj.FieldType.ToUpper() == "NUMBER")
                {
                    tableScript += $"\t{tableObj.FieldName} NUMBER(22,6),\n";

                }
            }
            tableScript = tableScript.Substring(0, tableScript.Length - 2);
            tableScript += $"\n); \n---- END TABLE {tableName} -------------------------------- \n\n\n";
            output += tableScript;
            return output;

        }

        public static void WriteFile(string rootFolder, string nameFile, string contents)
        {
            var pathFileGroovy = Path.Combine(rootFolder, nameFile);
            if (File.Exists(pathFileGroovy))
            {
                File.Delete(pathFileGroovy);
                Console.WriteLine("File deleted: " + nameFile);
            }
            using var streamWriter = File.CreateText(pathFileGroovy);
            streamWriter.WriteLine(contents);
        }
    }
}
