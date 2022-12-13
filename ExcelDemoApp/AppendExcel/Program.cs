using DocumentFormat.OpenXml.Spreadsheet;
using ExcelUtils;

namespace AppendExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\source\CsharpLibraries\CreatingExcelFilesInCsharp\AppendData.xlsx";

            ExcelHelper excel = new ExcelHelper(filePath);
            excel.AddSheet(1, "Sample Data");
            excel.Close();


            excel = new ExcelHelper(filePath, false);
            Populate(excel);
            excel.Close();
        }

        private static void Populate(ExcelHelper excel)
        {
            excel.Append(new Row());
            excel.Append(new Row());
            Row row = new Row();
            row.AppendStringRefCell("Order Date", "C3");
            row.AppendStringRefCell("Region", "D3");
            row.AppendStringRefCell("Rep", "E3");
            excel.Append(row);
        }
    }
}
