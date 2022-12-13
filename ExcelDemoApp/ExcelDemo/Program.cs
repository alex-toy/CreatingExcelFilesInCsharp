using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\source\CsharpLibraries\CreatingExcelFilesInCsharp\sampleData.xlsx";

            ExcelHelper excel = new ExcelHelper(filePath);
            excel.InitExcel();

            excel.AddSheet(1, "Sample Data");

            //excel.AddSheet(2, "Sample Data Test");



            for (int i = 0; i < 5; i++)
            {
                Row row = new Row();

                for (int j = 0; j < 5; j++)
                {
                    string cellValue = $"Cell - {(i * j).ToString()}";
                    row.AppendStringCell(cellValue);
                }
                excel.SheetData.Append(row);
            }

            excel.Close();
        }
    }
}
