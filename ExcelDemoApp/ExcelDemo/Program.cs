using DocumentFormat.OpenXml.Spreadsheet;
using ExcelUtils;

namespace ExcelDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\source\CsharpLibraries\CreatingExcelFilesInCsharp\sampleData.xlsx";

            ExcelHelper excel = new ExcelHelper(filePath);

            excel.AddSheet(1, "Sample Data");

            PopulateHeader(excel);
            Populate(excel);

            PopulateHeaderBis(excel);
            PopulateBis(excel);

            excel.Close();
        }

        private static void PopulateHeader(ExcelHelper excel)
        {
            Row row = new Row();
            char[] Cols = "BCDEFG".ToCharArray();
            for (int i = 0; i < Cols.Length; i++)
            {
                row.AppendStringRefCell($"Col {i}", $"{Cols[i]}1", excel);
            }
            excel.SheetData.Append(row);
        }

        private static void Populate(ExcelHelper excel)
        {
            for (int i = 1; i < 6; i++)
            {
                Row row = new Row();
                row.Append(new Cell());

                for (int j = 0; j < 6; j++)
                {
                    string cellValue = $"Cell - {(i * j)}";
                    row.AppendStringCell(cellValue);
                }
                excel.SheetData.Append(row);
            }
        }

        private static void PopulateHeaderBis(ExcelHelper excel)
        {
            excel.Append(new Row());
            Row row = new Row();
            char[] Cols = "EFGHI".ToCharArray();
            for (int i = 0; i < Cols.Length; i++)
            {
                row.AppendStringRefCell($"Col {i}", $"{Cols[i]}8", excel);
            }
            excel.SheetData.Append(row);
        }

        private static void PopulateBis(ExcelHelper excel)
        {
            for (int i = 9; i < 12; i++)
            {
                Row row = new Row();
                row.Append(new Cell());
                row.Append(new Cell());
                row.Append(new Cell());
                row.Append(new Cell());

                for (int j = 6; j < 11; j++)
                {
                    string cellValue = $"Cell - {(i * j)}";
                    row.AppendStringCell(cellValue);
                }
                excel.Append(row);
            }
        }
    }
}
