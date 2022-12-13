using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace ExcelDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\source\CsharpLibraries\CreatingExcelFilesInCsharp\sampleData.xlsx";
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookPart = spreadSheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            Sheets sheets = spreadSheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Sheet sheet = new Sheet()
            {
                Id = spreadSheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sample Data"
            };

            sheets.Append(sheet);

            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            for(int i = 0; i < 5; i++)
            {
                Row row = new Row();

                for(int j = 0; j < 5; j++)
                {
                    Cell cell = new Cell()
                    {
                        CellValue = new CellValue((i * j).ToString()),
                        DataType = CellValues.String
                    };
                    row.Append(cell);
                }

                sheetData.Append(row);
            }

            worksheetPart.Worksheet.Save();

            spreadSheetDocument.Close();
        }
    }
}
