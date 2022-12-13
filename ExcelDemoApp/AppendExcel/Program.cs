using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace AppendExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\source\CsharpLibraries\CreatingExcelFilesInCsharp\AppendData.xlsx";

            //InitExcelFile(filePath);

            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filePath, true);

            WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;

            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            char[] reference = "BCDEFGH".ToCharArray();

            int skipRowIndex = 3;

            Row row = new Row();
            row.Append(new Cell());
            //AppendStringRefCell(row, "today", reference[0].ToString() + skipRowIndex);
            AppendStringRefCell(row, "today", "B6");
            sheetData.Append(row);

            worksheetPart.Worksheet.Save();

            spreadSheetDocument.Close();
        }

        private static void InitExcelFile(string filePath)
        {
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

             WorksheetPart worksheetPart = GetWorksheetPart(spreadSheetDocument);

            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            sheetData.Append(new Row());
            Row row = new Row();
            row.Append(new Cell());
            AppendStringCell(row, "Order Date");
            AppendStringCell(row, "Region");
            AppendStringCell(row, "Rep");
            AppendStringCell(row, "Item");
            AppendStringCell(row, "Units");
            AppendStringCell(row, "Unit Cost");
            AppendStringCell(row, "Total");
            sheetData.Append(row);

            worksheetPart.Worksheet.Save();

            spreadSheetDocument.Close();
        }

        private static void AppendStringRefCell(Row row, string value, StringValue reference)
        {
            row.Append(new Cell()
            {
                CellValue = new CellValue(value),
                CellReference = reference,
                DataType = CellValues.String
            });
        }

        private static void AppendStringCell(Row row, string value)
        {
            row.Append(new Cell()
            {
                CellValue = new CellValue(value),
                DataType = CellValues.String
            });
        }

        private static WorksheetPart GetWorksheetPart(SpreadsheetDocument spreadSheetDocument)
        {
            WorkbookPart workbookPart = spreadSheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            Sheets sheets = spreadSheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            AddSheet(spreadSheetDocument, worksheetPart, sheets, 1, "Append Data");
            return worksheetPart;
        }

        private static void AddSheet(SpreadsheetDocument spreadSheetDocument, WorksheetPart worksheetPart, Sheets sheets, UInt32Value sheetId, string name)
        {
            Sheet sheet = new Sheet()
            {
                Id = spreadSheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = name
            };
            sheets.Append(sheet);
        }
    }
}
