using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelDemo
{
    public class ExcelHelper
    {
        public SpreadsheetDocument SpreadSheetDocument { get; set; }
        public WorkbookPart WorkbookPart { get; set; }
        public WorksheetPart WorksheetPart { get; set; }
        public Sheets Sheets { get; set; }
        public Worksheet Worksheet { get; set; }
        public SheetData SheetData { get; set; }

        public ExcelHelper(string filePath)
        {
            SpreadSheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        }

        public void InitExcel()
        {
            WorkbookPart = SpreadSheetDocument.AddWorkbookPart();
            WorkbookPart.Workbook = new Workbook();
            WorksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
            WorksheetPart.Worksheet = new Worksheet(new SheetData());
            Sheets = SpreadSheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Worksheet = WorksheetPart.Worksheet;
            SheetData = Worksheet.GetFirstChild<SheetData>();
        }

        public void AddSheet(UInt32Value sheetId, string name)
        {
            Sheet sheet = new Sheet()
            {
                Id = SpreadSheetDocument.WorkbookPart.GetIdOfPart(WorksheetPart),
                SheetId = sheetId,
                Name = name
            };
            Sheets.Append(sheet);
        }

        public void Close()
        {
            WorksheetPart.Worksheet.Save();
            SpreadSheetDocument.Close();
        }
    }
}
