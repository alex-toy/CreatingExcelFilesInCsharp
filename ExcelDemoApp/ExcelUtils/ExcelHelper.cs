using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace ExcelUtils
{
    public class ExcelHelper
    {
        public SpreadsheetDocument SpreadSheetDocument { get; set; }
        public WorkbookPart WorkbookPart { get; set; }
        public WorksheetPart WorksheetPart { get; set; }
        public Sheets Sheets { get; set; }
        public Worksheet Worksheet { get; set; }
        public SheetData SheetData { get; set; }
        public SharedStringTablePart SharedStringTablePart { get; set; }

        public ExcelHelper(string filePath, bool isCreate = true)
        {
            if (isCreate)
            {
                SpreadSheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
                InitExcel();
            }
            else
            {
                SpreadSheetDocument = SpreadsheetDocument.Open(filePath, true);
                ReInitExcel();
            }
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

        public void ReInitExcel()
        {
            WorkbookPart = SpreadSheetDocument.WorkbookPart;
            WorksheetPart = WorkbookPart.WorksheetParts.First();
            Worksheet = WorksheetPart.Worksheet;
            SheetData = Worksheet.GetFirstChild<SheetData>();
        }

        public void SetSharedString()
        {
            if (SpreadSheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                SharedStringTablePart = SpreadSheetDocument.WorkbookPart.SharedStringTablePart;
            }
            else
            {
                SharedStringTablePart = SpreadSheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }
        }

        public int InsertSharedStringItem(string value)
        {
            SetSharedString();

            if (SharedStringTablePart.SharedStringTable == null)
            {
                SharedStringTablePart.SharedStringTable  = new SharedStringTable();
            }

            int index = 0;

            foreach(SharedStringItem item in SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == value)
                {
                    return index;
                }
                index++;
            }

            SharedStringTablePart.SharedStringTable.AppendChild<SharedStringItem>(new SharedStringItem(value));
            SharedStringTablePart.SharedStringTable.Save();

            return index;
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

        public void Append(Row row)
        {
            SheetData.Append(row);
        }
    }

    public static class ExcelHelperExtensions
    {
        public static void Populate(this ExcelHelper excel)
        {

        }
    }
}
