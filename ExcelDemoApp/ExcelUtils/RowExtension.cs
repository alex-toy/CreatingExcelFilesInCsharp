using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelUtils
{
    public static class RowExtension
    {
        public static void AppendStringCell(this Row row, string value)
        {
            row.Append(new Cell()
            {
                CellValue = new CellValue(value),
                DataType = CellValues.String
            });
        }

        public static void AppendStringRefCell(this Row row, string value, StringValue reference, ExcelHelper excel = null)
        {
            int? indexValue = null;

            if (excel != null) indexValue = excel.InsertSharedStringItem(value);

            row.Append(new Cell()
            {
                CellValue = new CellValue(indexValue?.ToString() ?? value),
                CellReference = reference,
                DataType = indexValue != null ? new EnumValue<CellValues>(CellValues.SharedString) : CellValues.String
            });
        }
    }
}
