using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelDemo
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

        public static void AppendStringRefCell(this Row row, string value, StringValue reference)
        {
            row.Append(new Cell()
            {
                CellValue = new CellValue(value),
                CellReference = reference,
                DataType = CellValues.String
            });
        }
    }
}
