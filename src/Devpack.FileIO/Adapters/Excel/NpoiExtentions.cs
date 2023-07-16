using Devpack.Extensions.Types;
using NPOI.SS.UserModel;

namespace Devpack.FileIO.Adapters.Excel
{
    public static class NpoiExtentions
    {
        public static bool IsNullOrEmpty(this IRow row)
        {
            return row == null || !row.Cells.Exists(c => c.GetCellValue() != null);
        }

        public static void SetCellValue(this IRow row, int ColumnIndex, string value)
        {
            var cell = row.GetCell(ColumnIndex) ?? row.CreateCell(ColumnIndex);
            cell.SetCellValue(value);
        }

        public static object? GetCellValue(this ICell cell)
        {
            var valueType = cell.CellType != CellType.Formula ? cell.CellType : cell.CachedFormulaResultType;

            switch (valueType)
            {
                case CellType.String:
                    return cell.StringCellValue.EmptyAsNull();

                case CellType.Numeric:
                    {
                        if (DateUtil.IsCellDateFormatted(cell))
                            return cell.DateCellValue;
                        else
                            return cell.NumericCellValue;
                    }

                default:
                    return null;
            }
        }

        public static void CloneInRow(this ICell source, IRow targetRow)
        {
            var cell = targetRow.GetCell(source.ColumnIndex) ?? targetRow.CreateCell(source.ColumnIndex);

            var newCellStyle = targetRow.Sheet.Workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(source.CellStyle);
            
            cell.CellStyle = newCellStyle;

            var valueType = source.CellType != CellType.Formula 
                ? source.CellType 
                : source.CachedFormulaResultType;

            switch (valueType)
            {
                case CellType.String:
                    cell.SetCellValue(source.StringCellValue);
                    break;

                case CellType.Numeric:
                    cell.SetCellValue(source.NumericCellValue);
                    break;

                case CellType.Blank:
                    cell.SetCellType(CellType.Blank);
                    break;
            }
        }
    }
}