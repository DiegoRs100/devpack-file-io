using Devpack.Extensions.Types;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics.CodeAnalysis;

namespace Devpack.FileIO.Adapters.Excel
{
    public class ExcelHelper : IDisposable
    {
        private readonly IWorkbook _workbook;
        private readonly ISheet _sheet;

        public int RowsCount => _sheet.LastRowNum + 1;

        public ExcelHelper(Stream stream, int sheetIndex = 0)
        {
            _workbook = new XSSFWorkbook(stream);
            _sheet = _workbook.GetSheetAt(sheetIndex);
        }

        public ExcelHelper(Stream stream, string sheetName)
        {
            _workbook = new XSSFWorkbook(stream);
            _sheet = _workbook.GetSheet(sheetName);
        }

        [ExcludeFromCodeCoverage]
        public void SetCellSelection(int rowIndex, int columnIndex)
        {
            _sheet.SetActiveCell(rowIndex, columnIndex);
        }

        [ExcludeFromCodeCoverage]
        public void ProtectSheet(string sheetName, string password)
        {
            var sheet = _workbook.GetSheet(sheetName);
            sheet.ProtectSheet(password);
        }

        public IRow GetRow(int index, string? sheetName = null)
        {
            var sheet = sheetName!.IsNullOrEmpty()
                ? _sheet
                : _workbook.GetSheet(sheetName);

            return sheet.GetRow(index);
        }

        public IRow CloneRow(IRow row, int indexRow)
        {
            var newRow = _sheet.GetRow(indexRow) ?? _sheet.CreateRow(indexRow);

            foreach (var item in row.Cells)
                item.CloneInRow(newRow);

            return newRow;
        }

        public void RemoveRow(int index)
        {
            _sheet.RemoveRow(_sheet.GetRow(index));
        }

        public MemoryStream GetAsMemoryStream()
        {
            var stream = new MemoryStream();
            _workbook.Write(stream);

            stream.Position = 0;

            return stream;
        }

        public FileStream GetAsFileStream(string path)
        {
            var stream = new FileStream(path.ToString(), FileMode.OpenOrCreate, FileAccess.ReadWrite);
            _workbook.Write(stream);

            stream.Position = 0;

            return stream;
        }

        public void Dispose()
        {
            _workbook.Close();
        }
    }
}