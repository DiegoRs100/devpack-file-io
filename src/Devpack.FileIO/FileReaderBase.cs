using Devpack.Extensions.Types;
using Devpack.FileIO.Adapters.Excel;
using Devpack.FileIO.Interfaces;
using FluentValidation.Results;
using NPOI.SS.UserModel;

namespace Devpack.FileIO
{
    public abstract class FileReaderBase<TClass> : IDisposable
        where TClass : Importable<TClass>, new()
    {
        public int AtualLine { get; protected set; }
        protected int ErrorsCount;

        protected readonly IImportFileMapping<TClass> _importMapping;

        public abstract int RowsCount { get; }
        public abstract IDictionary<int, string> Header { get; }

        public bool IsValid => _validationResult.IsValid;

        protected readonly ValidationResult _validationResult = new();
        public IReadOnlyCollection<ValidationFailure> ValidationErrors => _validationResult.Errors;

        protected int _logFileErrorsColumIndex;
        protected readonly ExcelHelper _importLog;

        public bool HasImportErrors => ErrorsCount > 0;

        protected FileReaderBase(Stream importFile)
        {
            using var logMemoryStream = FileReaderBase<TClass>.CreateLogMemoryStream(importFile);

            _importLog = new ExcelHelper(logMemoryStream);
            _importMapping = new TClass().HeaderMapping;

            InitLogFile();
        }

        protected FileReaderBase(Stream importFile, string logFileSheetName)
        {
            using var logMemoryStream = FileReaderBase<TClass>.CreateLogMemoryStream(importFile);

            _importLog = new ExcelHelper(logMemoryStream, logFileSheetName);
            _importMapping = new TClass().HeaderMapping;

            InitLogFile();
        }

        public MemoryStream GetLogAsMemoryStream()
        {
            return _importLog.GetAsMemoryStream();
        }

        public void SaveLog(string path)
        {
            using var fileStream = _importLog.GetAsFileStream(path);
            fileStream.Close();
        }

        public virtual ValidationResult Validate()
        {
            var hasNotMappedRequiredTag = _importMapping.Tags
                .Any(h => h.Options == ImportOptions.Required && h.FileColumnIndex == null);

            if (hasNotMappedRequiredTag)
            {
                _validationResult.Errors.Add(new ValidationFailure(null, ImportFileErrors.InvalidHeaderMessage)
                {
                    ErrorCode = ImportFileErrors.InvalidHeaderKey
                });
            }

            if (Header.Values.IsDuplicated(h => h.ToLower()))
            {
                _validationResult.Errors.Add(new ValidationFailure(null, ImportFileErrors.DuplicatedTagsMessage)
                {
                    ErrorCode = ImportFileErrors.DuplicatedTagsKey
                });
            }

            return _validationResult;
        }

        protected void InitializeHeader()
        {
            foreach (var tag in Header)
            {
                var mappedTag = _importMapping.Tags.FirstOrDefault(t => t.Name.Match(tag.Value));
                mappedTag?.SetFileColumnIndex(tag.Key);
            }

            _logFileErrorsColumIndex = GetLogFileErrorsColumIndex();
        }

        private int GetLogFileErrorsColumIndex()
        {
            var headerRow = _importLog.GetRow(0);
            return headerRow?.Cells.Find(x => x.StringCellValue == "Erros")?.ColumnIndex ?? 0;
        }

        private void InitLogFile()
        {
            if (_importLog.RowsCount < 1)
                return;

            while (_importLog.RowsCount > 1)
                _importLog.RemoveRow(_importLog.RowsCount - 1);

            var headerRow = _importLog.GetRow(0);

            if (!headerRow.Any(c => c.CellType == CellType.String && c.StringCellValue == "Erros"))
            {
                var errorCell = headerRow.CreateCell(headerRow.Count());
                errorCell.SetCellValue("Erros");
            }
        }

        private static MemoryStream CreateLogMemoryStream(Stream importFile)
        {
            var logMemoryStream = new MemoryStream();
            importFile.CopyTo(logMemoryStream);

            importFile.Position = 0;
            logMemoryStream.Position = 0;

            return logMemoryStream;
        }

        public void Dispose()
        {
            _importLog?.Dispose();
        }
    }
}