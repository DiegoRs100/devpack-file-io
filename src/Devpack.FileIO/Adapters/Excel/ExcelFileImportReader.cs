using Devpack.Extensions.Types;
using Devpack.FileIO.Interfaces;
using Devpack.FileIO.Mapping;
using FluentValidation.Internal;
using FluentValidation.Results;
using NPOI.SS.UserModel;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Devpack.FileIO.Adapters.Excel
{
    public partial class ExcelFileImportReader<TClass> : FileReaderBase<TClass>, IFileImportReader<TClass>
        where TClass : Importable<TClass>, new()
    {
        private IRow AtualRow;

        private readonly ExcelHelper _excelHelper;
        private int? _totalValidRegisters;

        [GeneratedRegex("^(true|yes|y|sim|s|1)$", RegexOptions.IgnoreCase, "pt-BR")]
        private static partial Regex TrueBolleanRegex();

        [GeneratedRegex("^(false|no|n|nao|não|0)$", RegexOptions.IgnoreCase, "pt-BR")]
        private static partial Regex FalseBolleanRegex();

        public override int RowsCount => _excelHelper.RowsCount;
        public override IDictionary<int, string> Header => GetHeader();

        public ExcelFileImportReader(Stream importFile) : base(importFile)
        {
            _excelHelper = new ExcelHelper(importFile);
            AtualRow = _excelHelper.GetRow(AtualLine);

            InitializeHeader();
        }

        public ExcelFileImportReader(Stream importFile, string importFileSheetName)
            : base(importFile, importFileSheetName)
        {
            _excelHelper = new ExcelHelper(importFile, importFileSheetName);
            AtualRow = _excelHelper.GetRow(AtualLine);

            InitializeHeader();
        }

        // A cada chamada desse método, a próxima linha do arquivo é lida.
        // Desconsidera o cabeçalho.
        // Remove linhas em branco.
        // Retorna nulo quando a linha não existe.
        // Retorna um dicionário vázio quando a linha existe mas todos as suas células estão em branco.
        public TClass ReadLine()
        {
            AtualLine++;

            if (AtualLine > _excelHelper.RowsCount - 1)
                return default!;

            AtualRow = _excelHelper.GetRow(AtualLine);

            if (AtualRow.IsNullOrEmpty())
                return ReadLine();

            var entity = new TClass();

            foreach (var tag in _importMapping.Tags.Where(t => t.FileColumnIndex != null))
            {
                var cell = AtualRow.GetCell((int)tag.FileColumnIndex!);
                var cellValue = cell?.GetCellValue();

                if (cellValue == null)
                    continue;

                SetValueInOutputEntity(entity, tag, cellValue);
            }

            return entity;
        }

        public IEnumerable<TClass> ReadAllLines()
        {
            var result = new List<TClass>();
            TClass auxEntity;
            
            while (true)
            {
                auxEntity = ReadLine();

                if (auxEntity == null)
                    break;

                result.Add(auxEntity);
            }

            return result;
        }

        /// <summary>Insere o log considerando a ultima linha que foi lida no arquivo.</summary>
        public void SetErrorsInLogFile(IEnumerable<ValidationFailure> errors)
        {
            SetErrorsInLogFile(AtualRow, errors);
        }

        public void SetErrorsInLogFile(int rowIndex, IEnumerable<ValidationFailure> errors)
        {
            var row = _excelHelper.GetRow(rowIndex);
            SetErrorsInLogFile(row, errors);
        }

        public int CountTotalValidRegisters()
        {
            if (_totalValidRegisters.HasValue)
                return _totalValidRegisters.Value;

            _totalValidRegisters = 0;

            for (var i = 1; i < RowsCount; i++)
            {
                var row = _excelHelper.GetRow(i);

                if (!row.IsNullOrEmpty())
                    _totalValidRegisters++;
            }

            return _totalValidRegisters.Value;
        }

        public override ValidationResult Validate()
        {
            if (CountTotalValidRegisters() == 0)
            {
                _validationResult.Errors.Add(new ValidationFailure(null, ImportFileErrors.ErrorEmptySheetMessage)
                {
                    ErrorCode = ImportFileErrors.ErrorEmptySheetKey
                });
            }

            return base.Validate();
        }

        private void SetErrorsInLogFile(IRow row, IEnumerable<ValidationFailure> errors)
        {
            var text = string.Join(" | ", errors.Select(e => e.ErrorMessage));

            var newRow = _importLog.CloneRow(row, ++ErrorsCount);
            newRow.SetCellValue(_logFileErrorsColumIndex, text);

            newRow.Cells[^1].CellStyle.VerticalAlignment = VerticalAlignment.Center;
            newRow.Cells[^1].CellStyle.IsLocked = true;
        }

        private IDictionary<int, string> GetHeader()
        {
            var result = new Dictionary<int, string>();
            var headerRow = _excelHelper.GetRow(0);

            if (headerRow == null)
                return result;

            foreach (var item in headerRow)
                result.Add(item.ColumnIndex, item.StringCellValue);

            return result;
        }

        private void SetValueInOutputEntity(object entity, Tag<TClass> tag, object value)
        {
            var propertyMember = tag.Property.GetMember();
            var propertyInfo = typeof(TClass).GetProperty(propertyMember.Name)!;

            if (propertyInfo.PropertyType.In(typeof(bool), typeof(bool?)))
                SetCellValueInBoolean(propertyInfo, entity, value);

            else if (propertyInfo.PropertyType.In(typeof(DateTime), typeof(DateTime?)))
                SetCellValueInDateTime(propertyInfo, entity, value);

            else if (propertyInfo.PropertyType.In(typeof(int), typeof(int?), typeof(double), typeof(double?), typeof(decimal), typeof(decimal?))
                && value is double)
            {
                propertyInfo.SetValue(entity, value);
            }

            else if (propertyInfo.PropertyType.In(typeof(string)))
            {
                if (value is string)
                    propertyInfo.SetValue(entity, value);
                else
                    propertyInfo.SetValue(entity, value.ToString());
            }

            (entity as Importable<TClass>)!.SetSourceImportRowIndex(AtualRow.RowNum);
        }

        private static void SetCellValueInDateTime(PropertyInfo propertyInfo, object entity, object value)
        {
            if (value is DateTime)
            {
                propertyInfo.SetValue(entity, value);
                return;
            }

            if (value is string)
            {
                DateTime.TryParseExact(value.ToString(), new[]
                {
                    "dd-MM-yyyy HH:mm:ss",
                    "dd/MM/yyyy HH:mm:ss",
                    "yyyy-MM-dd HH:mm:ss",
                    "yyyy/MM/dd HH:mm:ss",
                    "dd-MM-yyyy HH:mm",
                    "dd/MM/yyyy HH:mm",
                    "yyyy-MM-dd HH:mm",
                    "yyyy/MM/dd HH:mm",
                    "dd-MM-yyyy",
                    "dd-MM-yyyy",
                    "yyyy-MM-dd",
                    "yyyy-MM-dd"
                }, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateTime);

                if (dateTime == DateTime.MinValue)
                    return;

                propertyInfo.SetValue(entity, dateTime);
            }
        }

        private static void SetCellValueInBoolean(PropertyInfo propertyInfo, object entity, object value)
        {
            if (value is string stringValue)
            {
                if (TrueBolleanRegex().IsMatch(stringValue.Trim()))
                    propertyInfo.SetValue(entity, true);

                else if (FalseBolleanRegex().IsMatch(stringValue.Trim()))
                    propertyInfo.SetValue(entity, false);
            }

            if (value is double doubleValue && doubleValue.In(0, 1))
                propertyInfo.SetValue(entity, Convert.ToBoolean(doubleValue));
        }

        new public void Dispose()
        {
            _excelHelper?.Dispose();

            base.Dispose();
        }
    }
}