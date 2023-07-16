using Devpack.Extensions.Types;
using Devpack.FileIO;
using Devpack.FileIO.Adapters.Excel;
using Devpack.FileIO.Interfaces;
using NPOI.SS.UserModel;

namespace Devpack.FileIO.Tests.Common.Extensions
{
    public static class ExcelFileImportReaderExtensions
    {
        public static IRow? GetAtualRow<TClass>(this ExcelFileImportReader<TClass> importReader)
            where TClass : Importable<TClass>, new()
        {
            return importReader.GetFieldValue("AtualRow") as IRow;
        }

        public static ExcelHelper? GetLogExcelHelper<TClass>(this ExcelFileImportReader<TClass> importReader)
            where TClass : Importable<TClass>, new()
        {
            return importReader.GetFieldValue("_importLog") as ExcelHelper;
        }

        public static IWorkbook? GetImportWorkbook<TClass>(this ExcelFileImportReader<TClass> importReader)
             where TClass : Importable<TClass>, new()
        {
            var excelHelper = importReader.GetFieldValue("_excelHelper") as ExcelHelper;
            return excelHelper!.GetFieldValue("_workbook") as IWorkbook;
        }

        public static IWorkbook? GetLogWorkbook<TClass>(this ExcelFileImportReader<TClass> importReader)
             where TClass : Importable<TClass>, new()
        {
            var importLog = importReader.GetFieldValue("_importLog") as ExcelHelper;
            return importLog!.GetFieldValue("_workbook") as IWorkbook;
        }

        public static IImportFileMapping<TClass>? GetImportMapping<TClass>(this ExcelFileImportReader<TClass> importReader)
             where TClass : Importable<TClass>, new()
        {
            return importReader.GetFieldValue("_importMapping") as IImportFileMapping<TClass>;
        }

        public static int LogFileErrorsColumIndex<TClass>(this ExcelFileImportReader<TClass> importReader)
             where TClass : Importable<TClass>, new()
        {
            return (int)importReader.GetFieldValue("_logFileErrorsColumIndex")!;
        }
    }
}