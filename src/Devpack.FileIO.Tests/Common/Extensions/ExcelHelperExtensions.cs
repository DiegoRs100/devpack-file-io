using Devpack.Extensions.Types;
using Devpack.FileIO.Adapters.Excel;
using NPOI.SS.UserModel;

namespace Devpack.FileIO.Tests.Common.Extensions
{
    public static class ExcelHelperExtensions
    {
        public static IWorkbook? GetWorkbook(this ExcelHelper excelHelper)
        {
            return excelHelper!.GetFieldValue("_workbook") as IWorkbook;
        }
    }
}