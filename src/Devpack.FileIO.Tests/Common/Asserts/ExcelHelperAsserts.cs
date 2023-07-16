using Devpack.Extensions.Types;
using Devpack.FileIO.Adapters.Excel;
using FluentAssertions;
using FluentAssertions.Streams;
using NPOI.SS.UserModel;

namespace Devpack.FileIO.Tests.Common.Asserts
{
    public static class ExcelHelperAsserts
    {
        public static AndConstraint<TAssertions> BeImportFile<TAssertions>(
            this StreamAssertions<Stream, TAssertions> objectAssertion)
            where TAssertions : StreamAssertions<Stream, TAssertions>
        {
            using var excelHelper = new ExcelHelper(objectAssertion.Subject);
            var workbook = excelHelper.GetFieldValue("_workbook") as IWorkbook;

            workbook!.NumberOfSheets.Should().Be(4);
            return new AndConstraint<TAssertions>((TAssertions)objectAssertion);
        }
    }
}