using Devpack.Extensions.Types;
using Devpack.FileIO.Adapters.Excel;
using Devpack.FileIO.Interfaces;
using Devpack.FileIO.Tests.Common.Factories;
using FluentAssertions;
using FluentAssertions.Primitives;
using FluentAssertions.Streams;
using NPOI.SS.UserModel;

namespace Devpack.FileIO.Tests.Common.Asserts
{
    public static class ExcelFileImportReaderAsserts
    {
        public static AndConstraint<TAssertions> BeImportFile<TAssertions>(
            this ObjectAssertions<object, TAssertions> objectAssertion)
            where TAssertions : ObjectAssertions<object, TAssertions>
        {
            var subject = objectAssertion.Subject as ExcelFileImportReader<EntityValidImportableTest>;
            var excelHelper = subject!.GetFieldValue("_excelHelper") as ExcelHelper;
            var workbook = excelHelper?.GetFieldValue("_workbook") as IWorkbook;

            workbook!.NumberOfSheets.Should().Be(4);
            return new AndConstraint<TAssertions>((TAssertions)objectAssertion);
        }

        public static AndConstraint<TAssertions> BeLogFile<TAssertions>(
            this StreamAssertions<Stream, TAssertions> objectAssertion)
            where TAssertions : StreamAssertions<Stream, TAssertions>
        {
            using var excelHelper = new ExcelHelper(objectAssertion.Subject);
            var workbook = excelHelper.GetFieldValue("_workbook") as IWorkbook;

            workbook!.NumberOfSheets.Should().Be(4);
            return new AndConstraint<TAssertions>((TAssertions)objectAssertion);
        }

        public static AndConstraint<TAssertions> UseWorksheet<TAssertions>(
            this ObjectAssertions<object, TAssertions> objectAssertion, string workSheetName)
            where TAssertions : ObjectAssertions<object, TAssertions>
        {
            var subject = objectAssertion.Subject as ExcelFileImportReader<EntityValidImportableTest>;
            var excelHelper = subject!.GetFieldValue("_excelHelper") as ExcelHelper;
            var workSheet = excelHelper!.GetFieldValue("_sheet") as ISheet;

            workSheet!.SheetName.Should().Be(workSheetName);
            return new AndConstraint<TAssertions>((TAssertions)objectAssertion);
        }

        public static AndConstraint<TAssertions> HasOnlyHeader<TAssertions>(
            this ObjectAssertions<object, TAssertions> objectAssertion)
            where TAssertions : ObjectAssertions<object, TAssertions>
        {
            var subject = objectAssertion.Subject as ExcelHelper;

            subject!.RowsCount.Should().Be(1);
            return new AndConstraint<TAssertions>((TAssertions)objectAssertion);
        }

        public static AndConstraint<TAssertions> MapHeaderCorrectly<TAssertions>(
            this ObjectAssertions<object, TAssertions> objectAssertion)
            where TAssertions : ObjectAssertions<object, TAssertions>
        {
            var subject = objectAssertion.Subject as IImportFileMapping<EntityValidImportableTest>;

            subject!.Tags.Should().NotContain(t => t.FileColumnIndex == null);
            subject!.Tags.FirstOrDefault(t => t.Name.Match("Coluna A"))?.FileColumnIndex.Should().Be(0);
            subject!.Tags.FirstOrDefault(t => t.Name.Match("Coluna B"))?.FileColumnIndex.Should().Be(1);
            subject!.Tags.FirstOrDefault(t => t.Name.Match("Coluna D"))?.FileColumnIndex.Should().Be(3);
            subject!.Tags.FirstOrDefault(t => t.Name.Match("Coluna C"))?.FileColumnIndex.Should().Be(4);
            subject!.Tags.FirstOrDefault(t => t.Name.Match("Coluna E"))?.FileColumnIndex.Should().Be(5);
            subject!.Tags.FirstOrDefault(t => t.Name.Match("Coluna F"))?.FileColumnIndex.Should().Be(6);
            subject!.Tags.FirstOrDefault(t => t.Name.Match("Coluna G"))?.FileColumnIndex.Should().Be(7);

            return new AndConstraint<TAssertions>((TAssertions)objectAssertion);
        }

        public static AndConstraint<TAssertions> BeEquivalentToImportFileRecord1<TAssertions>(
            this ObjectAssertions<object, TAssertions> objectAssertion)
            where TAssertions : ObjectAssertions<object, TAssertions>
        {
            var subject = objectAssertion.Subject as EntityValidImportableTest;

            subject!.ColunaA.Should().Be("Linha A1");
            subject!.ColunaB.Should().Be("Linha B1");
            subject!.ColunaC.Should().BeNull();
            subject!.ColunaD.Should().Be(1);
            subject!.ColunaE.Should().Be(new DateTime(2022, 04, 28, 17, 30, 0));
            subject!.ColunaF.Should().Be(new DateTime(2022, 04, 27, 17, 15, 0));
            subject!.ColunaG.Should().BeTrue();
            subject!.SourceImportRowIndex.Should().Be(1);

            return new AndConstraint<TAssertions>((TAssertions)objectAssertion);
        }

        public static AndConstraint<TAssertions> BeEquivalentToImportFileRecord2<TAssertions>(
            this ObjectAssertions<object, TAssertions> objectAssertion)
            where TAssertions : ObjectAssertions<object, TAssertions>
        {
            var subject = objectAssertion.Subject as EntityValidImportableTest;

            subject!.ColunaA.Should().Be("123456");
            subject!.ColunaB.Should().Be("Linha B2");
            subject!.ColunaC.Should().Be("Linha C2");
            subject!.ColunaD.Should().BeNull();
            subject!.ColunaE.Should().Be(new DateTime(2021, 05, 26));
            subject!.ColunaF.Should().Be(new DateTime(2022, 04, 27));
            subject!.ColunaG.Should().BeFalse();
            subject!.SourceImportRowIndex.Should().Be(2);

            return new AndConstraint<TAssertions>((TAssertions)objectAssertion);
        }

        public static AndConstraint<TAssertions> BeEquivalentToImportFileRecord3<TAssertions>(
            this ObjectAssertions<object, TAssertions> objectAssertion)
            where TAssertions : ObjectAssertions<object, TAssertions>
        {
            var subject = objectAssertion.Subject as EntityValidImportableTest;

            subject!.ColunaA.Should().Be("Linha A3");
            subject!.ColunaB.Should().Be("Linha B3");
            subject!.ColunaC.Should().Be("Linha C3");
            subject!.ColunaD.Should().Be(3);
            subject!.ColunaE.Should().BeNull();
            subject!.ColunaF.Should().Be(new DateTime(2022, 04, 27, 17, 15, 15));
            subject!.ColunaG.Should().BeFalse();
            subject!.SourceImportRowIndex.Should().Be(3);

            return new AndConstraint<TAssertions>((TAssertions)objectAssertion);
        }

        public static AndConstraint<TAssertions> BeEquivalentToImportFileRecord4<TAssertions>(
            this ObjectAssertions<object, TAssertions> objectAssertion)
            where TAssertions : ObjectAssertions<object, TAssertions>
        {
            var subject = objectAssertion.Subject as EntityValidImportableTest;

            subject!.ColunaA.Should().Be("Linha A5");
            subject!.ColunaB.Should().Be("Linha B5");
            subject!.ColunaC.Should().Be("Linha C5");
            subject!.ColunaD.Should().Be(5);
            subject!.ColunaE.Should().BeNull();
            subject!.ColunaF.Should().Be(new DateTime(2022, 04, 27, 17, 15, 0));
            subject!.ColunaG.Should().BeTrue();

            return new AndConstraint<TAssertions>((TAssertions)objectAssertion);
        }
    }
}