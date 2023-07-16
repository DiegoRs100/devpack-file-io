using Bogus;
using Devpack.Extensions.Types;
using Devpack.FileIO.Adapters.Excel;
using Devpack.FileIO.Tests.Common.Asserts;
using Devpack.FileIO.Tests.Common.Extensions;
using Devpack.FileIO.Tests.Common.Factories;
using Devpack.FileIO.Tests.Resources;
using FluentAssertions;
using FluentValidation.Results;
using NPOI;
using NPOI.SS.UserModel;
using Xunit;

namespace Devpack.FileIO.Tests
{
    public class ExcelFileImporteReaderTests
    {
        [Fact(DisplayName = "Deve retornar a quantidade de registros do arquivo de importação quando o método for chamado.")]
        [Trait("Category", "Helpers")]
        public void RowsCount()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            importReader.Should().BeImportFile();
        }

        [Fact(DisplayName = "Deve retornar as tags de importação quando o arquivo de importação possuir um header.")]
        [Trait("Category", "Helpers")]
        public void Header()
        {
            var ExpectedHeader = new Dictionary<int, string>()
            {
                { 0, "Coluná A" },
                { 1, "Coluna B" },
                { 2, "Discard" },
                { 3, "Coluna D" },
                { 4, "colunA C" },
                { 5, "Coluna E" },
                { 6, "Coluna F" },
                { 7, "Coluna G" },
                { 8, "Erros" },
            };
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();

            importReader.Header.Should().BeEquivalentTo(ExpectedHeader);
        }

        [Fact(DisplayName = "Deve intanciar a classe corretamente quando usado os (WorkSheets) default.")]
        [Trait("Category", "Helpers")]
        public void Constructor_WhenWorkSheetsDefault()
        {
            using var importData = new MemoryStream(Resource.ImportFile);
            var importReader = new ExcelFileImportReader<EntityValidImportableTest>(importData);

            importReader.Should().BeImportFile().And.UseWorksheet("Planilha1");
            importReader.GetAtualRow()!.RowNum.Should().Be(0);
        }

        [Fact(DisplayName = "Deve intanciar a classe corretamente quando usado os (WorkSheets) default.")]
        [Trait("Category", "Helpers")]
        public void Constructor_WhenCustomWorkSheets()
        {
            using var importData = new MemoryStream(Resource.ImportFile);
            var importReader = new ExcelFileImportReader<EntityValidImportableTest>(importData, "Planilha3");

            importReader.Should().BeImportFile().And.UseWorksheet("Planilha3");
            importReader.GetAtualRow()!.RowNum.Should().Be(0);
        }

        [Fact(DisplayName = "Deve retornar próxima linha do arquivo de importação quando existirem registros válidos.")]
        [Trait("Category", "Helpers")]
        public void ReadLine_WhenValidRegisters()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();

            var line1 = importReader.ReadLine();
            var line2 = importReader.ReadLine();
            var line3 = importReader.ReadLine();
            var line4 = importReader.ReadLine();

            line1.Should().BeEquivalentToImportFileRecord1();
            line2.Should().BeEquivalentToImportFileRecord2();
            line3.Should().BeEquivalentToImportFileRecord3();
            line4.Should().BeEquivalentToImportFileRecord4();
            line4.SourceImportRowIndex.Should().Be(5);
        }

        [Fact(DisplayName = "Deve desconsiderar próxima linha do arquivo de importação quando não existirem valores nela.")]
        [Trait("Category", "Helpers")]
        public void ReadLine_WhenBlankLine()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();

            importReader.ReadLine();
            importReader.ReadLine();

            var line3 = importReader.ReadLine();
            var line4 = importReader.ReadLine();

            line3.Should().BeEquivalentToImportFileRecord3();
            line3.SourceImportRowIndex.Should().Be(3);

            line4.Should().BeEquivalentToImportFileRecord4();
            line4.SourceImportRowIndex.Should().Be(5);
        }

        [Fact(DisplayName = "Deve retornar nulo quando não existirem mais linhas restantes no arquivo de importação.")]
        [Trait("Category", "Helpers")]
        public void ReadLine_WhenEndFile()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();

            for (var i = 0; i < 4; i++)
                importReader.ReadLine();

            var line5 = importReader.ReadLine();

            line5.Should().BeNull();
        }

        [Fact(DisplayName = "deve ler todas as linhas do arquivo e devolver uma lista de objetos quando o arquivo de importação tiver registros.")]
        [Trait("Category", "Helpers")]
        public void ReadAllLine()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();

            var entities = importReader.ReadAllLines().ToList();

            entities.Should().HaveCount(4);
            entities[0].Should().BeEquivalentToImportFileRecord1();
            entities[1].Should().BeEquivalentToImportFileRecord2();
            entities[2].Should().BeEquivalentToImportFileRecord3();
            entities[3].Should().BeEquivalentToImportFileRecord4();
        }

        [Fact(DisplayName = "Deve inserir um registro de erro no arquivo de log considerando a linha atual quando o método for chamado.")]
        [Trait("Category", "Helpers")]
        public void SetErrorsInLogFile()
        {
            var errors = new List<ValidationFailure>()
            {
                new ValidationFailure(null, "Erro 01"),
                new ValidationFailure(null, "Erro 02")
            };

            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            var logExcelHelper = importReader.GetLogExcelHelper();

            importReader.ReadLine();
            importReader.SetErrorsInLogFile(errors);

            var logLine1 = logExcelHelper!.GetRow(1);

            logLine1.Should().HaveCount(9);
            logLine1.Cells[0].GetCellValue().Should().Be("Linha A1");
            logLine1.Cells[1].GetCellValue().Should().Be("Linha B1");
            logLine1.Cells[2].GetCellValue().Should().BeNull();
            logLine1.Cells[3].GetCellValue().Should().Be(1);
            logLine1.Cells[4].GetCellValue().Should().BeNull();
            logLine1.Cells[5].GetCellValue().Should().BeEquivalentTo(new DateTime(2022, 4, 28, 17, 30, 0));
            logLine1.Cells[6].GetCellValue().Should().Be("27-04-2022 17:15");
            logLine1.Cells[7].GetCellValue().Should().Be("sim");
            logLine1.Cells[8].GetCellValue().Should().Be("Erro 01 | Erro 02");

            logLine1.Cells[8].CellStyle.IsLocked.Should().BeTrue();
            logLine1.Cells[8].CellStyle.VerticalAlignment.Should().Be(VerticalAlignment.Center);
        }

        [Fact(DisplayName = "Deve inserir um registro de erro no arquivo de log quando o index de uma linha específica for passado.")]
        [Trait("Category", "Helpers")]
        public void SetErrorsInLogFile_UsingRowIndex()
        {
            var errors = new List<ValidationFailure>()
            {
                new ValidationFailure(null, "Erro 01"),
                new ValidationFailure(null, "Erro 02")
            };

            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            var logExcelHelper = importReader.GetLogExcelHelper();

            importReader.ReadLine();
            var line2 = importReader.ReadLine();
            importReader.ReadLine();

            importReader.SetErrorsInLogFile(line2.SourceImportRowIndex, errors);

            var logLine1 = logExcelHelper!.GetRow(1);

            logLine1.Should().HaveCount(9);
            logLine1.Cells[0].GetCellValue().Should().Be(123456);
            logLine1.Cells[1].GetCellValue().Should().Be("Linha B2");
            logLine1.Cells[2].GetCellValue().Should().Be(2);
            logLine1.Cells[3].GetCellValue().Should().BeNull();
            logLine1.Cells[4].GetCellValue().Should().Be("Linha C2");
            logLine1.Cells[5].GetCellValue().Should().Be(new DateTime(2021, 05, 26));
            logLine1.Cells[6].GetCellValue().Should().Be("27-04-2022");
            logLine1.Cells[7].GetCellValue().Should().Be("não ");
            logLine1.Cells[8].GetCellValue().Should().Be("Erro 01 | Erro 02");

            logLine1.Cells[8].CellStyle.IsLocked.Should().BeTrue();
            logLine1.Cells[8].CellStyle.VerticalAlignment.Should().Be(VerticalAlignment.Center);
        }

        [Fact(DisplayName = "Deve retornar a quantidade de registros válidos quando o arquivo de importação for carregado corretamente.")]
        [Trait("Category", "Helpers")]
        public void CountTotalValidRegisters()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            importReader.CountTotalValidRegisters().Should().Be(4);
        }

        [Fact(DisplayName = "Não deve ler o arquivo de importação novamente quando já existe um valor setado no contador.")]
        [Trait("Category", "Helpers")]
        public void CountTotalValidRegisters_WhenAlreadyPopulated()
        {
            var faker = new Faker();
            var count = faker.Random.Number(200, 300);

            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            importReader.SetFieldValue("_totalValidRegisters", count);

            importReader.CountTotalValidRegisters().Should().Be(count);
        }

        [Fact(DisplayName = "Deve destruir os objetos de importação e log quando o método for chamado.")]
        [Trait("Category", "Helpers")]
        public void Dispose()
        {
            var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();

            var importFileWorkbook = importReader.GetImportWorkbook();
            var importLogWorkbook = importReader.GetLogWorkbook();

            importReader!.Dispose();

            importFileWorkbook.Invoking(er => er!.Write(new MemoryStream()))
                .Should().Throw<POIXMLException>()
                .WithInnerException<NullReferenceException>();

            importLogWorkbook.Invoking(er => er!.Write(new MemoryStream()))
                .Should().Throw<POIXMLException>()
                .WithInnerException<NullReferenceException>();
        }
    }
}