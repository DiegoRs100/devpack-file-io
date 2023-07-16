using Devpack.FileIO.Adapters.Excel;
using Devpack.FileIO.Tests.Common.Asserts;
using Devpack.FileIO.Tests.Common.Extensions;
using Devpack.FileIO.Tests.Common.Factories;
using Devpack.FileIO.Tests.Resources;
using FluentAssertions;
using FluentValidation.Results;
using NPOI;
using Xunit;

namespace Devpack.FileIO.Tests
{
    public class FileReaderBaseTests
    {
        [Fact(DisplayName = "Deve retornar verdadeiro quando Não houverem erros de validação na importação do arquivo.")]
        [Trait("Category", "Helpers")]
        public void IsValid_BeTrue()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            importReader.IsValid.Should().BeTrue();
        }

        [Fact(DisplayName = "Deve retornar falso quando houverem erros de validação na importação do arquivo.")]
        [Trait("Category", "Helpers")]
        public void IsValid_BeFalse()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateInvalidImportReader();
            importReader.Validate();

            importReader.IsValid.Should().BeFalse();
        }

        [Fact(DisplayName = "Deve retornar verdadeiro quando existirem erros nos dados importados.")]
        [Trait("Category", "Helpers")]
        public void HasImportErrors_BeTrue()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            importReader.SetErrorsInLogFile(new List<ValidationFailure> { new ValidationFailure(null, "test") });

            importReader.HasImportErrors.Should().BeTrue();
        }

        [Fact(DisplayName = "Deve retornar falso quando não existirem erros nos dados importados.")]
        [Trait("Category", "Helpers")]
        public void HasImportErrors_BeFalse()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            importReader.HasImportErrors.Should().BeFalse();
        }

        [Fact(DisplayName = "Deve intanciar a classe corretamente quando usado o WorkSheetName default.")]
        [Trait("Category", "Helpers")]
        public void Constructor_WhenLogFileWorkSheetDefault()
        {
            using var importData = new MemoryStream(Resource.ImportFile);
            var importReader = new ExcelFileImportReader<EntityValidImportableTest>(importData);

            var importMapping = importReader.GetImportMapping();
            var logFileErrorsColumIndex = importReader.LogFileErrorsColumIndex();
            var importLog = importReader.GetLogExcelHelper();

            using var logFileMemoryStream = importReader.GetLogAsMemoryStream();

            logFileMemoryStream.Should().BeLogFile();
            logFileErrorsColumIndex.Should().Be(8);
            importLog!.Should().HasOnlyHeader();
            importMapping.Should().MapHeaderCorrectly();
        }

        [Fact(DisplayName = "Deve intanciar a classe corretamente quando passado um WorkSheetName customizado.")]
        [Trait("Category", "Helpers")]
        public void Constructor_WhenCustomLogFileWorkSheet_And_FileHasRows()
        {
            using var importData = new MemoryStream(Resource.ImportFile);
            var importReader = new ExcelFileImportReader<EntityValidImportableTest>(importData, "Planilha1");

            var importMapping = importReader.GetImportMapping();
            var logFileErrorsColumIndex = importReader.LogFileErrorsColumIndex();
            var importLog = importReader.GetLogExcelHelper();

            using var logFileMemoryStream = importReader.GetLogAsMemoryStream();

            logFileMemoryStream.Should().BeLogFile();
            logFileErrorsColumIndex.Should().Be(8);
            importLog!.Should().HasOnlyHeader();
            importMapping.Should().MapHeaderCorrectly();
        }

        [Fact(DisplayName = "Deve escrever uma coluna de erros no arquivo de log quando o arquivo de importação não tiver uma.")]
        [Trait("Category", "Helpers")]
        public void Constructor_WhenImportFilesNotContainsErrorColumn()
        {
            using var importData = new MemoryStream(Resource.ImportFile);
            var importReader = new ExcelFileImportReader<EntityValidImportableTest>(importData, "Planilha2");

            var logFileErrorsColumIndex = importReader.LogFileErrorsColumIndex();
            var importLog = importReader.GetLogExcelHelper();
            using var logFileMemoryStream = importReader.GetLogAsMemoryStream();

            logFileErrorsColumIndex.Should().Be(2);
            importLog!.Should().HasOnlyHeader();
        }

        [Fact(DisplayName = "Deve obter o (MemoryStream) do arquivo de log quando o método for chamado.")]
        [Trait("Category", "Helpers")]
        public void GetLogAsMemoryStream()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            using var logFileMemoryStream = importReader.GetLogAsMemoryStream();

            logFileMemoryStream.Should().BeLogFile();
        }

        [Fact(DisplayName = "Deve salvar um arquivo de log no disco quando o caminho for passado.")]
        [Trait("Category", "Helpers")]
        public void SaveLog()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();

            importReader.SaveLog("logTest.xlsx");
            using var fileStream = new FileStream("logTest.xlsx", FileMode.Open, FileAccess.Read);

            fileStream.Should().BeLogFile();
        }

        [Fact(DisplayName = "Deve gerar erros de validação quando o arquivo de importação não for válido.")]
        [Trait("Category", "Helpers")]
        public void Validate_NotValid()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateInvalidImportReader("Planilha2");

            importReader.Validate().Errors.Should().HaveCount(3)
                .And.Contain(e => e.ErrorCode == "EmptySheet" && e.ErrorMessage == "The supplied file contains no records.")
                .And.Contain(e => e.ErrorCode == "InvalidHeader" && e.ErrorMessage == "The supplied file does not have all the required columns.")
                .And.Contain(e => e.ErrorCode == "DuplicatedTags" && e.ErrorMessage == "The supplied file has duplicate identifiers in the header.");
        }

        [Fact(DisplayName = "Deve retornar uma lista vazia quando o arquivo de importação for válido.")]
        [Trait("Category", "Helpers")]
        public void Validate_Valid()
        {
            using var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            importReader.Validate().IsValid.Should().BeTrue();
        }

        [Fact(DisplayName = "Deve destruir o objeto de log quando o método for chamado.")]
        [Trait("Category", "Helpers")]
        public void Dispose()
        {
            var importReader = ExcelFileImportReaderFactory.CreateValidImportReader();
            var importLogWorkbook = importReader.GetLogWorkbook();

            importReader!.Dispose();

            importLogWorkbook.Invoking(er => er!.Write(new MemoryStream()))
                .Should().Throw<POIXMLException>()
                .WithInnerException<NullReferenceException>();
        }
    }
}