using Devpack.FileIO.Adapters.Excel;
using Devpack.FileIO.Tests.Common.Asserts;
using Devpack.FileIO.Tests.Common.Extensions;
using Devpack.FileIO.Tests.Common.Factories;
using Devpack.FileIO.Tests.Resources;
using FluentAssertions;
using NPOI;
using Xunit;

namespace Devpack.FileIO.Tests
{
    public class ExcelHelperTests
    {
        [Fact(DisplayName = "Deve assumir a primeira WorkSheet do arquivo quando não for passado um index customizado no construtor.")]
        [Trait("Category", "Helpers")]
        public void Constructor_Success_WhenDefaultWorkSheet()
        {
            // Testa simultaneamente o contrutor e a propriedade RowsCount

            using var importFileMemoryStream = new MemoryStream(Resource.ImportFile);
            var excelHelper = new ExcelHelper(importFileMemoryStream);

            excelHelper.RowsCount.Should().Be(6);
        }

        [Fact(DisplayName = "Deve assumir a WorkSheet indicada quando for passado um index customizado no construtor.")]
        [Trait("Category", "Helpers")]
        public void Constructor_Success_WhenCustomWorkSheet()
        {
            // Testa simultaneamente o contrutor e a propriedade RowsCount

            using var importFileMemoryStream = new MemoryStream(Resource.ImportFile);
            var excelHelper = new ExcelHelper(importFileMemoryStream, 2);

            excelHelper.RowsCount.Should().Be(3);
        }

        [Fact(DisplayName = "Deve assumir a WorkSheet indicada quando for passado um nome de WorkSheet válido.")]
        [Trait("Category", "Helpers")]
        public void Constructor_Success_WhenHasWorkSheetName()
        {
            // Testa simultaneamente o contrutor e a propriedade RowsCount

            using var importFileMemoryStream = new MemoryStream(Resource.ImportFile);
            var excelHelper = new ExcelHelper(importFileMemoryStream, "Planilha3");

            excelHelper.RowsCount.Should().Be(3);
        }

        [Fact(DisplayName = "Deve retornar a linha da WorkSheet default quando o método for chamado sem passar um nome de WorkSheet.")]
        [Trait("Category", "Helpers")]
        public void GetRow_WhenDefaltSheet()
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            var row = excelHelper.GetRow(1);

            row.RowNum.Should().Be(1);
            row.GetCell(0).StringCellValue.Should().Be("Linha A1");
        }

        [Fact(DisplayName = "Deve retornar a linha da WorkSheet indicada quando o método for chamado passando um nome de WorkSheet.")]
        [Trait("Category", "Helpers")]
        public void GetRow_WhenCustomSheet()
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            var row = excelHelper.GetRow(1, "Planilha3");

            row.RowNum.Should().Be(1);
            row.GetCell(0).StringCellValue.Should().Be("Linha 2");
        }

        [Theory(DisplayName = "Deve clonar uma linha quando a linha origem for válida.")]
        [InlineData(2)]
        [InlineData(10)]
        [Trait("Category", "Helpers")]
        public void CloneRow(int targetRowIndex)
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();

            var sourceRow = excelHelper.GetRow(1);
            var newRow = excelHelper.CloneRow(sourceRow, targetRowIndex);

            var originalCells = sourceRow.Cells.Select(c => new { c.CellType, c.CellStyle.DataFormat, Value = c.GetCellValue() });
            var clonedCells = newRow.Cells.Select(c => new { c.CellType, c.CellStyle.DataFormat, Value = c.GetCellValue() });

            newRow.RowNum.Should().Be(targetRowIndex);
            originalCells.Should().BeEquivalentTo(clonedCells);
        }

        [Fact(DisplayName = "Deve remover uma linha do arwquivo quando a linha em questão existir.")]
        [Trait("Category", "Helpers")]
        public void RemoveRow()
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            excelHelper.RemoveRow(1);

            excelHelper.GetRow(1).Should().BeNull();
        }

        [Fact(DisplayName = "Deve escrever o arquivo excel na memória quando o método for chamado.")]
        [Trait("Category", "Helpers")]
        public void GetAsMemoryStream()
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            using var outputMemoryStream = excelHelper.GetAsMemoryStream();

            outputMemoryStream.Should().BeImportFile();
        }

        [Fact(DisplayName = "Deve escrever o arquivo excel em um FileStream quando um path for passado.")]
        [Trait("Category", "Helpers")]
        public void GetAsFileStream()
        {
            var path = Guid.NewGuid().ToString();

            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            using var outputFileStream = excelHelper.GetAsFileStream(path);

            outputFileStream.Should().BeImportFile();
        }

        [Fact(DisplayName = "Deve retornar dados da primeira linha do arquivo quando o mesmo possuir registros.")]
        [Trait("Category", "Helpers")]
        public void GetHeader()
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

        [Fact(DisplayName = "Deve finalizar o workbook de edição do excel quando o mesmo estiver instanciado.")]
        [Trait("Category", "Helpers")]
        public void Dispose()
        {
            var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            var workbook = excelHelper.GetWorkbook();

            excelHelper!.Dispose();

            workbook.Invoking(er => er!.Write(new MemoryStream()))
                .Should().Throw<POIXMLException>()
                .WithInnerException<NullReferenceException>();
        }
    }
}