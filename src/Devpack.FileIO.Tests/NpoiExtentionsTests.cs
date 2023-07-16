using Devpack.Extensions.Types;
using Devpack.FileIO.Adapters.Excel;
using Devpack.FileIO.Tests.Common.Factories;
using FluentAssertions;
using NPOI.SS.UserModel;
using Xunit;

namespace Devpack.FileIO.Tests
{
    public class NpoiExtentionsTests
    {
        [Fact(DisplayName = "Deve retornar verdadeiro quando a linha passada nula.")]
        [Trait("Category", "Extensions")]
        public void IsNullOrEmpty_BeTrue_WhenNull()
        {
            IRow row = null!;
            row.IsNullOrEmpty().Should().BeTrue();
        }

        [Fact(DisplayName = "Deve retornar verdadeiro quando a linha passada não possuir valores.")]
        [Trait("Category", "Extensions")]
        public void IsNullOrEmpty_BeTrue_WhenEmpty()
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            var row = excelHelper.GetRow(4);

            row.IsNullOrEmpty().Should().BeTrue();
        }

        [Theory(DisplayName = "Deve pupular uma célula quando um valor válido for passado.")]
        [InlineData(2)]
        [InlineData(10)]
        [Trait("Category", "Helpers")]
        public void SetCellValue(int colIndex)
        {
            var cellValue = Guid.NewGuid().ToString();

            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            var row = excelHelper.GetRow(2);

            row.SetCellValue(colIndex, cellValue);

            var cell = row.GetCell(colIndex);
            cell.StringCellValue.Should().Be(cellValue);
        }

        [Fact(DisplayName = "Deve retornar o valor da célula quando a mesma for do tipo texto.")]
        [Trait("Category", "Helpers")]
        public void GetCellValue_WhenString()
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            var cell = excelHelper.GetRow(1).GetCell(0);

            cell.GetCellValue().Should().Be("Linha A1");
        }

        [Fact(DisplayName = "Deve retornar o valor da célula quando a mesma for do tipo numérico.")]
        [Trait("Category", "Helpers")]
        public void GetCellValue_WhenNumber()
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            var cell = excelHelper.GetRow(1).GetCell(3);

            cell.GetCellValue().Should().Be(1);
        }

        [Fact(DisplayName = "Deve retornar o valor da célula quando a mesma for do tipo data.")]
        [Trait("Category", "Helpers")]
        public void GetCellValue_WhenDateTime()
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            var cell = excelHelper.GetRow(1).GetCell(5);

            cell.GetCellValue().Should().BeEquivalentTo(new DateTime(2022, 4, 28, 17, 30, 0));
        }

        [Fact(DisplayName = "Deve retornar o valor da célula quando a mesma for uma fórmula.")]
        [Trait("Category", "Helpers")]
        public void GetCellValue_WhenFormule()
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();
            var cell = excelHelper.GetRow(5).GetCell(5);

            cell.GetCellValue().Should().Be("Linha A1");
        }

        [Fact(DisplayName = "Deve retornar nulo quando a célula não tiver um formato conhecido.")]
        [Trait("Category", "Helpers")]
        public void GetCellValue_WhenUndefinedFormat()
        {
            using var excelHelper = ExcelHelperFactory.CreateDefaultImportFiles();

            var cell = excelHelper.GetRow(0).GetCell(0);
            cell.SetCellType(CellType.Blank);

            cell.GetCellValue().Should().BeNull();
        }

        [Theory(DisplayName = "Deve clonar uma célula quando a mesma for do tipo texto.")]
        [InlineData(1)]
        [InlineData(2)]
        [Trait("Category", "Helpers")]
        public void CloneInRow_WhenString(int targetRowIndex)
        {
            using var excelHelper = ExcelHelperFactory.CreateImportFiles("Planilha4");

            var sourceCell = excelHelper.GetRow(0).GetCell(1);
            var targetRow = excelHelper.GetRow(targetRowIndex);

            sourceCell.CloneInRow(targetRow);

            var targetCell = targetRow.GetCell(sourceCell.ColumnIndex);

            targetCell.CellType.Should().Be(CellType.String);
            targetCell.GetCellValue().Should().Be(sourceCell.GetCellValue());
        }

        [Theory(DisplayName = "Deve clonar uma célula quando a mesma for do tipo numérico.")]
        [InlineData(1)]
        [InlineData(2)]
        [Trait("Category", "Helpers")]
        public void CloneInRow_WhenNumber(int targetRowIndex)
        {
            using var excelHelper = ExcelHelperFactory.CreateImportFiles("Planilha4");

            var sourceCell = excelHelper.GetRow(0).GetCell(2);
            var targetRow = excelHelper.GetRow(targetRowIndex);

            sourceCell.CloneInRow(targetRow);

            var targetCell = targetRow.GetCell(sourceCell.ColumnIndex);

            targetCell.CellType.Should().Be(CellType.Numeric);
            targetCell.GetCellValue().Should().Be(sourceCell.GetCellValue());
        }

        [Theory(DisplayName = "Deve clonar uma célula quando a mesma for do tipo data.")]
        [InlineData(1)]
        [InlineData(2)]
        [Trait("Category", "Helpers")]
        public void CloneInRow_WhenDateTime(int targetRowIndex)
        {
            using var excelHelper = ExcelHelperFactory.CreateImportFiles("Planilha4");

            var sourceCell = excelHelper.GetRow(0).GetCell(3);
            var targetRow = excelHelper.GetRow(targetRowIndex);

            sourceCell.CloneInRow(targetRow);

            var targetCell = targetRow.GetCell(sourceCell.ColumnIndex);

            targetCell.CellType.Should().Be(CellType.Numeric);
            targetCell.GetCellValue().Should().Be(sourceCell.GetCellValue());
        }

        [Theory(DisplayName = "Deve clonar uma célula quando a mesma for uma fórmula.")]
        [InlineData(1)]
        [InlineData(2)]
        [Trait("Category", "Helpers")]
        public void CloneInRow_WhenFormule(int targetRowIndex)
        {
            using var excelHelper = ExcelHelperFactory.CreateImportFiles("Planilha4");

            var sourceCell = excelHelper.GetRow(0).GetCell(4);
            var targetRow = excelHelper.GetRow(targetRowIndex);

            sourceCell.CloneInRow(targetRow);

            var targetCell = targetRow.GetCell(sourceCell.ColumnIndex);

            targetCell.CellType.Should().Be(CellType.String);
            targetCell.GetCellValue().Should().Be(excelHelper.GetRow(0).GetCell(1).GetCellValue());
        }

        [Theory(DisplayName = "Deve clonar uma célula quando a mesma estiver vazia.")]
        [InlineData(1)]
        [InlineData(2)]
        [Trait("Category", "Helpers")]
        public void CloneInRow_WhenBlankData(int targetRowIndex)
        {
            using var excelHelper = ExcelHelperFactory.CreateImportFiles("Planilha4");

            var sourceCell = excelHelper.GetRow(0).GetCell(1);
            var targetRow = excelHelper.GetRow(targetRowIndex);
            sourceCell.SetCellType(CellType.Blank);

            sourceCell.CloneInRow(targetRow);

            var targetCell = targetRow.GetCell(sourceCell.ColumnIndex);

            targetCell.CellType.Should().Be(CellType.Blank);
            targetCell.GetCellValue().Should().Be(sourceCell.GetCellValue());
        }
    }
}