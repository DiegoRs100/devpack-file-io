using Bogus;
using Devpack.FileIO.Tests.Common.Factories;
using FluentAssertions;
using Xunit;

namespace Devpack.FileIO.Tests
{
    public class ImportableTests
    {
        [Fact(DisplayName = "Deve popular a propriedade (SourceImportRowIndex) quando o método for chamado.")]
        [Trait("Category", "Models")]
        public void SetErrorsInLogFile()
        {
            var faker = new Faker();
            var rowIndex = faker.Random.Number(1, 100);
            var importable = new EntityValidImportableTest();

            importable.SetSourceImportRowIndex(rowIndex);

            importable.SourceImportRowIndex.Should().Be(rowIndex);
        }
    }
}