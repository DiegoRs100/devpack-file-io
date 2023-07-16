using Bogus;
using Devpack.FileIO.Mapping;
using Devpack.FileIO.Tests.Common.Factories;
using FluentAssertions;
using System.Linq.Expressions;
using Xunit;

namespace Devpack.FileIO.Tests
{
    public class TagTets
    {
        [Fact(DisplayName = "Deve instanciar corretamente as propriedades quando o objeto for instanciado.")]
        [Trait("Category", "Models")]
        public void Constructor()
        {
            var faker = new Faker();

            var name = Guid.NewGuid().ToString();
            Expression<Func<EntityValidImportableTest, object?>> property = ev => ev.ColunaA;
            var options = faker.Random.Enum<ImportOptions>();

            var tag = new Tag<EntityValidImportableTest>(name, property, options);

            tag.Name.Should().Be(name);
            tag.Property.Should().Be(property);
            tag.Options.Should().Be(options);
        }

        [Fact(DisplayName = "Deve popular corretamente a propriedade (FileColumnIndex) quando o método for chamado.")]
        [Trait("Category", "Models")]
        public void SetFileColumnIndex()
        {
            var index = new Faker().Random.Number(1, 99);

            var tag = new Tag<EntityValidImportableTest>(Guid.NewGuid().ToString(), ev => ev.ColunaA, ImportOptions.Required);
            tag.SetFileColumnIndex(index);

            tag.FileColumnIndex.Should().Be(index);
        }
    }
}