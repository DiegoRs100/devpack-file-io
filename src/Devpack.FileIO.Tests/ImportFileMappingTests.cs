using Devpack.FileIO.Mapping;
using Devpack.FileIO.Tests.Common.Factories;
using FluentAssertions;
using System.Linq.Expressions;
using Xunit;

namespace Devpack.FileIO.Tests
{
    public class ImportFileMappingTests
    {
        [Fact(DisplayName = "Deve adicionar uma Tag requerida quando a tag ainda não existir dentro do mapeamento.")]
        [Trait("Category", "Helpers")]
        public void MapRequired_WhenValidTag()
        {
            // Valida em conjunto a propriedade (Tags)

            var name = Guid.NewGuid().ToString();
            Expression<Func<EntityValidImportableTest, object?>> property = ev => ev.ColunaA;

            var mapping = new ImportFileMapping<EntityValidImportableTest>();
            mapping.MapRequired(name, property);

            mapping.Tags.Count.Should().Be(1);

            mapping.Tags.Should().Contain(t =>
                t.Name == name
                && t.Property == property
                && t.Options == ImportOptions.Required);
        }

        [Fact(DisplayName = "Deve estourar exception quando uma Tag requerida duplicada for mapeada.")]
        [Trait("Category", "Helpers")]
        public void MapRequired_WhenDuplicatedTag()
        {
            var name = Guid.NewGuid().ToString();

            var mapping = new ImportFileMapping<EntityValidImportableTest>();
            mapping.MapRequired(name, ev => ev.ColunaA);

            mapping.Invoking(m => m.MapRequired(name, ev => ev.ColunaA))
                .Should().Throw<ArgumentException>()
                .WithMessage($"The tag already exists in the import mapping. (Parameter '{name}')");
        }

        [Fact(DisplayName = "Deve adicionar uma Tag opcional quando a tag ainda não existir dentro do mapeamento.")]
        [Trait("Category", "Helpers")]
        public void MapOptional_WhenValidTag()
        {
            // Valida em conjunto a propriedade (Tags)

            var name = Guid.NewGuid().ToString();
            Expression<Func<EntityValidImportableTest, object?>> property = ev => ev.ColunaA;

            var mapping = new ImportFileMapping<EntityValidImportableTest>();
            mapping.MapOptional(name, property);

            mapping.Tags.Count.Should().Be(1);

            mapping.Tags.Should().Contain(t =>
                t.Name == name
                && t.Property == property
                && t.Options == ImportOptions.Optional);
        }

        [Fact(DisplayName = "Deve estourar exception quando uma Tag opcional duplicada for mapeada.")]
        [Trait("Category", "Helpers")]
        public void MapOptional_WhenDuplicatedTag()
        {
            var name = Guid.NewGuid().ToString();

            var mapping = new ImportFileMapping<EntityValidImportableTest>();
            mapping.MapOptional(name, ev => ev.ColunaA);

            mapping.Invoking(m => m.MapRequired(name, ev => ev.ColunaA))
                .Should().Throw<ArgumentException>()
                .WithMessage($"The tag already exists in the import mapping. (Parameter '{name}')");
        }
    }
}