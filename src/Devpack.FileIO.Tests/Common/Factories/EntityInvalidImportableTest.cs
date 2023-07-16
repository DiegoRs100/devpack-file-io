using Devpack.FileIO;
using Devpack.FileIO.Interfaces;
using Devpack.FileIO.Mapping;

namespace Devpack.FileIO.Tests.Common.Factories
{
    public class EntityInvalidImportableTest : Importable<EntityInvalidImportableTest>
    {
        private readonly IImportFileMapping<EntityInvalidImportableTest> _headerMapping =
            new ImportFileMapping<EntityInvalidImportableTest>().MapRequired("Coluna TESTE", p => p.ColunaA);

        public override IImportFileMapping<EntityInvalidImportableTest> HeaderMapping => _headerMapping;

        public string ColunaA { get; set; } = default!;
    }
}