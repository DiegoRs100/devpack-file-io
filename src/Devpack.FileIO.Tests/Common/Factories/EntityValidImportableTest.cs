using Devpack.FileIO.Interfaces;
using Devpack.FileIO.Mapping;

namespace Devpack.FileIO.Tests.Common.Factories
{
    public class EntityValidImportableTest : Importable<EntityValidImportableTest>
    {
        private readonly IImportFileMapping<EntityValidImportableTest> _headerMapping = new ImportFileMapping<EntityValidImportableTest>()
            .MapRequired("Coluná A", p => p.ColunaA)
            .MapRequired("Coluna B", p => p.ColunaB)
            .MapOptional("Coluna C", p => p.ColunaC)
            .MapRequired("Coluna D", p => p.ColunaD)
            .MapRequired("Coluna E", p => p.ColunaE)
            .MapRequired("Coluna F", p => p.ColunaF)
            .MapRequired("Coluna G", p => p.ColunaG);

        public override IImportFileMapping<EntityValidImportableTest> HeaderMapping => _headerMapping;

        public string ColunaA { get; set; } = default!;
        public string ColunaB { get; set; } = default!;
        public string ColunaC { get; set; } = default!;
        public double? ColunaD { get; set; } = default!;
        public DateTime? ColunaE { get; set; } = default!;
        public DateTime? ColunaF { get; set; } = default!;
        public bool? ColunaG { get; set; } = default!;
    }
}