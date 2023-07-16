using Devpack.FileIO.Mapping;

namespace Devpack.FileIO.Interfaces
{
    public interface IImportFileMapping<TClass> where TClass : IImportable
    {
        IReadOnlyCollection<Tag<TClass>> Tags { get; }
    }
}