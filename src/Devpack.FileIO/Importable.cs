using Devpack.FileIO.Interfaces;

namespace Devpack.FileIO
{
    public abstract class Importable<TClass> : IImportable where TClass : IImportable
    {
        public int SourceImportRowIndex { get; private set; }
        public abstract IImportFileMapping<TClass> HeaderMapping { get; }

        public void SetSourceImportRowIndex(int rowIndex)
        {
            SourceImportRowIndex = rowIndex;
        }
    }
}