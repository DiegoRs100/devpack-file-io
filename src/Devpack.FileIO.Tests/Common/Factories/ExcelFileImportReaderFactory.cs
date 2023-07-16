using Devpack.FileIO.Adapters.Excel;
using Devpack.FileIO.Tests.Resources;

namespace Devpack.FileIO.Tests.Common.Factories
{
    public static class ExcelFileImportReaderFactory
    {
        public static ExcelFileImportReader<EntityValidImportableTest> CreateValidImportReader()
        {
            using var importData = new MemoryStream(Resource.ImportFile);
            return new ExcelFileImportReader<EntityValidImportableTest>(importData);
        }

        public static ExcelFileImportReader<EntityValidImportableTest> CreateValidImportReader(string importSheetName)
        {
            using var importData = new MemoryStream(Resource.ImportFile);
            return new ExcelFileImportReader<EntityValidImportableTest>(importData, importSheetName);
        }

        public static ExcelFileImportReader<EntityInvalidImportableTest> CreateInvalidImportReader()
        {
            using var importData = new MemoryStream(Resource.ImportFile);
            return new ExcelFileImportReader<EntityInvalidImportableTest>(importData);
        }

        public static ExcelFileImportReader<EntityInvalidImportableTest> CreateInvalidImportReader(string importSheetName)
        {
            using var importData = new MemoryStream(Resource.ImportFile);
            return new ExcelFileImportReader<EntityInvalidImportableTest>(importData, importSheetName);
        }
    }
}