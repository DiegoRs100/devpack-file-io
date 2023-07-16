using Devpack.FileIO.Adapters.Excel;
using Devpack.FileIO.Tests.Resources;

namespace Devpack.FileIO.Tests.Common.Factories
{
    public static class ExcelHelperFactory
    {
        public static ExcelHelper CreateDefaultImportFiles()
        {
            using var importFileMemoryStream = new MemoryStream(Resource.ImportFile);
            return new ExcelHelper(importFileMemoryStream);
        }

        public static ExcelHelper CreateImportFiles(string worksheetName)
        {
            using var importFileMemoryStream = new MemoryStream(Resource.ImportFile);
            return new ExcelHelper(importFileMemoryStream, worksheetName);
        }
    }
}