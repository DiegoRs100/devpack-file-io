namespace Devpack.FileIO.Tests.Common
{
    public class ImportData : IDisposable
    {
        public MemoryStream ImportFileMemoryStream { get; private set; }
        public MemoryStream LogFileMemoryStream { get; private set; }

        public ImportData(MemoryStream importFileMemoryStream, MemoryStream logFileMemoryStream)
        {
            ImportFileMemoryStream = importFileMemoryStream;
            LogFileMemoryStream = logFileMemoryStream;
        }

        public void Dispose()
        {
            ImportFileMemoryStream.Dispose();
            LogFileMemoryStream.Dispose();
        }
    }
}