using FluentValidation.Results;

namespace Devpack.FileIO.Interfaces
{
    public interface IFileImportReader<TClass> : IDisposable where TClass : IImportable
    {
        IReadOnlyCollection<ValidationFailure> ValidationErrors { get; }
        int AtualLine { get; }
        bool HasImportErrors { get; }

        TClass ReadLine();
        void SetErrorsInLogFile(IEnumerable<ValidationFailure> errors);
        void SetErrorsInLogFile(int rowIndex, IEnumerable<ValidationFailure> errors);
        int CountTotalValidRegisters();
        MemoryStream GetLogAsMemoryStream();
        void SaveLog(string path);
        ValidationResult Validate();
    }
}