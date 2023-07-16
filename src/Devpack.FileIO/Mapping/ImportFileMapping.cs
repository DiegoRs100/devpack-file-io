using Devpack.Extensions.Types;
using Devpack.FileIO.Interfaces;
using System.Linq.Expressions;

namespace Devpack.FileIO.Mapping
{
    public class ImportFileMapping<TClass> : IImportFileMapping<TClass> where TClass : IImportable
    {
        private readonly List<Tag<TClass>> _tags = new();
        public IReadOnlyCollection<Tag<TClass>> Tags => _tags;

        public ImportFileMapping<TClass> MapRequired(string name, Expression<Func<TClass, object?>> property)
        {
            ValidateTagName(name);

            _tags.Add(new Tag<TClass>(name, property!, ImportOptions.Required));
            return this;
        }

        public ImportFileMapping<TClass> MapOptional(string name, Expression<Func<TClass, object?>> property)
        {
            ValidateTagName(name);

            _tags.Add(new Tag<TClass>(name, property!, ImportOptions.Optional));
            return this;
        }

        private void ValidateTagName(string name)
        {
            if (_tags.Exists(t => t.Name.Match(name)))
                throw new ArgumentException("The tag already exists in the import mapping.", name);
        }
    }
}