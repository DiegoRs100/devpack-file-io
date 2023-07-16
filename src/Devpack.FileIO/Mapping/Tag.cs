using Devpack.FileIO.Interfaces;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Devpack.FileIO.Tests")]
namespace Devpack.FileIO.Mapping
{
    public class Tag<TClass> where TClass : IImportable
    {
        internal string Name { get; private set; }
        internal Expression<Func<TClass, object?>> Property { get; private set; }
        internal ImportOptions Options { get; private set; }
        internal int? FileColumnIndex { get; private set; }

        internal Tag(string name, Expression<Func<TClass, object?>> property, ImportOptions options)
        {
            Name = name;
            Property = property;
            Options = options;
        }

        internal void SetFileColumnIndex(int columnIndex)
        {
            FileColumnIndex = columnIndex;
        }
    }
}