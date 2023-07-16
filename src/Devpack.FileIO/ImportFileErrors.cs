namespace Devpack.FileIO
{
    public struct ImportFileErrors
    {
        public const string ErrorEmptySheetKey = "EmptySheet";
        internal const string ErrorEmptySheetMessage = "The supplied file contains no records.";

        public const string InvalidHeaderKey = "InvalidHeader";
        internal const string InvalidHeaderMessage = "The supplied file does not have all the required columns.";

        public const string DuplicatedTagsKey = "DuplicatedTags";
        internal const string DuplicatedTagsMessage = "The supplied file has duplicate identifiers in the header.";
    }
}