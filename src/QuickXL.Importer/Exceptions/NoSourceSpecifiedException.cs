namespace QuickXL.Importer.Exceptions
{
    internal sealed class NoSourceSpecifiedException : Exception
    {
        public NoSourceSpecifiedException() : base("No input source specified. Call FromFile() or FromStream() before Import().")
        {

        }
    }
}
