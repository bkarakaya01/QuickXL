namespace QuickXL.Core.Extensions.IO
{
    public static class IOExtensions
    {
        public static void Reset(this Stream stream)
        {
            stream.Position = 0;
        }
    }
}
