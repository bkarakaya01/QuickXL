namespace QuickXL.Core.Result
{
    public class XLValidationResult
    {
        public bool Success { get; set; }
        public IList<string>? Errors { get; set; }

        public static XLValidationResult SUCCESS = 
            new() 
            { 
                Success = true 
            };
    }
}
