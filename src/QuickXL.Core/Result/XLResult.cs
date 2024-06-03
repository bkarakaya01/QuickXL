namespace QuickXL.Core.Result;

public sealed record XLResult
{    
    public bool Succeeded { get; set; }
    public Stream? Data { get; set; }
    public IList<XLResultError>? Errors { get; set; }


    public static XLResult Success(Stream data) =>
        new()
        {
            Succeeded = true,
            Data = data
        };

    public static XLResult Failure(IList<XLResultError> errors) =>
        new()
        {
            Succeeded = false,            
            Errors = errors
        };

    public static explicit operator XLResult(Exception exception)
    {
        return Failure(
        [
            new(exception.GetType().Name, exception.Message)
        ]);
    }
}
