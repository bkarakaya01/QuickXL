namespace QuickXL.Core.Result;

public sealed record XLResultError
{
    public XLResultError(string code, string message)
    {
        Code = code;
        Message = message;
    }
    public string Code { get; set; }
    public string Message { get; set; }
}

