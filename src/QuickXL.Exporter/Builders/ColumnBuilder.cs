namespace QuickXL;


/// <summary>
/// Transitional class for builder pattern, intents to shae user's build mechanism.
/// </summary>
/// <typeparam name="TDto"></typeparam>
public class ColumnBuilder<TDto> where TDto : class, new()
{
    public required ExportBuilder<TDto> ExportBuilder { get; set; }

    public ColumnBuilder<TDto> AddColumn(string header, Func<TDto, object> propertySelector)
    {        
        ExportBuilder.HeaderPropertySelectors.Add(header, propertySelector);

        return this;
    }
}

