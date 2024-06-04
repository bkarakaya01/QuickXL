namespace QuickXL.Exporter.Tests.SampleProvider;

internal static class SampleDataProvider
{
    internal static List<SampleDto> Create()
    {
        return
        [
            new()
            {
                Id = 1,
                Name = "One"
            },
            new()
            {
                Id = 2,
                Name = "Two"
            },
            new()
            {
                Id = 3,
                Name = "Three"
            },
            new()
            {
                Id = 4,
                Name = "Four"
            },
            new()
            {
                Id = 5,
                Name = "Five"
            }
        ];
    }
}
