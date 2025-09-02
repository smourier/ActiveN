namespace ActiveN.Hosting;

public interface IAggregable
{
    bool SupportsAggregation { get; }
    nint OuterUnknown { get; set; }
    object? Wrapper { get; set; }
}
