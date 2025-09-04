namespace ActiveN.Hosting;

public interface IAggregable
{
    bool SupportsAggregation { get; }
    nint Wrapper { get; set; }
    IReadOnlyList<Type> AggregableInterfaces { get; }
}
