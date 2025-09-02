namespace ActiveN;

public class DispatchMember
{
    public DispatchMember(int dispId, DispatchCategory category, MemberInfo? info)
    {
        ArgumentNullException.ThrowIfNull(category);
        DispId = dispId;
        Category = category;
        Info = info;
    }

    public int DispId { get; }
    public DispatchCategory Category { get; }
    public MemberInfo? Info { get; } // if null, means the method/property is in the TLB but not in the actual type

    public override string ToString() => $"{DispId}: {Info?.Name} ({Info?.MemberType}) [{Category}]";
}