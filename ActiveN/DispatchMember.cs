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
    public virtual Guid? PropertyPageId { get; set; }
    public virtual string? DefaultString { get; set; }

    public override string ToString() => $"{DispId}: {Info?.Name} ({Info?.MemberType}) [{Category}]";

    private Type GetMemberType()
    {
        if (Info is PropertyInfo pi)
            return pi.PropertyType;

        if (Info is MethodInfo mi)
            return mi.ReturnType;

        throw new NotSupportedException();
    }

    public virtual object? GetDefaultValue()
    {
        var type = GetMemberType();
        if (type == typeof(string))
            return DefaultString;

        Conversions.TryChangeObjectType(DefaultString, type, out var value);
        return value;
    }
}