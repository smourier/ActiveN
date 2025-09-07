namespace ActiveN;

[AttributeUsage(AttributeTargets.Class, Inherited = true, AllowMultiple = false)]
public sealed class MiscStatusAttribute(OLEMISC value) : Attribute
{
    public OLEMISC Value { get; } = value;
}
