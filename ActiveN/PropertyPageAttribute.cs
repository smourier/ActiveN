namespace ActiveN;

[AttributeUsage(AttributeTargets.Method | AttributeTargets.Property, Inherited = true)]
public sealed class PropertyPageAttribute(string clsid) : Attribute
{
    public string Clsid { get; } = clsid;
    public Guid Guid
    {
        get
        {
            if (Guid.TryParse(Clsid, out var guid))
                return guid;

            return BaseDispatch.DefaultPropertyPageId;
        }
    }

    public string? DefaultString { get; set; }
}
