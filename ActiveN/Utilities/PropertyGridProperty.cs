namespace ActiveN.Utilities;

[GeneratedComClass]
public partial class PropertyGridProperty : IPropertyGridProperty
{
    public PropertyGridProperty(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        Name = name;
    }

    public string Name { get; }
    public string? Category { get; set; }
    public object? Value { get; set; }
    public object? DefaultValue { get; set; }
    public PropertyGridPropertyOptions Options { get; set; }
    public string? TypeName { get; set; }

    public override string ToString() => Name;

    HRESULT IPropertyGridProperty.GetCategory(out PWSTR name) { name = new Pwstr(Category); return Constants.S_OK; }
    HRESULT IPropertyGridProperty.GetName(out PWSTR name) { name = new Pwstr(Name); return Constants.S_OK; }
    HRESULT IPropertyGridProperty.GetOptions(out PropertyGridPropertyOptions name) { name = Options; return Constants.S_OK; }
    HRESULT IPropertyGridProperty.GetTypeName(out PWSTR name) { name = new Pwstr(TypeName); return Constants.S_OK; }

    HRESULT IPropertyGridProperty.GetDefaultValue(out VARIANT value)
    {
        using var pv = new Variant(DefaultValue);
        value = pv.Detach();
        return Constants.S_OK;
    }

    HRESULT IPropertyGridProperty.GetValue(out VARIANT value)
    {
        using var pv = new Variant(Value);
        value = pv.Detach();
        return Constants.S_OK;
    }

    HRESULT IPropertyGridProperty.SetValue(VARIANT value)
    {
        using var pv = Variant.Attach(ref value, false);
        Value = pv.Value;
        return Constants.S_OK;
    }
}
