namespace ActiveN.Utilities;

[GeneratedComClass]
public partial class PropertyGridObject : IPropertyGridObject
{
    public PropertyGridObject([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] Type type, object? instance)
    {
        ArgumentNullException.ThrowIfNull(type);
        ArgumentNullException.ThrowIfNull(instance);

        foreach (var prop in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (prop.GetIndexParameters().Length > 0)
                continue;

            var typeName = prop.PropertyType.FullName;
            if (typeName == null)
                continue;

            var browsable = prop.GetCustomAttribute<BrowsableAttribute>()?.Browsable ?? true;
            if (!browsable)
                continue;

            var isReadonly = !prop.CanWrite || prop.GetCustomAttribute<ReadOnlyAttribute>()?.IsReadOnly == true;
            var options = isReadonly ? PropertyGridPropertyOptions.ReadOnly : PropertyGridPropertyOptions.None;

            var pgp = new PropertyGridProperty(prop.Name)
            {
                Category = prop.GetCustomAttribute<CategoryAttribute>()?.Category.Nullify(),
                Options = options,
                TypeName = typeName,
                Value = prop.GetValue(instance),
                DefaultValue = prop.GetCustomAttribute<DefaultValueAttribute>()?.Value
            };
            Properties.Add(pgp);
        }
    }

    public IList<PropertyGridProperty> Properties { get; } = [];

    HRESULT IPropertyGridObject.GetProperties(out VARIANT properties)
    {
        using var pv = new Variant(Properties.Cast<object>().ToArray());
        properties = pv.Detach();
        return Constants.S_OK;
    }
}
