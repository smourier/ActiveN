namespace ActiveN.Utilities;

#if NETFRAMEWORK
[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("8214bd98-1dbb-4fb1-810a-02e9d563dc8f")]
#else
[GeneratedComInterface, Guid("8214bd98-1dbb-4fb1-810a-02e9d563dc8f")]
#endif
public partial interface IPropertyGridProperty
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetName(out PWSTR name);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetCategory(out PWSTR name);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetTypeName(out PWSTR name);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetOptions(out PropertyGridPropertyOptions name);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetDefaultValue(out VARIANT value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT SetValue(VARIANT value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetValue(out VARIANT value);
}
