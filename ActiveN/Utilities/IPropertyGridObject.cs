namespace ActiveN.Utilities;

#if NETFRAMEWORK
[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("dc933288-bccb-4f4d-af8c-5aff61750a18")]
#else
[GeneratedComInterface, Guid("dc933288-bccb-4f4d-af8c-5aff61750a18")]
#endif
public partial interface IPropertyGridObject
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetProperties(out VARIANT properties);
}
