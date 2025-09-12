namespace ActiveN.Samples.WebView2;

[Guid("9d51ea97-30dd-4e70-99ae-a5085fbf8095")]
[GeneratedComInterface]
public partial interface IWebView2Control : IDispatch
{
#pragma warning disable IDE1006 // Naming Styles
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_Enabled(out BOOL value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT set_Enabled(BOOL value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_Caption(out BSTR value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT set_Caption(BSTR value);
#pragma warning restore IDE1006 // Naming Styles
}
