namespace ActiveN.Samples.PdfView;

[Guid("63b16fe2-2faa-4498-9c3e-023e720f8cae")]
[GeneratedComInterface]
public partial interface IPdfViewControl : IDispatch
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

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_ShowControls(out BOOL value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT set_ShowControls(BOOL value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_HWND(out nint value);
#pragma warning restore IDE1006 // Naming Styles

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OpenFile(BSTR filePath);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT CloseFile();

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT MovePage(int delta);
}
