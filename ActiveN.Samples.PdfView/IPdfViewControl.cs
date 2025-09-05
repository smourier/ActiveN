namespace ActiveN.Samples.PdfView;

[Guid("63b16fe2-2faa-4498-9c3e-023e720f8cae")]
[GeneratedComInterface]
public partial interface IPdfViewControl : IDispatch
{
#pragma warning disable IDE1006 // Naming Styles
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_Enabled(out VARIANT_BOOL value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT set_Enabled(VARIANT_BOOL value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_BackColor(out uint value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT set_BackColor(uint value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_ShowControls(out VARIANT_BOOL value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT set_ShowControls(VARIANT_BOOL value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_PageCount(out int value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_IsPasswordProtected(out VARIANT_BOOL value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_FilePath(out BSTR value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_HWND(out nint value);
#pragma warning restore IDE1006 // Naming Styles

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OpenFile(BSTR filePath);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT OpenStream(IStream stream);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT CloseFile();

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT MovePage(int delta);
}
