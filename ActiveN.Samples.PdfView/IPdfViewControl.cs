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
    HRESULT get_HWND(out nint value);
#pragma warning restore IDE1006 // Naming Styles
}
