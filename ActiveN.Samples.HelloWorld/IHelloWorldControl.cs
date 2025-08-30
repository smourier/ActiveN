namespace ActiveN.Samples.HelloWorld;

[Guid("00000002-2126-40c8-a2f3-8fa83f8cc1f6")]
[GeneratedComInterface]
public partial interface IHelloWorldControl : IDispatch
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

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_CurrentDateTime(out double value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT ComputePi(out double value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Add(double left, double right, out double value);
}
