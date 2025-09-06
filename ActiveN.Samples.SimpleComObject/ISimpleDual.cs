namespace ActiveN.Samples.SimpleComObject;

[Guid("381ca3ad-ae05-4e00-8f80-e39ed374500b")]
[GeneratedComInterface]
public partial interface ISimpleDual : IDispatch
{
#pragma warning disable IDE1006 // Naming Styles

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_Name(out BSTR value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT set_Name(BSTR value);
#pragma warning restore IDE1006 // Naming Styles

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Multiply(double left, double right, out double ret);
}
