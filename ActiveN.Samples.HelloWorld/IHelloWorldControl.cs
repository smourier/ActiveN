namespace ActiveN.Samples.HelloWorld;

[Guid("00000002-2126-40c8-a2f3-8fa83f8cc1f6")]
[GeneratedComInterface]
#pragma warning disable SYSLIB1230 // Specifying 'GeneratedComInterfaceAttribute' on an interface that has a base interface defined in another assembly is not supported
public partial interface IHelloWorldControl : IDispatch
#pragma warning restore SYSLIB1230 // Specifying 'GeneratedComInterfaceAttribute' on an interface that has a base interface defined in another assembly is not supported
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT CurrentDateTime(out double ret);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT HWND(out nint ret);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT ComputePi(out double ret);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Add(double left, double right, out double ret);
}
