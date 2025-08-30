namespace ActiveN.Samples.PdfView;

// note we *must* redefine IDispatch here because although it's defined in DirectN, we must have it
// in the same assembly as .NET Core ComWrappers Source generation is quite broken at that level, if you ask me
// https://learn.microsoft.com/en-us/dotnet/standard/native-interop/comwrappers-source-generation#derived-interfaces-across-assembly-boundaries
[GeneratedComInterface, Guid("00020400-0000-0000-c000-000000000046")]
public partial interface IDispatch
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetTypeInfoCount(out uint pctinfo);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetTypeInfo(uint iTInfo, uint lcid, [MarshalUsing(typeof(UniqueComInterfaceMarshaller<ITypeInfo>))] out ITypeInfo ppTInfo);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetIDsOfNames(in Guid riid, [In][MarshalUsing(CountElementName = nameof(cNames))] PWSTR[] rgszNames, uint cNames, uint lcid, [In][Out][MarshalUsing(CountElementName = nameof(cNames))] int[] rgDispId);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint /* optional VARIANT* */ pVarResult, nint /* optional EXCEPINFO* */ pExcepInfo, nint /* optional uint* */ puArgErr);
}
