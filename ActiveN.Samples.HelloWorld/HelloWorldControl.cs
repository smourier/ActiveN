
namespace ActiveN.Samples.HelloWorld;

// TODO: generate another GUID, change ProgId and DisplayName
// This GUID *must* match the one in the corresponding .idl coclass
[Guid("7883986f-b61e-474e-91b5-7822308e3b7d")]
[ProgId("ActiveN.Samples.HelloWorld.HelloWorldControl")]
[DisplayName("ActiveN Hello World Control")]
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
[GeneratedComClass]
public partial class HelloWorldControl : BaseControl, IHelloWorldControl
{
    protected override ComRegistration ComRegistration => ComHosting.Instance;

    public static new HRESULT RegisterType(ComRegistrationContext context)
    {
        ComRegistration.Trace($"Register type {typeof(HelloWorldControl).FullName}...");
        return BaseControl.RegisterType(context);
    }

    public static new HRESULT UnregisterType(ComRegistrationContext context)
    {
        ComRegistration.Trace($"Unregister type {typeof(HelloWorldControl).FullName}...");
        return BaseControl.UnregisterType(context);
    }

    public HRESULT ComputePi(out double ret)
    {
        ret = Math.PI;
        ComRegistration.Trace();
        return Constants.S_OK;
    }

    public HRESULT GetIDsOfNames(in Guid riid, PWSTR[] rgszNames, uint cNames, uint lcid, int[] rgDispId)
    {
        ComRegistration.Trace();
        throw new NotImplementedException();
    }

    public HRESULT GetTypeInfo(uint iTInfo, uint lcid, out ITypeInfo ppTInfo)
    {
        ComRegistration.Trace();
        throw new NotImplementedException();
    }

    public HRESULT GetTypeInfoCount(out uint pctinfo)
    {
        ComRegistration.Trace();
        throw new NotImplementedException();
    }

    public HRESULT Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr)
    {
        ComRegistration.Trace();
        throw new NotImplementedException();
    }
}
