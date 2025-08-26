namespace ActiveN.Samples.HelloWorld;

// TODO: generate another GUID, change ProgId and DisplayName
// This GUID *must* match the one in the corresponding .idl coclass
[Guid("7883986f-b61e-474e-91b5-7822308e3b7d")]
[ProgId("ActiveN.Samples.HelloWorld.HelloWorldControl")]
[DisplayName("ActiveN Hello World Control")]
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
[GeneratedComClass]
public partial class HelloWorldControl : BaseControl
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
}
