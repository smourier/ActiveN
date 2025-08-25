namespace ActiveN.Samples.HelloWorld;

// TODO: generate another GUID
[Guid("7883986f-b61e-474e-91b5-7822308e3b7d")]
[GeneratedComClass]
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
public partial class HelloWorldControl : BaseControl
{
    public void DoSomething()
    {
        BaseComRegistration.Trace("Hello from HelloWorldControl.DoSomething");
    }

    public void DoSomethingElse()
    {
        BaseComRegistration.Trace("Hello from HelloWorldControl.DoSomething");
    }

    public static new HRESULT RegisterType(ComRegistrationContext context)
    {
        BaseComRegistration.Trace($"Register type {typeof(HelloWorldControl).FullName}...");
        return BaseControl.RegisterType(context);
    }

    public static new HRESULT UnregisterType(ComRegistrationContext context)
    {
        BaseComRegistration.Trace($"Unregister type {typeof(HelloWorldControl).FullName}...");
        return BaseControl.UnregisterType(context);
    }
}
