namespace ActiveN.Samples.HelloWorld;

// TODO: generate another GUID, change ProgId and DisplayName
// This GUID *must* match the one in the corresponding .idl coclass
[Guid("7883986f-b61e-474e-91b5-7822308e3b7d")]
[ProgId("ActiveN.Samples.HelloWorld.HelloWorldControl")]
[DisplayName("ActiveN Hello World Control")]
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
[GeneratedComClass]
#pragma warning disable CA1822 // Mark members as static; no since we're dealing with COM instance methods & properties
public partial class HelloWorldControl : BaseControl, IHelloWorldControl
{
    protected override ComRegistration ComRegistration => ComHosting.Instance;
    protected override HWND GetWindowHandle()
    {
        throw new NotImplementedException();
    }

    // note this is necesary to avoid trimming Task<T>.Result for AOT publishing
    // all Task<T> results should be unwrapped here
    // so you can return any type needed by public methods and properties returning Tasks
    protected override object? GetTaskResult(Task task)
    {
        if (task is Task<string> s)
            return s.Result;

        return null;
    }

    // COM visible public properties exposed through IHelloWorldControl IDL should go here
    public HRESULT CurrentDateTime(out double ret)
    {
        ret = DateTime.Now.ToOADate();
        return Constants.S_OK;
    }

    public HRESULT HWND(out nint ret)
    {
        ret = GetWindowHandle();
        return Constants.S_OK;
    }

    // this is not in the IDL, but it works (IDispatch + automatic dispid)
    public long TickCount => Environment.TickCount64;

    // COM visible public methods exposed through IHelloWorldControl IDL should go here
    public Task<string> GetInfoAsync(int delay) => Task.Run(async () =>
    {
        await Task.Delay(delay).ConfigureAwait(false); // simulate some async work
        return $"Hello from ActiveN HelloWorldControl! Date is {DateTime.Now}.";
    });

    // this would be valid too: public double ComputePi()
    public HRESULT ComputePi(out double ret)
    {
        ret = Math.PI;
        TracingUtilities.Trace();
        return Constants.S_OK;
    }

    // this would be valid too: public HRESULT ComputePi(double left, double right, out double sum)
    public HRESULT Add(double left, double right, out double sum)
    {
        sum = left + right;
        TracingUtilities.Trace();
        return Constants.S_OK;
    }

    // COM registration
    public static new HRESULT RegisterType(ComRegistrationContext context)
    {
        TracingUtilities.Trace($"Register type {typeof(HelloWorldControl).FullName}...");
        return BaseControl.RegisterType(context);
    }

    public static new HRESULT UnregisterType(ComRegistrationContext context)
    {
        TracingUtilities.Trace($"Unregister type {typeof(HelloWorldControl).FullName}...");
        return BaseControl.UnregisterType(context);
    }
}
#pragma warning restore CA1822 // Mark members as static
