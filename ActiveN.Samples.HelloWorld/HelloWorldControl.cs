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
    #region Mandatory overrides
    protected override ComRegistration ComRegistration => ComHosting.Instance;
    protected override Window CreateWindow(HWND parentHandle, RECT rect) => new HelloWorldWindow(parentHandle, GetDefaultWindowStyle(parentHandle), rect);

    // note this is necesary to avoid trimming Task<T>.Result for AOT publishing
    // all Task<T> results should be unwrapped here
    // so you can return any type needed by public methods and properties returning Tasks
    protected override object? GetTaskResult(Task task)
    {
        // string here is at least for the GetInfoAsync method below
        if (task is Task<string> s)
            return s.Result;

        return null;
    }
    #endregion

    protected override void Draw(HDC hdc, RECT bounds)
    {
        if (hdc == 0)
            return;

        HelloWorldWindow.Paint(hdc, bounds);
    }

    #region IDispatch implementation, static (IDL/TLB) and dynamic (reflection)
#pragma warning disable CA1822 // Mark members as static; no since we're dealing with COM instance methods & properties

    // COM visible public properties exposed through IHelloWorldControl IDL should go here
    public DateTime CurrentDate => DateTime.Now;
    HRESULT IHelloWorldControl.get_CurrentDateTime(out double value) { value = CurrentDate.ToOADate(); return Constants.S_OK; }

    public bool Enabled { get; set; } = true;
    HRESULT IHelloWorldControl.get_Enabled(out BOOL value) { value = Enabled; return Constants.S_OK; }
    HRESULT IHelloWorldControl.set_Enabled(BOOL value) { Enabled = value; return Constants.S_OK; }

    public string Caption { get; set; } = "Hello World";
    HRESULT IHelloWorldControl.get_Caption(out BSTR value) { value = new BSTR(Marshal.StringToBSTR(Caption)); return Constants.S_OK; }
    HRESULT IHelloWorldControl.set_Caption(BSTR value) { Caption = value.ToString() ?? string.Empty; return Constants.S_OK; }

    public HWND HWND => GetWindowHandle();
    HRESULT IHelloWorldControl.get_HWND(out nint value) { value = HWND; return Constants.S_OK; }

    // example of explicit dispid
    // priority for dispids is:
    // 1. look for explicit DispId in TypeLib/IDL
    // 2. look for explicit DispId as a .NET attribute
    // 3. assign dispids automatically starting from AutoDispidsBase (0x10000 by default)
    [DispId(0x20000)]
    [ComAliasName("__id")] // example of alias name for IDispatch (Office seems to like this attribute)
    public int Id { get; set; }

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

#pragma warning restore CA1822 // Mark members as static
    #endregion

    #region COM Registration support

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

    #endregion

    #region IDispatch support

    HRESULT IDispatch.GetIDsOfNames(in Guid riid, PWSTR[] rgszNames, uint cNames, uint lcid, int[] rgDispId) =>
        GetIDsOfNames(in riid, rgszNames, cNames, lcid, rgDispId);

    HRESULT IDispatch.Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr) =>
        Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

    HRESULT IDispatch.GetTypeInfo(uint iTInfo, uint lcid, out ITypeInfo ppTInfo) =>
        GetTypeInfo(iTInfo, lcid, out ppTInfo);

    HRESULT IDispatch.GetTypeInfoCount(out uint pctinfo) =>
        GetTypeInfoCount(out pctinfo);

    #endregion
}
