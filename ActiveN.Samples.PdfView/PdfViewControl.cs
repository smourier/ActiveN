namespace ActiveN.Samples.PdfView;

// TODO: generate another GUID, change ProgId and DisplayName
// This GUID *must* match the one in the corresponding .idl coclass
[Guid("d1803953-c348-426e-bd42-5c47f26d9caa")]
[ProgId("ActiveN.Samples.PdfView.PdfViewControl")]
[DisplayName("ActiveN Pdf View Control")]
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
[GeneratedComClass]
public partial class PdfViewControl : BaseControl, IPdfViewControl
{
    #region Mandatory overrides
    protected override ComRegistration ComRegistration => ComHosting.Instance;
    protected override Window CreateWindow(HWND parentHandle, RECT rect) => new PdfViewWindow(parentHandle, GetDefaultWindowStyle(parentHandle), rect);
    #endregion

    // note this is necesary to avoid trimming Task<T>.Result for AOT publishing
    // all Task<T> results should be unwrapped here
    // so you can return any type needed by public methods and properties returning Tasks
    protected override object? GetTaskResult(Task task)
    {
        if (task is Task<string> s)
            return s.Result;

        return null;
    }

    #region IDispatch implementation, static (IDL/TLB) and dynamic (reflection)
#pragma warning disable CA1822 // Mark members as static; no since we're dealing with COM instance methods & properties

    // COM visible public properties exposed through IPdfView IDL should go here
    public bool Enabled { get; set; } = true;
    HRESULT IPdfViewControl.get_Enabled(out BOOL value) { value = Enabled; return Constants.S_OK; }
    HRESULT IPdfViewControl.set_Enabled(BOOL value) { Enabled = value; return Constants.S_OK; }

    public HWND HWND => GetWindowHandle();
    HRESULT IPdfViewControl.get_HWND(out nint value) { value = HWND; return Constants.S_OK; }

    // example of explicit dispid
    // priority for dispids is:
    // 1. look for explicit DispId in TypeLib/IDL
    // 2. look for explicit DispId as a .NET attribute
    // 3. assign dispids automatically starting from AutoDispidsBase (0x10000 by default)
    [DispId(0x20000)]
    [ComAliasName("__id")] // example of alias name for IDispatch (Excel likes this attribute)
    public int Id { get; set; }

#pragma warning restore CA1822 // Mark members as static
    #endregion

    #region COM Registration support

    // COM registration
    public static new HRESULT RegisterType(ComRegistrationContext context)
    {
        TracingUtilities.Trace($"Register type {typeof(PdfViewControl).FullName}...");
        return BaseControl.RegisterType(context);
    }

    public static new HRESULT UnregisterType(ComRegistrationContext context)
    {
        TracingUtilities.Trace($"Unregister type {typeof(PdfViewControl).FullName}...");
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
