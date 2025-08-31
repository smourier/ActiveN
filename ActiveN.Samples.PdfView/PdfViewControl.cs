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
    private readonly DispatchConnectionPoint _eventsConnectionPoint;

    public PdfViewControl()
        : base()
    {
        // we need that to be able to use Windows.UI.Composition APIs
        // and no, we can't use Windows.System.DispatcherQueueController.CreateOnDedicatedThread()
        // even implicitly, or compositor will raise an access denied error
        DispatcherQueueController = new WindowsDispatcherQueueController();
        DispatcherQueueController.EnsureOnCurrentThread();

        // setup connection point for our (IDispatch) events
        // the advantage of IDispatch-based events is that we don't need to define interfaces or classes, 
        // we just need the Guid and Dispids
        var IID_IPdfViewControlEvents = new Guid("48c606f1-d597-467d-8a38-1c02fd7e019d"); // this must match the one in the .idl
        _eventsConnectionPoint = new DispatchConnectionPoint(IID_IPdfViewControlEvents);
        AddConnectionPoint(_eventsConnectionPoint);
    }

    public WindowsDispatcherQueueController DispatcherQueueController { get; }
    public new PdfViewWindow? Window => (PdfViewWindow?)base.Window;

    #region Mandatory overrides
    protected override ComRegistration ComRegistration => ComHosting.Instance;
    protected override Window CreateWindow(HWND parentHandle, RECT rect)
    {
        var window = new PdfViewWindow(parentHandle, GetDefaultWindowStyle(parentHandle), rect);

        // bind .NET events to forward to COM clients
        window.FileOpened += (s, e) => _eventsConnectionPoint.InvokeMember((int)PdfViewControlEventsDispIds.FileOpened);
        window.FileClosed += (s, e) => _eventsConnectionPoint.InvokeMember((int)PdfViewControlEventsDispIds.FileClosed);
        window.PageChanged += (s, e) => _eventsConnectionPoint.InvokeMember((int)PdfViewControlEventsDispIds.PageChanged);
        return window;
    }
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

    public string Caption { get => Window?.FilePath ?? string.Empty; set { } }
    HRESULT IPdfViewControl.get_Caption(out BSTR value) { value = new BSTR(Marshal.StringToBSTR(Caption)); return Constants.S_OK; }
    HRESULT IPdfViewControl.set_Caption(BSTR value) { Caption = value.ToString() ?? string.Empty; return Constants.S_OK; }

    public bool ShowControls { get => Window?.ShowControls ?? true; set { if (Window != null) Window.ShowControls = value; } }
    HRESULT IPdfViewControl.get_ShowControls(out BOOL value) { value = ShowControls; return Constants.S_OK; }
    HRESULT IPdfViewControl.set_ShowControls(BOOL value) { ShowControls = value; return Constants.S_OK; }

    public HWND HWND => GetWindowHandle();
    HRESULT IPdfViewControl.get_HWND(out nint value) { value = HWND; return Constants.S_OK; }

    public void OpenFile(string filePath) => Window?.OpenFile(filePath);
    HRESULT IPdfViewControl.OpenFile(BSTR filePath)
    {
        if (filePath.Value == 0)
            return Constants.E_POINTER;

        var str = filePath.ToString();
        if (string.IsNullOrWhiteSpace(str))
            return Constants.E_INVALIDARG;

        OpenFile(str);
        return Constants.S_OK;
    }

    public void CloseFile() => Window?.CloseFile();
    HRESULT IPdfViewControl.CloseFile() { CloseFile(); return Constants.S_OK; }

    public void MovePage(int delta) => Window?.MovePage(delta);
    HRESULT IPdfViewControl.MovePage(int delta) { MovePage(delta); return Constants.S_OK; }

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
