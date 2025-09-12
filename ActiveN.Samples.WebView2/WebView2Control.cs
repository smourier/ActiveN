// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN.Samples.WebView2;

// TODO: generate another GUID, change ProgId and DisplayName
// This GUID *must* match the one in the corresponding .idl coclass
[Guid("27ef2dff-595f-4724-aea1-d16f7730dc33")]
[ProgId("ActiveN.Samples.WebView2.WebView2Control")]
[DisplayName("ActiveN WebView2 Control")]
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
[GeneratedComClass]
public partial class WebView2Control : BaseControl, IWebView2Control
{
    #region Mandatory overrides
    protected override ComRegistration ComRegistration => ComHosting.Instance;
    protected override Guid DispatchInterfaceId => typeof(IWebView2Control).GUID;
    protected override Window CreateWindow(HWND parentHandle, RECT rect) => new WebView2Window(parentHandle, GetDefaultWindowStyle(parentHandle), rect);

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

        WebView2Window.Paint(hdc, bounds);
    }

    #region IDispatch implementation, static (IDL/TLB) and dynamic (reflection)
#pragma warning disable CA1822 // Mark members as static; no since we're dealing with COM instance methods & properties

    // COM visible public properties exposed through IWebView2Control IDL should go here
    public bool Enabled { get; set; } = true;
    HRESULT IWebView2Control.get_Enabled(out BOOL value) { value = Enabled; return Constants.S_OK; }
    HRESULT IWebView2Control.set_Enabled(BOOL value) { Enabled = value; return Constants.S_OK; }

    public string Caption { get; set; } = "WebView2";
    HRESULT IWebView2Control.get_Caption(out BSTR value) { value = new BSTR(Marshal.StringToBSTR(Caption)); return Constants.S_OK; }
    HRESULT IWebView2Control.set_Caption(BSTR value) { Caption = value.ToString() ?? string.Empty; return Constants.S_OK; }

#pragma warning restore CA1822 // Mark members as static
    #endregion

    #region COM Registration support

    // COM registration
    public static new HRESULT RegisterType(ComRegistrationContext context)
    {
        TracingUtilities.Trace($"Register type {typeof(WebView2Control).FullName}...");
        return BaseControl.RegisterType(context);
    }

    public static new HRESULT UnregisterType(ComRegistrationContext context)
    {
        TracingUtilities.Trace($"Unregister type {typeof(WebView2Control).FullName}...");
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
