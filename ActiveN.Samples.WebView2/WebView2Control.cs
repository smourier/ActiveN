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
    public const string Category = "WebView2";

    public WebView2Control()
    {
        PropertyPagesIds = [typeof(WebView2ControlPage).GUID];
    }

    #region Mandatory overrides
    protected override ComRegistration ComRegistration => ComHosting.Instance;
    protected override Guid DispatchInterfaceId => typeof(IWebView2Control).GUID;
    protected override Window CreateWindow(HWND parentHandle, RECT rect) => new WebView2Window(parentHandle, GetDefaultWindowStyle(parentHandle), rect);
    #endregion

    public new WebView2Window? Window => (WebView2Window?)base.Window; // disposed by BaseControl

    protected override void Draw(HDC hdc, RECT bounds)
    {
        if (hdc == 0)
            return;

        WebView2Window.Paint(hdc, bounds);
    }

    #region IDispatch implementation, static (IDL/TLB) and dynamic (reflection)
#pragma warning disable CA1822 // Mark members as static; no since we're dealing with COM instance methods & properties

    // COM visible public properties exposed through IWebView2Control IDL should go here
    [DispId(unchecked((int)DISPID.STDPROPID_XOBJ_NAME))]
    public string Name { get => GetStockProperty<string>() ?? nameof(WebView2Control); set => SetStockProperty(value); }

    [Category(Category)]
    public string BrowserVersion => WebView2Window.BrowserVersion;

    [Category(Category)]
    public string Source => Window?.Source ?? string.Empty;

    [Category(Category)]
    public string StatusBarText => Window?.StatusBarText ?? string.Empty;

    [Category(Category)]
    public string DocumentTitle => Window?.DocumentTitle ?? string.Empty;

    HRESULT IWebView2Control.get_BrowserVersion(out BSTR value) { value = new BSTR(Marshal.StringToBSTR(BrowserVersion)); return DirectN.Constants.S_OK; }
    HRESULT IWebView2Control.get_Source(out BSTR value) { value = new BSTR(Marshal.StringToBSTR(Source)); return DirectN.Constants.S_OK; }
    HRESULT IWebView2Control.get_StatusBarText(out BSTR value) { value = new BSTR(Marshal.StringToBSTR(StatusBarText)); return DirectN.Constants.S_OK; }
    HRESULT IWebView2Control.get_DocumentTitle(out BSTR value) { value = new BSTR(Marshal.StringToBSTR(DocumentTitle)); return DirectN.Constants.S_OK; }

    HRESULT IWebView2Control.GoBack() { Window?.GoBack(); return DirectN.Constants.S_OK; }
    HRESULT IWebView2Control.GoForward() { Window?.GoForward(); return DirectN.Constants.S_OK; }
    HRESULT IWebView2Control.Reload() { Window?.Reload(); return DirectN.Constants.S_OK; }
    HRESULT IWebView2Control.Navigate(BSTR uri)
    {
        if (uri.Value == 0)
            return DirectN.Constants.E_POINTER;

        Window?.Navigate(Marshal.PtrToStringBSTR(uri.Value) ?? string.Empty);
        return DirectN.Constants.S_OK;
    }

    HRESULT IWebView2Control.NavigateToString(BSTR htmlContent)
    {
        if (htmlContent.Value == 0)
            return DirectN.Constants.E_POINTER;

        Window?.NavigateToString(Marshal.PtrToStringBSTR(htmlContent.Value) ?? string.Empty);
        return DirectN.Constants.S_OK;
    }

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
