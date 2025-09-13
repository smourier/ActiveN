// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

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
    private readonly DispatchConnectionPoint _eventsConnectionPoint; // disposed by BaseControl
    private WindowsDispatcherQueueController? _dispatcherQueueController;
    private Icon? _drawIcon;

    public PdfViewControl()
        : base()
    {
        // we need that to be able to use Windows.UI.Composition APIs
        // and no, we can't use Windows.System.DispatcherQueueController.CreateOnDedicatedThread()
        // even implicitly, or compositor will raise an access denied error
        _dispatcherQueueController = new WindowsDispatcherQueueController();
        _dispatcherQueueController.EnsureOnCurrentThread();

        // setup connection point for our (IDispatch) events
        // the advantage of IDispatch-based events is that we don't need to define interfaces or classes, 
        // we just need the Guid and Dispids
        var IID_IPdfViewControlEvents = new Guid("48c606f1-d597-467d-8a38-1c02fd7e019d"); // this must match the one in the .idl
        _eventsConnectionPoint = new DispatchConnectionPoint(IID_IPdfViewControlEvents);
        AddConnectionPoint(_eventsConnectionPoint);
        PropertyPagesIds = [typeof(PdfViewControlPage).GUID];
    }

    #region Mandatory overrides
    protected override ComRegistration ComRegistration => ComHosting.Instance;
    protected override Guid DispatchInterfaceId => typeof(IPdfViewControl).GUID;
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

    #region other overrides
    public new PdfViewWindow? Window => (PdfViewWindow?)base.Window; // disposed by BaseControl

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        Interlocked.Exchange(ref _dispatcherQueueController, null)?.Dispose();
        Interlocked.Exchange(ref _drawIcon, null)?.Dispose();
    }

    protected override void SetStockProperty(object? value, [CallerMemberName] string? name = null)
    {
        base.SetStockProperty(value, name);
        if (Window != null)
        {
            switch (name)
            {
                case nameof(BackColor):
                    Window.BackgroundColor = BackColor;
                    break;

                case nameof(ShowControls):
                    Window.ShowControls = ShowControls;
                    break;
            }
        }
    }

    protected override void Draw(HDC hdc, RECT bounds)
    {
        if (hdc == 0)
            return;

        if ((Window != null && Window.PageCount > 0) && InUserMode)
            return; // nothing to do, window will paint itself

        // we can get here if we've not been activated yet 
        // draw our icon centered & maxed in the bounds
        const int iconSize = 256;
        _drawIcon ??= Icon.Load(Functions.GetModuleHandleW(PWSTR.From(ComRegistration.DllPath)), 1, iconSize);
        if (_drawIcon != null)
        {
            TracingUtilities.Trace($"Drawing icon (size: {iconSize}) {_drawIcon.Handle} in {bounds}");

            var size = new SIZE(iconSize, iconSize);
            var factor = size.GetScaleFactor(bounds.Width, bounds.Height);
            var w = (int)(size.cx * factor.width);
            var h = (int)(size.cy * factor.height);
            var x = bounds.left + (bounds.Width - w) / 2;
            var y = bounds.top + (bounds.Height - h) / 2;
            TracingUtilities.Trace($"Drawing icon at {x},{y} {w}x{h}");
            Functions.DrawIconEx(hdc, x, y, _drawIcon.Handle, w, h, 0, HBRUSH.Null, DI_FLAGS.DI_NORMAL);
        }
    }
    #endregion

    #region IDispatch implementation, static (IDL/TLB) and dynamic (reflection)
#pragma warning disable CA1822 // Mark members as static; no since we're dealing with COM instance methods & properties

    // COM visible public properties exposed through IPdfView IDL should go here
    // types of properties here must convertible into variants or must implement from IValueGet
    public bool Enabled { get => GetStockProperty(true); set => SetStockProperty(value); }
    HRESULT IPdfViewControl.get_Enabled(out VARIANT_BOOL value) { value = Enabled; return Constants.S_OK; }
    HRESULT IPdfViewControl.set_Enabled(VARIANT_BOOL value) { Enabled = value; return Constants.S_OK; }

    [DispId(unchecked((int)DISPID.STDPROPID_XOBJ_NAME))]
    public string Name { get => GetStockProperty<string>() ?? nameof(PdfViewControl); set => SetStockProperty(value); }

    public string FilePath => Window?.FilePath ?? string.Empty;
    HRESULT IPdfViewControl.get_FilePath(out BSTR value) { value = new BSTR(Marshal.StringToBSTR(FilePath)); return Constants.S_OK; }

    public int PageCount => Window?.PageCount ?? -1;
    HRESULT IPdfViewControl.get_PageCount(out int value) { value = PageCount; return Constants.S_OK; }

    public int CurrentPageIndex => (int?)(Window?.CurrentPdfPage?.Index) ?? -1;
    HRESULT IPdfViewControl.get_CurrentPageIndex(out int value) { value = CurrentPageIndex; return Constants.S_OK; }

    public bool IsPasswordProtected => Window?.IsPasswordProtected ?? false;
    HRESULT IPdfViewControl.get_IsPasswordProtected(out VARIANT_BOOL value) { value = IsPasswordProtected; return Constants.S_OK; }

    // category (for host that support a property grid editor like VB/VBA)
    // can match PROPCAT_XXX or be a custom string
    [Category("Appearance")]
    [PropertyPage(PdfViewControlPage.ClsidString)]
    public bool ShowControls { get => GetStockProperty(true); set => SetStockProperty(value); }
    HRESULT IPdfViewControl.get_ShowControls(out VARIANT_BOOL value) { value = ShowControls; return Constants.S_OK; }
    HRESULT IPdfViewControl.set_ShowControls(VARIANT_BOOL value) { ShowControls = value; return Constants.S_OK; }

    public OLE_COLOR BackColor { get => GetStockProperty<OLE_COLOR>(PdfViewWindow.DefaultBackgroundColor); set => SetStockProperty(value); }
    HRESULT IPdfViewControl.get_BackColor(out OLE_COLOR value) { value = BackColor; return Constants.S_OK; }
    HRESULT IPdfViewControl.set_BackColor(OLE_COLOR value) { BackColor = value; return Constants.S_OK; }

    public HWND HWND => GetWindowHandle();
    HRESULT IPdfViewControl.get_HWND(out nint value) { value = HWND; return Constants.S_OK; }

    [DispId(unchecked((int)DISPID.DISPID_CAPTION))]
    public string Caption { get; set; } = $"ActiveN {nameof(PdfViewControl)}";

    public void OpenFile(string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        EnsureWindow(HWND.Null, new RECT(0, 0, 0, 0));
        Window?.OpenFile(filePath).Wait();
    }

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

    public void OpenStream(IStream stream) => Window?.OpenStream(new StreamOnIStream(stream));
    HRESULT IPdfViewControl.OpenStream(IStream stream)
    {
        if (stream == null)
            return Constants.E_POINTER;

        OpenStream(stream);
        return Constants.S_OK;
    }

    public PdfViewPage GetPage(int pageNumber)
    {
        // shows how to throw a nice error to COM clients
        var window = Window;
        if (window == null)
            TracingUtilities.ThrowArgument("Control was not initialized.");

        if (window.PdfDocument == null)
            TracingUtilities.ThrowArgument("No document was opened.");

        return new(this, window.PdfDocument.GetPage((uint)pageNumber));
    }

    HRESULT IPdfViewControl.GetPage(int pageNumber, out IPdfViewPage page) { page = GetPage(pageNumber); return Constants.S_OK; }

    public void CloseFile() => Window?.CloseFile();
    HRESULT IPdfViewControl.CloseFile() { CloseFile(); return Constants.S_OK; }

    public void MovePage(int delta) => Window?.MovePage(delta);
    HRESULT IPdfViewControl.MovePage(int delta) { MovePage(delta); return Constants.S_OK; }

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
