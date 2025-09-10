namespace ActiveN;

[GeneratedComClass]
public abstract partial class BasePropertyPage : ICustomQueryInterface, IDisposable,
    IPropertyPage
{
    private IComObject<IPropertyPageSite>? _site;
    private Window? _window;
    private readonly List<nint> _objects = [];

    protected virtual bool IsDirty { get; set; }
    protected virtual Window? Window => _window;
    protected IReadOnlyList<nint> Objects => _objects;

    protected virtual string Title => GetType().Name;
    protected virtual SIZE InitialSize => new(600, 400);

    protected virtual void OnSendStatus(uint status)
    {
        _site?.Object?.OnStatusChange(status);
    }

    protected virtual void Activate(HWND parent, RECT rect, bool model)
    {
        if (_objects.Count == 0)
            return;

        var obj = DirectN.Extensions.Com.ComObject.FromPointer<IDispatch>(_objects[0]);
        if (obj == null)
            return;

        // this may be an outer from aggregation, so we must QI for the real object
        TracingUtilities.Trace($" object[0]: {obj.Object}");
        PropertyGrid.Show(obj);
    }

    protected virtual void DisposeObjects()
    {
        foreach (var unk in _objects)
        {
            if (unk != 0)
            {
                Marshal.Release(unk);
            }
        }
        _objects.Clear();
    }

    CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out nint ppv) => GetInterface(ref iid, out ppv);
    protected virtual CustomQueryInterfaceResult GetInterface(ref Guid iid, out nint ppv)
    {
        ppv = 0;
        TracingUtilities.Trace($"iid: {iid.GetName()}");
        return CustomQueryInterfaceResult.NotHandled;
    }

    protected virtual HRESULT NotImplemented([CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        TracingUtilities.Trace($"E_NOIMPL", methodName, filePath);
        return Constants.E_NOTIMPL;
    }

    protected virtual void DisposeWindow()
    {
        TracingUtilities.Trace($"window: {_window}");
        var window = Interlocked.Exchange(ref _window, null);
        window?.Dispose();
    }

    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            Interlocked.Exchange(ref _site, null)?.Dispose();
            DisposeWindow();
            DisposeObjects();
        }
    }

    HRESULT IPropertyPage.Activate(HWND hWndParent, in RECT pRect, BOOL bModal)
    {
        TracingUtilities.Trace($"hWndParent: {hWndParent}, pRect: {pRect}, bModal: {bModal}");
        var rc = pRect;
        return TracingUtilities.WrapErrors(() =>
        {
            Activate(hWndParent, rc, bModal);
            return Constants.S_OK;
        });
    }

    HRESULT IPropertyPage.Apply()
    {
        TracingUtilities.Trace();
        return Constants.S_OK;
    }

    HRESULT IPropertyPage.Deactivate()
    {
        TracingUtilities.Trace();
        return Constants.S_OK;
    }

    unsafe HRESULT IPropertyPage.GetPageInfo(out PROPPAGEINFO pPageInfo)
    {
        var pageInfo = new PROPPAGEINFO();
        var hr = TracingUtilities.WrapErrors(() =>
        {
            pageInfo.cb = (uint)sizeof(PROPPAGEINFO);
            pageInfo.pszTitle = new PWSTR(Marshal.StringToCoTaskMemUni(Title));
            pageInfo.size = InitialSize;
            return Constants.S_OK;
        }, () =>
        {
            if (pageInfo.pszTitle.Value != 0)
            {
                Marshal.FreeCoTaskMem(pageInfo.pszTitle.Value);
                pageInfo.pszTitle = PWSTR.Null;
            }
        });
        pPageInfo = pageInfo;
        return hr;
    }

    HRESULT IPropertyPage.Help(PWSTR pszHelpDir) => NotImplemented();
    HRESULT IPropertyPage.IsPageDirty()
    {
        TracingUtilities.Trace($"IsDirty: {IsDirty}");
        return IsDirty ? Constants.S_OK : Constants.S_FALSE;
    }

    HRESULT IPropertyPage.Move(in RECT pRect)
    {
        TracingUtilities.Trace($"{pRect}");
        return Constants.S_OK;
    }

    HRESULT IPropertyPage.SetObjects(uint cObjects, nint[] ppUnk)
    {
        TracingUtilities.Trace($"cObjects: {cObjects}, ppUnk: {ppUnk?.Length}");

        DisposeObjects();
        for (var i = 0; i < cObjects; i++)
        {
            var unk = ppUnk?[i] ?? 0;
            if (unk == 0)
                continue;

            _objects.Add(unk);
        }
        return Constants.S_OK;
    }

    HRESULT IPropertyPage.SetPageSite(IPropertyPageSite pPageSite) => TracingUtilities.WrapErrors(() =>
    {
        Interlocked.Exchange(ref _site, null)?.Dispose();
        _site = _site != null ? new ComObject<IPropertyPageSite>(pPageSite) : null;

        TracingUtilities.Trace($"PropertyPageSite: {_site}");
        return Constants.S_OK;
    });

    HRESULT IPropertyPage.Show(uint nCmdShow)
    {
        TracingUtilities.Trace($"nCmdShow: {(SHOW_WINDOW_CMD)nCmdShow}");
        return Constants.S_OK;
    }

    HRESULT IPropertyPage.TranslateAccelerator(in MSG pMsg)
    {
        TracingUtilities.Trace($"{MessageDecoder.Decode(pMsg)}");
        // WM_KEYDOWN = WM_KEYFIRST WM_UNICHAR = WM_KEYLAST
        // WM_MOUSEMOVE = WM_MOUSEFIRST WM_MOUSEHWHEEL = WM_MOUSELAST
        if ((pMsg.message < MessageDecoder.WM_KEYDOWN || pMsg.message > MessageDecoder.WM_UNICHAR) &&
            (pMsg.message < MessageDecoder.WM_MOUSEMOVE || pMsg.message > MessageDecoder.WM_MOUSEHWHEEL))
            return Constants.S_FALSE;

        var window = _window;
        if (window == null || window.Handle == 0)
            return Constants.S_FALSE;

        return Functions.IsDialogMessageW(window.Handle, pMsg) ? Constants.S_OK : Constants.S_FALSE;
    }
}
