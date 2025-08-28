using IServiceProvider = DirectN.IServiceProvider;

namespace ActiveN;

[GeneratedComClass]
public abstract partial class BaseControl : BaseDispatch,
    IOleObject,
    IOleControl,
    IOleWindow,
    IDataObject,
    IObjectWithSite,
    IOleInPlaceActiveObject,
    IOleInPlaceObject,
    IQuickActivate,
    IServiceProvider,
    IProvideClassInfo,
    IPersistStreamInit,
    IViewObject2,
    IViewObjectEx,
    IPointerInactive,
    IConnectionPointContainer,
    ISpecifyPropertyPages,
    ICustomQueryInterface
{
    private readonly ConcurrentDictionary<Guid, IConnectionPoint> _connectionPoints = new();
    private readonly ComObject<IOleAdviseHolder> _adviseHolder;
    private ComObject<IOleClientSite>? _clientSite;
    private IComObject<IObjectWithSite>? _site;
    private ComObject<IAdviseSink>? _adviseSink;
    private PropertyNotifySinkConnectionPoint _connectionPoint;
    private bool _isDirty;
    private SIZE _extent;

    protected BaseControl()
    {
        TracingUtilities.Trace($"Created {GetType().FullName} ({GetType().GUID:B})");
        Functions.CreateOleAdviseHolder(out var obj).ThrowOnError();
        _adviseHolder = new ComObject<IOleAdviseHolder>(obj);
        _connectionPoint = new PropertyNotifySinkConnectionPoint();
        AddConnectionPoint(_connectionPoint);
    }

    protected virtual POINTERINACTIVE PointerActivationPolicy => POINTERINACTIVE.POINTERINACTIVE_ACTIVATEONENTRY;
    protected virtual OLEMISC MiscStatus => OLEMISC.OLEMISC_RECOMPOSEONRESIZE | OLEMISC.OLEMISC_CANTLINKINSIDE | OLEMISC.OLEMISC_INSIDEOUT | OLEMISC.OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC.OLEMISC_SETCLIENTSITEFIRST;
    protected virtual VIEWSTATUS ViewStatus => VIEWSTATUS.VIEWSTATUS_OPAQUE;

    protected virtual void AddConnectionPoint(IConnectionPoint connectionPoint)
    {
        ArgumentNullException.ThrowIfNull(connectionPoint);
        if (connectionPoint is BaseConnectionPoint target)
        {
            if (target._container != null)
                throw new ArgumentException("Connection point already has a container", nameof(connectionPoint));

            target._container = this;
        }

        connectionPoint.GetConnectionInterface(out var iid).ThrowOnError();
        if (!_connectionPoints.TryAdd(iid, connectionPoint))
            throw new ArgumentException($"Connection point with iid {iid} is already registered", nameof(connectionPoint));
    }

    CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out nint ppv) => GetInterface(ref iid, out ppv);
    protected virtual CustomQueryInterfaceResult GetInterface(ref Guid iid, out nint ppv)
    {
        ppv = 0;
        //// don't log these
        //if (iid == typeof(IOleLink).GUID || iid == typeof(IPersistStorage).GUID)
        //    return CustomQueryInterfaceResult.NotHandled;

        //// keep an eye on these
        //if (iid == typeof(IRunnableObject).GUID)
        //{
        //    TracingUtilities.Trace(typeof(IRunnableObject).Name);
        //}
        //else if (iid != typeof(IOleObject).GUID && iid != typeof(IProvideClassInfo).GUID && iid != typeof(IOleObject).GUID &&
        //    iid != typeof(IPersistStreamInit).GUID && iid != typeof(IViewObject2).GUID && iid != typeof(IViewObjectEx).GUID &&
        //    iid != typeof(IOleControl).GUID && iid != typeof(IPointerInactive).GUID)
        //{
        //    TracingUtilities.Trace($"GetInterface: {GuidNames.GetInterfaceIdName(iid)}");
        //}
        TracingUtilities.Trace($"iid: {iid.GetName()}");
        return CustomQueryInterfaceResult.NotHandled;
    }

    protected virtual void Close(OLECLOSE close)
    {
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _adviseHolder?.Dispose();
            _clientSite?.Dispose();
            _site?.Dispose();
        }
        base.Dispose(disposing);
    }

    protected static HRESULT RegisterType(ComRegistrationContext context)
    {
        ArgumentNullException.ThrowIfNull(context);
        TracingUtilities.Trace($"Register type {typeof(BaseControl).FullName}...");

        // add the "Control" subkey to indicate that this is an ActiveX control
        using var key = ComRegistration.EnsureWritableSubKey(context.RegistryRoot, Path.Combine(ComRegistration.ClsidRegistryKey, context.GUID.ToString("B")));

        if (context.ImplementedCategories.Contains(ControlCategories.CATID_Control))
        {
            key.CreateSubKey("Control", false)?.Dispose();
        }

        if (context.ImplementedCategories.Contains(ControlCategories.CATID_Insertable))
        {
            key.CreateSubKey("Insertable", false)?.Dispose();
        }

        if (context.ImplementedCategories.Contains(ControlCategories.CATID_Programmable))
        {
            key.CreateSubKey("Programmable", false)?.Dispose();
        }

        if (context.ImplementedCategories?.Count > 0)
        {
            using var cats = key.CreateSubKey("Implemented Categories", true);
            foreach (var cat in context.ImplementedCategories)
            {
                cats.CreateSubKey($"{cat:B}", false)?.Dispose();
            }
        }

        using var miscStatus = key.CreateSubKey("MiscStatus", true);
        miscStatus.SetValue(null, ((int)context.MiscStatus).ToString());

        // we currently only support bitmaps with id 1
        if (ResourceUtilities.HasEmbeddedBitmap(context.Registration.DllPath))
        {
            using var bitmap = key.CreateSubKey("ToolboxBitmap32", true);
            bitmap.SetValue(null, $"{context.Registration.DllPath},1");
        }

        return Constants.S_OK;
    }

    protected static HRESULT UnregisterType(ComRegistrationContext context)
    {
        ArgumentNullException.ThrowIfNull(context);
        TracingUtilities.Trace($"Unregister type {typeof(BaseControl).FullName}...");
        return Constants.S_OK;
    }

    protected virtual HRESULT NotImplemented([CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        TracingUtilities.Trace($"E_NOIMPL", methodName, filePath);
        return Constants.E_NOTIMPL;
    }

    HRESULT IOleObject.SetClientSite(IOleClientSite pClientSite)
    {
        _clientSite?.Dispose();
        _clientSite = pClientSite != null ? new ComObject<IOleClientSite>(pClientSite) : null;
        TracingUtilities.Trace($"Site: {pClientSite}");
        return Constants.S_OK;
    }

    HRESULT IOleObject.GetClientSite(out IOleClientSite ppClientSite)
    {
        ppClientSite = _clientSite?.Object!;
        TracingUtilities.Trace($"Site: {ppClientSite}");
        return Constants.S_OK;
    }

    HRESULT IOleObject.SetHostNames(PWSTR szContainerApp, PWSTR szContainerObj)
    {
        TracingUtilities.Trace($"szContainerApp: '{szContainerApp}' szContainerObj: '{szContainerObj}'");
        return Constants.S_OK;
    }

    HRESULT IOleObject.Close(uint dwSaveOption) => ComRegistration.WrapErrors(() => Close((OLECLOSE)dwSaveOption));

    HRESULT IOleObject.SetMoniker(uint dwWhichMoniker, IMoniker pmk) => NotImplemented();

    HRESULT IOleObject.GetMoniker(uint dwAssign, uint dwWhichMoniker, out IMoniker ppmk) { ppmk = null!; return NotImplemented(); }

    HRESULT IOleObject.InitFromData(IDataObject pDataObject, BOOL fCreation, uint dwReserved) => NotImplemented();

    HRESULT IOleObject.GetClipboardData(uint dwReserved, out IDataObject ppDataObject) { ppDataObject = null!; return NotImplemented(); }

    unsafe HRESULT IOleObject.DoVerb(int iVerb, nint lpmsg, IOleClientSite pActiveSite, int lindex, HWND hwndParent, nint lprcPosRect)
    {
        var verb = (OLEIVERB)iVerb;
        var hr = Constants.S_OK;
        TracingUtilities.Trace($"iVerb: {verb} lpmsg: {lpmsg} pActiveSite: {pActiveSite} lindex: {lindex} hwndParent: {hwndParent} lprcPosRect: {lprcPosRect}");
        if (lprcPosRect != 0)
        {
            var rcPosRect = *(RECT*)lprcPosRect;
            TracingUtilities.Trace($"rcPosRect: {rcPosRect}");
        }

        if (verb == OLEIVERB.OLEIVERB_INPLACEACTIVATE)
        {
            if (lprcPosRect != 0)
            {
                var rcPosRect = *(RECT*)lprcPosRect;
                _extent = rcPosRect.Size;
                TracingUtilities.Trace($"extent: {_extent}");
            }
            hr = pActiveSite.ShowObject();
        }
        TracingUtilities.Trace($"hr: {hr}");
        return hr;
    }

    HRESULT IOleObject.EnumVerbs(out IEnumOLEVERB ppEnumOleVerb)
    {
        TracingUtilities.Trace();
        ppEnumOleVerb = new EnumVerbs([]);
        return Constants.S_OK;
    }

    HRESULT IOleObject.Update() => NotImplemented();

    HRESULT IOleObject.IsUpToDate() => NotImplemented();

    HRESULT IOleObject.GetUserClassID(out Guid pClsid) { pClsid = new(); return NotImplemented(); }

    HRESULT IOleObject.GetUserType(uint dwFormOfType, out PWSTR pszUserType)
    {
        var hr = Functions.OleRegGetUserType(GetType().GUID, dwFormOfType, out pszUserType);
        TracingUtilities.Trace($"dwFormOfType: {dwFormOfType} pszUserType: '{pszUserType}' hr: {hr}");
        return hr;
    }

    HRESULT IOleObject.SetExtent(DVASPECT dwDrawAspect, in SIZE psizel)
    {
        TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect} psizel: {psizel}");
        if (dwDrawAspect != DVASPECT.DVASPECT_CONTENT)
            return Constants.DV_E_DVASPECT;

        _extent = psizel;
        return Constants.S_OK;
    }

    HRESULT IOleObject.GetExtent(DVASPECT dwDrawAspect, out SIZE psizel)
    {
        TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect}");
        if (dwDrawAspect != DVASPECT.DVASPECT_CONTENT)
        {
            psizel = new();
            return Constants.DV_E_DVASPECT;
        }

        psizel = _extent;
        return Constants.S_OK;
    }

    HRESULT IOleObject.Advise(IAdviseSink pAdvSink, out uint pdwConnection)
    {
        var hr = _adviseHolder.Object.Advise(pAdvSink, out pdwConnection);
        TracingUtilities.Trace($"dwConnection {pdwConnection}: {hr}");
        return hr;
    }

    HRESULT IOleObject.Unadvise(uint dwConnection)
    {
        var hr = _adviseHolder.Object.Unadvise(dwConnection);
        TracingUtilities.Trace($"dwConnection {dwConnection}: {hr}");
        return hr;
    }

    HRESULT IOleObject.EnumAdvise(out IEnumSTATDATA ppenumAdvise)
    {
        var hr = _adviseHolder.Object.EnumAdvise(out ppenumAdvise);
        TracingUtilities.Trace($"ppenumAdvise {ppenumAdvise}: {hr}");
        return hr;
    }

    HRESULT IOleObject.GetMiscStatus(DVASPECT dwAspect, out OLEMISC pdwStatus)
    {
        pdwStatus = MiscStatus;
        TracingUtilities.Trace($"dwAspect: {dwAspect} pdwStatus: {pdwStatus}");
        return Constants.S_OK;
    }

    HRESULT IOleObject.SetColorScheme(in LOGPALETTE pLogpal) => NotImplemented();

    HRESULT IProvideClassInfo.GetClassInfo(out ITypeInfo ppTI)
    {
        var ti = EnsureTypeInfo();
        ppTI = ti != null ? ti.Object : null!;
        TracingUtilities.Trace($"ppTI {ppTI}");
        return ti != null ? Constants.S_OK : Constants.E_UNEXPECTED;
    }

    HRESULT IPersistStreamInit.IsDirty()
    {
        TracingUtilities.Trace($"IsDirty: {_isDirty}");
        return _isDirty ? Constants.S_OK : Constants.S_FALSE;
    }

    HRESULT IPersistStreamInit.Load(IStream pStm)
    {
        TracingUtilities.Trace($"pStm {pStm}");
        _isDirty = false;
        return Constants.S_OK;
    }

    HRESULT IPersistStreamInit.Save(IStream pStm, BOOL fClearDirty)
    {
        _isDirty = !fClearDirty;
        TracingUtilities.Trace($"pStm {pStm} fClearDirty: {fClearDirty}");
        return Constants.S_OK;
    }

    HRESULT IPersistStreamInit.GetSizeMax(out ulong pCbSize) { pCbSize = 0; return NotImplemented(); }

    HRESULT IPersistStreamInit.InitNew()
    {
        _isDirty = true;
        TracingUtilities.Trace();
        return Constants.S_OK;
    }

    HRESULT IPersist.GetClassID(out Guid pClassID)
    {
        pClassID = GetType().GUID;
        TracingUtilities.Trace();
        return Constants.S_OK;
    }

    HRESULT IViewObject2.GetExtent(DVASPECT dwDrawAspect, int lindex, nint ptd, out SIZE lpsizel)
    {
        TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect} lindex: {lindex} ptd: {ptd}");
        lpsizel = _extent;
        return Constants.S_OK;
    }

    HRESULT IViewObject.Draw(DVASPECT dwDrawAspect, int lindex, nint pvAspect, nint ptd, HDC hdcTargetDev, HDC hdcDraw, nint lprcBounds, nint lprcWBounds, nint pfnContinue, nuint dwContinue)
    {
        TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect} lindex: {lindex} pvAspect: {pvAspect} ptd: {ptd} hdcTargetDev: {hdcTargetDev} hdcDraw: {hdcDraw} lprcBounds: {lprcBounds} lprcWBounds: {lprcWBounds} pfnContinue: {pfnContinue} dwContinue: {dwContinue}");
        return Constants.S_OK;
    }

    HRESULT IViewObject.GetColorSet(DVASPECT dwDrawAspect, int lindex, nint pvAspect, nint ptd, HDC hicTargetDev, out nint ppColorSet) { ppColorSet = 0; return NotImplemented(); }
    HRESULT IViewObject.Freeze(DVASPECT dwDrawAspect, int lindex, nint pvAspect, out uint pdwFreeze) { pdwFreeze = 0; return NotImplemented(); }
    HRESULT IViewObject.Unfreeze(uint dwFreeze) => NotImplemented();
    HRESULT IViewObject.SetAdvise(DVASPECT aspects, uint advf, IAdviseSink pAdvSink)
    {
        _adviseSink?.Dispose();
        _adviseSink = pAdvSink != null ? new ComObject<IAdviseSink>(pAdvSink) : null;
        TracingUtilities.Trace($"Sink: {pAdvSink}");
        return Constants.S_OK;
    }

    HRESULT IViewObject.GetAdvise(nint pAspects, nint pAdvf, out IAdviseSink ppAdvSink)
    {
        ppAdvSink = _adviseSink?.Object!;
        TracingUtilities.Trace($"Site: {ppAdvSink}");
        return Constants.S_OK;
    }

    HRESULT IViewObjectEx.GetRect(uint dwAspect, out RECTL pRect) { pRect = new(); return NotImplemented(); }
    HRESULT IViewObjectEx.GetViewStatus(out uint pdwStatus)
    {
        pdwStatus = (uint)ViewStatus;
        TracingUtilities.Trace($"pdwStatus: {ViewStatus}");
        return Constants.S_OK;
    }

    HRESULT IViewObjectEx.QueryHitPoint(uint dwAspect, in RECT pRectBounds, POINT ptlLoc, int lCloseHint, out uint pHitResult)
    {
        var aspect = (DVASPECT)dwAspect;
        //TracingUtilities.Trace($"dwAspect: {aspect} pRectBounds: {pRectBounds} ptlLoc: {ptlLoc} lCloseHint: {lCloseHint}");

        if (aspect == DVASPECT.DVASPECT_CONTENT)
        {
            var result = Functions.PtInRect(pRectBounds, ptlLoc) ? TXTHITRESULT.TXTHITRESULT_HIT : TXTHITRESULT.TXTHITRESULT_NOHIT;
            //TracingUtilities.Trace($"pHitResult: {result}");
            pHitResult = (uint)result;
            return Constants.S_OK;
        }

        pHitResult = 0;
        TracingUtilities.Trace($"E_FAIL");
        return Constants.E_FAIL;
    }

    HRESULT IViewObjectEx.QueryHitRect(uint dwAspect, in RECT pRectBounds, in RECT pRectLoc, int lCloseHint, out uint pHitResult) { pHitResult = 0; return NotImplemented(); }
    HRESULT IViewObjectEx.GetNaturalExtent(DVASPECT dwAspect, int lindex, nint ptd, HDC hicTargetDev, nint pExtentInfo, nint pSizel) { pSizel = 0; return NotImplemented(); }

    HRESULT IOleControl.GetControlInfo(ref CONTROLINFO pCI) { pCI = new(); return NotImplemented(); }
    HRESULT IOleControl.OnMnemonic(in MSG pMsg)
    {
        TracingUtilities.Trace($"pMSg: {MessageDecoder.Decode(pMsg)}");
        return Constants.S_OK;
    }

    HRESULT IOleControl.OnAmbientPropertyChange(int dispID)
    {
        TracingUtilities.Trace($"dispID: {dispID}");
        return Constants.S_OK;
    }

    HRESULT IOleControl.FreezeEvents(BOOL bFreeze)
    {
        TracingUtilities.Trace($"bFreeze: {bFreeze}");
        return Constants.S_OK;
    }

    HRESULT IPointerInactive.GetActivationPolicy(out POINTERINACTIVE pdwPolicy)
    {
        pdwPolicy = PointerActivationPolicy;
        TracingUtilities.Trace($"pdwPolicy: {pdwPolicy}");
        return Constants.S_OK;
    }

    HRESULT IPointerInactive.OnInactiveMouseMove(in RECT pRectBounds, int x, int y, uint grfKeyState)
    {
        TracingUtilities.Trace($"pRectBounds: {pRectBounds} x: {x} y: {y} grfKeyState: {grfKeyState}");
        return Constants.S_OK;
    }

    HRESULT IPointerInactive.OnInactiveSetCursor(in RECT pRectBounds, int x, int y, uint dwMouseMsg, BOOL fSetAlways)
    {
        TracingUtilities.Trace($"pRectBounds: {pRectBounds} x: {x} y: {y} dwMouseMsg: {dwMouseMsg} fSetAlways: {fSetAlways}");
        return Constants.S_OK;
    }

    HRESULT ISpecifyPropertyPages.GetPages(out CAUUID pPages)
    {
        pPages.pElems = 0;
        pPages.cElems = 0;
        TracingUtilities.Trace($"cElems: {pPages.cElems}");
        return Constants.S_OK;
    }

    HRESULT IConnectionPointContainer.EnumConnectionPoints(out IEnumConnectionPoints ppEnum)
    {
        ppEnum = new EnumConnectionPoints([.. _connectionPoints]);
        return Constants.S_OK;
    }

    HRESULT IConnectionPointContainer.FindConnectionPoint(in Guid riid, out IConnectionPoint ppCP)
    {
        _connectionPoints.TryGetValue(riid, out var cp);
        ppCP = cp!;
        return cp == null ? Constants.CONNECT_E_NOCONNECTION : Constants.S_OK;
    }

    HRESULT IOleWindow.ContextSensitiveHelp(BOOL fEnterMode) => NotImplemented();
    HRESULT IOleWindow.GetWindow(out HWND phwnd)
    {
        TracingUtilities.Trace();
        phwnd = GetWindowHandle();
        TracingUtilities.Trace($"phwnd: {phwnd}");
        return Constants.S_OK;
    }

    HRESULT IDataObject.GetData(in FORMATETC pformatetcIn, out STGMEDIUM pmedium) { pmedium = new(); return NotImplemented(); }
    HRESULT IDataObject.GetDataHere(in FORMATETC pformatetc, ref STGMEDIUM pmedium) => NotImplemented();
    HRESULT IDataObject.QueryGetData(in FORMATETC pformatetc) => NotImplemented();
    HRESULT IDataObject.GetCanonicalFormatEtc(in FORMATETC pformatectIn, out FORMATETC pformatetcOut) { pformatetcOut = new(); return NotImplemented(); }
    HRESULT IDataObject.SetData(in FORMATETC pformatetc, in STGMEDIUM pmedium, BOOL fRelease) => NotImplemented();
    HRESULT IDataObject.EnumFormatEtc(uint dwDirection, out IEnumFORMATETC ppenumFormatEtc) { ppenumFormatEtc = null!; return NotImplemented(); }
    HRESULT IDataObject.DAdvise(in FORMATETC pformatetc, uint advf, IAdviseSink pAdvSink, out uint pdwConnection) { pdwConnection = 0; return NotImplemented(); }
    HRESULT IDataObject.DUnadvise(uint dwConnection) => NotImplemented();
    HRESULT IDataObject.EnumDAdvise(out IEnumSTATDATA ppenumAdvise) { ppenumAdvise = null!; return NotImplemented(); }

    HRESULT IObjectWithSite.SetSite(nint pUnkSite)
    {
        _site?.Dispose();
        _site = DirectN.Extensions.Com.ComObject.FromPointer<IObjectWithSite>(pUnkSite);
        TracingUtilities.Trace($"Site: {pUnkSite}");
        return Constants.S_OK;
    }

    HRESULT IObjectWithSite.GetSite(in Guid riid, out nint ppvSite)
    {
        TracingUtilities.Trace($"riid: {riid.GetName()}");
        if (_site == null)
        {
            ppvSite = 0;
            return Constants.E_NOINTERFACE;
        }

        ppvSite = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(_site);
        return ppvSite != 0 ? Constants.S_OK : Constants.E_NOINTERFACE;
    }

    HRESULT IOleInPlaceActiveObject.TranslateAccelerator(nint lpmsg) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.OnFrameWindowActivate(BOOL fActivate) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.OnDocWindowActivate(BOOL fActivate) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.ResizeBorder(in RECT prcBorder, IOleInPlaceUIWindow pUIWindow, BOOL fFrameWindow) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.EnableModeless(BOOL fEnable) => NotImplemented();
    HRESULT IOleInPlaceObject.InPlaceDeactivate() => NotImplemented();
    HRESULT IOleInPlaceObject.UIDeactivate() => NotImplemented();
    HRESULT IOleInPlaceObject.SetObjectRects(in RECT lprcPosRect, in RECT lprcClipRect) => NotImplemented();
    HRESULT IOleInPlaceObject.ReactivateAndUndo() => NotImplemented();
    unsafe HRESULT IQuickActivate.QuickActivate(in QACONTAINER pQaContainer, ref QACONTROL pQaControl)
    {
        TracingUtilities.Trace($"pQaContainer: {pQaContainer} pQacontrol size: {pQaControl.cbSize}");
        pQaControl.cbSize = (uint)sizeof(QACONTROL);
        pQaControl.dwMiscStatus = (uint)MiscStatus;
        pQaControl.dwViewStatus = (uint)ViewStatus;
        pQaControl.dwPointerActivationPolicy = (uint)PointerActivationPolicy;
        TracingUtilities.Trace($"dwMiscStatus: {MiscStatus} dwViewStatus: {ViewStatus} dwPointerActivationPolicy: {PointerActivationPolicy}");
        return Constants.S_OK;
    }

    HRESULT IQuickActivate.SetContentExtent(in SIZE pSizel)
    {
        TracingUtilities.Trace($"pSizel: {pSizel}");
        _extent = pSizel;
        return Constants.S_OK;
    }

    HRESULT IQuickActivate.GetContentExtent(out SIZE pSizel)
    {
        TracingUtilities.Trace($"pSizel: {_extent}");
        pSizel = _extent;
        return Constants.S_OK;
    }

    HRESULT IServiceProvider.QueryService(in Guid guidService, in Guid riid, out nint ppvObject)
    {
        TracingUtilities.Trace($"guidService: {guidService.GetName()} riid: {riid.GetName()}");
        ppvObject = 0;
        return Constants.E_NOINTERFACE;
    }
}
