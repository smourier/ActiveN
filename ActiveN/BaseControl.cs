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
    private ComObject<IAdviseSink>? _adviseSink;
    private bool _isDirty;

    protected BaseControl()
    {
        TracingUtilities.Trace($"Created {GetType().FullName} ({GetType().GUID:B})");
        Functions.CreateOleAdviseHolder(out var obj).ThrowOnError();
        _adviseHolder = new ComObject<IOleAdviseHolder>(obj);
    }

    protected virtual POINTERINACTIVE PointerActivationPolicy => POINTERINACTIVE.POINTERINACTIVE_ACTIVATEONENTRY;
    protected virtual OLEMISC MiscStatus => OLEMISC.OLEMISC_RECOMPOSEONRESIZE | OLEMISC.OLEMISC_CANTLINKINSIDE | OLEMISC.OLEMISC_INSIDEOUT | OLEMISC.OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC.OLEMISC_SETCLIENTSITEFIRST;

    protected virtual void AddConnectionPoint(IConnectionPoint connectionPoint)
    {
        ArgumentNullException.ThrowIfNull(connectionPoint);
        if (connectionPoint is ConnectionPoint target)
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
        TracingUtilities.Trace($"GetInterface: {iid.GetName()}");
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
        }
        base.Dispose(disposing);
    }

    protected static HRESULT RegisterType(ComRegistrationContext context)
    {
        ArgumentNullException.ThrowIfNull(context);
        TracingUtilities.Trace($"Register type {typeof(BaseControl).FullName}...");

        // add the "Control" subkey to indicate that this is an ActiveX control
        using var key = ComRegistration.EnsureWritableSubKey(context.RegistryRoot, Path.Combine(ComRegistration.ClsidRegistryKey, context.GUID.ToString("B")));

        if (context.ImplementedCategories.Contains(Categories.CATID_Control))
        {
            key.CreateSubKey("Control", false)?.Dispose();
        }

        if (context.ImplementedCategories.Contains(Categories.CATID_Insertable))
        {
            key.CreateSubKey("Insertable", false)?.Dispose();
        }

        if (context.ImplementedCategories.Contains(Categories.CATID_Programmable))
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

    HRESULT IOleObject.DoVerb(int iVerb, nint lpmsg, IOleClientSite pActiveSite, int lindex, HWND hwndParent, nint lprcPosRect)
    {
        var verb = (OLEIVERB)iVerb;
        TracingUtilities.Trace($"iVerb: {verb} lpmsg: {lpmsg} pActiveSite: {pActiveSite} lindex: {lindex} hwndParent: {hwndParent} lprcPosRect: {lprcPosRect}");
        return Constants.S_OK;
    }

    HRESULT IOleObject.EnumVerbs(out IEnumOLEVERB ppEnumOleVerb)
    {
        ppEnumOleVerb = new Verbs([]);
        return Constants.S_OK;
    }

    HRESULT IOleObject.Update() => NotImplemented();

    HRESULT IOleObject.IsUpToDate() => NotImplemented();

    HRESULT IOleObject.GetUserClassID(out Guid pClsid) { pClsid = new(); return NotImplemented(); }

    HRESULT IOleObject.GetUserType(uint dwFormOfType, out PWSTR pszUserType)
    {
        var hr = Functions.OleRegGetUserType(GetType().GUID, dwFormOfType, out pszUserType);
        return hr;
    }

    HRESULT IOleObject.SetExtent(DVASPECT dwDrawAspect, in SIZE psizel) => NotImplemented();

    HRESULT IOleObject.GetExtent(DVASPECT dwDrawAspect, out SIZE psizel) { psizel = new(); return NotImplemented(); }

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
        var hr = Functions.OleRegGetMiscStatus(GetType().GUID, (uint)dwAspect, out var status);
        pdwStatus = (OLEMISC)status;
        return hr;
    }

    HRESULT IOleObject.SetColorScheme(in LOGPALETTE pLogpal) => NotImplemented();

    HRESULT IProvideClassInfo.GetClassInfo(out ITypeInfo ppTI)
    {
        var ti = EnsureTypeInfo();
        ppTI = ti != null ? ti.Object : null!;
        TracingUtilities.Trace($"ppTI {ppTI}");
        return ti != null ? Constants.S_OK : Constants.E_UNEXPECTED;
    }

    HRESULT IPersistStreamInit.IsDirty() => _isDirty ? Constants.S_OK : Constants.S_FALSE;

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

    HRESULT IViewObject2.GetExtent(DVASPECT dwDrawAspect, int lindex, in DVTARGETDEVICE ptd, out SIZE lpsizel) { lpsizel = new(); return NotImplemented(); }
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
        var status = VIEWSTATUS.VIEWSTATUS_OPAQUE;
        pdwStatus = (uint)status;
        TracingUtilities.Trace($"pdwStatus: {status}");
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
    HRESULT IViewObjectEx.GetNaturalExtent(DVASPECT dwAspect, int lindex, in DVTARGETDEVICE ptd, HDC hicTargetDev, in DVEXTENTINFO pExtentInfo, out SIZE pSizel) { pSizel = new(); return NotImplemented(); }

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
        ppEnum = new ConnectionsPoints([.. _connectionPoints]);
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
        phwnd = GetWindowHandle();
        TracingUtilities.Trace($"phwnd: {phwnd}");
        return Constants.S_OK;
    }

    HRESULT IDataObject.GetData(in FORMATETC pformatetcIn, out STGMEDIUM pmedium)
    {
        throw new NotImplementedException();
    }

    HRESULT IDataObject.GetDataHere(in FORMATETC pformatetc, ref STGMEDIUM pmedium)
    {
        throw new NotImplementedException();
    }

    HRESULT IDataObject.QueryGetData(in FORMATETC pformatetc)
    {
        throw new NotImplementedException();
    }

    HRESULT IDataObject.GetCanonicalFormatEtc(in FORMATETC pformatectIn, out FORMATETC pformatetcOut)
    {
        throw new NotImplementedException();
    }

    HRESULT IDataObject.SetData(in FORMATETC pformatetc, in STGMEDIUM pmedium, BOOL fRelease)
    {
        throw new NotImplementedException();
    }

    HRESULT IDataObject.EnumFormatEtc(uint dwDirection, out IEnumFORMATETC ppenumFormatEtc)
    {
        throw new NotImplementedException();
    }

    HRESULT IDataObject.DAdvise(in FORMATETC pformatetc, uint advf, IAdviseSink pAdvSink, out uint pdwConnection)
    {
        throw new NotImplementedException();
    }

    HRESULT IDataObject.DUnadvise(uint dwConnection)
    {
        throw new NotImplementedException();
    }

    HRESULT IDataObject.EnumDAdvise(out IEnumSTATDATA ppenumAdvise)
    {
        throw new NotImplementedException();
    }

    HRESULT IObjectWithSite.SetSite(nint pUnkSite)
    {
        throw new NotImplementedException();
    }

    HRESULT IObjectWithSite.GetSite(in Guid riid, out nint ppvSite)
    {
        throw new NotImplementedException();
    }

    HRESULT IOleInPlaceActiveObject.TranslateAccelerator(nint lpmsg)
    {
        throw new NotImplementedException();
    }

    HRESULT IOleInPlaceActiveObject.OnFrameWindowActivate(BOOL fActivate)
    {
        throw new NotImplementedException();
    }

    HRESULT IOleInPlaceActiveObject.OnDocWindowActivate(BOOL fActivate)
    {
        throw new NotImplementedException();
    }

    HRESULT IOleInPlaceActiveObject.ResizeBorder(in RECT prcBorder, IOleInPlaceUIWindow pUIWindow, BOOL fFrameWindow)
    {
        throw new NotImplementedException();
    }

    HRESULT IOleInPlaceActiveObject.EnableModeless(BOOL fEnable)
    {
        throw new NotImplementedException();
    }

    HRESULT IOleInPlaceObject.InPlaceDeactivate()
    {
        throw new NotImplementedException();
    }

    HRESULT IOleInPlaceObject.UIDeactivate()
    {
        throw new NotImplementedException();
    }

    HRESULT IOleInPlaceObject.SetObjectRects(in RECT lprcPosRect, in RECT lprcClipRect)
    {
        throw new NotImplementedException();
    }

    HRESULT IOleInPlaceObject.ReactivateAndUndo()
    {
        throw new NotImplementedException();
    }

    HRESULT IQuickActivate.QuickActivate(in QACONTAINER pQaContainer, ref QACONTROL pQaControl)
    {
        throw new NotImplementedException();
    }

    HRESULT IQuickActivate.SetContentExtent(in SIZE pSizel)
    {
        throw new NotImplementedException();
    }

    HRESULT IQuickActivate.GetContentExtent(out SIZE pSizel)
    {
        throw new NotImplementedException();
    }

    HRESULT IServiceProvider.QueryService(in Guid guidService, in Guid riid, out nint ppvObject)
    {
        throw new NotImplementedException();
    }

    [GeneratedComClass]
    private sealed partial class Verbs(IReadOnlyList<OLEVERB> verbs) : IEnumOLEVERB
    {
        private int _index = -1;

        public HRESULT Clone(out IEnumOLEVERB enumerator)
        {
            enumerator = new Verbs(verbs);
            return Constants.S_OK;
        }

        public HRESULT Next(uint count, OLEVERB[] outVerbs, nint outFetched)
        {
            var max = Math.Max(0, Math.Min(outVerbs.Length - (_index + 1), count));
            var fetched = max;
            if (fetched > 0)
            {
                for (var i = _index + 1; i < fetched; i++)
                {
                    outVerbs[i] = verbs[i];
                    _index++;
                }
            }

            if (outFetched != 0)
            {
                Marshal.WriteInt32(outFetched, (int)fetched);
            }
            return (fetched == count) ? Constants.S_OK : Constants.S_FALSE;
        }

        public HRESULT Reset() => _index = -1;
        public HRESULT Skip(uint count)
        {
            var max = (uint)Math.Max(0, Math.Min(verbs.Count - (_index + 1), count));
            if (max > 0)
            {
                _index += (int)max;
            }
            return (max == count) ? Constants.S_OK : Constants.S_FALSE;
        }
    }

    [GeneratedComClass]
    private sealed partial class ConnectionsPoints(IReadOnlyList<KeyValuePair<Guid, IConnectionPoint>> connections) : IEnumConnectionPoints
    {
        private int _index = -1;

        public HRESULT Clone(out IEnumConnectionPoints enumerator)
        {
            enumerator = new ConnectionsPoints(connections);
            return Constants.S_OK;
        }

        public HRESULT Next(uint count, nint[] points, out uint fetched)
        {
            var max = (uint)Math.Max(0, Math.Min(connections.Count - (_index + 1), count));
            fetched = max;
            if (fetched > 0)
            {
                for (var i = _index + 1; i < fetched; i++)
                {
                    points[i] = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance<IConnectionPoint>(connections[i].Value);
                    _index++;
                }
            }
            return (fetched == count) ? Constants.S_OK : Constants.S_FALSE;
        }

        public HRESULT Reset() => _index = -1;
        public HRESULT Skip(uint count)
        {
            var max = (uint)Math.Max(0, Math.Min(connections.Count - (_index + 1), count));
            if (max > 0)
            {
                _index += (int)max;
            }
            return (max == count) ? Constants.S_OK : Constants.S_FALSE;
        }
    }
}
