namespace ActiveN;

[GeneratedComClass]
public abstract partial class BaseControl : BaseDispatch,
    IOleObject,
    IOleControl,
    IOleWindow,
    IObjectSafety,
    IDataObject,
    IObjectWithSite,
    IOleInPlaceActiveObject,
    IOleInPlaceObject,
    IOleInPlaceObjectWindowless,
    IQuickActivate,
    DirectN.IServiceProvider,
    IProvideClassInfo,
    IPersistStreamInit,
    IViewObject2,
    IViewObjectEx,
    IPointerInactive,
    ISupportErrorInfo,
    IConnectionPointContainer,
    ISpecifyPropertyPages,
    ICustomQueryInterface
{
    private readonly ConcurrentDictionary<Guid, IConnectionPoint> _connectionPoints = new();
    private readonly ComObject<IOleAdviseHolder> _adviseHolder;
    private readonly PropertyNotifySinkConnectionPoint _connectionPoint;
    private IComObject<IOleClientSite>? _clientSite;
    private IComObject<IOleInPlaceSite>? _inPlaceSite;
    private IComObject<IOleInPlaceSiteEx>? _inPlaceSiteEx;
    private IComObject<IOleInPlaceSiteWindowless>? _inPlaceSiteWindowless;
    private IComObject<IObjectWithSite>? _site;
    private IComObject<IAdviseSink>? _adviseSink;
    private IComObject<IAdviseSinkEx>? _adviseSinkEx;
    private bool _isDirty;
    private SIZE _extent;
    private int _freezeCount;
    private Window? _window;

    protected BaseControl()
    {
        TracingUtilities.Trace($"Created {GetType().FullName} ({GetType().GUID:B})");
        CurrentSafetyOptions = SupportedSafetyOptions;
        Functions.CreateOleAdviseHolder(out var obj).ThrowOnError();
        _adviseHolder = new ComObject<IOleAdviseHolder>(obj);
        _connectionPoint = new PropertyNotifySinkConnectionPoint();
        AddConnectionPoint(_connectionPoint);
    }

    protected abstract Window CreateWindow(HWND parentHandle, RECT rect);

    protected virtual POINTERINACTIVE PointerActivationPolicy => POINTERINACTIVE.POINTERINACTIVE_ACTIVATEONENTRY;
    protected virtual OLEMISC MiscStatus => OLEMISC.OLEMISC_RECOMPOSEONRESIZE | OLEMISC.OLEMISC_CANTLINKINSIDE | OLEMISC.OLEMISC_INSIDEOUT | OLEMISC.OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC.OLEMISC_SETCLIENTSITEFIRST;
    protected virtual VIEWSTATUS ViewStatus => VIEWSTATUS.VIEWSTATUS_OPAQUE | VIEWSTATUS.VIEWSTATUS_SOLIDBKGND;
    protected virtual uint SupportedSafetyOptions => Constants.INTERFACESAFE_FOR_UNTRUSTED_CALLER | Constants.INTERFACESAFE_FOR_UNTRUSTED_DATA;
    protected virtual uint CurrentSafetyOptions { get; set; }
    protected virtual Window? Window => _window;
    protected override HWND GetWindowHandle() => _window?.Handle ?? HWND.Null;
    protected virtual ControlState State { get; private set; }

    protected virtual void ChangeState(ControlState newState)
    {
        TracingUtilities.Trace($"Changing state from {State} to {newState}");
        if (newState == State)
            return;

        var oldState = State;
        State = newState;
        switch (newState)
        {
            case ControlState.InplaceActive:
                _inPlaceSite?.Object.OnInPlaceActivate().ThrowOnError();
                if (oldState == ControlState.UIActive)
                {
                    _inPlaceSite?.Object.OnUIDeactivate(false).ThrowOnError();
                }
                break;

            case ControlState.Active:
                // nothing to do
                break;

            case ControlState.UIActive:
                _inPlaceSite?.Object.OnUIActivate().ThrowOnError();
                break;

            case ControlState.Running:
                if (oldState == ControlState.UIActive)
                {
                    _inPlaceSite?.Object.OnUIDeactivate(false).ThrowOnError();
                    _inPlaceSite?.Object.OnInPlaceDeactivate().ThrowOnError();
                }
                else if (oldState == ControlState.InplaceActive)
                {
                    _inPlaceSite?.Object.OnInPlaceDeactivate().ThrowOnError();
                }
                break;
        }
    }

    protected virtual HRESULT UIActivate(HWND hwndParent, RECT pos)
    {
        ChangeState(ControlState.UIActive);
        return Constants.S_OK;
    }

    protected virtual HRESULT DiscardUndoState(HWND hwndParent, RECT pos) => NotImplemented();

    protected virtual HRESULT InplaceActivate(HWND hwndParent, RECT pos)
    {
        var window = EnsureWindow(hwndParent, pos);
        TracingUtilities.Trace($"window: {window}");

        if (_inPlaceSite != null)
        {
            TracingUtilities.Trace($"calling OnInPlaceActivate");
            _inPlaceSite.Object.OnInPlaceActivate().ThrowOnError();

            // in the case of hosted in VS, if .NET < 10, this fails if the control has no parent (Winforms Control)
            //var info = new OLEINPLACEFRAMEINFO { cb = (uint)sizeof(OLEINPLACEFRAMEINFO) };
            //_inPlaceSite.Object.GetWindowContext(out var frameObj, out var uiWindowObj, out var rcPos, out var clip, ref info).ThrowOnError();
            //using var frame = frameObj != null ? new ComObject<IOleInPlaceFrame>(frameObj) : null;
            //using var uiWindow = uiWindowObj != null ? new ComObject<IOleInPlaceUIWindow>(uiWindowObj) : null;
            //TracingUtilities.Trace($"frameObj: {frame} uiWindo: {uiWindow} rcPos: {rcPos} clip: {clip} info hwnd: {info.hwndFrame}");
        }

        //pActiveSite.ShowObject().ThrowOnError();
        ChangeState(ControlState.InplaceActive);
        return Constants.S_OK;
    }

    protected virtual HRESULT Open(HWND hwndParent, RECT pos)
    {
        TracingUtilities.Trace($"hwndParent: {hwndParent} pos: {pos}");
        return InplaceActivate(hwndParent, pos);
    }

    protected virtual HRESULT Hide(HWND hwndParent, RECT pos)
    {
        TracingUtilities.Trace($"hwndParent: {hwndParent} pos: {pos}");
        ChangeState(ControlState.Running);
        return Constants.S_OK;
    }

    protected virtual WINDOW_STYLE GetDefaultWindowStyle(HWND parentHandle)
    {
        var style = WINDOW_STYLE.WS_VISIBLE | WINDOW_STYLE.WS_CLIPCHILDREN | WINDOW_STYLE.WS_CLIPSIBLINGS;
        if (parentHandle == HWND.Null)
        {
            style |= WINDOW_STYLE.WS_POPUP;
        }
        else
        {
            style |= WINDOW_STYLE.WS_CHILD;
        }
        return style;
    }

    protected virtual Window? EnsureWindow(HWND parentHandle, RECT rect) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"parentHandle: {parentHandle} rect: {rect} window: {_window} null: {_window == null}");
        if (_window == null)
        {
            var window = CreateWindow(parentHandle, rect);
            window.Show();
            _window = window;
        }
        return _window;
    });

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

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _adviseSink?.Dispose();
            _adviseSinkEx = null;
            _adviseSinkEx?.Dispose();
            _adviseSinkEx = null;
            _adviseHolder?.Dispose();
            _clientSite?.Dispose();
            _clientSite = null;
            _inPlaceSite?.Dispose();
            _inPlaceSite = null;
            _inPlaceSiteEx?.Dispose();
            _inPlaceSiteEx = null;
            _inPlaceSiteWindowless?.Dispose();
            _inPlaceSiteWindowless = null;
            _site?.Dispose();
            _site = null;
            _window?.Dispose();
            _window = null;
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

    protected static HRESULT UnregisterType(ComRegistrationContext context) => TracingUtilities.WrapErrors(() =>
    {
        ArgumentNullException.ThrowIfNull(context);
        TracingUtilities.Trace($"Unregister type {typeof(BaseControl).FullName}...");
        return Constants.S_OK;
    });

    protected virtual HRESULT NotImplemented([CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        TracingUtilities.Trace($"E_NOIMPL", methodName, filePath);
        return Constants.E_NOTIMPL;
    }

    HRESULT IOleObject.SetClientSite(IOleClientSite pClientSite) => TracingUtilities.WrapErrors(() =>
    {
        _clientSite?.Dispose();
        _clientSite = pClientSite != null ? new ComObject<IOleClientSite>(pClientSite) : null;

        _inPlaceSite?.Dispose();
        _inPlaceSite = _clientSite.As<IOleInPlaceSite>();

        _inPlaceSiteEx?.Dispose();
        _inPlaceSiteEx = _clientSite.As<IOleInPlaceSiteEx>();

        _inPlaceSiteWindowless?.Dispose();
        _inPlaceSiteWindowless = _clientSite.As<IOleInPlaceSiteWindowless>();

        TracingUtilities.Trace($"ClientSite: {_clientSite}");
        TracingUtilities.Trace($"InPlaceSite: {_inPlaceSite}");
        TracingUtilities.Trace($"InPlaceSiteEx: {_inPlaceSiteEx}");
        TracingUtilities.Trace($"InPlaceSiteWindowless: {_inPlaceSiteWindowless}");
        return Constants.S_OK;
    });

    HRESULT IOleObject.GetClientSite(out IOleClientSite ppClientSite)
    {
        IOleClientSite? clientSite = null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            clientSite = _clientSite?.Object!;
            TracingUtilities.Trace($"Site: {clientSite}");
            return Constants.S_OK;
        });
        ppClientSite = clientSite!;
        return hr;
    }

    HRESULT IOleObject.SetHostNames(PWSTR szContainerApp, PWSTR szContainerObj) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"szContainerApp: '{szContainerApp}' szContainerObj: '{szContainerObj}'");
        return Constants.S_OK;
    });

    HRESULT IOleObject.Close(uint dwSaveOption) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"dwSaveOption: {(OLECLOSE)dwSaveOption}");
        return Constants.S_OK;
    });

    HRESULT IOleObject.InitFromData(IDataObject pDataObject, BOOL fCreation, uint dwReserved) => NotImplemented();
    HRESULT IOleObject.GetClipboardData(uint dwReserved, out IDataObject ppDataObject) { ppDataObject = null!; return NotImplemented(); }
    HRESULT IOleObject.SetMoniker(uint dwWhichMoniker, IMoniker pmk) => NotImplemented();
    HRESULT IOleObject.GetMoniker(uint dwAssign, uint dwWhichMoniker, out IMoniker ppmk)
    {
        IMoniker? mk = null;
        var hr = NotImplemented();
        ppmk = mk!;
        return hr;
    }

    unsafe HRESULT IOleObject.DoVerb(int iVerb, nint lpmsg, IOleClientSite pActiveSite, int lindex, HWND hwndParent, nint lprcPosRect) => TracingUtilities.WrapErrors(() =>
    {
        var verb = (OLEIVERB)iVerb;

        // default if no rect provided
        var pos = RECT.Sized(0, 0, 100, 100);
        if (lprcPosRect != 0)
        {
            pos = *(RECT*)lprcPosRect;
            TracingUtilities.Trace($"rcPosRect: {pos}");
        }

        TracingUtilities.Trace($"iVerb: {verb} lpmsg: {lpmsg} pActiveSite: {pActiveSite} lindex: {lindex} hwndParent: {hwndParent} lprcPosRect: {pos}");
        var hr = verb switch
        {
            OLEIVERB.OLEIVERB_INPLACEACTIVATE => InplaceActivate(hwndParent, pos),
            OLEIVERB.OLEIVERB_SHOW or OLEIVERB.OLEIVERB_PRIMARY or OLEIVERB.OLEIVERB_OPEN => Open(hwndParent, pos),
            OLEIVERB.OLEIVERB_HIDE => Hide(hwndParent, pos),
            OLEIVERB.OLEIVERB_UIACTIVATE => UIActivate(hwndParent, pos),
            OLEIVERB.OLEIVERB_DISCARDUNDOSTATE => DiscardUndoState(hwndParent, pos),
            _ => Constants.E_NOTIMPL,
        };
        TracingUtilities.Trace($"hr: {hr}");
        return hr;
    });

    HRESULT IOleObject.EnumVerbs(out IEnumOLEVERB ppEnumOleVerb)
    {
        TracingUtilities.Trace();
        ppEnumOleVerb = new EnumVerbs([]);
        return Constants.S_OK;
    }

    HRESULT IOleObject.Update() => NotImplemented();
    HRESULT IOleObject.IsUpToDate() => NotImplemented();
    HRESULT IOleObject.GetUserClassID(out Guid pClsid)
    {
        var clsid = Guid.Empty;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            clsid = GetType().GUID;
            return Constants.S_OK;
        });
        pClsid = clsid;
        return hr;
    }

    HRESULT IOleObject.GetUserType(uint dwFormOfType, out PWSTR pszUserType)
    {
        var userType = PWSTR.Null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var hr = Functions.OleRegGetUserType(GetType().GUID, dwFormOfType, out var userType);
            TracingUtilities.Trace($"dwFormOfType: {dwFormOfType} pszUserType: '{userType}' hr: {hr}");
            return hr;
        });
        pszUserType = userType;
        return hr;
    }

    HRESULT IOleObject.SetExtent(DVASPECT dwDrawAspect, in SIZE psizel)
    {
        var size = psizel;
        return TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect} psizel: {size}");
            if (dwDrawAspect != DVASPECT.DVASPECT_CONTENT)
                return Constants.DV_E_DVASPECT;

            _extent = size;
            return Constants.S_OK;
        });
    }

    HRESULT IOleObject.GetExtent(DVASPECT dwDrawAspect, out SIZE psizel)
    {
        var size = new SIZE();
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect}");
            if (dwDrawAspect != DVASPECT.DVASPECT_CONTENT)
                return Constants.DV_E_DVASPECT;

            size = _extent;
            return Constants.S_OK;
        });
        psizel = size;
        return hr;
    }

    HRESULT IOleObject.Advise(IAdviseSink pAdvSink, out uint pdwConnection)
    {
        var connection = 0u;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var hr = _adviseHolder.Object.Advise(pAdvSink, out connection);
            TracingUtilities.Trace($"dwConnection {connection}: {hr}");
            return hr;
        });
        pdwConnection = connection;
        return hr;
    }

    HRESULT IOleObject.Unadvise(uint dwConnection) => TracingUtilities.WrapErrors(() =>
    {
        var hr = _adviseHolder.Object.Unadvise(dwConnection);
        TracingUtilities.Trace($"dwConnection {dwConnection}: {hr}");
        return hr;
    });

    HRESULT IOleObject.EnumAdvise(out IEnumSTATDATA ppenumAdvise)
    {
        TracingUtilities.Trace();
        IEnumSTATDATA? enumAdvise = null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var hr = _adviseHolder.Object.EnumAdvise(out enumAdvise);
            TracingUtilities.Trace($"ppenumAdvise {enumAdvise}: {hr}");
            return hr;
        });
        ppenumAdvise = enumAdvise!;
        return hr;
    }

    HRESULT IOleObject.SetColorScheme(in LOGPALETTE pLogpal) => NotImplemented();
    HRESULT IOleObject.GetMiscStatus(DVASPECT dwAspect, out OLEMISC pdwStatus)
    {
        var status = new OLEMISC();
        var hr = TracingUtilities.WrapErrors(() =>
        {
            status = MiscStatus;
            TracingUtilities.Trace($"dwAspect: {dwAspect} pdwStatus: {status}");
            return Constants.S_OK;
        });
        pdwStatus = status;
        return hr;
    }

    HRESULT IProvideClassInfo.GetClassInfo(out ITypeInfo ppTI)
    {
        ITypeInfo? pti = null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var ti = EnsureTypeInfo();
            TracingUtilities.Trace($"ppTI {ti}");
            pti = ti?.Object;
            return pti != null ? Constants.S_OK : Constants.E_UNEXPECTED;
        });
        ppTI = pti!;
        return hr;
    }

    HRESULT IPersistStreamInit.IsDirty() => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"IsDirty: {_isDirty}");
        return _isDirty ? Constants.S_OK : Constants.S_FALSE;
    });

    HRESULT IPersistStreamInit.Load(IStream pStm) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"pStm {pStm}");
        _isDirty = false;
        return Constants.S_OK;
    });

    HRESULT IPersistStreamInit.GetSizeMax(out ulong pCbSize) { pCbSize = 0; return NotImplemented(); }
    HRESULT IPersistStreamInit.Save(IStream pStm, BOOL fClearDirty) => TracingUtilities.WrapErrors(() =>
    {
        _isDirty = !fClearDirty;
        TracingUtilities.Trace($"pStm {pStm} fClearDirty: {fClearDirty}");
        return Constants.S_OK;
    });

    HRESULT IPersistStreamInit.InitNew() => TracingUtilities.WrapErrors(() =>
    {
        _isDirty = true;
        TracingUtilities.Trace();
        return Constants.S_OK;
    });

    HRESULT IPersist.GetClassID(out Guid pClassID)
    {
        var classID = Guid.Empty;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            classID = GetType().GUID;
            TracingUtilities.Trace();
            return Constants.S_OK;
        });
        pClassID = classID;
        return hr;
    }

    HRESULT IViewObject2.GetExtent(DVASPECT dwDrawAspect, int lindex, nint ptd, out SIZE lpsizel)
    {
        var size = new SIZE();
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect} lindex: {lindex} ptd: {ptd}");
            size = _extent;
            return Constants.S_OK;
        });
        lpsizel = size;
        return hr;
    }

    HRESULT IViewObject.Draw(
        DVASPECT dwDrawAspect,
        int lindex,
        nint pvAspect,
        nint ptd,
        HDC hdcTargetDev,
        HDC hdcDraw,
        nint lprcBounds,
        nint lprcWBounds,
        nint pfnContinue,
        nuint dwContinue) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect} lindex: {lindex} pvAspect: {pvAspect} ptd: {ptd} hdcTargetDev: {hdcTargetDev} hdcDraw: {hdcDraw} lprcBounds: {lprcBounds} lprcWBounds: {lprcWBounds} pfnContinue: {pfnContinue} dwContinue: {dwContinue}");
        if (dwDrawAspect == DVASPECT.DVASPECT_DOCPRINT)
            return Constants.DV_E_DVASPECT;

        return Constants.S_OK;
    });

    HRESULT IViewObject.GetColorSet(DVASPECT dwDrawAspect, int lindex, nint pvAspect, nint ptd, HDC hicTargetDev, out nint ppColorSet) { ppColorSet = 0; return NotImplemented(); }
    HRESULT IViewObject.Freeze(DVASPECT dwDrawAspect, int lindex, nint pvAspect, out uint pdwFreeze) { pdwFreeze = 0; return NotImplemented(); }
    HRESULT IViewObject.Unfreeze(uint dwFreeze) => NotImplemented();
    HRESULT IViewObject.SetAdvise(DVASPECT aspects, uint advf, IAdviseSink pAdvSink) => TracingUtilities.WrapErrors(() =>
    {
        _adviseSink?.Dispose();
        _adviseSink = pAdvSink != null ? new ComObject<IAdviseSink>(pAdvSink) : null;
        _adviseSinkEx?.Dispose();
        _adviseSinkEx = pAdvSink != null ? new ComObject<IAdviseSinkEx>(pAdvSink) : null;
        TracingUtilities.Trace($"Sink: {_adviseSink}");
        TracingUtilities.Trace($"SinkEx: {_adviseSinkEx}");
        return Constants.S_OK;
    });

    HRESULT IViewObject.GetAdvise(nint pAspects, nint pAdvf, out IAdviseSink ppAdvSink)
    {
        IAdviseSink? sink = null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            sink = _adviseSink?.Object!;
            TracingUtilities.Trace($"Site: {sink}");
            return Constants.S_OK;
        });
        ppAdvSink = sink!;
        return hr;
    }

    HRESULT IViewObjectEx.GetRect(uint dwAspect, out RECTL pRect) { pRect = new(); return NotImplemented(); }
    HRESULT IViewObjectEx.GetViewStatus(out uint pdwStatus)
    {
        var status = 0u;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            status = (uint)ViewStatus;
            TracingUtilities.Trace($"pdwStatus: {ViewStatus}");
            return Constants.S_OK;
        });
        pdwStatus = status;
        return hr;
    }

    HRESULT IViewObjectEx.QueryHitPoint(uint dwAspect, in RECT pRectBounds, POINT ptlLoc, int lCloseHint, out uint pHitResult)
    {
        var hitResult = 0u;
        var rcBounds = pRectBounds;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var aspect = (DVASPECT)dwAspect;
            //TracingUtilities.Trace($"dwAspect: {aspect} pRectBounds: {pRectBounds} ptlLoc: {ptlLoc} lCloseHint: {lCloseHint}");

            if (aspect == DVASPECT.DVASPECT_CONTENT)
            {
                var result = Functions.PtInRect(rcBounds, ptlLoc) ? TXTHITRESULT.TXTHITRESULT_HIT : TXTHITRESULT.TXTHITRESULT_NOHIT;
                //TracingUtilities.Trace($"pHitResult: {result}");
                hitResult = (uint)result;
                return Constants.S_OK;
            }

            hitResult = 0;
            TracingUtilities.Trace($"E_FAIL");
            return Constants.E_FAIL;
        });
        pHitResult = hitResult;
        return hr;
    }

    HRESULT IViewObjectEx.QueryHitRect(uint dwAspect, in RECT pRectBounds, in RECT pRectLoc, int lCloseHint, out uint pHitResult) { pHitResult = 0; return NotImplemented(); }
    unsafe HRESULT IViewObjectEx.GetNaturalExtent(DVASPECT dwAspect, int lindex, nint ptd, HDC hicTargetDev, nint pExtentInfo, nint pSizel) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"dwAspect: {dwAspect} lindex: {lindex} ptd: {ptd} hicTargetDev: {hicTargetDev} pExtentInfo: {pExtentInfo} pSizel: {pSizel}");
        if (pSizel == 0 || pExtentInfo == 0)
            return Constants.E_POINTER;

        if (dwAspect != DVASPECT.DVASPECT_CONTENT)
            return Constants.DV_E_DVASPECT;

        var size = new SIZE { cx = 1000, cy = 1000 };
        *(SIZE*)pSizel = size;
        TracingUtilities.Trace($"sizel: {size}");

        var extentInfo = *(DVEXTENTINFO*)pExtentInfo;
        var mode = (DVEXTENTMODE)extentInfo.dwExtentMode;
        TracingUtilities.Trace($"extentInfo sizelProposed: {extentInfo.sizelProposed} mode: {mode}");

        var p = (SIZE*)pSizel;
        p->cx = -1;
        p->cy = -1;
        return Constants.S_OK;
    });

    HRESULT IOleControl.GetControlInfo(ref CONTROLINFO pCI) { pCI = new(); return NotImplemented(); }
    HRESULT IOleControl.OnMnemonic(in MSG pMsg)
    {
        var msg = pMsg;
        return TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"pMSg: {MessageDecoder.Decode(msg)}");
            return Constants.S_OK;
        });
    }

    HRESULT IOleControl.OnAmbientPropertyChange(int dispID) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"dispID: {dispID}");
        return Constants.S_OK;
    });

    HRESULT IOleControl.FreezeEvents(BOOL bFreeze) => TracingUtilities.WrapErrors(() =>
    {
        if (bFreeze)
        {
            _freezeCount++;
        }
        else if (_freezeCount > 0)
        {
            _freezeCount--;
        }

        TracingUtilities.Trace($"bFreeze: {bFreeze} count:{_freezeCount}");
        return Constants.S_OK;
    });

    HRESULT IPointerInactive.GetActivationPolicy(out POINTERINACTIVE pdwPolicy)
    {
        var policy = (POINTERINACTIVE)0;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            policy = PointerActivationPolicy;
            TracingUtilities.Trace($"pdwPolicy: {policy}");
            return Constants.S_OK;
        });
        pdwPolicy = policy;
        return hr;
    }

    HRESULT IPointerInactive.OnInactiveMouseMove(in RECT pRectBounds, int x, int y, uint grfKeyState)
    {
        var rcBounds = pRectBounds;
        return TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"pRectBounds: {rcBounds} x: {x} y: {y} grfKeyState: {grfKeyState}");
            return Constants.S_OK;
        });
    }

    HRESULT IPointerInactive.OnInactiveSetCursor(in RECT pRectBounds, int x, int y, uint dwMouseMsg, BOOL fSetAlways)
    {
        var rcBounds = pRectBounds;
        return TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"pRectBounds: {rcBounds} x: {x} y: {y} dwMouseMsg: {dwMouseMsg} fSetAlways: {fSetAlways}");
            return Constants.S_OK;
        });
    }

    HRESULT ISpecifyPropertyPages.GetPages(out CAUUID pPages)
    {
        var pages = new CAUUID();
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"cElems: {pages.cElems}");
            return Constants.S_OK;
        });
        pPages = pages;
        return hr;
    }

    HRESULT IConnectionPointContainer.EnumConnectionPoints(out IEnumConnectionPoints ppEnum)
    {
        TracingUtilities.Trace();
        ppEnum = new EnumConnectionPoints([.. _connectionPoints]);
        return Constants.S_OK;
    }

    HRESULT IConnectionPointContainer.FindConnectionPoint(in Guid riid, out IConnectionPoint ppCP)
    {
        TracingUtilities.Trace($"iid:{riid.GetName()}");
        _connectionPoints.TryGetValue(riid, out var cp);
        ppCP = cp!;
        return cp == null ? Constants.CONNECT_E_NOCONNECTION : Constants.S_OK;
    }

    HRESULT IOleWindow.ContextSensitiveHelp(BOOL fEnterMode) => NotImplemented();
    HRESULT IOleWindow.GetWindow(out HWND phwnd)
    {
        var hwnd = HWND.Null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            hwnd = GetWindowHandle();
            TracingUtilities.Trace($"window:{_window?.GetType().Name} phwnd: {hwnd}");
            return Constants.S_OK;
        });
        phwnd = hwnd;
        return hr;
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

    HRESULT IObjectWithSite.SetSite(nint pUnkSite) => TracingUtilities.WrapErrors(() =>
    {
        _site?.Dispose();
        _site = DirectN.Extensions.Com.ComObject.FromPointer<IObjectWithSite>(pUnkSite);
        TracingUtilities.Trace($"Site: {pUnkSite}");
        return Constants.S_OK;
    });

    HRESULT IObjectWithSite.GetSite(in Guid riid, out nint ppvSite)
    {
        var site = nint.Zero;
        var iid = riid;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"riid: {iid.GetName()}");
            if (_site == null)
                return Constants.E_NOINTERFACE;

            site = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(_site);
            return site != 0 ? Constants.S_OK : Constants.E_NOINTERFACE;
        });
        ppvSite = site;
        return hr;
    }

    HRESULT IOleInPlaceActiveObject.TranslateAccelerator(nint lpmsg) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.OnFrameWindowActivate(BOOL fActivate) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.OnDocWindowActivate(BOOL fActivate) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.ResizeBorder(in RECT prcBorder, IOleInPlaceUIWindow pUIWindow, BOOL fFrameWindow) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.EnableModeless(BOOL fEnable) => NotImplemented();
    HRESULT IOleInPlaceObject.InPlaceDeactivate() => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"State: {State}");
        ChangeState(ControlState.Running);
        return Constants.S_OK;
    });

    HRESULT IOleInPlaceObject.UIDeactivate() => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"State: {State}");
        if (State != ControlState.Running)
        {
            ChangeState(ControlState.InplaceActive);
        }
        return Constants.S_OK;
    });

    HRESULT IOleInPlaceObject.SetObjectRects(in RECT lprcPosRect, in RECT lprcClipRect)
    {
        var pos = lprcPosRect;
        var clip = lprcClipRect;
        return TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"lprcPosRect: {pos} lprcClipRect: {clip} window: {_window}");
            if (_window != null)
            {
                var tempRgn = new HRGN();
                if (Functions.IntersectRect(out var rc, pos, clip) && !Functions.EqualRect(rc, pos))
                {
                    Functions.OffsetRect(ref rc, -pos.left, -pos.top);
                    tempRgn = Functions.CreateRectRgnIndirect(rc);
                }

                _ = Functions.SetWindowRgn(_window.Handle, tempRgn, true);
                _window.SetWindowPos(HWND.Null, pos.left, pos.top, pos.Width, pos.Height, SET_WINDOW_POS_FLAGS.SWP_NOZORDER | SET_WINDOW_POS_FLAGS.SWP_NOACTIVATE);
            }
            return Constants.S_OK;
        });
    }

    HRESULT IOleInPlaceObject.ReactivateAndUndo() => NotImplemented();
    unsafe HRESULT IQuickActivate.QuickActivate(in QACONTAINER pQaContainer, ref QACONTROL pQaControl)
    {
        var container = pQaContainer;
        var control = pQaControl;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"pQaContainer: {container} pQacontrol size: {control.cbSize}");
            TracingUtilities.Trace($"container flags: {(QACONTAINERFLAGS)container.dwAmbientFlags} dwAppearance: {container.dwAppearance} lcid: {container.lcid} pClientSite: {container.pClientSite} pAdviseSink: {container.pAdviseSink} pPropertyNotifySink: {container.pPropertyNotifySink} pUnkEventSink: {container.pUnkEventSink} colorFore: {container.colorFore:X8} colorBack: {container.colorBack:X8} pFont: {container.pFont} pUndoMgr: {container.pUndoMgr} pSrvProvider: {container.pServiceProvider}");

            if (container.pClientSite != 0)
            {
                _clientSite?.Dispose();
                _clientSite = DirectN.Extensions.Com.ComObject.FromPointer<IOleClientSite>(container.pClientSite);

                _inPlaceSite?.Dispose();
                _inPlaceSite = _clientSite.As<IOleInPlaceSite>();

                _inPlaceSiteEx?.Dispose();
                _inPlaceSiteEx = _clientSite.As<IOleInPlaceSiteEx>();

                _inPlaceSiteWindowless?.Dispose();
                _inPlaceSiteWindowless = _clientSite.As<IOleInPlaceSiteWindowless>();
            }

            TracingUtilities.Trace($"ClientSite: {_clientSite}");
            TracingUtilities.Trace($"InPlaceSite: {_inPlaceSite}");
            TracingUtilities.Trace($"InPlaceSiteEx: {_inPlaceSiteEx}");
            TracingUtilities.Trace($"InPlaceSiteWindowless: {_inPlaceSiteWindowless}");

            if (container.pAdviseSink != 0)
            {
                _adviseSink?.Dispose();
                _adviseSink = DirectN.Extensions.Com.ComObject.FromPointer<IAdviseSink>(container.pAdviseSink);
                _adviseSinkEx?.Dispose();
                _adviseSinkEx = DirectN.Extensions.Com.ComObject.FromPointer<IAdviseSinkEx>(container.pAdviseSink);
            }

            TracingUtilities.Trace($"Sink: {_adviseSink}");
            TracingUtilities.Trace($"SinkEx: {_adviseSinkEx}");

            control.dwMiscStatus = (uint)MiscStatus;
            control.dwViewStatus = (uint)ViewStatus;
            control.dwPointerActivationPolicy = (uint)PointerActivationPolicy;
            TracingUtilities.Trace($"control dwMiscStatus: {MiscStatus} dwViewStatus: {ViewStatus} dwPointerActivationPolicy: {PointerActivationPolicy}");
            return Constants.S_OK;
        });
        pQaControl = control;
        return hr;
    }

    HRESULT IQuickActivate.SetContentExtent(in SIZE pSizel)
    {
        var size = pSizel;
        return TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"pSizel: {size}");
            _extent = size;
            return Constants.S_OK;
        });
    }

    HRESULT IQuickActivate.GetContentExtent(out SIZE pSizel)
    {
        var size = new SIZE();
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"pSizel: {_extent}");
            size = _extent;
            return Constants.S_OK;
        });
        pSizel = size;
        return hr;
    }

    HRESULT DirectN.IServiceProvider.QueryService(in Guid guidService, in Guid riid, out nint ppvObject)
    {
        var obj = nint.Zero;
        var siid = guidService;
        var iid = riid;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"guidService: {siid.GetName()} riid: {iid.GetName()}");
            return Constants.E_NOINTERFACE;
        });
        ppvObject = obj;
        return hr;
    }

    HRESULT IOleInPlaceObjectWindowless.GetDropTarget(out IDropTarget ppDropTarget) { ppDropTarget = null!; return NotImplemented(); }
    HRESULT IOleInPlaceObjectWindowless.OnWindowMessage(uint msg, WPARAM wParam, LPARAM lParam, out LRESULT plResult)
    {
        var result = LRESULT.Null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"msg: {msg} wParam: {wParam} lParam: {lParam}");
            return Constants.S_FALSE;
        });
        plResult = result;
        return hr;
    }

    HRESULT ISupportErrorInfo.InterfaceSupportsErrorInfo(in Guid riid)
    {
        var iid = riid;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"iid: {iid.GetName()}");
            return Constants.S_OK;
        });
        return hr;
    }

    HRESULT IObjectSafety.GetInterfaceSafetyOptions(in Guid riid, out uint pdwSupportedOptions, out uint pdwEnabledOptions)
    {
        var iid = riid;
        var supportedOptions = 0u;
        var enabledOptions = 0u;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"iid: {iid.GetName()}");
            supportedOptions = SupportedSafetyOptions;
            enabledOptions = CurrentSafetyOptions;
            return Constants.S_OK;
        });
        pdwSupportedOptions = supportedOptions;
        pdwEnabledOptions = enabledOptions;
        return hr;
    }

    HRESULT IObjectSafety.SetInterfaceSafetyOptions(in Guid riid, uint dwOptionSetMask, uint dwEnabledOptions)
    {
        var iid = riid;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"iid: {iid.GetName()} dwOptionSetMask:{dwOptionSetMask} dwEnabledOptions:{dwEnabledOptions}");
            CurrentSafetyOptions = dwEnabledOptions; // we currently do no check on all this
            return Constants.S_OK;
        });
        return hr;
    }
}
