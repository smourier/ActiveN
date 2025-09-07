namespace ActiveN;

[GeneratedComClass]
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.Interfaces)]
// warning: * any* interface added here must also be added to IAggregable.AggregatableInterfaces
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
    IProvideClassInfo2,
    IPersistStreamInit,
    IViewObject2,
    IViewObjectEx,
    IPointerInactive,
    IPerPropertyBrowsing,
    ISupportErrorInfo,
    IConnectionPointContainer,
    ISpecifyPropertyPages,
    ICategorizeProperties,
    IRunnableObject,
    IAggregable
{
    private readonly ConcurrentDictionary<Guid, IConnectionPoint> _connectionPoints = new();
    private readonly ConcurrentBag<Guid> _aggregableInterfacesIids;
    private readonly PropertyNotifySinkConnectionPoint _connectionPoint;
    private IComObject<IOleAdviseHolder>? _adviseHolder;
    private IComObject<IDataAdviseHolder>? _dataAdviseHolder;
    private IComObject<IOleClientSite>? _clientSite;
    private IComObject<IOleInPlaceSite>? _inPlaceSite;
    private IComObject<IOleInPlaceSiteEx>? _inPlaceSiteEx;
    private IComObject<IOleInPlaceSiteWindowless>? _inPlaceSiteWindowless;
    private IComObject<IObjectWithSite>? _site;
    private IComObject<IAdviseSink>? _adviseSink;
    private Window? _window;
    private bool _changingExtent;

    public event EventHandler<ValueEventArgs<DISPID>>? AmbientPropertyChanged;

    protected BaseControl()
    {
        TracingUtilities.Trace($"Created {GetType().FullName} ({GetType().GUID:B})");
        CurrentSafetyOptions = SupportedSafetyOptions;
        Functions.CreateOleAdviseHolder(out var obj).ThrowOnError();
        HiMetricExtent = GetOriginalExtent();
        _adviseHolder = new ComObject<IOleAdviseHolder>(obj);
        Functions.CreateDataAdviseHolder(out var dah).ThrowOnError();
        _dataAdviseHolder = new ComObject<IDataAdviseHolder>(dah);
        _connectionPoint = new PropertyNotifySinkConnectionPoint();
        AddConnectionPoint(_connectionPoint);

        AggregableInterfaces = ClassFactory.GetComAggregatedInterfaces(GetType());
        _aggregableInterfacesIids = [.. AggregableInterfaces.Select(i => i.GUID)];
    }

    protected abstract Window CreateWindow(HWND parentHandle, RECT rect);
    protected override HWND GetWindowHandle() => _window?.Handle ?? HWND.Null;
    protected IComObject<IOleAdviseHolder>? AdviseHolder => _adviseHolder;
    protected IComObject<IDataAdviseHolder>? DataAdviseHolder => _dataAdviseHolder;
    protected IComObject<IOleClientSite>? ClientSite => _clientSite;
    protected IComObject<IOleInPlaceSite>? InPlaceSite => _inPlaceSite;
    protected IComObject<IOleInPlaceSiteEx>? InPlaceSiteEx => _inPlaceSiteEx;
    protected IComObject<IOleInPlaceSiteWindowless>? InPlaceSiteWindowless => _inPlaceSiteWindowless;
    protected IComObject<IObjectWithSite>? Site => _site;
    protected IComObject<IAdviseSink>? AdviseSink => _adviseSink;

    protected virtual POINTERINACTIVE PointerActivationPolicy => POINTERINACTIVE.POINTERINACTIVE_ACTIVATEONENTRY;
    protected virtual OLEMISC MiscStatus =>
        OLEMISC.OLEMISC_RECOMPOSEONRESIZE |
        OLEMISC.OLEMISC_CANTLINKINSIDE |
        OLEMISC.OLEMISC_INSIDEOUT |
        OLEMISC.OLEMISC_ACTIVATEWHENVISIBLE |
        OLEMISC.OLEMISC_SETCLIENTSITEFIRST |
        //OLEMISC.OLEMISC_IGNOREACTIVATEWHENVISIBLE |
        OLEMISC.OLEMISC_RENDERINGISDEVICEINDEPENDENT;

    protected virtual VIEWSTATUS ViewStatus => VIEWSTATUS.VIEWSTATUS_OPAQUE | VIEWSTATUS.VIEWSTATUS_SOLIDBKGND;
    protected virtual uint SupportedSafetyOptions => Constants.INTERFACESAFE_FOR_UNTRUSTED_CALLER | Constants.INTERFACESAFE_FOR_UNTRUSTED_DATA;
    protected virtual uint CurrentSafetyOptions { get; set; }
    protected virtual Window? Window => _window;
    protected virtual SIZE? HiMetricNaturalExtent => null;
    protected virtual ControlState State { get; private set; }
    protected virtual bool IsDirty { get; set; }
    protected virtual int FreezeCount { get; set; }
    protected virtual SIZE HiMetricExtent { get; set; }
    protected virtual bool SupportsAggregation => true;
    protected virtual nint AggregationWrapper { get; set; }
    protected virtual CTRLINFO KeyboardBehavior { get; set; } // CTRLINFO.CTRLINFO_EATS_RETURN | CTRLINFO.CTRLINFO_EATS_ESCAPE; call OnControlInfoChanged when changed
    protected virtual IReadOnlyList<OleVerb> Verbs { get; set; } = [];
    protected virtual AcceleratorTable? KeyboardAccelerators { get; set; } // call OnControlInfoChanged when changed
    protected virtual IReadOnlyList<Type> AggregableInterfaces { get; }

    protected bool InUserMode => GetAmbientProperty(DISPID.DISPID_AMBIENT_USERMODE, false);
    protected bool InDesignMode => !InUserMode;

    bool IAggregable.SupportsAggregation => SupportsAggregation;
    IReadOnlyList<Type> IAggregable.AggregableInterfaces => AggregableInterfaces;
    nint IAggregable.Wrapper { get => AggregationWrapper; set => AggregationWrapper = value; }

    protected virtual uint GetDpi()
    {
        var handle = GetWindowHandle();
        if (handle == 0)
        {
            _inPlaceSite?.Object.GetWindow(out handle);
            if (handle == 0)
            {
                handle = SystemUtilities.CurrentProcess.MainWindowHandle;
                if (handle == 0)
                {
                    handle = Window.FromProcess(SystemUtilities.CurrentProcess).FirstOrDefault(w => w.IsVisible && w.IsTopLevel)?.Handle ?? 0;
                }
            }
        }

        var dpi = handle != 0 ? DpiUtilities.GetDpiForWindow(handle).width : 96;
        TracingUtilities.Trace($"found handle: 0x{handle:X} dpi: {dpi}");
        return dpi;
    }

    protected virtual SIZE GetOriginalExtent()
    {
        var dpi = GetDpi();
        var size = 192.PixelToHiMetric(dpi);
        TracingUtilities.Trace($"dpi: {dpi} size: {size}");
        return new SIZE(size, size);
    }

    protected override void OnStockPropertyChanged(int dispId)
    {
        TracingUtilities.Trace($"dispId: {dispId} FreezeCount: {FreezeCount}");
        base.OnStockPropertyChanged(dispId);
        if (FreezeCount == 0 && !FireOnRequestEdit(dispId))
        {
            TracingUtilities.Trace($"OnRequestEdit failed for dispId: {dispId}");
            return;
        }

        IsDirty = true;

        if (FreezeCount == 0)
        {
            FireOnChanged(dispId);
        }

        SendOnViewChange();
        SendOnDataChanged();
    }

    protected virtual bool FireOnRequestEdit(int dispId)
    {
        var cp = _connectionPoints.FirstOrDefault(c => c.Key == typeof(IPropertyNotifySink).GUID).Value;
        if (cp == null)
            return true;

        foreach (var connection in EnumConnections.EnumerateConnections(cp))
        {
            var com = connection.As<IPropertyNotifySink>();
            if (com == null)
                continue;

            var hr = com.Object.OnRequestEdit(dispId);
            if (hr.IsFalse)
            {
                TracingUtilities.Trace($"OnRequestEdit failed: {hr}");
                return false;
            }
        }
        return true;
    }

    protected virtual void FireOnChanged(int dispId)
    {
        var cp = _connectionPoints.FirstOrDefault(c => c.Key == typeof(IPropertyNotifySink).GUID).Value;
        if (cp == null)
            return;

        foreach (var connection in EnumConnections.EnumerateConnections(cp))
        {
            var com = connection.As<IPropertyNotifySink>();
            com?.Object.OnChanged(dispId);
        }
    }

    protected T? GetAmbientProperty<T>(DISPID dispid, T? defaultValue = default) => GetAmbientProperty((int)dispid, defaultValue);
    protected T? GetAmbientProperty<T>(int dispid, T? defaultValue = default)
    {
        if (!TryGetAmbientProperty<T>(dispid, out var value))
            return defaultValue;

        return value;
    }

    protected bool TryGetAmbientProperty<T>(DISPID dispid, out T? value) => TryGetAmbientProperty<T>((int)dispid, out value);
    protected virtual bool TryGetAmbientProperty<T>(int dispid, out T? value)
    {
        if (!TryGetAmbientObjectProperty(dispid, out var v))
        {
            value = default;
            return true;
        }
        return Conversions.TryChangeType(v, out value);
    }

    protected object? GetAmbientObjectProperty(DISPID dispid, object? defaultValue = null) => GetAmbientObjectProperty((int)dispid, defaultValue);
    protected object? GetAmbientObjectProperty(int dispid, object? defaultValue = null)
    {
        if (!TryGetAmbientObjectProperty(dispid, out object? value))
            return defaultValue;

        return value;
    }

    protected bool TryGetAmbientObjectProperty(DISPID dispid, out object? value) => TryGetAmbientObjectProperty((int)dispid, out value);
    protected unsafe virtual bool TryGetAmbientObjectProperty(int dispid, out object? value)
    {
        value = null;
        var site = _site;
        if (site == null)
            return false;

        using var disp = _site.As<IDispatch>();
        if (disp == null)
            return false;

        var p = new DISPPARAMS();
        var v = new VARIANT();
        var hr = disp.Object.Invoke(dispid, Guid.Empty, 0, DISPATCH_FLAGS.DISPATCH_PROPERTYGET, p, (nint)(&v), 0, 0);
        if (hr.IsError)
        {
            TracingUtilities.Trace($"dispid: {dispid} hr: {hr}");
            return false;
        }

        using var variant = Variant.Attach(ref v);
        value = variant.Value;
        TracingUtilities.Trace($"dispid: {dispid} value: {value}");
        return true;
    }

    protected virtual void Draw(HDC hdc, RECT bounds)
    {
        TracingUtilities.Trace($"hdc: {hdc} bounds: {bounds}");
    }

    protected HRESULT DoVerb(int verbId, MSG? msg, IOleClientSite activeSite, HWND hwndParent)
    {
        var id = (OLEIVERB)verbId;
        TracingUtilities.Trace($"iVerb: {id} lpmsg: {msg} pActiveSite: {activeSite} hwndParent: {hwndParent}");
        foreach (var verb in Verbs)
        {
            if (verb.Verb.lVerb == id)
                return verb.Invoke(msg, activeSite, hwndParent);
        }
        return Constants.E_NOTIMPL;
    }

    protected virtual void SetWindowPos(RECT position)
    {
        _window?.SetWindowPos(HWND.Null, position.left, position.top, position.Width, position.Height, SET_WINDOW_POS_FLAGS.SWP_NOZORDER | SET_WINDOW_POS_FLAGS.SWP_NOACTIVATE);
    }

    protected virtual HRESULT Save(IStream stream, bool clearDirty)
    {
        if (clearDirty)
        {
            IsDirty = false;
        }
        return Constants.S_OK;
    }

    protected virtual HRESULT Load(IStream stream)
    {
        IsDirty = false;
        return Constants.S_OK;
    }

    protected virtual void SendOnDataChanged(ADVF advf = 0)
    {
        var dataObjectPtr = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance<IDataObject>(this);
        if (dataObjectPtr == 0)
            return;

        using var dataObject = DirectN.Extensions.Com.ComObject.FromPointer<IDataObject>(dataObjectPtr)!;
        if (dataObject == null)
            return;

        _dataAdviseHolder?.Object.SendOnDataChange(dataObject.Object, 0, (uint)advf);
    }

    protected virtual void OnControlInfoChanged()
    {
        using var container = _clientSite.As<IOleControlSite>();
        container?.Object.OnControlInfoChanged();
    }

    protected virtual void SendOnClose()
    {
        _adviseHolder?.Object.SendOnClose();
    }

    protected virtual void SendOnSave()
    {
        _adviseHolder?.Object.SendOnSave();
    }

    protected virtual void SendOnViewChange()
    {
        _adviseSink?.Object.OnViewChange((uint)DVASPECT.DVASPECT_CONTENT, -1);
    }

    protected virtual HRESULT Close(OLECLOSE option)
    {
        if (State == ControlState.InplaceActive)
        {
            InPlaceDeactivate();
        }

        if (_window != null)
        {
            DisposeWindow();
        }

        SendOnClose();

        if ((option == OLECLOSE.OLECLOSE_SAVEIFDIRTY || option == OLECLOSE.OLECLOSE_PROMPTSAVE) && IsDirty)
        {
            _clientSite?.Object.SaveObject();
            SendOnSave();
            SendOnDataChanged();
        }

        ChangeState(ControlState.Loaded);
        return Constants.S_OK;
    }

    protected virtual void ChangeState(ControlState newState)
    {
        TracingUtilities.Trace($"Changing state from {State} to {newState}");
        if (newState == State)
            return;

        var oldState = State;
        State = newState;

        var inPlaceSite = _inPlaceSite;
        if (inPlaceSite != null)
        {
            HRESULT hr;
            switch (newState)
            {
                case ControlState.InplaceActive:
                    hr = inPlaceSite.Object.OnInPlaceActivate();
                    if (hr.IsError)
                    {
                        TracingUtilities.Trace($"OnInPlaceActivate: {hr}");
                    }

                    if (oldState == ControlState.UIActive)
                    {
                        // VB fails with E_FAIL here, probably not a real problem
                        hr = inPlaceSite.Object.OnUIDeactivate(false);
                        if (hr.IsError)
                        {
                            TracingUtilities.Trace($"OnUIDeactivate: {hr}");
                        }
                    }
                    break;

                case ControlState.Active:
                    // nothing to do
                    break;

                case ControlState.UIActive:
                    hr = inPlaceSite.Object.OnUIActivate();
                    if (hr.IsError)
                    {
                        TracingUtilities.Trace($"OnUIActivate: {hr}");
                    }
                    break;

                case ControlState.Running:
                    if (oldState == ControlState.UIActive)
                    {
                        hr = inPlaceSite.Object.OnUIDeactivate(false);
                        if (hr.IsError)
                        {

                            TracingUtilities.Trace($"OnUIDeactivate: {hr}");
                        }

                        hr = inPlaceSite.Object.OnInPlaceDeactivate();
                        if (hr.IsError)
                        {
                            TracingUtilities.Trace($"OnInPlaceDeactivate: {hr}");
                        }
                    }
                    else if (oldState == ControlState.InplaceActive)
                    {
                        hr = inPlaceSite.Object.OnInPlaceDeactivate();
                        if (hr.IsError)
                        {
                            TracingUtilities.Trace($"OnInPlaceDeactivate: {hr}");
                        }
                    }
                    break;
            }
        }
    }

    protected virtual HRESULT DiscardUndoState(HWND hwndParent, RECT pos) => NotImplemented();

    protected virtual HRESULT UIActivate(HWND hwndParent, RECT pos)
    {
        var window = EnsureWindow(hwndParent, pos);
        TracingUtilities.Trace($"window: {window}");
        ChangeState(ControlState.UIActive);

        window?.Show();
        ((IOleInPlaceObject)this).SetObjectRects(pos, pos);
        return Constants.S_OK;
    }

    protected virtual HRESULT InplaceActivate(HWND hwndParent, RECT pos)
    {
        var window = EnsureWindow(hwndParent, pos);
        TracingUtilities.Trace($"window: {window}");
        ChangeState(ControlState.InplaceActive);

        window?.Show();
        ((IOleInPlaceObject)this).SetObjectRects(pos, pos);
        return Constants.S_OK;
    }

    protected virtual HRESULT InPlaceDeactivate()
    {
        ChangeState(ControlState.Running);
        return Constants.S_OK;
    }

    protected virtual HRESULT Open(HWND hwndParent, RECT pos)
    {
        TracingUtilities.Trace($"hwndParent: {hwndParent} pos: {pos}");
        _window?.Show();
        return InplaceActivate(hwndParent, pos);
    }

    protected virtual HRESULT Hide(HWND hwndParent, RECT pos)
    {
        TracingUtilities.Trace($"hwndParent: {hwndParent} pos: {pos}");
        _window?.Hide();
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
        TracingUtilities.Trace($"parentHandle: {parentHandle} rect: {rect} window: {_window}");
        if (_window == null)
        {
            var window = CreateWindow(parentHandle, rect);
            if (window != null)
            {
                // our window can be destroyed by parent as we're probably a child of it, so we need to track that
                window.Destroyed += OnWindowDestroyed;
                window.Resized += OnWindowResized;
                window.FocusChanged += OnWindowFocusChanged;
                window.Show();
                _clientSite?.Object.OnShowWindow(true);
            }
            _window = window;
        }
        return _window;
    });

    protected virtual void OnWindowDestroyed(object? sender, EventArgs e)
    {
        _clientSite?.Object.OnShowWindow(false);
        DisposeWindow();
    }

    protected virtual void OnWindowResized(object? sender, ValueEventArgs<(WindowResizedType ResizedType, SIZE Size)> e)
    {
        TracingUtilities.Trace($"sender: {sender} type: {e.Value.ResizedType} size: {e.Value.Size}");
        SendOnViewChange();
        SendOnDataChanged();
    }

    protected virtual void OnWindowFocusChanged(object? sender, ValueEventArgs<bool> e)
    {
        TracingUtilities.Trace($"sender: {sender}");
        using var container = _clientSite.As<IOleControlSite>();
        container?.Object.OnFocus(e.Value).ThrowOnError();
    }

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
            throw new ArgumentException($"Connection point with iid {iid.GetName()} is already registered", nameof(connectionPoint));
    }

    protected override CustomQueryInterfaceResult GetInterface(ref Guid iid, out nint ppv)
    {
        ppv = 0;
        TracingUtilities.Trace($"iid: {iid.GetName()}");

        if (State != ControlState.InplaceActive && State != ControlState.UIActive && iid == typeof(IOleInPlaceObject).GUID)
        {
            TracingUtilities.Trace($"iid: {iid.GetName()} not allowed in inplace or UI active state. Current state: {State}");
            return CustomQueryInterfaceResult.Failed;
        }

        if (!_aggregableInterfacesIids.Contains(iid))
        {
            var wrapper = AggregationWrapper;
            if (wrapper != 0)
            {
                var hr = Aggregable.OuterQueryInterface(wrapper, iid, out ppv);
                TracingUtilities.Trace($"iid: {iid.GetName()} hr: {hr}");
                return hr.IsError ? CustomQueryInterfaceResult.Failed : CustomQueryInterfaceResult.Handled;
            }

            TracingUtilities.Trace($"iid: {iid.GetName()} wrapper was not set");
        }

        return CustomQueryInterfaceResult.NotHandled;
    }

    protected virtual void DisposeConnectionPoints()
    {
        foreach (var cp in _connectionPoints.Values)
        {
            if (cp is BaseConnectionPoint target)
            {
                target._container = null;
            }

            if (cp is IDisposable disposable)
            {
                TracingUtilities.Trace($"disposable: {disposable}");
                disposable.Dispose();
            }
        }
        _connectionPoints.Clear();
    }

    protected virtual void DisposeWindow()
    {
        TracingUtilities.Trace($"window: {_window}");
        var window = Interlocked.Exchange(ref _window, null);
        if (window != null)
        {
            window.Destroyed -= OnWindowDestroyed;
            window.Resized -= OnWindowResized;
            window.FocusChanged -= OnWindowFocusChanged;
            window.Dispose();
        }
    }

    protected override void Dispose(bool disposing) => TracingUtilities.WrapErrors(() =>
    {
        if (disposing)
        {
            Interlocked.Exchange(ref _adviseHolder, null)?.Dispose();
            Interlocked.Exchange(ref _dataAdviseHolder, null)?.Dispose();
            Interlocked.Exchange(ref _clientSite, null)?.Dispose();
            Interlocked.Exchange(ref _inPlaceSite, null)?.Dispose();
            Interlocked.Exchange(ref _inPlaceSiteEx, null)?.Dispose();
            Interlocked.Exchange(ref _inPlaceSiteWindowless, null)?.Dispose();
            Interlocked.Exchange(ref _site, null)?.Dispose();
            Interlocked.Exchange(ref _adviseSink, null)?.Dispose();
            DisposeWindow();
            DisposeConnectionPoints();
        }
        base.Dispose(disposing);
        return Constants.S_OK;
    });

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
        Interlocked.Exchange(ref _clientSite, null)?.Dispose();
        _clientSite = pClientSite != null ? new ComObject<IOleClientSite>(pClientSite) : null;

        Interlocked.Exchange(ref _inPlaceSite, null)?.Dispose();
        _inPlaceSite = _clientSite.As<IOleInPlaceSite>();

        Interlocked.Exchange(ref _inPlaceSiteEx, null)?.Dispose();
        _inPlaceSiteEx = _clientSite.As<IOleInPlaceSiteEx>();

        Interlocked.Exchange(ref _inPlaceSiteWindowless, null)?.Dispose();
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
        var option = (OLECLOSE)dwSaveOption;
        TracingUtilities.Trace($"dwSaveOption: {option}");
        return Close(option);
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
        // note: posRect is in pixels
        var pos = RECT.Sized(0, 0, 192, 192);
        if (lprcPosRect != 0)
        {
            pos = *(RECT*)lprcPosRect;
            TracingUtilities.Trace($"rcPosRect: {pos}");
        }

        TracingUtilities.Trace($"iVerb: {verb} lpmsg: {lpmsg} pActiveSite: {pActiveSite} lindex: {lindex} hwndParent: {hwndParent} lprcPosRect: {lprcPosRect} pos: {pos}");
        var hr = verb switch
        {
            OLEIVERB.OLEIVERB_INPLACEACTIVATE => InplaceActivate(hwndParent, pos),
            OLEIVERB.OLEIVERB_SHOW or OLEIVERB.OLEIVERB_PRIMARY or OLEIVERB.OLEIVERB_OPEN => Open(hwndParent, pos),
            OLEIVERB.OLEIVERB_HIDE => Hide(hwndParent, pos),
            OLEIVERB.OLEIVERB_UIACTIVATE => UIActivate(hwndParent, pos),
            OLEIVERB.OLEIVERB_DISCARDUNDOSTATE => DiscardUndoState(hwndParent, pos),
            _ => Constants.E_NOTIMPL,
        };
        return hr;
    });

    HRESULT IOleObject.EnumVerbs(out IEnumOLEVERB ppEnumOleVerb)
    {
        TracingUtilities.Trace($"verbs: {Verbs.Count}");
        ppEnumOleVerb = new EnumVerbs(Verbs);
        return Constants.S_OK;
    }

    HRESULT IOleObject.Update()
    {
        TracingUtilities.Trace();
        return Constants.S_OK;
    }

    HRESULT IOleObject.IsUpToDate()
    {
        return Constants.S_OK;
    }

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
            var form = (USERCLASSTYPE)dwFormOfType;
            var hr = Functions.OleRegGetUserType(GetType().GUID, dwFormOfType, out userType);
            TracingUtilities.Trace($"dwFormOfType: {form} pszUserType: '{userType}' hr: {hr}");
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
            var dpi = GetDpi();
            TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect} psizel: {size} pixels: {size.HiMetricToPixel(dpi)}");
            if (dwDrawAspect != DVASPECT.DVASPECT_CONTENT)
                return Constants.DV_E_DVASPECT;

            if (_changingExtent)
                return Constants.S_OK;

            _changingExtent = true;
            try
            {
                HiMetricExtent = size;
                var rc = new RECT(0, 0, HiMetricExtent.cx.HiMetricToPixel(dpi), HiMetricExtent.cy.HiMetricToPixel(dpi));
                SetWindowPos(rc);
                SendOnViewChange();

                IsDirty = true;

                // needed or not?
                //_inPlaceSite?.Object.OnPosRectChange(rc);
            }
            finally
            {
                _changingExtent = false;
            }
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

            size = HiMetricExtent;
            TracingUtilities.Trace($"psizel: {size} pixels: {size.HiMetricToPixel(GetDpi())}");
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
            var hr = _adviseHolder?.Object.Advise(pAdvSink, out connection) ?? Constants.E_NOTIMPL;
            TracingUtilities.Trace($"dwConnection {connection}: {hr}");
            return hr;
        });
        pdwConnection = connection;
        return hr;
    }

    HRESULT IOleObject.Unadvise(uint dwConnection) => TracingUtilities.WrapErrors(() =>
    {
        var hr = _adviseHolder?.Object.Unadvise(dwConnection) ?? Constants.E_NOTIMPL;
        TracingUtilities.Trace($"dwConnection {dwConnection}: {hr}");
        return hr;
    });

    HRESULT IOleObject.EnumAdvise(out IEnumSTATDATA ppenumAdvise)
    {
        TracingUtilities.Trace();
        IEnumSTATDATA? enumAdvise = null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var hr = _adviseHolder?.Object.EnumAdvise(out enumAdvise) ?? Constants.E_NOTIMPL;
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
            pti = ti?.Object;
            TracingUtilities.Trace($"ppTI: {pti}");
            return pti != null ? Constants.S_OK : Constants.E_UNEXPECTED;
        });
        ppTI = pti!;
        return hr;
    }

    HRESULT IProvideClassInfo2.GetGUID(uint dwGuidKind, out Guid pGUID)
    {
        var guid = Guid.Empty;
        var kind = (GUIDKIND)dwGuidKind;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"dwGuidKind: {kind}");
            if (kind == GUIDKIND.GUIDKIND_DEFAULT_SOURCE_DISP_IID)
            {
                if (_connectionPoints.IsEmpty)
                    return Constants.E_NOINTERFACE;

                // return the first connection point's IID
                var diid = _connectionPoints.Values.OfType<BaseConnectionPoint>().FirstOrDefault(c => c.IsIDispatch)?.InterfaceId;
                TracingUtilities.Trace($"dwGuidKind: {kind} diid: {diid?.GetName()}");
                if (diid.HasValue)
                {
                    guid = diid.Value;
                    return Constants.S_OK;
                }
            }
            return Constants.E_INVALIDARG;
        });
        pGUID = guid;
        return hr;
    }

    HRESULT IPersistStreamInit.IsDirty()
    {
        TracingUtilities.Trace($"IsDirty: {IsDirty}");
        return IsDirty ? Constants.S_OK : Constants.S_FALSE;
    }

    HRESULT IPersistStreamInit.Load(IStream pStm) => TracingUtilities.WrapErrors(() =>
    {
        var hr = Load(pStm);
        TracingUtilities.Trace($"pStm {pStm} hr: {hr}");
        ChangeState(ControlState.Loaded);
        return hr;
    });

    HRESULT IPersistStreamInit.GetSizeMax(out ulong pCbSize) { pCbSize = 0; return NotImplemented(); }
    HRESULT IPersistStreamInit.Save(IStream pStm, BOOL fClearDirty) => TracingUtilities.WrapErrors(() =>
    {
        var hr = Save(pStm, fClearDirty);
        TracingUtilities.Trace($"pStm {pStm} fClearDirty: {fClearDirty} hr: {hr}");
        if (hr.IsSuccess)
        {
            SendOnSave();
        }
        return hr;
    });

    HRESULT IPersistStreamInit.InitNew() => TracingUtilities.WrapErrors(() =>
    {
        IsDirty = true;
        TracingUtilities.Trace();
        return Constants.S_OK;
    });

    HRESULT IPersist.GetClassID(out Guid pClassID)
    {
        var classID = Guid.Empty;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            classID = GetType().GUID;
            TracingUtilities.Trace($"pClassID: {classID}");
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
            if (dwDrawAspect != DVASPECT.DVASPECT_CONTENT)
                return Constants.DV_E_DVASPECT;

            size = HiMetricExtent;
            TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect} lindex: {lindex} ptd: {ptd} size: {HiMetricExtent}");
            return Constants.S_OK;
        });
        lpsizel = size;
        return hr;
    }

    unsafe HRESULT IViewObject.Draw(
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
        TracingUtilities.Trace($"dwDrawAspect: {dwDrawAspect} lindex: {lindex} pvAspect: {pvAspect} ptd: {ptd} hdcTargetDev: {hdcTargetDev} hdcDraw: {hdcDraw} lprcBounds: {lprcBounds} lprcWBounds: {lprcWBounds} pfnContinue: {pfnContinue} dwContinue: {dwContinue} window: {Window}");
        if (dwDrawAspect == DVASPECT.DVASPECT_DOCPRINT)
            return Constants.DV_E_DVASPECT;

        // note if we're being called as non active
        // we could get a window as parent but it can use a bad window (tested with Excel)
        // instead we'll draw something simple or just do nothing
        if (hdcDraw != 0 && lprcBounds != 0)
        {
            var rcBounds = *(RECT*)lprcBounds;
            Draw(hdcDraw, rcBounds);
        }
        return Constants.S_OK;
    });

    HRESULT IViewObject.GetColorSet(DVASPECT dwDrawAspect, int lindex, nint pvAspect, nint ptd, HDC hicTargetDev, out nint ppColorSet) { ppColorSet = 0; return NotImplemented(); }
    HRESULT IViewObject.Freeze(DVASPECT dwDrawAspect, int lindex, nint pvAspect, out uint pdwFreeze) { pdwFreeze = 0; return NotImplemented(); }
    HRESULT IViewObject.Unfreeze(uint dwFreeze) => NotImplemented();
    HRESULT IViewObject.SetAdvise(DVASPECT aspects, uint advf, IAdviseSink pAdvSink) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"Sink: {pAdvSink}");
        Interlocked.Exchange(ref _adviseSink, null)?.Dispose();
        _adviseSink = pAdvSink != null ? new ComObject<IAdviseSink>(pAdvSink) : null;
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

        var extentInfo = *(DVEXTENTINFO*)pExtentInfo;
        var mode = (DVEXTENTMODE)extentInfo.dwExtentMode;
        TracingUtilities.Trace($"extentInfo sizelProposed: {extentInfo.sizelProposed} mode: {mode}");

        var natural = HiMetricNaturalExtent;
        if (natural == null)
            return Constants.E_NOTIMPL;

        *(SIZE*)pSizel = natural.Value;
        TracingUtilities.Trace($"sizel: {natural.Value}");
        return Constants.S_OK;
    });

    HRESULT IOleControl.GetControlInfo(ref CONTROLINFO pCI)
    {
        var accelerators = KeyboardAccelerators;
        if (accelerators != null && !accelerators.IsDisposed && accelerators.Count > 0)
        {
            pCI.cAccel = (ushort)accelerators.Count;
            pCI.hAccel = accelerators.Handle;
        }
        else
        {
            pCI.cAccel = 0;
            pCI.hAccel = HACCEL.Null;
        }

        pCI.dwFlags = (uint)KeyboardBehavior;
        TracingUtilities.Trace($"pCI.cAccel: {pCI.cAccel} pCI.hAccel: {pCI.hAccel} pCI.dwFlags: {pCI.dwFlags}");
        return Constants.S_OK;
    }

    HRESULT IOleControl.OnMnemonic(in MSG pMsg)
    {
        var msg = pMsg;
        return TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"pMSg: {MessageDecoder.Decode(msg)}");
            return Constants.S_OK;
        });
    }

    protected virtual void OnAmbientPropertyChanged(object sender, ValueEventArgs<DISPID> e) => AmbientPropertyChanged?.Invoke(this, e);
    protected virtual void OnAmbientPropertyChanged(DISPID dispId)
    {
        // https://learn.microsoft.com/en-us/windows/win32/com/standard-properties
        if (dispId == DISPID.DISPID_AMBIENT_USERMODE && Window != null)
        {
            // this is mostly necessary for Word
            if (InUserMode)
            {
                Window.Show();
            }
            else
            {
                Window.Hide();
            }
        }
        OnAmbientPropertyChanged(this, new(dispId));
    }

    HRESULT IOleControl.OnAmbientPropertyChange(int dispId) => TracingUtilities.WrapErrors(() =>
    {
        var dispid = (DISPID)dispId;
        TracingUtilities.Trace($"dispID: {dispid}");
        OnAmbientPropertyChanged(dispid);
        return Constants.S_OK;
    });

    HRESULT IOleControl.FreezeEvents(BOOL bFreeze) => TracingUtilities.WrapErrors(() =>
    {
        if (bFreeze)
        {
            FreezeCount++;
        }
        else if (FreezeCount > 0)
        {
            FreezeCount--;
        }

        TracingUtilities.Trace($"bFreeze: {bFreeze} count:{FreezeCount}");
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

    HRESULT IDataObject.GetDataHere(in FORMATETC pformatetc, ref STGMEDIUM pmedium) => NotImplemented();
    HRESULT IDataObject.GetCanonicalFormatEtc(in FORMATETC pformatectIn, out FORMATETC pformatetcOut) { pformatetcOut = new(); return NotImplemented(); }
    HRESULT IDataObject.SetData(in FORMATETC pformatetc, in STGMEDIUM pmedium, BOOL fRelease) => NotImplemented();
    HRESULT IDataObject.EnumFormatEtc(uint dwDirection, out IEnumFORMATETC ppenumFormatEtc) { ppenumFormatEtc = null!; return NotImplemented(); }

    HRESULT IDataObject.QueryGetData(in FORMATETC pformatetc)
    {
        var tymed = (TYMED)pformatetc.tymed;
        TracingUtilities.Trace($"tymed: {tymed} format: {pformatetc.cfFormat} ('{Clipboard.GetFormatName(pformatetc.cfFormat)}') aspect: {(DVASPECT)pformatetc.dwAspect}");
        if (tymed.HasFlag(TYMED.TYMED_MFPICT) || tymed.HasFlag(TYMED.TYMED_ENHMF))
            return Constants.S_OK;

        return Constants.DV_E_TYMED;
    }

    protected virtual unsafe HRESULT DrawMetaFile(TYMED tymed, ref STGMEDIUM medium, Action<HDC> draw)
    {
        ArgumentNullException.ThrowIfNull(draw);
        if (tymed.HasFlag(TYMED.TYMED_MFPICT))
        {
            var hdc = Functions.CreateMetaFileW(PWSTR.Null);
            if (hdc == 0)
                return Constants.E_UNEXPECTED;

            draw(hdc);

            var hMF = Functions.CloseMetaFile(hdc);
            if (hdc == 0)
            {
                Functions.DeleteMetaFile(hMF);
                return Constants.E_UNEXPECTED;
            }

            var size = sizeof(METAFILEPICT);
            var pict = (METAFILEPICT*)Marshal.AllocHGlobal(size);
            Unsafe.InitBlockUnaligned(pict, 0, (uint)size);

            // For MM_ANISOTROPIC pictures, xExt and yExt can be zero when no suggested size is supplied.
            pict->hMF = hMF;
            pict->mm = (int)HDC_MAP_MODE.MM_ANISOTROPIC;

            medium.tymed = (uint)TYMED.TYMED_MFPICT;
            medium.u.hGlobal = (nint)pict;
            return Constants.S_OK;
        }

        if (tymed.HasFlag(TYMED.TYMED_ENHMF))
        {
            var hdc = Functions.CreateEnhMetaFileW(HDC.Null, PWSTR.Null, 0, PWSTR.Null);
            if (hdc == 0)
                return Constants.E_UNEXPECTED;

            draw(hdc);

            var hMF = Functions.CloseEnhMetaFile(hdc);
            if (hdc == 0)
            {
                Functions.DeleteEnhMetaFile(hMF);
                return Constants.E_UNEXPECTED;
            }

            medium.tymed = (uint)TYMED.TYMED_ENHMF;
            medium.u.hEnhMetaFile = hMF;
            return Constants.S_OK;
        }
        throw new NotSupportedException();
    }

    // this supports Word preview for example
    unsafe HRESULT IDataObject.GetData(in FORMATETC pformatetcIn, out STGMEDIUM pmedium)
    {
        var medium = new STGMEDIUM();
        var format = pformatetcIn;
        var tymed = (TYMED)pformatetcIn.tymed;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"tymed: {tymed} format: {format.cfFormat} ('{Clipboard.GetFormatName(format.cfFormat)}') aspect: {(DVASPECT)format.dwAspect}");
            if (format.dwAspect != (uint)DVASPECT.DVASPECT_CONTENT)
                return Constants.DV_E_DVASPECT;

            if (tymed.HasFlag(TYMED.TYMED_MFPICT) || tymed.HasFlag(TYMED.TYMED_ENHMF))
                return DrawMetaFile(tymed, ref medium, hdc =>
                {
                    var dpi = GetDpi();
                    var rc = new RECT(0, 0, HiMetricExtent.cx.HiMetricToPixel(dpi), HiMetricExtent.cy.HiMetricToPixel(dpi));
                    TracingUtilities.Trace($"hdc: {hdc} dpi: {dpi} rc: {rc}");
                    _ = Functions.SaveDC(hdc);
                    Functions.SetWindowOrgEx(hdc, 0, 0, 0);
                    Functions.SetWindowExtEx(hdc, rc.Width, rc.Height, 0);
                    Draw(hdc, rc);
                    Functions.RestoreDC(hdc, -1);
                    SendOnViewChange();
                });

            return Constants.DV_E_FORMATETC;
        });
        pmedium = medium;
        return hr;
    }

    HRESULT IDataObject.DAdvise(in FORMATETC pformatetc, uint advf, IAdviseSink pAdvSink, out uint pdwConnection)
    {
        pdwConnection = 0;
        var hr = _dataAdviseHolder?.Object.Advise(this, pformatetc, advf, pAdvSink, out pdwConnection);
        TracingUtilities.Trace($"pformatetc: {pformatetc} advf: {(ADVF)advf} pAdvSink: {pAdvSink} connection: {pdwConnection}");
        return hr ?? Constants.E_NOTIMPL;
    }

    HRESULT IDataObject.DUnadvise(uint dwConnection)
    {
        var hr = _dataAdviseHolder?.Object.Unadvise(dwConnection);
        TracingUtilities.Trace($"dwConnection: {dwConnection}");
        return hr ?? Constants.E_NOTIMPL;
    }

    HRESULT IDataObject.EnumDAdvise(out IEnumSTATDATA ppenumAdvise)
    {
        IEnumSTATDATA? enumAdvise = null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var hr = _dataAdviseHolder?.Object.EnumAdvise(out enumAdvise) ?? Constants.E_NOTIMPL;
            TracingUtilities.Trace($"ppenumAdvise {enumAdvise}: {hr}");
            return hr;
        });
        ppenumAdvise = enumAdvise!;
        return hr;
    }

    HRESULT IObjectWithSite.SetSite(nint pUnkSite) => TracingUtilities.WrapErrors(() =>
    {
        Interlocked.Exchange(ref _site, null)?.Dispose();
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

    unsafe HRESULT IOleInPlaceActiveObject.TranslateAccelerator(nint lpmsg) => TracingUtilities.WrapErrors(() =>
    {
        if (lpmsg == 0)
            return Constants.E_POINTER;

        var msg = *(MSG*)lpmsg;
        TracingUtilities.Trace($"lpmsg: {MessageDecoder.Decode(msg)}");

        using var container = _clientSite.As<IOleControlSite>();
        if (container != null)
        {
            var km = KeyboardUtilities.GetKEYMODIFIERS();
            var hr = container.Object.TranslateAccelerator(msg, km);
            TracingUtilities.Trace($"km: {km} ok: {hr.IsOk}");
            return hr.IsSuccess ? Constants.S_OK : Constants.S_FALSE;
        }

        return Constants.S_FALSE;
    });

    HRESULT IOleInPlaceActiveObject.OnFrameWindowActivate(BOOL fActivate) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.OnDocWindowActivate(BOOL fActivate) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.ResizeBorder(in RECT prcBorder, IOleInPlaceUIWindow pUIWindow, BOOL fFrameWindow) => NotImplemented();
    HRESULT IOleInPlaceActiveObject.EnableModeless(BOOL fEnable) => NotImplemented();
    HRESULT IOleInPlaceObject.InPlaceDeactivate() => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"State: {State}");
        var hr = InPlaceDeactivate();
        return hr;
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
            var window = _window;
            TracingUtilities.Trace($"lprcPosRect: {pos} lprcClipRect: {clip} window: {window}");
            if (window != null)
            {
                //var tempRgn = new HRGN();
                //if (Functions.IntersectRect(out var rc, pos, clip) && !Functions.EqualRect(rc, pos))
                //{
                //    Functions.OffsetRect(ref rc, -pos.left, -pos.top);
                //    tempRgn = Functions.CreateRectRgnIndirect(rc);
                //}

                //_ = Functions.SetWindowRgn(window.Handle, tempRgn, true);
                SetWindowPos(pos);
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
                Interlocked.Exchange(ref _clientSite, null)?.Dispose();
                _clientSite = DirectN.Extensions.Com.ComObject.FromPointer<IOleClientSite>(container.pClientSite);

                Interlocked.Exchange(ref _inPlaceSite, null)?.Dispose();
                _inPlaceSite = _clientSite.As<IOleInPlaceSite>();

                Interlocked.Exchange(ref _inPlaceSiteEx, null)?.Dispose();
                _inPlaceSiteEx = _clientSite.As<IOleInPlaceSiteEx>();

                Interlocked.Exchange(ref _inPlaceSiteWindowless, null)?.Dispose();
                _inPlaceSiteWindowless = _clientSite.As<IOleInPlaceSiteWindowless>();
            }

            TracingUtilities.Trace($"ClientSite: {_clientSite}");
            TracingUtilities.Trace($"InPlaceSite: {_inPlaceSite}");
            TracingUtilities.Trace($"InPlaceSiteEx: {_inPlaceSiteEx}");
            TracingUtilities.Trace($"InPlaceSiteWindowless: {_inPlaceSiteWindowless}");

            if (container.pAdviseSink != 0)
            {
                Interlocked.Exchange(ref _adviseSink, null)?.Dispose();
                _adviseSink = DirectN.Extensions.Com.ComObject.FromPointer<IAdviseSink>(container.pAdviseSink);
            }

            if (container.pPropertyNotifySink != 0)
            {
                ((IConnectionPointContainer)this).FindConnectionPoint(typeof(IPropertyNotifySink).GUID, out var cpObj);
                if (cpObj is BaseConnectionPoint bcp)
                {
                    ((IConnectionPoint)bcp).Advise(container.pPropertyNotifySink, out var cookie).ThrowOnError();
                    control.dwPropNotifyCookie = cookie;
                }
                else
                {
                    using var cp = cpObj != null ? new ComObject<IConnectionPoint>(cpObj) : null;
                    if (cp != null)
                    {
                        cp.Object.Advise(container.pPropertyNotifySink, out var cookie).ThrowOnError();
                        control.dwPropNotifyCookie = cookie;
                    }
                }
            }

            if (container.pUnkEventSink != 0)
            {
                ((IProvideClassInfo2)this).GetGUID((uint)GUIDKIND.GUIDKIND_DEFAULT_SOURCE_DISP_IID, out var iid);
                if (iid != Guid.Empty)
                {
                    ((IConnectionPointContainer)this).FindConnectionPoint(iid, out var cpObj);
                    if (cpObj is BaseConnectionPoint bcp)
                    {
                        ((IConnectionPoint)bcp).Advise(container.pUnkEventSink, out var cookie).ThrowOnError();
                        control.dwEventCookie = cookie;
                    }
                    else
                    {
                        using var cp = cpObj != null ? new ComObject<IConnectionPoint>(cpObj) : null;
                        if (cp != null)
                        {
                            cp.Object.Advise(container.pUnkEventSink, out var cookie).ThrowOnError();
                            control.dwEventCookie = cookie;
                        }
                    }
                }
            }

            TracingUtilities.Trace($"Sink: {_adviseSink}");

            control.dwMiscStatus = (uint)MiscStatus;
            control.dwViewStatus = (uint)ViewStatus;
            control.dwPointerActivationPolicy = (uint)PointerActivationPolicy;
            TracingUtilities.Trace($"control dwMiscStatus: {MiscStatus} dwViewStatus: {ViewStatus} dwPointerActivationPolicy: {PointerActivationPolicy} dwEventCookie: {control.dwEventCookie} dwPropNotifyCookie: {control.dwPropNotifyCookie}");
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
            HiMetricExtent = size;
            TracingUtilities.Trace($"pSizel: {size} pixels: {HiMetricExtent.HiMetricToPixel(GetDpi())}");
            return ((IOleObject)this).SetExtent(DVASPECT.DVASPECT_CONTENT, size);
        });
    }

    HRESULT IQuickActivate.GetContentExtent(out SIZE pSizel)
    {
        var size = new SIZE();
        var hr = TracingUtilities.WrapErrors(() =>
        {
            size = HiMetricExtent;
            TracingUtilities.Trace($"pSizel: {size} pixels: {size.HiMetricToPixel(GetDpi())}");
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

    HRESULT IRunnableObject.GetRunningClass(out Guid lpClsid)
    {
        TracingUtilities.Trace();
        // is this ok?
        lpClsid = Guid.Empty;
        //lpClsid = GetType().GUID;
        return Constants.E_UNEXPECTED;
    }

    HRESULT IRunnableObject.Run(IBindCtx pbc)
    {
        TracingUtilities.Trace($"pbc: {pbc}");
        return Constants.S_OK;
    }

    BOOL IRunnableObject.IsRunning()
    {
        TracingUtilities.Trace("true");
        return true;
    }

    HRESULT IRunnableObject.LockRunning(BOOL fLock, BOOL fLastUnlockCloses)
    {
        TracingUtilities.Trace($"fLock: {fLock} fLastUnlockCloses: {fLastUnlockCloses}");
        return Constants.S_OK;
    }

    HRESULT IRunnableObject.SetContainedObject(BOOL fContained)
    {
        TracingUtilities.Trace($"fContained: {fContained}");
        return Constants.S_OK;
    }

    unsafe HRESULT ISpecifyPropertyPages.GetPages(out CAUUID pPages)
    {
        var pages = new CAUUID();
        var hr = TracingUtilities.WrapErrors(() =>
        {
            pages = GetPages();
            return Constants.S_OK;
        });
        pPages = pages;
        return hr;
    }

    HRESULT ICategorizeProperties.MapPropertyToCategory(DISPID dispId, out PROPCAT ppropcat)
    {
        var propcat = PROPCAT.PROPCAT_Misc;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            propcat = MapPropertyToCategory((int)dispId);
            TracingUtilities.Trace($"dispId: {dispId} (0x:{dispId:X}) cat: {propcat}");
            return Constants.S_OK;
        });
        ppropcat = propcat;
        return hr;
    }

    HRESULT ICategorizeProperties.GetCategoryName(PROPCAT propcat, uint lcid, out BSTR pbstrName)
    {
        var bstr = BSTR.Null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var name = GetCategoryName(propcat, lcid);
            bstr = new BSTR(Marshal.StringToBSTR(name));
            TracingUtilities.Trace($"propcat: {propcat} lcid: {lcid} name: '{name}'");
            return Constants.S_OK;
        });
        pbstrName = bstr;
        return hr;
    }

    HRESULT IPerPropertyBrowsing.GetDisplayString(int dispId, out BSTR pBstr)
    {
        var bstr = BSTR.Null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var name = GetDisplayString(dispId);
            TracingUtilities.Trace($"dispId: {dispId} (0x:{dispId:X}) name: '{name}'");
            if (name != null)
            {
                bstr = new BSTR(Marshal.StringToBSTR(name));
                return Constants.S_OK;
            }
            return Constants.S_FALSE;
        });
        pBstr = bstr;
        return hr;
    }

    HRESULT IPerPropertyBrowsing.MapPropertyToPage(int dispId, out Guid pClsid)
    {
        var clsid = DefaultPropertyPageId;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var id = MapPropertyToPage(dispId);
            TracingUtilities.Trace($"dispId: {dispId} (0x:{dispId:X}) clsid: {id}");
            if (id != null)
            {
                clsid = id.Value;
                return Constants.S_OK;
            }
            return Constants.PERPROP_E_NOPAGEAVAILABLE;
        });
        pClsid = clsid;
        return hr;
    }

    unsafe HRESULT IPerPropertyBrowsing.GetPredefinedStrings(int dispId, nint pCaStringsOut, nint pCaCookiesOut)
    {
        var strElemsCount = 0u;
        var strElems = nint.Zero;

        var cookiesElemsCount = 0u;
        var cookiesElems = nint.Zero;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            var hr = Constants.E_UNEXPECTED;
            if (PredefinedStrings.TryGetValue(dispId, out var list) && list.Count > 0)
            {
                if (pCaStringsOut != 0)
                {
                    strElemsCount = (uint)list.Count;
                    strElems = Marshal.AllocCoTaskMem(sizeof(PWSTR) * list.Count);
                }

                if (pCaCookiesOut != 0)
                {
                    cookiesElemsCount = (uint)list.Count;
                    cookiesElems = Marshal.AllocCoTaskMem(sizeof(uint) * list.Count);
                }

                if (pCaCookiesOut != 0 || pCaStringsOut != 0)
                {
                    var strArray = (PWSTR*)strElems;
                    var cookiesArray = (uint*)cookiesElems;
                    for (var i = 0; i < list.Count; i++)
                    {
                        var predefined = list[i];
                        if (cookiesArray != null)
                        {
                            cookiesArray[i] = predefined.Id;
                        }

                        if (strArray != null)
                        {
                            strArray[i] = new PWSTR(Marshal.StringToCoTaskMemUni(predefined.Name));
                        }
                    }
                    hr = Constants.S_OK;
                }
            }
            TracingUtilities.Trace($"dispId: {dispId} (0x:{dispId:X}) strings: {strElemsCount} cookies: {cookiesElemsCount}");
            return hr;
        });

        // some containers (Word) seems to send null for pCaStringsOut and/or pCaCookiesOut here...
        if (pCaStringsOut != 0)
        {
            *(CALPOLESTR*)pCaStringsOut = new CALPOLESTR { cElems = strElemsCount, pElems = strElems };
        }

        if (pCaCookiesOut != 0)
        {
            *(CADWORD*)pCaCookiesOut = new CADWORD { cElems = cookiesElemsCount, pElems = cookiesElems };
        }
        return hr;
    }

    HRESULT IPerPropertyBrowsing.GetPredefinedValue(int dispId, uint dwCookie, out VARIANT pVarOut)
    {
        TracingUtilities.Trace($"dispId: {dispId} (0x:{dispId:X}) dwCookie: {dwCookie}");
        var variant = new VARIANT();
        var hr = TracingUtilities.WrapErrors(() =>
        {
            if (PredefinedStrings.TryGetValue(dispId, out var list) && list.Count > 0)
            {
                var predefined = list.FirstOrDefault(p => p.Id == dwCookie);
                if (predefined != null)
                {
                    using var v = new Variant(predefined.Value);
                    variant = v.Detach();
                    TracingUtilities.Trace($" pVarOut: {variant}");
                    return Constants.S_OK;
                }
            }
            return Constants.E_INVALIDARG;
        });
        pVarOut = variant;
        return hr;
    }
}
