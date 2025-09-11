﻿namespace ActiveN;

[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
[GeneratedComClass]
public abstract partial class BaseDispatch : IDisposable, ICustomQueryInterface,
    IProvideClassInfo,
    IVsPerPropertyBrowsing,
    IVSMDPerPropertyBrowsing,
    IProvidePropertyBuilder,
    ICategorizeProperties,
    INotifyPropertyChanged
{
    private static readonly ConcurrentDictionary<Type, DispatchType> _cache = new();

    private ComObject<ITypeInfo>? _typeInfo;
    private ComObject<ITypeInfo>? _dispatchInterfaceInfo;
    private bool _typeInfosLoaded;
    private bool _disposedValue;

    public event PropertyChangedEventHandler? PropertyChanged; // stock property change notifications

    public bool IsDisposed => _disposedValue;
    protected virtual IDictionary<string, object?> StockProperties { get; } = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
    public virtual Guid? PropertyPageId { get; set; }
    protected virtual object? GetTaskResult(Task task) => null;
    protected virtual HWND GetWindowHandle() => HWND.Null;
    protected virtual IEnumerable<Guid> PropertyPagesIds { get; set; } = [];
    protected abstract ComRegistration ComRegistration { get; }
    protected abstract Guid DispatchInterfaceId { get; }
    protected virtual IDictionary<int, IReadOnlyList<PredefinedString>> PredefinedStrings { get; set; } = new Dictionary<int, IReadOnlyList<PredefinedString>>();
    protected virtual int AutoDispidsBase => 0x10000;

    protected virtual void SetStockProperty(object? value, [CallerMemberName] string? name = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (StockProperties.TryGetValue(name, out var current))
        {
            if (Equals(current, value))
                return;

            StockProperties[name] = value;
            OnStockPropertyChanged(name);
        }
        else
        {
            StockProperties[name] = value;
            OnStockPropertyChanged(name);
        }
    }

    protected virtual T? GetStockProperty<T>(T? defaultValue = default, [CallerMemberName] string? name = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (!StockProperties.TryGetValue(name, out var current))
            return defaultValue;

        if (!Conversions.TryChangeType<T>(current, out var typed))
            return defaultValue;

        return typed;
    }

    protected virtual object? GetStockObjectProperty([CallerMemberName] string? name = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (StockProperties.TryGetValue(name, out var current))
            return current;

        return null;
    }

    protected virtual DispatchType CreateType([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)] Type type) => new(type);
    protected virtual void OnStockPropertyChanged(object sender, PropertyChangedEventArgs e) => PropertyChanged?.Invoke(this, e);
    protected virtual void OnStockPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(propertyName);

        var type = GetDispatchType();
        if (type.TryGetDispId(propertyName, out var dispId))
        {
            OnStockPropertyChanged(dispId);
        }
        OnStockPropertyChanged(this, new PropertyChangedEventArgs(propertyName));
    }

    protected virtual void OnStockPropertyChanged(int dispId)
    {
    }

    CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out nint ppv) => GetInterface(ref iid, out ppv);
    protected virtual CustomQueryInterfaceResult GetInterface(ref Guid iid, out nint ppv)
    {
        ppv = 0;
        TracingUtilities.Trace($"iid: {iid.GetName()}");
        return CustomQueryInterfaceResult.NotHandled;
    }

    protected unsafe virtual CAUUID GetPages()
    {
        var pages = new CAUUID();
        var iids = PropertyPagesIds.ToArray();
        pages.cElems = (uint)iids.Length;
        var guids = (Guid*)Marshal.AllocCoTaskMem(sizeof(Guid) * (int)pages.cElems);
        for (var i = 0; i < iids.Length; i++)
        {
            TracingUtilities.Trace($"add: {iids[i]}");
            guids[i] = iids[i];
        }
        pages.pElems = (nint)guids;
        return pages;
    }

    protected IComObject<ITypeInfo>? GetTypeInfo()
    {
        EnsureTypeInfoLoaded();
        return _typeInfo;
    }

    protected IComObject<ITypeInfo>? GetDispatchInterfaceInfo()
    {
        EnsureTypeInfoLoaded();
        return _dispatchInterfaceInfo;
    }

    protected virtual void EnsureTypeInfoLoaded()
    {
        if (_typeInfosLoaded)
            return;

        var reg = ComRegistration ?? throw new InvalidOperationException("ComRegistration is not set");
        using var typeLib = TypeLib.LoadTypeLib(reg.DllPath, throwOnError: false);
        TracingUtilities.Trace($"Loaded type lib for type : {GetType().GUID:B} from: {reg.DllPath} typeLib: {typeLib}");
        if (typeLib != null)
        {
            var hr = typeLib.Object.GetTypeInfoOfGuid(GetType().GUID, out var ti);
            TracingUtilities.Trace($"GetTypeInfoOfGuid: {GetType().GUID:B} hr: {hr}");
            _typeInfo = ti != null ? new ComObject<ITypeInfo>(ti) : null;

            TracingUtilities.Trace($"Loaded type lib for type : {DispatchInterfaceId:B} from: {reg.DllPath} typeLib: {typeLib}");
            hr = typeLib.Object.GetTypeInfoOfGuid(DispatchInterfaceId, out var iti);
            TracingUtilities.Trace($"GetTypeInfoOfGuid: {DispatchInterfaceId:B} hr: {hr}");
            _dispatchInterfaceInfo = iti != null ? new ComObject<ITypeInfo>(iti) : null;
        }
        _typeInfosLoaded = true;
    }

    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // dispose managed state (managed objects)
                _typeInfo?.Dispose();
                _dispatchInterfaceInfo?.Dispose();
            }

            // free unmanaged resources (unmanaged objects) and override finalizer
            // set large fields to null
            _disposedValue = true;
        }
    }

    protected virtual DispatchType GetDispatchType()
    {
        var type = GetType();
        if (!_cache.TryGetValue(type, out var dispatchType))
        {
            // we can load dispids from type info if they are declared in idl (Standard dispatch ID constants in olectl.h, like DISPID_HWND, etc.)
            // or build them using reflection
            dispatchType = CreateType(type) ?? throw new InvalidOperationException();

            EnsureTypeInfoLoaded();
            var ti = GetTypeInfo();
            if (ti != null)
            {
                // add dispids for methods and properties declared in the IDL
                dispatchType.AddTypeInfoDispids(ti.Object);
            }

            // add dispids for public methods and properties using reflection (not already added by type info).
            // AutoDispidsBase is the base value.
            dispatchType.AddReflectionDispids(AutoDispidsBase);

            _cache[type] = dispatchType;
        }
        return dispatchType;
    }

    protected virtual HRESULT GetIDsOfNames(in Guid riid, PWSTR[] rgszNames, uint cNames, uint lcid, int[] rgDispId) => TracingUtilities.WrapErrors(() =>
    {
        if (rgszNames == null || rgszNames.Length == 0 || rgszNames.Length != cNames)
            return Constants.E_INVALIDARG;

        for (var i = 0; i < cNames; i++)
        {
            var name = rgszNames[i].ToString();
            if (name == null)
            {
                rgDispId[i] = -1;
                continue;
            }

            TracingUtilities.Trace($"name '{name}'");
            var dispatchType = GetDispatchType();
            if (dispatchType.TryGetDispId(name, out var dispId))
            {
                rgDispId[i] = dispId;
            }
            else
            {
                rgDispId[i] = unchecked((int)DISPID.DISPID_UNKNOWN);
            }
            TracingUtilities.Trace($"name '{name}' => {(DISPID)rgDispId[i]} (0x{rgDispId[i]:X8})");
        }

        if (rgDispId.Any(id => id == -1))
            return Constants.DISP_E_UNKNOWNNAME;

        return Constants.S_OK;
    });

    protected virtual void SetResult(object? value, nint pVarResult)
    {
        if (pVarResult == 0)
            return;

        TracingUtilities.Trace($"SetResult value: {value} type: {value?.GetType().FullName}");
        if (value is Variant v)
        {
            v.DetachTo(pVarResult);
            return;
        }

        if (value is VARIANT variant)
        {
            unsafe
            {
                *(VARIANT*)pVarResult = variant;
            }
            return;
        }

        if (value != null)
        {
            var type = value.GetType();
            if (type.IsClass)
            {
                // if it's a com object, we need to pass the wrapper
                var details = StrategyBasedComWrappers.DefaultIUnknownInterfaceDetailsStrategy.GetComExposedTypeDetails(type.TypeHandle);
                if (details != null)
                {
                    // favor IDispatch for late-bound clients
                    var unk = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance<IDispatch>(value);
                    VARENUM varType;
                    if (unk != 0)
                    {
                        varType = VARENUM.VT_DISPATCH;
                    }
                    else
                    {
                        unk = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(value, throwOnError: true);
                        varType = VARENUM.VT_UNKNOWN;
                    }

                    TracingUtilities.Trace($"returning COM object {value} as {(varType == VARENUM.VT_DISPATCH ? "IDispatch" : "IUnknown")}");
                    using var vunk = new Variant(unk, varType);
                    vunk.DetachTo(pVarResult);
                    return;
                }
            }
        }

        using var va = new Variant(value);
        va.DetachTo(pVarResult);
    }

    // we generally don't support out arguments in IDispatch Invoke, except for IDL methods equipped with [out, retval)
    // so in .NET that return nothing or an HRESULT equivalent and that has a last parameter marked as out
    private static bool IsOutRetval(MethodInfo method)
    {
        if (method.ReturnType != typeof(void) &&
            method.ReturnType != typeof(uint) &&
            method.ReturnType != typeof(int) &&
            method.ReturnType != typeof(HRESULT))
            return false;

        var parameters = method.GetParameters();
        return parameters.Length > 0 && parameters[^1].Attributes.HasFlag(ParameterAttributes.Out);
    }

    protected virtual unsafe HRESULT Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr)
    {
        try
        {
            TracingUtilities.Trace($"dispIdMember: {(DISPID)dispIdMember} (0x{dispIdMember:X8} wFlags: {wFlags} cArgs: {pDispParams.cArgs} pVarResult: {pVarResult} pExcepInfo: {pExcepInfo} puArgErr: {puArgErr}");
            var dispatchType = GetDispatchType();

            var member = dispatchType.GetMemberInfo(dispIdMember);
            // note we can return DISP_E_MEMBERNOTFOUND for a method/property that exists in the TLB but not in the actual type
            if (member == null)
            {
                TracingUtilities.Trace($"dispIdMember: {(DISPID)dispIdMember} was not found => DISP_E_MEMBERNOTFOUND.");
                return Constants.DISP_E_MEMBERNOTFOUND;
            }

            TracingUtilities.Trace($"member: {member.Name} ({member.MemberType})");
            if (member is PropertyInfo property)
            {
                if (wFlags.HasFlag(DISPATCH_FLAGS.DISPATCH_PROPERTYGET))
                {
                    var value = property.GetValue(this);
                    TracingUtilities.Trace($"get value: {value}");
                    SetResult(value, pVarResult);
                    return Constants.S_OK;
                }

                if (wFlags.HasFlag(DISPATCH_FLAGS.DISPATCH_PROPERTYPUT) || wFlags.HasFlag(DISPATCH_FLAGS.DISPATCH_PROPERTYPUTREF))
                {
                    if (pDispParams.cArgs != 1)
                        return Constants.DISP_E_BADPARAMCOUNT;

                    var varArgs = (VARIANT*)pDispParams.rgvarg;
                    var setValue = Variant.Unwrap(varArgs[0]);
                    TracingUtilities.Trace($"set value: {setValue} type: {setValue?.GetType().FullName} property type {property.PropertyType.FullName}");

                    if (setValue != null && !setValue.GetType().IsAssignableTo(property.PropertyType))
                    {
                        if (Conversions.TryChangeObjectType(setValue, property.PropertyType, out var converted))
                        {
                            setValue = converted;
                            TracingUtilities.Trace($"set converted value: {setValue} ({setValue?.GetType().FullName})");
                        }
                    }

                    property.SetValue(this, setValue);
                    return Constants.S_OK;
                }
            }
            else if (member is MethodInfo method)
            {
                if (wFlags.HasFlag(DISPATCH_FLAGS.DISPATCH_METHOD))
                {
                    var variantsToDispose = new List<Variant>();
                    try
                    {
                        var isOutRevalMethod = IsOutRetval(method);
                        var lastParameterIsOutRetval = false;
                        var arguments = method.GetParameters();

                        if (arguments.Length != pDispParams.cArgs)
                        {
                            // [out, reval] parameters handling
                            lastParameterIsOutRetval = arguments.Length == (pDispParams.cArgs + 1) && isOutRevalMethod;
                            if (!lastParameterIsOutRetval)
                            {
                                TracingUtilities.Trace($"member: {member.Name} expected parameters count {arguments.Length}, provided {pDispParams.cArgs} : => DISP_E_BADPARAMCOUNT.");
                                return Constants.DISP_E_BADPARAMCOUNT;
                            }
                        }

                        var methodArgs = method.GetParameters();
                        TracingUtilities.Trace($"methodArgs.Length: {methodArgs.Length} pDispParams.cArgs: {pDispParams.cArgs} isOutRetval: {isOutRevalMethod} lastParameterIsOutRetval: {lastParameterIsOutRetval}");
                        var varArgs = (VARIANT*)pDispParams.rgvarg;
                        var args = new List<object?>();
                        if (pDispParams.cArgs > 0)
                        {
                            for (var i = 0; i < pDispParams.cArgs; i++)
                            {
                                // note arguments are stored in in reverse order
                                var index = (int)(pDispParams.cArgs - i - 1);
                                var varArg = varArgs[index];

                                // no need to unwrap variant if the method expects one
                                if (methodArgs.Length > i && methodArgs[i].ParameterType == typeof(VARIANT))
                                {
                                    args.Add(varArg);
                                    TracingUtilities.Trace($"arg {i}/{pDispParams.cArgs} is VARIANT vt: {varArg.Anonymous.Anonymous.vt}");
                                }
                                else if (methodArgs.Length > i && methodArgs[i].ParameterType == typeof(Variant))
                                {
                                    var v = new Variant(varArg);
                                    variantsToDispose.Add(v);
                                    args.Add(v);
                                    TracingUtilities.Trace($"arg {i}/{pDispParams.cArgs} is VARIANT vt: {varArg.Anonymous.Anonymous.vt}");
                                }
                                else
                                {
                                    var value = Variant.Unwrap(varArg);
                                    args.Add(value);
                                    TracingUtilities.Trace($"arg {i}/{pDispParams.cArgs} value: {value} type: {value?.GetType().FullName}");
                                }
                            }
                        }

                        if (lastParameterIsOutRetval)
                        {
                            args.Add(null);
                        }

                        var array = args.ToArray();
                        var result = method.Invoke(this, array);
                        if (result is Task task)
                        {
                            // we need to run message loop, on the UI thread, until the task is completed
                            const uint WM_COMPLETED = MessageDecoder.WM_APP + 1234;
                            var windowHandle = GetWindowHandle();
                            if (windowHandle == HWND.Null)
                                throw new InvalidOperationException("Cannot invoke async method when there is no window handle");

                            // note this code avoids using reflection on Task<T> to avoid trimming issues with AOT publishing
                            var completed = false;
                            var awaiter = task.GetAwaiter();
                            awaiter.OnCompleted(() =>
                            {
                                completed = true;
                                Functions.PostMessageW(windowHandle, WM_COMPLETED, 0, 0);
                            });

                            // note if the window enters a modal loop (like moving it using the caption bar),
                            // the invoke call will not return until the modal loop is exited. Not sure how to avoid this...
                            while (!completed)
                            {
                                RunMessageLoop(msg => msg.message == WM_COMPLETED);
                            }

                            result = GetTaskResult(task);
                        }

                        if (lastParameterIsOutRetval)
                        {
                            SetResult(array[^1], pVarResult);
                        }

                        if (result is HRESULT hr)
                            return hr;

                        SetResult(result, pVarResult);
                        return Constants.S_OK;
                    }
                    finally
                    {
                        foreach (var v in variantsToDispose)
                        {
                            v.Dispose();
                        }
                    }
                }
            }
            throw new InvalidOperationException();
        }
        catch (Exception ex)
        {
            TracingUtilities.Trace($"Error: {ex}");
            if (pExcepInfo != 0)
            {
                var excepInfo = new EXCEPINFO
                {
                    scode = unchecked((int)Constants.E_FAIL),
                    bstrDescription = new Bstr(ex.GetInterestingExceptionMessage()),
                    bstrSource = new Bstr(GetType().FullName),
                };

                *(EXCEPINFO*)pExcepInfo = excepInfo;
            }
            TracingUtilities.SetError(ex.GetInterestingExceptionMessage());
            return Constants.DISP_E_EXCEPTION;
        }
    }

    protected virtual HRESULT GetTypeInfo(uint iTInfo, uint lcid, out ITypeInfo ppTInfo)
    {
        ITypeInfo? ti = null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            TracingUtilities.Trace($"iTInfo {iTInfo}");
            if (iTInfo != 0)
                return Constants.DISP_E_BADINDEX;

            EnsureTypeInfoLoaded();
            var tio = GetDispatchInterfaceInfo();
            if (tio != null)
            {
                ti = tio.Object;
                unsafe
                {
                    ti.GetTypeAttr(out var ta).ThrowOnError();
                    var atts = (TYPEATTR*)ta;
                    TracingUtilities.Trace($" funcs: {atts->cFuncs} vars: {atts->cVars}");
                }
                return Constants.S_OK;
            }
            return Constants.E_NOTIMPL;
        });
        ppTInfo = ti!;
        return hr;
    }

    protected virtual HRESULT GetTypeInfoCount(out uint pctinfo)
    {
        var count = 0u;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            EnsureTypeInfoLoaded();
            var ti = GetTypeInfo();
            count = ti != null ? 1u : 0u;
            TracingUtilities.Trace($"pctinfo {count}");
            return ti != null ? Constants.S_OK : Constants.E_NOTIMPL;
        });
        pctinfo = count;
        return hr;
    }

    protected virtual void RunMessageLoop(Func<MSG, bool> exitLoopFunc)
    {
        ArgumentNullException.ThrowIfNull(exitLoopFunc);
        do
        {
            if (Functions.PeekMessageW(out var msg, HWND.Null, 0, 0, PEEK_MESSAGE_REMOVE_TYPE.PM_REMOVE))
            {
                if (exitLoopFunc(msg))
                    return;

                if (msg.message == MessageDecoder.WM_QUIT)
                {
                    // repost
                    Functions.PostQuitMessage(0);
                    return;
                }

                Functions.TranslateMessage(msg);
                Functions.DispatchMessageW(msg);
            }
        } while (true);
    }

    protected virtual PROPCAT MapPropertyToCategory(int dispid)
    {
        var type = GetDispatchType();
        return type.GetMember((int)dispid)?.Category.Category ?? PROPCAT.PROPCAT_Misc;
    }

    protected virtual string GetCategoryName(PROPCAT propcat, uint lcid)
    {
        var type = GetDispatchType();
        var category = type.GetCategory(propcat);
        return category.GetLocalizedName(lcid);
    }

    protected virtual string? GetDisplayString(int dispId)
    {
        var type = GetDispatchType();
        var member = type.GetMember(dispId);
        return member?.DefaultString;
    }

    protected virtual Guid? MapPropertyToPage(int dispId)
    {
        if ((DISPID)dispId == DISPID.MEMBERID_NIL)
            return PropertyPageId;

        var type = GetDispatchType();
        var member = type.GetMember(dispId);
        return member?.PropertyPageId;
    }

    HRESULT IProvideClassInfo.GetClassInfo(out ITypeInfo ppTI)
    {
        ITypeInfo? pti = null;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            EnsureTypeInfoLoaded();
            var ti = GetTypeInfo();
            pti = ti?.Object;
            TracingUtilities.Trace($"ppTI: {pti}");
            return pti != null ? Constants.S_OK : Constants.E_UNEXPECTED;
        });
        ppTI = pti!;
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

    HRESULT IVsPerPropertyBrowsing.HideProperty(int dispId, out BOOL pfHide)
    {
        TracingUtilities.Trace($"dispId: {dispId}");
        // we already filter properties in the property browser using Browsable
        pfHide = false;
        return Constants.S_OK;
    }

    HRESULT IVsPerPropertyBrowsing.DisplayChildProperties(int dispId, out BOOL pfDisplay)
    {
        TracingUtilities.Trace($"dispId: {dispId}");
        pfDisplay = true;
        return Constants.S_OK;
    }

    HRESULT IVsPerPropertyBrowsing.GetLocalizedPropertyInfo(int dispId, uint localeID, out BSTR pbstrLocalizedName, out BSTR pbstrLocalizeDescription)
    {
        TracingUtilities.Trace($"dispId: {dispId} localeID: {localeID}");
        pbstrLocalizedName = BSTR.Null;
        pbstrLocalizeDescription = BSTR.Null;
        return Constants.S_OK;
    }

    HRESULT IVsPerPropertyBrowsing.HasDefaultValue(int dispId, out BOOL fDefault)
    {
        TracingUtilities.Trace($"dispId: {dispId}");
        var type = GetDispatchType();
        var member = type.GetMember(dispId);
        fDefault = member?.DefaultString != null;
        return Constants.S_OK;
    }

    HRESULT IVsPerPropertyBrowsing.IsPropertyReadOnly(int dispId, out BOOL fReadOnly)
    {
        TracingUtilities.Trace($"dispId: {dispId}");
        var type = GetDispatchType();
        var member = type.GetMember(dispId);
        fReadOnly = member?.IsReadOnly == true;
        return Constants.S_OK;
    }

    HRESULT IVsPerPropertyBrowsing.GetClassName(out BSTR pbstrClassName)
    {
        pbstrClassName = new Bstr(GetType().FullName);
        TracingUtilities.Trace($"class name: '{pbstrClassName}'");
        return Constants.S_OK;
    }

    HRESULT IVsPerPropertyBrowsing.CanResetPropertyValue(int dispId, out BOOL pfCanReset)
    {
        TracingUtilities.Trace($"dispId: {dispId}");
        var type = GetDispatchType();
        var member = type.GetMember(dispId);
        pfCanReset = member?.DefaultString != null;
        return Constants.S_OK;
    }

    HRESULT IVsPerPropertyBrowsing.ResetPropertyValue(int dispId) => TracingUtilities.WrapErrors(() =>
    {
        TracingUtilities.Trace($"dispId: {dispId}");
        var type = GetDispatchType();
        var member = type.GetMember(dispId);
        if (member == null)
            return Constants.E_NOTIMPL;

        var prop = member.Info as PropertyInfo;
        if (prop == null)
            return Constants.E_NOTIMPL;

        var canReset = member.DefaultString != null;
        if (!canReset)
            return Constants.E_NOTIMPL;

        var value = member.GetDefaultValue();
        prop.SetValue(this, value);
        return Constants.S_OK;
    });

    unsafe HRESULT IVSMDPerPropertyBrowsing.GetPropertyAttributes(int dispId, out uint pceltAttrs, out nint ppbstrTypeNames, out nint ppvarAttrValues)
    {
        TracingUtilities.Trace($"dispId: {dispId}");
        pceltAttrs = 0;
        ppbstrTypeNames = 0;
        ppvarAttrValues = 0;

        // attributes for the type (we could set things like ReadOnly, etc.)
        if ((DISPID)dispId == DISPID.MEMBERID_NIL)
            return Constants.S_OK;

        var type = GetDispatchType();
        var category = type.GetMember(dispId)?.Category;
        if (category != null)
        {
            var name = category.GetLocalizedName((uint)Thread.CurrentThread.CurrentUICulture.LCID).Nullify();
            if (name != null)
            {
                var bstrs = (BSTR*)Marshal.AllocCoTaskMem(sizeof(BSTR));
                bstrs[0] = new BSTR(Marshal.StringToBSTR(typeof(CategoryAttribute).FullName));
                ppbstrTypeNames = (nint)bstrs;

                var vars = (VARIANT*)Marshal.AllocCoTaskMem(sizeof(VARIANT));
                vars[0] = new Variant(name).Detach();
                ppvarAttrValues = (nint)vars;

                pceltAttrs = 1;
            }
        }
        return Constants.S_OK;
    }

    HRESULT IProvidePropertyBuilder.MapPropertyToBuilder(int dispId, out CTLBLDTYPE pdwCtlBldType, out BSTR pbstrGuidBldr, out VARIANT_BOOL builderAvailable)
    {
        TracingUtilities.Trace($"dispId: {dispId}");
        pdwCtlBldType = (CTLBLDTYPE)0;
        pbstrGuidBldr = BSTR.Null;
        builderAvailable = false;
        return Constants.E_NOTIMPL;
    }

    HRESULT IProvidePropertyBuilder.ExecuteBuilder(int dispId, BSTR bstrGuidBldr, IDispatch pdispApp, HWND hwndBldrOwner, in VARIANT pvarValue, out VARIANT_BOOL pbActionCommitted)
    {
        TracingUtilities.Trace($"dispId: {dispId} bstrGuidBldr: {bstrGuidBldr} pdispApp: {pdispApp} hwndBldrOwner: {hwndBldrOwner} pvarValue: {pvarValue.Anonymous.Anonymous.vt}");
        pbActionCommitted = false;
        return Constants.E_NOTIMPL;
    }

    public class PredefinedString(uint id, string name)
    {
        public uint Id { get; } = id;
        public string Name { get; } = name;
        public virtual object? Value { get; set; }

        public override string ToString() => $"{Id} '{Name}': {Value}";
    }
}