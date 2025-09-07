namespace ActiveN;

[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
public abstract partial class BaseDispatch : IDisposable, ICustomQueryInterface
{
    private static readonly ConcurrentDictionary<Type, DispatchType> _cache = new();

    private ComObject<ITypeInfo>? _typeInfo;
    private bool _typeInfoLoaded;
    private bool _disposedValue;
    public bool IsDisposed => _disposedValue;

    protected virtual object? GetTaskResult(Task task) => null;
    protected virtual HWND GetWindowHandle() => HWND.Null;
    protected abstract ComRegistration ComRegistration { get; }

    protected virtual int AutoDispidsBase => 0x10000;

    protected virtual DispatchType CreateType([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)] Type type) => new(type);

    CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out nint ppv) => GetInterface(ref iid, out ppv);
    protected virtual CustomQueryInterfaceResult GetInterface(ref Guid iid, out nint ppv)
    {
        ppv = 0;
        TracingUtilities.Trace($"iid: {iid.GetName()}");
        return CustomQueryInterfaceResult.NotHandled;
    }

    protected virtual ComObject<ITypeInfo>? EnsureTypeInfo()
    {
        if (!_typeInfoLoaded)
        {
            var reg = ComRegistration ?? throw new InvalidOperationException("ComRegistration is not set");
            using var typeLib = TypeLib.LoadTypeLib(reg.DllPath, throwOnError: false);
            TracingUtilities.Trace($"Loaded type lib for type : {GetType().GUID:B} from: {reg.DllPath} typeLib: {typeLib}");
            if (typeLib != null)
            {
                var hr = typeLib.Object.GetTypeInfoOfGuid(GetType().GUID, out var ti);
                TracingUtilities.Trace($"GetTypeInfoOfGuid: {GetType().GUID:B} hr: {hr}");
                _typeInfo = ti != null ? new ComObject<ITypeInfo>(ti) : null;
            }
            _typeInfoLoaded = true;
        }
        return _typeInfo;
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

            var ti = EnsureTypeInfo();
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
                rgDispId[i] = -1;
            }
            TracingUtilities.Trace($"name '{name}' => {rgDispId[i]} (0x{rgDispId[i]:X8})");
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
            TracingUtilities.Trace($"dispIdMember: {dispIdMember} (0x{dispIdMember:X8}) wFlags: {wFlags} cArgs: {pDispParams.cArgs} pVarResult: {pVarResult} pExcepInfo: {pExcepInfo} puArgErr: {puArgErr}");
            var dispatchType = GetDispatchType();

            var member = dispatchType.GetMemberInfo(dispIdMember);
            // note we can return DISP_E_MEMBERNOTFOUND for a method/property that exists in the TLB but not in the actual type
            if (member == null)
            {
                TracingUtilities.Trace($"dispIdMember: {dispIdMember} was not found => DISP_E_MEMBERNOTFOUND.");
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

            var tio = EnsureTypeInfo();
            if (tio != null)
            {
                ti = tio.Object;
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
            var ti = EnsureTypeInfo();
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
}