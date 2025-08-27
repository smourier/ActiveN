namespace ActiveN;

[GeneratedComClass]
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
public abstract partial class BaseDispatch : IDispatch, IDisposable
{
    private static readonly ConcurrentDictionary<Type, DispatchType> _cache = new();

    private ComObject<ITypeInfo>? _typeInfo;
    private bool _typeInfoLoaded;
    private bool _disposedValue;

    protected abstract object? GetTaskResult(Task task);
    protected abstract HWND GetWindowHandle();
    protected abstract ComRegistration ComRegistration { get; }
    protected virtual int AutoDispidsBase => 0x10000;

    protected virtual ComObject<ITypeInfo>? EnsureTypeInfo()
    {
        if (!_typeInfoLoaded)
        {
            var reg = ComRegistration ?? throw new InvalidOperationException("ComRegistration is not set");
            using var typeLib = TypeLib.LoadTypeLib(reg.DllPath, throwOnError: false);
            if (typeLib != null)
            {
                var hr = typeLib.Object.GetTypeInfoOfGuid(GetType().GUID, out var ti);
                ComRegistration.Trace($"GetTypeInfoOfGuid: {GetType().GUID:B} hr:{hr}");
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

    // // override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
    // ~BaseControl()
    // {
    //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
    //     Dispose(disposing: false);
    // }

    HRESULT IDispatch.GetIDsOfNames(in Guid riid, PWSTR[] rgszNames, uint cNames, uint lcid, int[] rgDispId) =>
        GetIDsOfNames(in riid, rgszNames, cNames, lcid, rgDispId);

    protected virtual HRESULT GetIDsOfNames(in Guid riid, PWSTR[] rgszNames, uint cNames, uint lcid, int[] rgDispId)
    {
        if (rgszNames == null || rgszNames.Length == 0 || rgszNames.Length != cNames)
            return Constants.E_INVALIDARG;

        var type = GetType();
        for (var i = 0; i < cNames; i++)
        {
            var name = rgszNames[i].ToString();
            if (name == null)
            {
                rgDispId[i] = -1;
                continue;
            }

            ComRegistration.Trace($"name '{name}'");

            if (!_cache.TryGetValue(type, out var dispatchType))
            {
                // we can load dispids from type info if they are declared in idl (Standard dispatch ID constants in olectl.h, like DISPID_HWND, etc.)
                // or build them using reflection
                dispatchType = new DispatchType(type);

                var ti = EnsureTypeInfo();
                if (ti != null)
                {
                    // add dispids for methods and properties declared in the IDL
                    dispatchType.AddTypeInfoDispids(ti.Object);
                }

                // add dispids for public methods and properties using reflection (not already added by type info).
                // AutoDispidsBase is the base value.
                dispatchType.AddAutoDispids(AutoDispidsBase);

                _cache[type] = dispatchType;
            }

            if (dispatchType.TryGetDispId(name, out var dispId))
            {
                rgDispId[i] = dispId;
            }
            else
            {
                rgDispId[i] = -1;
            }
            ComRegistration.Trace($"name '{name}' => {rgDispId[i]} (0x{rgDispId[i]:X8})");
        }

        if (rgDispId.Any(id => id == -1))
            return Constants.DISP_E_UNKNOWNNAME;

        return Constants.S_OK;
    }

    HRESULT IDispatch.Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr) =>
        Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

    protected virtual unsafe HRESULT Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr)
    {
        try
        {
            ComRegistration.Trace($"dispIdMember: {dispIdMember} wFlags: {wFlags} cArgs: {pDispParams.cArgs} pVarResult: {pVarResult} pExcepInfo: {pExcepInfo} puArgErr: {puArgErr}");
            var type = GetType();
            // note we can return DISP_E_MEMBERNOTFOUND for a method/property that exists in the TLB but not in the actual type
            if (!_cache.TryGetValue(type, out var dispatchType) || dispatchType.GetMemberInfo(dispIdMember) is not MemberInfo member)
            {
                ComRegistration.Trace($"dispIdMember: {dispIdMember} was not found => DISP_E_MEMBERNOTFOUND.");
                return Constants.DISP_E_MEMBERNOTFOUND;
            }

            ComRegistration.Trace($"member: {member.Name} ({member.MemberType})");
            if (member is PropertyInfo property)
            {
                if (wFlags.HasFlag(DISPATCH_FLAGS.DISPATCH_PROPERTYGET))
                {
                    var value = property.GetValue(this);
                    ComRegistration.Trace($"get value: {value}");
                    if (pVarResult != 0)
                    {
                        using var v = new Variant(value);
                        var detached = v.Detach();
                        *(VARIANT*)pVarResult = detached;
                    }
                    return Constants.S_OK;
                }

                if (wFlags.HasFlag(DISPATCH_FLAGS.DISPATCH_PROPERTYPUT) || wFlags.HasFlag(DISPATCH_FLAGS.DISPATCH_PROPERTYPUTREF))
                {
                    if (pDispParams.cArgs != 1)
                        return Constants.DISP_E_BADPARAMCOUNT;

                    var varArgs = (VARIANT*)pDispParams.rgvarg;
                    var setValue = Variant.Unwrap(varArgs[0]);
                    property.SetValue(this, setValue);
                    return Constants.S_OK;
                }
            }
            else if (member is MethodInfo method)
            {
                if (wFlags.HasFlag(DISPATCH_FLAGS.DISPATCH_METHOD))
                {
                    var arguments = method.GetParameters();
                    var hasOutRetval = false;

                    if (arguments.Length != pDispParams.cArgs)
                    {
                        // out, reval parameters handling
                        hasOutRetval = arguments.Length == (pDispParams.cArgs + 1) && arguments[^1].Attributes.HasFlag(ParameterAttributes.Out);
                        if (!hasOutRetval)
                        {
                            ComRegistration.Trace($"member: {member.Name} expected parameters count {arguments.Length}, provided {pDispParams.cArgs} : => DISP_E_BADPARAMCOUNT.");
                            return Constants.DISP_E_BADPARAMCOUNT;
                        }
                    }

                    var outParameters = new Dictionary<int, int>();
                    var varArgs = (VARIANT*)pDispParams.rgvarg;
                    var args = new List<object?>();
                    if (pDispParams.cArgs > 0)
                    {
                        for (var i = 0; i < pDispParams.cArgs; i++)
                        {
                            // note arguments are stored in in reverse order
                            var index = (int)(pDispParams.cArgs - i - 1);
                            var varArg = varArgs[index];
                            if (varArg.Anonymous.Anonymous.vt.HasFlag(VARENUM.VT_BYREF))
                            {
                                // out parameters
                                args.Add(null);
                                outParameters.Add(i, index);
                            }
                            else
                            {
                                var value = Variant.Unwrap(varArg);
                                args.Add(value);
                            }
                        }
                    }

                    if (hasOutRetval)
                    {
                        args.Add(null);
                    }

                    var array = args.ToArray();
                    VARENUM? resultType = null;
                    var result = method.Invoke(this, array);
                    if (result is Task task)
                    {
                        // we need to run message loop, on the UI thread, until the task is completed
                        const uint WM_COMPLETED = MessageDecoder.WM_APP + 1234;
                        var windowHandle = GetWindowHandle();

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
                    else if (result is HRESULT hr)
                    {
                        hr.ThrowOnError();
                        result = 0;
                        resultType = VARENUM.VT_ERROR;
                    }

                    // fill out parameters the best we can
                    foreach (var kv in outParameters)
                    {
                        var outParamValue = array[kv.Key];
                        using var v = new Variant(outParamValue);
                        var final = v;

                        var va = varArgs[kv.Value];
                        var requiredType = va.Anonymous.Anonymous.vt & ~VARENUM.VT_BYREF;
                        if (requiredType != v.VarType)
                        {
                            // try to convert
                            var newV = v.ChangeType(requiredType);
                            if (newV != null)
                            {
                                final = newV;
                            }
                        }

                        final.DetachToByRef((nint)(&va));
                    }

                    if (hasOutRetval)
                    {
                        result = array[^1]; // out, retval
                    }

                    if (pVarResult != 0)
                    {
                        using var v = new Variant(result, resultType);
                        var detached = v.Detach();
                        *(VARIANT*)pVarResult = detached;
                    }
                    return Constants.S_OK;
                }
            }
            throw new InvalidOperationException();
        }
        catch (Exception ex)
        {
            ComRegistration.Trace($"Exception: {ex}");
            if (pExcepInfo != 0)
            {
                var excepInfo = new EXCEPINFO
                {
                    scode = unchecked((int)Constants.E_FAIL),
                    bstrDescription = new Bstr(ex.GetAllMessages()),
                };

                *(EXCEPINFO*)pExcepInfo = excepInfo;
            }
            return Constants.DISP_E_EXCEPTION;
        }
    }

    HRESULT IDispatch.GetTypeInfo(uint iTInfo, uint lcid, out ITypeInfo ppTInfo)
    {
        ComRegistration.Trace($"iTInfo {iTInfo}");
        ppTInfo = null!;
        var ti = EnsureTypeInfo();
        if (ti != null)
        {
            if (iTInfo != 0)
                return Constants.DISP_E_BADINDEX;

            ppTInfo = ti.Object;
            return Constants.S_OK;
        }
        return Constants.E_NOTIMPL;
    }

    HRESULT IDispatch.GetTypeInfoCount(out uint pctinfo)
    {
        var ti = EnsureTypeInfo();
        pctinfo = ti != null ? 1u : 0u;
        ComRegistration.Trace($"pctinfo {pctinfo}");
        return ti != null ? Constants.S_OK : Constants.E_NOTIMPL;
    }

    private static void RunMessageLoop(Func<MSG, bool> exitLoopFunc)
    {
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

    private sealed class DispatchType([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)] Type type)
    {
        private readonly Dictionary<string, DispatchMember> _membersByName = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<int, DispatchMember> _memberByDispIds = [];

        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicMethods)]
        public Type Type { get; } = type;

        public unsafe void AddTypeInfoDispids(ITypeInfo typeInfo)
        {
            var typeAttr = TypeLib.GetAttributes(typeInfo);
            if (!typeAttr.HasValue)
                return;

            for (uint i = 0; i < typeAttr.Value.cImplTypes; i++)
            {
                if (typeInfo.GetRefTypeOfImplType(i, out var href).IsError)
                    continue;

                if (typeInfo.GetRefTypeInfo(href, out var obj).IsError || obj == null)
                    continue;

                var typeName = TypeLib.GetName(obj, -1);
                ComRegistration.Trace($"type: '{typeName}'");

                using var refTypeInfo = new ComObject<ITypeInfo>(obj);
                var refTypeAttr = TypeLib.GetAttributes(refTypeInfo.Object);
                if (!refTypeAttr.HasValue)
                    continue;

                for (uint f = 0; f < refTypeAttr.Value.cFuncs; f++)
                {
                    var funcDesc = TypeLib.GetFuncDesc(refTypeInfo.Object, f);
                    if (!funcDesc.HasValue)
                        continue;

                    // skip restricted (QueryInterface, AddRef, Invoke, etc.)
                    if (funcDesc.Value.wFuncFlags.HasFlag(FUNCFLAGS.FUNCFLAG_FRESTRICTED))
                        continue;

                    var name = TypeLib.GetName(refTypeInfo.Object, funcDesc.Value.memid);
                    ComRegistration.Trace($"funcDesc: id:{funcDesc.Value.memid} name:'{name}' kind:{funcDesc.Value.funckind} invkind:{funcDesc.Value.invkind} params:{funcDesc.Value.cParams} paramsOpt:{funcDesc.Value.cParamsOpt} flags:{funcDesc.Value.wFuncFlags}");
                    if (name == null)
                        continue;

                    // if null, means the method/property is in the TLB but not in the actual type
                    var memberInfo = (MemberInfo?)Type.GetProperty(name, BindingFlags.Public | BindingFlags.Instance)
                        ?? Type.GetMethod(name, BindingFlags.Public | BindingFlags.Instance);

                    var member = new DispatchMember(funcDesc.Value.memid, memberInfo);
                    _membersByName[memberInfo?.Name ?? name] = member;
                    _memberByDispIds[member.DispId] = member;
                }
            }
        }

        public void AddAutoDispids(int autoDispidsBase)
        {
            // note we don't support overloaded methods & properties
            // add only members not already added by type info
            var methods = Type.GetMethods(BindingFlags.Public | BindingFlags.Instance);
            for (var i = 0; i < methods.Length; i++)
            {
                var method = methods[i];
                if (_membersByName.ContainsKey(method.Name))
                    continue;

                var member = new DispatchMember(autoDispidsBase + _membersByName.Count, method);
                _membersByName[method.Name] = member;
                _memberByDispIds[member.DispId] = member;
            }

            var properties = Type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            for (var i = 0; i < properties.Length; i++)
            {
                var property = properties[i];
                if (_membersByName.ContainsKey(property.Name))
                    continue;

                var member = new DispatchMember(autoDispidsBase + _membersByName.Count, property);
                _membersByName[property.Name] = member;
                _memberByDispIds[member.DispId] = member;
            }

#if DEBUG
            foreach (var kv in _memberByDispIds)
            {
                ComRegistration.Trace($"dispid: {kv.Key} => {kv.Value}");
            }

            foreach (var kv in _membersByName)
            {
                ComRegistration.Trace($"name: '{kv.Key}' => {kv.Value}");
            }
#endif
        }

        public MemberInfo? GetMemberInfo(int dispId)
        {
            if (!_memberByDispIds.TryGetValue(dispId, out var member))
                return null;

            return member.Info;
        }

        public bool TryGetDispId(string name, out int dispId)
        {
            dispId = 0;
            if (!_membersByName.TryGetValue(name, out var member))
                return false;

            dispId = member.DispId;
            return true;
        }
    }

    private sealed class DispatchMember(int dispId, MemberInfo? info)
    {
        public int DispId { get; } = dispId;
        public MemberInfo? Info { get; } = info; // if null, means the method/property is in the TLB but not in the actual type

        public override string ToString() => $"{DispId}: {Info?.Name} ({Info?.MemberType})";
    }
}