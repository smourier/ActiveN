namespace ActiveN.Hosting;

// this implements something similar to this https://learn.microsoft.com/en-us/windows/win32/com/aggregation
// note: due to GeneratedComInterface, AOT and trimming limitations,
// we cannot use reflection to build vtables at runtime, dynamically, etc.
public unsafe class Aggregable
{
    // note we'll loose that memory until process exit
    private static readonly ConcurrentDictionary<string, nint> _aggregatedVTables = new();

    private static int? _maximumNumberOfMethodsPerInterface;
    public static int MaximumNumberOfMethodsPerInterface
    {
        // 255 should be enough for everyone (inludes IUnknown 3 methods)
        // testing found DirectN.ID3D11DeviceContext4 wich has 146!
        get => _maximumNumberOfMethodsPerInterface ?? 255;
        set
        {
            // can only be set once, before first use
            if (_maximumNumberOfMethodsPerInterface.HasValue)
                throw new InvalidOperationException();

            _maximumNumberOfMethodsPerInterface = value;
        }
    }

    public static nint Aggregate(nint outer, IAggregable innerAggregable)
    {
        ArgumentNullException.ThrowIfNull(innerAggregable);
        if (!innerAggregable.SupportsAggregation)
            throw new InvalidOperationException("inner does not support aggregation");

        var inner = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(innerAggregable);
        if (inner == 0)
            throw new InvalidOperationException();

        // build a new vtable
        var headerSize = sizeof(AggregatedClass);
        var vtbl = (AggregatedClass*)Marshal.AllocCoTaskMem(headerSize);
        Unsafe.InitBlockUnaligned(vtbl, 0, (uint)headerSize);

        vtbl->IUnknownVtbl.QueryInterface = &IUnknownQueryInterface;
        vtbl->IUnknownVtbl.AddRef = &IUnknownAddRef;
        vtbl->IUnknownVtbl.Release = &IUnknownRelease;
        vtbl->lpVtbl = &vtbl->IUnknownVtbl;
        vtbl->refCount = 1;
        vtbl->inner = inner;
        vtbl->outer = outer;
        vtbl->innerTypeHandle = innerAggregable.GetType().TypeHandle.Value;
        return (nint)vtbl;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static HRESULT IUnknownQueryInterface(nint thisPtr, Guid* riid, nint* ppv)
    {
        TracingUtilities.Trace($"this: 0x{thisPtr:X} riid: {(riid != null ? riid->GetName() : null)}");
        if (riid == null || ppv == null)
            return Constants.E_POINTER;
        *ppv = 0;

        var iid = *riid;
        var cls = (AggregatedClass*)thisPtr;
        if (iid == typeof(IUnknown).GUID)
        {
            *ppv = thisPtr;
            Interlocked.Increment(ref cls->refCount);
            return Constants.S_OK;
        }

        HRESULT hr = Marshal.QueryInterface(cls->inner, iid, out var ifaceUnk);
        try
        {
            if (hr.IsError || ifaceUnk == 0)
            {
                TracingUtilities.Trace($"inner QI failed: hr: {hr} ppv: 0x{ifaceUnk:X}");
                return Constants.E_NOINTERFACE;
            }

            // build or reuse a vtable for that interface and that type
            var key = $"{iid}.{cls->innerTypeHandle}"; // cache key
            if (!_aggregatedVTables.TryGetValue(key, out var ptr))
            {
                // allocate the size of the derived vtable (outer unknown + derived vtable) with max methods
                var rt = RuntimeTypeHandle.FromIntPtr(cls->innerTypeHandle);
                var type = Type.GetTypeFromHandle(rt);
                if (type == null)
                    return Constants.E_UNEXPECTED;

                // unfortunately we cannot get the actual number of methods in the interface
                var maxMethods = MaximumNumberOfMethodsPerInterface;
                var outerSize = sizeof(OuterClass) - sizeof(IUnknownVTable) + maxMethods * sizeof(nint);
                var vtbl = (OuterClass*)RuntimeHelpers.AllocateTypeAssociatedMemory(type, outerSize);
                Unsafe.InitBlockUnaligned(vtbl, 0, (uint)outerSize);
                vtbl->IDerivedUnknownVtbl.QueryInterface = &OuterQueryInterface;
                vtbl->IDerivedUnknownVtbl.AddRef = &OuterAddRef;
                vtbl->IDerivedUnknownVtbl.Release = &OuterRelease;
                vtbl->lpVtbl = &vtbl->IDerivedUnknownVtbl;
                vtbl->outer = cls->outer;

                // copy methods from inner vtable (after the 3 IUnknown methods)
                var innerVTable = *(nint**)ifaceUnk;
                for (var i = 3; i < maxMethods; i++)
                {
                    var p = innerVTable[i];
                    TracingUtilities.Trace($" copy inner vtbl method[{i}]: 0x{p:X}");
                    if (p == 0) // hopefully this means the end of the vtable. experience shows it seems ok
                        break;

                    ((nint*)&vtbl->IDerivedUnknownVtbl)[i] = p;
                }

                ptr = (nint)vtbl;
                _aggregatedVTables[key] = ptr;
            }

            *ppv = ptr;
            Interlocked.Increment(ref cls->refCount);
        }
        finally
        {
            if (ifaceUnk != 0)
            {
                Marshal.Release(ifaceUnk);
            }
        }

        TracingUtilities.Trace($"this: 0x{thisPtr:X} ppv: 0x{*ppv:X}");
        return Constants.S_OK;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint IUnknownAddRef(nint thisPtr)
    {
        var unk = (AggregatedClass*)thisPtr;
        var ui = (uint)Interlocked.Increment(ref unk->refCount);
        TracingUtilities.Trace($"this: 0x{thisPtr:X} refCount: {-1 + unk->refCount} => {unk->refCount}");
        return ui;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint IUnknownRelease(nint thisPtr)
    {
        var unk = (AggregatedClass*)thisPtr;
        var c = Interlocked.Decrement(ref unk->refCount);
        TracingUtilities.Trace($"this: 0x{thisPtr:X} refCount: {1 + unk->refCount} => {unk->refCount}");
        if (c == 0)
        {
            Marshal.FreeCoTaskMem(thisPtr);
        }
        return (uint)c;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static HRESULT OuterQueryInterface(nint thisPtr, Guid* riid, nint* ppv)
    {
        TracingUtilities.Trace($"this: 0x{thisPtr:X} riid: {(riid != null ? riid->GetName() : null)}");
        if (riid == null || ppv == null)
            return Constants.E_POINTER;

        var unk = (OuterClass*)thisPtr;
        HRESULT hr = Marshal.QueryInterface(unk->outer, *riid, out var ppvo);
        if (hr.IsSuccess && ppv != null)
        {
            *ppv = ppvo;
        }
        TracingUtilities.Trace($"outer  0x{thisPtr:X} ppv: 0x{ppvo:X} hr: {hr}");
        return hr;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint OuterAddRef(nint thisPtr)
    {
        var unk = (OuterClass*)thisPtr;
        var ui = Marshal.AddRef(unk->outer);
        TracingUtilities.Trace($"outer 0x{thisPtr:X} => {ui}");
        return (uint)ui;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint OuterRelease(nint thisPtr)
    {
        var unk = (OuterClass*)thisPtr;
        var c = Marshal.Release(unk->outer);
        TracingUtilities.Trace($"outer 0x{thisPtr:X} => {c}");
        return (uint)c;
    }

    private struct AggregatedClass
    {
        public IUnknownVTable* lpVtbl; // *must* be first: pointer to vtable
        public IUnknownVTable IUnknownVtbl; // embedded IUnknown-only vtable
        public nint inner;
        public nint outer;
        public nint innerTypeHandle;
        public int refCount;
    }

    private struct OuterClass
    {
        public IUnknownVTable* lpVtbl; // *must* be first: pointer to vtable
        public nint outer;
        // *warning*: from here, we have a variable size, up to MaximumNumberOfMethodsPerInterface (including 3 IUnknown methods)
        public IUnknownVTable IDerivedUnknownVtbl; // IUnknown-derived vtable starts here
        // so *don't* add any ther field here
    }

    private struct IUnknownVTable
    {
        private void* _queryInterface;
        private void* _addRef;
        private void* _release;

        public delegate* unmanaged[Stdcall]<nint, Guid*, nint*, HRESULT> QueryInterface { readonly get => (delegate* unmanaged[Stdcall]<nint, Guid*, nint*, HRESULT>)_queryInterface; set => _queryInterface = value; }
        public delegate* unmanaged[Stdcall]<nint, uint> AddRef { readonly get => (delegate* unmanaged[Stdcall]<nint, uint>)_addRef; set => _addRef = value; }
        public delegate* unmanaged[Stdcall]<nint, uint> Release { readonly get => (delegate* unmanaged[Stdcall]<nint, uint>)_release; set => _release = value; }
    }
}
