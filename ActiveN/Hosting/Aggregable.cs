namespace ActiveN.Hosting;

// this implements something similar to this https://learn.microsoft.com/en-us/windows/win32/com/aggregation
// but we cannot really implement code as described in doc above
// because .NET source generated ComWrappers seem to store private thing in the *this* pointer (see ComWrappers, ComInterfaceDispatch, ComInterfaceEntry cryptic code)
// if we do that, it crashes in QueryInterface in various ways (AV, stack overflow, etc.)
// hopefully that should not be needed for Release/AddRef, and QueryInterface is now handled here by OuterQueryInterface
// note: due to GeneratedComInterface, AOT and trimming limitations, we cannot use reflection to build vtables at runtime, dynamically, etc.
public unsafe class Aggregable
{
    private static readonly ConcurrentDictionary<WrapperGuid, object?> _tempIidStack = new();

    public static nint Aggregate(nint outer, IAggregable innerAggregable)
    {
        ArgumentNullException.ThrowIfNull(innerAggregable);
        if (!innerAggregable.SupportsAggregation)
            throw new InvalidOperationException("inner does not support aggregation");

        var inner = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(innerAggregable);
        if (inner == 0)
            throw new InvalidOperationException();

        // build a "wrapper" new vtable
        var headerSize = sizeof(Aggregated);
        var vtbl = (Aggregated*)Marshal.AllocCoTaskMem(headerSize);
        Unsafe.InitBlockUnaligned(vtbl, 0, (uint)headerSize);

        vtbl->IUnknownVtbl.QueryInterface = &IUnknownQueryInterface;
        vtbl->IUnknownVtbl.AddRef = &IUnknownAddRef;
        vtbl->IUnknownVtbl.Release = &IUnknownRelease;
        vtbl->lpVtbl = &vtbl->IUnknownVtbl;
        vtbl->refCount = 1;
        vtbl->inner = inner;
        vtbl->outer = outer;

        var interfaceIds = innerAggregable.AggregableInterfaces.Select(i => i.GUID).ToArray();
        vtbl->aggregableInterfaceIidsCount = interfaceIds.Length;

        var aggregableInterfaceIidsFieldSize = sizeof(Guid) * interfaceIds.Length;
        vtbl->aggregableInterfaceIids = (Guid*)Marshal.AllocCoTaskMem(aggregableInterfaceIidsFieldSize);
        Unsafe.InitBlockUnaligned(vtbl->aggregableInterfaceIids, 0, (uint)aggregableInterfaceIidsFieldSize);
        for (int i = 0; i < interfaceIds.Length; i++)
        {
            vtbl->aggregableInterfaceIids[i] = interfaceIds[i];
        }

        return (nint)vtbl;
    }

    public static HRESULT OuterQueryInterface(nint wrapper, in Guid iid, out nint ppv)
    {
        if (wrapper == 0)
            throw new ArgumentException(null, nameof(wrapper));

        var cls = (Aggregated*)wrapper;

        HRESULT hr = Marshal.QueryInterface(cls->outer, iid, out ppv);
        TracingUtilities.Trace($"outer: 0x{cls->outer:X} hr: 0x{hr:X} iid:{iid.GetName()} ifaceUnk: 0x{ppv:X}");
        // remember non served iids so we don't go infinite loop
        return hr;
    }

    private struct WrapperGuid
    {
        public Guid Iid;
        public nint Wrapper;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static HRESULT IUnknownQueryInterface(nint thisPtr, Guid* riid, nint* ppv)
    {
        TracingUtilities.Trace($"this: 0x{thisPtr:X} riid: {(riid != null ? riid->GetName() : null)}");
        if (riid == null || ppv == null)
            return Constants.E_POINTER;
        *ppv = 0;

        var iid = *riid;
        var cls = (Aggregated*)thisPtr;
        if (iid == typeof(IUnknown).GUID)
        {
            *ppv = thisPtr;
            Interlocked.Increment(ref cls->refCount);
            return Constants.S_OK;
        }

        // if we know for sure the requested iid is one of the aggregable interfaces, we forward to inner
        HRESULT hr;
        for (var i = 0; i < cls->aggregableInterfaceIidsCount; i++)
        {
            if (iid == cls->aggregableInterfaceIids[i])
            {
                // forward QueryInterface to inner
                hr = Marshal.QueryInterface(cls->inner, iid, out var ifaceUnk);
                TracingUtilities.Trace($"inner: 0x{cls->inner:X} hr: 0x{hr:X} iid:{iid.GetName()} ifaceUnk: 0x{ifaceUnk:X}");
                if (hr.IsSuccess && ifaceUnk != 0)
                {
                    *ppv = ifaceUnk;
                    Interlocked.Increment(ref cls->refCount);
                    //add a reference to outer
                    Marshal.AddRef(cls->outer);
                    return Constants.S_OK;
                }
                return Constants.E_NOINTERFACE; // huh?
            }
        }

        // else before we ask outer, we need to protect against infinite loops
        // as the outer may call back to us in its QueryInterface implementation, like TSTCON tool does
        var wg = new WrapperGuid { Iid = iid, Wrapper = thisPtr };

        if (!_tempIidStack.TryAdd(wg, null))
        {
            TracingUtilities.Trace($"outer: 0x{cls->outer:X} iid:{iid.GetName()} already in stack");
            return Constants.E_NOINTERFACE;
        }

        try
        {
            // forward QueryInterface to outer
            hr = OuterQueryInterface(thisPtr, iid, out var outerUnk);
            *ppv = outerUnk; // error or not
        }
        finally
        {
            _tempIidStack.Remove(wg, out _);
        }
        return hr;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint IUnknownAddRef(nint thisPtr)
    {
        var unk = (Aggregated*)thisPtr;
        var ui = (uint)Interlocked.Increment(ref unk->refCount);
        TracingUtilities.Trace($"this: 0x{thisPtr:X} refCount: {-1 + unk->refCount} => {unk->refCount}");
        return ui;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint IUnknownRelease(nint thisPtr)
    {
        var unk = (Aggregated*)thisPtr;
        var c = Interlocked.Decrement(ref unk->refCount);
        TracingUtilities.Trace($"this: 0x{thisPtr:X} refCount: {1 + unk->refCount} => {unk->refCount}");
        if (c == 0)
        {
            Marshal.FreeCoTaskMem((nint)unk->aggregableInterfaceIids);
            Marshal.FreeCoTaskMem(thisPtr);
        }
        return (uint)c;
    }

    private struct Aggregated
    {
        public IUnknownVTable* lpVtbl; // *must* be first: pointer to vtable
        public IUnknownVTable IUnknownVtbl; // embedded IUnknown-only vtable
        public nint inner;
        public nint outer;
        public int refCount;
        public int aggregableInterfaceIidsCount;
        public Guid* aggregableInterfaceIids;
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
