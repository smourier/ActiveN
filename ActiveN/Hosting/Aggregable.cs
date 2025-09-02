namespace ActiveN.Hosting;

// this more or less implements https://learn.microsoft.com/en-us/windows/win32/com/aggregation
public unsafe class Aggregable
{
    public static nint Aggregate(nint outer, nint inner)
    {
        if (outer == 0)
            return inner;

        // Assume inner points to an object whose first pointer is its IUnknown vtable
        var headerSize = sizeof(AggregatedUnknown);
        var mem = (AggregatedUnknown*)Marshal.AllocCoTaskMem(headerSize);
        Unsafe.InitBlockUnaligned(mem, 0, (uint)headerSize);

        var originalVtbl = *(IUnknownVTable**)inner;

        // Copy original vtable so we can patch only IUnknown slots (or keep pointer and build a tiny custom one)
        mem->patchedVtbl = *originalVtbl;
        mem->patchedVtbl.QueryInterface = &PatchedQueryInterface;
        mem->patchedVtbl.AddRef = &PatchedAddRef;
        mem->patchedVtbl.Release = &PatchedRelease;

        mem->lpVtbl = &mem->patchedVtbl;
        mem->originalVtbl = originalVtbl;
        mem->inner = inner;
        mem->outer = outer;
        mem->refCount = 1;

        return (nint)mem;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static HRESULT PatchedQueryInterface(nint thisPtr, Guid* riid, nint* ppv)
    {
        TracingUtilities.Trace($"this: {thisPtr:X}, riid: {riid->GetName()}");
        if (ppv == null)
            return Constants.E_POINTER;
        *ppv = 0;

        var hdr = (AggregatedUnknown*)thisPtr;
        if (*riid == typeof(IUnknown).GUID)
        {
            // return controlling unknown
            *ppv = hdr->outer != 0 ? hdr->outer : thisPtr;
            Interlocked.Increment(ref hdr->refCount);
            return Constants.S_OK;
        }

        // forward to inner original QI
        return hdr->originalVtbl->QueryInterface(hdr->inner, riid, ppv);
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint PatchedAddRef(nint thisPtr)
    {
        var hdr = (AggregatedUnknown*)thisPtr;
        TracingUtilities.Trace($"this: {thisPtr:X} refCount: {hdr->refCount}");
        return (uint)Interlocked.Increment(ref hdr->refCount);
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint PatchedRelease(nint thisPtr)
    {
        var hdr = (AggregatedUnknown*)thisPtr;
        TracingUtilities.Trace($"this: {thisPtr:X} refCount: {hdr->refCount}");
        var c = Interlocked.Decrement(ref hdr->refCount);
        if (c == 0)
        {
            hdr->originalVtbl->Release(hdr->inner);
            Marshal.FreeCoTaskMem((nint)hdr);
        }
        return (uint)c;
    }

    private unsafe struct AggregatedUnknown
    {
        public IUnknownVTable* lpVtbl;          // First field: pointer to active vtable
        public IUnknownVTable patchedVtbl;      // Embedded patched vtable
        public IUnknownVTable* originalVtbl;    // Original inner IUnknown vtable
        public nint inner;                      // Inner object (controlling unknown of inner)
        public nint outer;                      // Outer (controlling) if any
        public int refCount;
    }

    private struct IUnknownVTable
    {
        private void* _queryInterface;
        private void* _addRef;
        private void* _release;
        public int _ref;

        public delegate* unmanaged[Stdcall]<nint, Guid*, nint*, HRESULT> QueryInterface { readonly get => (delegate* unmanaged[Stdcall]<nint, Guid*, nint*, HRESULT>)_queryInterface; set => _queryInterface = value; }
        public delegate* unmanaged[Stdcall]<nint, uint> AddRef { readonly get => (delegate* unmanaged[Stdcall]<nint, uint>)_addRef; set => _addRef = value; }
        public delegate* unmanaged[Stdcall]<nint, uint> Release { readonly get => (delegate* unmanaged[Stdcall]<nint, uint>)_release; set => _release = value; }
    }
}
