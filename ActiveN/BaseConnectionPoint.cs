namespace ActiveN;

#pragma warning disable SYSLIB1097 // Add 'GeneratedComClassAttribute' to enable passing objects of this type to COM
public abstract partial class BaseConnectionPoint : IConnectionPoint, ICustomQueryInterface, IDisposable
#pragma warning restore SYSLIB1097 // Add 'GeneratedComClassAttribute' to enable passing objects of this type to COM
{
    private ConcurrentDictionary<uint, IComObject> _sinks = new();
    private uint _cookie;
    internal IConnectionPointContainer? _container;

    public IReadOnlyDictionary<uint, IComObject> Sinks => _sinks;
    public abstract Guid InterfaceId { get; }
    public virtual bool IsIDispatch => false;
    protected abstract IComObject GetFromPointer(nint ptr);

    public override string ToString() => InterfaceId.ToString();

    CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out nint ppv) => GetInterface(ref iid, out ppv);
    protected virtual CustomQueryInterfaceResult GetInterface(ref Guid iid, out nint ppv)
    {
        ppv = 0;
        TracingUtilities.Trace($"iid: {iid.GetName()}");
        return CustomQueryInterfaceResult.NotHandled;
    }

    HRESULT IConnectionPoint.Advise(nint sink, out uint pcookie)
    {
        TracingUtilities.Trace($"sink: {sink:X}, container: {_container}");
        var cookie = 0u;
        var hr = TracingUtilities.WrapErrors(() =>
        {
            if (sink == 0)
            {
                cookie = 0;
                return Constants.E_POINTER;
            }

            var sinkObj = GetFromPointer(sink);
            if (sinkObj == null)
            {
                cookie = 0;
                return Constants.CONNECT_E_CANNOTCONNECT;
            }

            cookie = Interlocked.Increment(ref _cookie);
            _sinks[cookie] = sinkObj;
            TracingUtilities.Trace($"cookie: {cookie}");
            return Constants.S_OK;
        });
        pcookie = cookie;
        return hr;
    }

    HRESULT IConnectionPoint.EnumConnections(out IEnumConnections enumerator)
    {
        TracingUtilities.Trace();
        enumerator = new EnumConnections([.. _sinks]);
        return Constants.S_OK;
    }

    HRESULT IConnectionPoint.GetConnectionInterface(out Guid interfaceId)
    {
        interfaceId = InterfaceId;
        TracingUtilities.Trace($"{interfaceId.GetName()}");
        return Constants.S_OK;
    }

    HRESULT IConnectionPoint.GetConnectionPointContainer(out IConnectionPointContainer container)
    {
        TracingUtilities.Trace($"{_container}");
        container = _container!;
        return container != null ? Constants.S_OK : Constants.E_FAIL;
    }

    HRESULT IConnectionPoint.Unadvise(uint cookie)
    {
        TracingUtilities.Trace($"cookie: {cookie}");
        if (!_sinks.TryRemove(cookie, out var sink))
            return Constants.E_UNEXPECTED;

        sink?.Dispose();
        return Constants.S_OK;
    }

    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            var sinks = Interlocked.Exchange(ref _sinks, new());
            foreach (var kv in sinks)
            {
                try
                {
                    TracingUtilities.Trace($"left sink: {kv.Value}");
                    kv.Value.Dispose();
                }
                catch
                {
                    // continue
                }
            }
            _sinks.Clear();
            // dispose managed state (managed objects)
        }

        // free unmanaged resources (unmanaged objects) and override finalizer
        // set large fields to null
    }
}
