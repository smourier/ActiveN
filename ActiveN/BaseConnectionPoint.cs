namespace ActiveN;

#pragma warning disable SYSLIB1097 // Add 'GeneratedComClassAttribute' to enable passing objects of this type to COM
public abstract partial class BaseConnectionPoint : IConnectionPoint, IDisposable
#pragma warning restore SYSLIB1097 // Add 'GeneratedComClassAttribute' to enable passing objects of this type to COM
{
    private readonly ConcurrentDictionary<uint, IComObject> _sinks = new();
    private uint _cookie;
    internal IConnectionPointContainer? _container;

    public IReadOnlyDictionary<uint, IComObject> Sinks => _sinks;
    public abstract Guid InterfaceId { get; }
    protected abstract IComObject GetFromPointer(nint ptr);

    public override string ToString() => InterfaceId.ToString();

    HRESULT IConnectionPoint.Advise(nint sink, out uint cookie)
    {
        if (sink == 0)
        {
            cookie = 0;
            return Constants.E_POINTER;
        }

        var disp = GetFromPointer(sink);
        if (disp == null)
        {
            cookie = 0;
            return Constants.CONNECT_E_CANNOTCONNECT;
        }

        cookie = Interlocked.Increment(ref _cookie);
        _sinks[cookie] = disp;
        return Constants.S_OK;
    }

    HRESULT IConnectionPoint.EnumConnections(out IEnumConnections enumerator)
    {
        enumerator = new EnumConnections([.. _sinks]);
        return Constants.S_OK;
    }

    HRESULT IConnectionPoint.GetConnectionInterface(out Guid interfaceId)
    {
        interfaceId = InterfaceId;
        return Constants.S_OK;
    }

    HRESULT IConnectionPoint.GetConnectionPointContainer(out IConnectionPointContainer container)
    {
        container = _container!;
        return container != null ? Constants.S_OK : Constants.E_FAIL;
    }

    HRESULT IConnectionPoint.Unadvise(uint cookie)
    {
        if (!_sinks.TryRemove(cookie, out var disp))
            return Constants.E_UNEXPECTED;

        disp?.Dispose();
        return Constants.S_OK;
    }

    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            foreach (var kv in _sinks)
            {
                try
                {
                    kv.Value.Dispose();
                }
                catch
                {
                    // continue
                }
            }
            // dispose managed state (managed objects)
        }

        // free unmanaged resources (unmanaged objects) and override finalizer
        // set large fields to null
    }
}
