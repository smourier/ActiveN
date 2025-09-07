namespace ActiveN;

[GeneratedComClass]
public partial class EnumConnections(KeyValuePair<uint, IComObject>[] connections) : IEnumConnections
{
    protected int Index { get; set; } = -1;
    public KeyValuePair<uint, IComObject>[] Connections => connections ?? throw new ArgumentNullException(nameof(connections));

    public HRESULT Clone(out IEnumConnections enumerator)
    {
        TracingUtilities.Trace();
        enumerator = new EnumConnections(Connections);
        return Constants.S_OK;
    }

    public HRESULT Next(uint count, CONNECTDATA[] rgcd, out uint fetched)
    {
        TracingUtilities.Trace($"count: {count}");
        var max = (uint)Math.Max(0, Math.Min(Connections.Length - (Index + 1), count));
        fetched = max;
        if (fetched > 0)
        {
            for (var i = Index + 1; i < fetched; i++)
            {
                rgcd[i].dwCookie = Connections[i].Key;
                rgcd[i].pUnk = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(Connections[i].Value);
                Marshal.AddRef(rgcd[i].pUnk);
                Index++;
            }
        }
        return (fetched == count) ? Constants.S_OK : Constants.S_FALSE;
    }

    public virtual HRESULT Reset()
    {
        TracingUtilities.Trace();
        Index = -1;
        return Constants.S_OK;
    }

    public HRESULT Skip(uint count)
    {
        TracingUtilities.Trace($"count: {count}");
        var max = (uint)Math.Max(0, Math.Min(Connections.Length - (Index + 1), count));
        if (max > 0)
        {
            Index += (int)max;
        }
        return (max == count) ? Constants.S_OK : Constants.S_FALSE;
    }

    public static IEnumerable<IComObject> EnumerateConnections(IConnectionPoint cp)
    {
        ArgumentNullException.ThrowIfNull(cp);

        if (cp is BaseConnectionPoint bcp)
        {
            foreach (var sink in bcp.Sinks)
            {
                yield return sink.Value;
            }
            yield break;
        }

        cp.EnumConnections(out var enumConnectionsObj);
        if (enumConnectionsObj == null)
            yield break;

        using var ecp = new ComObject<IEnumConnections>(enumConnectionsObj);
        var cd = new CONNECTDATA[1];
        while (ecp.Object.Next(1, cd, out var fetched) == Constants.S_OK && fetched == 1)
        {
            using var sink = DirectN.Extensions.Com.ComObject.FromPointer<IUnknown>(cd[0].pUnk);
            if (sink == null)
                continue;

            yield return sink;
        }
    }
}
