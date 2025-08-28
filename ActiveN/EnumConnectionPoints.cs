namespace ActiveN;

[GeneratedComClass]
public partial class EnumConnectionPoints(IReadOnlyList<KeyValuePair<Guid, IConnectionPoint>> connectionPoints) : IEnumConnectionPoints
{
    protected int Index { get; set; } = -1;
    public IReadOnlyList<KeyValuePair<Guid, IConnectionPoint>> ConnectionPoints => connectionPoints ?? throw new ArgumentNullException(nameof(connectionPoints));

    public virtual HRESULT Clone(out IEnumConnectionPoints enumerator)
    {
        TracingUtilities.Trace();
        enumerator = new EnumConnectionPoints(ConnectionPoints);
        return Constants.S_OK;
    }

    public virtual HRESULT Next(uint count, nint[] points, out uint fetched)
    {
        TracingUtilities.Trace($"count: {count}");
        ArgumentNullException.ThrowIfNull(points);
        var max = (uint)Math.Max(0, Math.Min(ConnectionPoints.Count - (Index + 1), count));
        fetched = max;
        if (fetched > 0)
        {
            for (var i = Index + 1; i < fetched; i++)
            {
                points[i] = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance<IConnectionPoint>(ConnectionPoints[i].Value);
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

    public virtual HRESULT Skip(uint count)
    {
        TracingUtilities.Trace($"count: {count}");
        var max = (uint)Math.Max(0, Math.Min(ConnectionPoints.Count - (Index + 1), count));
        if (max > 0)
        {
            Index += (int)max;
        }
        return (max == count) ? Constants.S_OK : Constants.S_FALSE;
    }
}
