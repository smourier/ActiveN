namespace ActiveN;

[GeneratedComClass]
public partial class EnumVerbs(IReadOnlyList<OLEVERB> verbs) : IEnumOLEVERB
{
    protected int Index { get; set; } = -1;
    public IReadOnlyList<OLEVERB> Verbs => verbs ?? throw new ArgumentNullException(nameof(verbs));

    public virtual HRESULT Clone(out IEnumOLEVERB enumerator)
    {
        TracingUtilities.Trace();
        enumerator = new EnumVerbs(Verbs);
        return Constants.S_OK;
    }

    public virtual HRESULT Next(uint count, OLEVERB[] outVerbs, nint outFetched)
    {
        TracingUtilities.Trace($"count: {count}");
        ArgumentNullException.ThrowIfNull(outVerbs);
        var max = Math.Max(0, Math.Min(verbs.Count - (Index + 1), count));
        var fetched = max;
        if (fetched > 0)
        {
            for (var i = Index + 1; i < fetched; i++)
            {
                outVerbs[i] = Verbs[i];
                Index++;
            }
        }

        if (outFetched != 0)
        {
            Marshal.WriteInt32(outFetched, (int)fetched);
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
        var max = (uint)Math.Max(0, Math.Min(Verbs.Count - (Index + 1), count));
        if (max > 0)
        {
            Index += (int)max;
        }
        return (max == count) ? Constants.S_OK : Constants.S_FALSE;
    }
}
