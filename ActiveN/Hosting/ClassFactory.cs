namespace ActiveN.Hosting;

[GeneratedComClass]
public partial class ClassFactory : IClassFactory
{
    HRESULT IClassFactory.CreateInstance(nint pUnkOuter, in Guid riid, out nint ppvObject)
    {
        ComHosting.Trace($"pUnkOuter:{pUnkOuter} riid:{riid}");
        if (pUnkOuter != 0)
        {
            ppvObject = 0;
            return Constants.CLASS_E_NOAGGREGATION;
        }

        var instance = CreateInstance(riid);
        var unk = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(instance, riid);
        ppvObject = unk;
        return unk == 0 ? Constants.E_NOINTERFACE : Constants.S_OK;
    }

    HRESULT IClassFactory.LockServer(BOOL fLock)
    {
        ComHosting.Trace($"lock:{fLock}");
        return Constants.S_OK;
    }

    // the list of types that are COM types
    public static Type[] ComTypes { get; } =
    [
    ];

    // cannot build COM abstract class
    protected virtual object CreateInstance(in Guid riid) => throw new NotSupportedException("Must be implemented by derived class.");
}
