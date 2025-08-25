namespace ActiveN.Hosting;

[GeneratedComClass]
public partial class ClassFactory(Guid clsid, ComRegistration registration) : IClassFactory
{
    public Guid Clsid { get; } = clsid;
    public ComRegistration ComRegistration { get; } = registration ?? throw new ArgumentNullException(nameof(registration));

    public override string ToString() => $"ClassFactory: {Clsid:B}";

    HRESULT IClassFactory.CreateInstance(nint pUnkOuter, in Guid riid, out nint ppvObject)
    {
        //BaseComHosting.Trace($"pUnkOuter:{pUnkOuter} riid:{riid}");
        var hr = ComRegistration.CreateInstance(this, pUnkOuter, riid, out var instance);
        if (hr.IsError)
        {
            ppvObject = 0;
            return hr;
        }

        var unk = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(instance, riid);
        ppvObject = unk;
        return unk == 0 ? Constants.E_NOINTERFACE : Constants.S_OK;
    }

    HRESULT IClassFactory.LockServer(BOOL fLock)
    {
        //BaseComHosting.Trace($"lock:{fLock}");
        return Constants.S_OK;
    }
}
