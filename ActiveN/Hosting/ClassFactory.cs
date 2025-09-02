﻿namespace ActiveN.Hosting;

[GeneratedComClass]
public partial class ClassFactory(Guid clsid, ComRegistration registration) : IClassFactory, ICustomQueryInterface
{
    public Guid Clsid { get; } = clsid;
    public ComRegistration ComRegistration { get; } = registration ?? throw new ArgumentNullException(nameof(registration));

    public override string ToString() => $"{Clsid:B}";

    CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out nint ppv) => GetInterface(ref iid, out ppv);
    protected virtual CustomQueryInterfaceResult GetInterface(ref Guid iid, out nint ppv)
    {
        ppv = 0;
        TracingUtilities.Trace($"iid: {iid.GetName()}");
        return CustomQueryInterfaceResult.NotHandled;
    }

    HRESULT IClassFactory.CreateInstance(nint pUnkOuter, in Guid riid, out nint ppvObject)
    {
        var ppv = nint.Zero;
        var iid = riid;
        TracingUtilities.Trace($"pUnkOuter: {pUnkOuter} clsid: {Clsid:B} riid: {iid.GetName()}");
        var hr = TracingUtilities.WrapErrors(() =>
        {
            if (pUnkOuter != 0 && iid != typeof(IUnknown).GUID)
                return Constants.CLASS_E_NOAGGREGATION;

            var hr = ComRegistration.CreateInstance(this, iid, out var instance);
            if (hr.IsError)
                return hr;

            if (pUnkOuter != 0)
            {
                if (instance is not IAggregable aggregatable || !aggregatable.SupportsAggregation)
                    return Constants.CLASS_E_NOAGGREGATION;

                aggregatable.OuterUnknown = pUnkOuter;
            }

            ppv = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(instance, iid);
            ppv = Aggregable.Aggregate(pUnkOuter, ppv);
            return ppv == 0 ? Constants.E_NOINTERFACE : Constants.S_OK;
        });
        ppvObject = ppv;
        return hr;
    }

    HRESULT IClassFactory.LockServer(BOOL fLock)
    {
        TracingUtilities.Trace($"lock: {fLock}");
        return Constants.S_OK;
    }
}
