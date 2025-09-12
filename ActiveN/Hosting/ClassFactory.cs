// Copyright (c) Aelyo Softworks S.A.S.. All rights reserved.
// See LICENSE in the project root for license information.

namespace ActiveN.Hosting;

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

            var hr = ComRegistration.CreateInstance(this, out var instance);
            if (hr.IsError)
                return hr;

            if (pUnkOuter == 0)
            {
                // no aggregation, just return the requested interface
                ppv = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(instance, iid);
            }
            else
            {
                if (instance is not IAggregable aggregable || !aggregable.SupportsAggregation)
                    return Constants.CLASS_E_NOAGGREGATION;

                ppv = Aggregable.Aggregate(pUnkOuter, aggregable);
                aggregable.Wrapper = ppv;
                TracingUtilities.Trace($"aggregated: 0x{ppv:X}");
                foreach (var type in aggregable.AggregableInterfaces)
                {
                    TracingUtilities.Trace($"  - {type.FullName} ({type.GUID.GetName()})");
                }
            }
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

    public static IReadOnlyList<Type> GetComAggregatedInterfaces([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.Interfaces)] Type type)
    {
        ArgumentNullException.ThrowIfNull(type);
        var list = new List<Type>();
        foreach (var cp in type.GetInterfaces())
        {
            if (cp.IsDefined(typeof(GeneratedComInterfaceAttribute), true))
            {
                list.Add(cp);
            }
        }
        return list;
    }
}
