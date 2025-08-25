namespace ActiveN;

[GeneratedComClass]
public abstract partial class BaseControl :
    IOleControl,
    ICustomQueryInterface
{
    CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out nint ppv) => GetInterface(ref iid, out ppv);
    protected virtual CustomQueryInterfaceResult GetInterface(ref Guid iid, out nint ppv)
    {
        ppv = 0;
        ComRegistration.Trace($"GetInterface: {iid:B}");
        return CustomQueryInterfaceResult.NotHandled;
    }

    HRESULT IOleControl.FreezeEvents(BOOL bFreeze)
    {
        throw new NotImplementedException();
    }

    HRESULT IOleControl.GetControlInfo(ref CONTROLINFO pCI)
    {
        throw new NotImplementedException();
    }

    HRESULT IOleControl.OnAmbientPropertyChange(int dispID)
    {
        throw new NotImplementedException();
    }

    HRESULT IOleControl.OnMnemonic(in MSG pMsg)
    {
        throw new NotImplementedException();
    }

    protected static HRESULT RegisterType(ComRegistrationContext context)
    {
        ArgumentNullException.ThrowIfNull(context);
        ComRegistration.Trace($"Register type {typeof(BaseControl).FullName}...");

        // add the "Control" subkey to indicate that this is an ActiveX control
        using var key = ComRegistration.EnsureWritableSubKey(context.RegistryRoot, Path.Combine(ComRegistration.ClsidRegistryKey, context.GUID.ToString("B")));
        key.CreateSubKey("Control", false)?.Dispose();
        return Constants.S_OK;
    }

    protected static HRESULT UnregisterType(ComRegistrationContext context)
    {
        ArgumentNullException.ThrowIfNull(context);
        ComRegistration.Trace($"Unregister type {typeof(BaseControl).FullName}...");
        return Constants.S_OK;
    }
}
