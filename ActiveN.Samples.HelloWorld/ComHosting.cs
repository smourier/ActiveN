namespace ActiveN.Samples.HelloWorld;

// this class *must* be named ComHosting and reside in assembly root namespace or in assembly root namespace + ".Hosting"
public class ComHosting : ComRegistration
{
    public static ComHosting Instance { get; } = new();

    public ComHosting() : base([
        // TODO: add your COM visible types here
        new ComRegistrationType(typeof(HelloWorldControl))
        ])
    {
    }

    // these are the standard COM DLL exports that *must* be declared

    // create registry entries for all types supported in this module.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllRegisterServer))]
    public static uint DllRegisterServer() => WrapErrors(Instance.RegisterServer);

    // remove entries created through DllRegisterServer.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllUnregisterServer))]
    public static uint DllUnregisterServer() => WrapErrors(Instance.UnregisterServer);

    // determines whether the module is in use. If not, the caller can unload the DLL from memory.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]
    public static uint DllCanUnloadNow() => WrapErrors(Instance.CanUnloadNow);

    // retrieves the class object from a DLL object handler or object application.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllGetClassObject))]
    public static uint DllGetClassObject(nint rclsid, nint riid, nint ppv) => WrapErrors(() => Instance.GetClassObject(rclsid, riid, ppv));

    // handles installation and setup for a module.
    // this one is optional but very useful to pass any command line arguments during install/uninstall
    [UnmanagedCallersOnly(EntryPoint = nameof(DllInstall))]
    public static uint DllInstall(bool install, nint cmdLinePtr) => WrapErrors(() => Instance.Install(install, cmdLinePtr));

    // this is a custom export to initialize thunking support, only in debug builds
#if DEBUG
    [UnmanagedCallersOnly(EntryPoint = nameof(DllThunkInit))]
    public static uint DllThunkInit(nint thunkDllPathPtr) => WrapErrors(() => Instance.ThunkInit(thunkDllPathPtr));
#endif
}
