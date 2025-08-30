namespace ActiveN.Samples.PdfView;

// this class *must* be named ComHosting and reside in assembly root namespace or in assembly root namespace + ".Hosting"
// for AotNetComHost.*.dll to find it (this is only for DEBUG builds)
public class ComHosting : ComRegistration
{
    public static ComHosting Instance { get; } = new();

    public ComHosting() : base([
        // TODO: add your COM visible types here
        new ComRegistrationType(typeof(PdfViewControl))
        ])
    {
    }

    public override bool InstallInHkcu => true;
    public override bool CanUnload => true;

    // these are the standard COM DLL exports that *must* be declared

    // create registry entries for all types supported in this module.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllRegisterServer))]
    public static uint DllRegisterServer() => Instance.RegisterServer().UValue;

    // remove entries created through DllRegisterServer.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllUnregisterServer))]
    public static uint DllUnregisterServer() => Instance.UnregisterServer().UValue;

    // determines whether the module is in use. If not, the caller can unload the DLL from memory.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]
    public static uint DllCanUnloadNow() => Instance.CanUnloadNow().UValue;

    // retrieves the class object from a DLL object handler or object application.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllGetClassObject))]
    public static uint DllGetClassObject(nint rclsid, nint riid, nint ppv) => Instance.GetClassObject(rclsid, riid, ppv).UValue;

    // handles installation and setup for a module.
    // this one is optional but very useful to pass any command line arguments during install/uninstall
    [UnmanagedCallersOnly(EntryPoint = nameof(DllInstall))]
    public static uint DllInstall(bool install, nint cmdLinePtr) => Instance.Install(install, cmdLinePtr).UValue;

    // this is a custom export to initialize thunking support, only in debug builds
#if DEBUG
    [UnmanagedCallersOnly(EntryPoint = nameof(DllThunkInit))]
    public static uint DllThunkInit(nint thunkDllPathPtr) => Instance.ThunkInit(thunkDllPathPtr).UValue;
#endif
}
