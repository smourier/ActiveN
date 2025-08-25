namespace ActiveN.Hosting;

public abstract partial class BaseComRegistration
{
    private readonly Lazy<string> _dllPath;

    [UnconditionalSuppressMessage("Trimming", "IL2026:Members annotated with 'RequiresUnreferencedCodeAttribute' require dynamic access otherwise can break functionality when trimming application code", Justification = "Will mostly work in debug only")]
    protected BaseComRegistration(IReadOnlyList<ComRegistrationType> comTypes)
    {
        ArgumentNullException.ThrowIfNull(comTypes);
        ComTypes = comTypes;

        // add all known assemblies that might have GUIDs to trace
        var assemblies = new HashSet<Assembly>
        {
            typeof(IUnknown).Assembly, // DirectN
            Assembly.GetExecutingAssembly(),
        };

        var entry = Assembly.GetEntryAssembly();
        if (entry != null)
        {
            assemblies.Add(entry);
        }

        foreach (var type in comTypes)
        {
            assemblies.Add(type.Type.Assembly);
        }

        foreach (var asm in assemblies)
        {
            GuidNames.AddClassesGuids(asm);
        }

        _dllPath = new Lazy<string>(GetDllPath);
    }

    public IReadOnlyList<ComRegistrationType> ComTypes { get; }
    public string DllPath => _dllPath.Value;
    public virtual bool InstallInHkcu { get; protected set; } = false;
    public virtual string? ThreadingModel { get; protected set; } // default is Both
    public string RegistrationDllPath => ThunkDllPath ?? DllPath;
    public string? ThunkDllPath { get; protected set; }

#if DIRECTN_DEBUG
    public bool EnableComObjectTraces
    {
        get => DirectN.Extensions.Com.ComObject.EnableTraces;
        set => DirectN.Extensions.Com.ComObject.EnableTraces = value;
    }
#endif

    protected virtual ComRegistrationContext CreateRegistrationContext(RegistryKey root) => new(this, root);

    protected static uint WrapErrors(Func<HRESULT> action)
    {
        try
        {
            return (uint)action();
        }
        catch (SecurityException se)
        {
            // transform this one as a well-known access denied
            Trace($"Ex:{se}");
            return (uint)Constants.E_ACCESSDENIED;
        }
        catch (Exception ex)
        {
            Trace($"Ex:{ex}");
            return (uint)ex.HResult;
        }
    }

    protected virtual HRESULT RegisterServer()
    {
        Trace($"Path:{DllPath}");
        var root = InstallInHkcu ? Registry.CurrentUser : Registry.LocalMachine;
        foreach (var type in ComTypes)
        {
            RegisterType(type, root);
        }
        return Constants.S_OK;
    }

    protected virtual void RegisterType(ComRegistrationType type, RegistryKey registryRoot)
    {
        ArgumentNullException.ThrowIfNull(type);
        ArgumentNullException.ThrowIfNull(registryRoot);
        Trace($"Type:{type.Type.FullName} guid:{type.Type.GUID}");
        RegisterInProcessComObject(registryRoot, type.Type, RegistrationDllPath, ThreadingModel);

        var method = type.Type.GetMethod(nameof(RegisterType), BindingFlags.Public | BindingFlags.Static);
        method?.Invoke(null, [CreateRegistrationContext(registryRoot) ?? throw new InvalidOperationException()]);
    }

    protected virtual HRESULT UnregisterServer()
    {
        Trace($"Path:{DllPath}");
        var root = InstallInHkcu ? Registry.CurrentUser : Registry.LocalMachine;
        foreach (var type in ComTypes)
        {
            UnregisterType(type, root);
        }
        return Constants.S_OK;
    }

    protected virtual void UnregisterType(ComRegistrationType type, RegistryKey registryRoot)
    {
        ArgumentNullException.ThrowIfNull(type);
        ArgumentNullException.ThrowIfNull(registryRoot);
        UnregisterComObject(registryRoot, type.Type);

        var method = type.Type.GetMethod(nameof(UnregisterType), BindingFlags.Public | BindingFlags.Static);
        method?.Invoke(null, [CreateRegistrationContext(registryRoot) ?? throw new InvalidOperationException()]);
    }

    protected virtual void GetClassObject(Guid clsid, Guid iid, out object? ppv)
    {
        Trace($"Path:{DllPath} CLSID:{clsid} IID:{iid}");
        foreach (var type in ComTypes)
        {
            Trace($"Type:{type.Type.FullName} guid:{type.Type.GUID}");
            if (clsid == type.Type.GUID && iid == typeof(IClassFactory).GUID)
            {
                ppv = new BaseClassFactory();
                Trace($"ppv:{ppv}");
                return;
            }
        }
        ppv = null;
    }

    protected unsafe HRESULT GetClassObject(nint rclsid, nint riid, nint ppv)
    {
        var clsid = *(Guid*)rclsid;
        var iid = *(Guid*)riid;
        GetClassObject(clsid, iid, out var obj);
        var unk = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(obj, iid);
        *(nint*)ppv = unk;
        Trace($"unk:{obj}");
        return unk == 0 ? Constants.E_NOINTERFACE : Constants.S_OK;
    }

    protected virtual HRESULT CanUnloadNow()
    {
        Trace($"Path:{DllPath}");
        return (uint)Constants.S_FALSE;
    }

    // call for example
    // "regsvr32 /i:user /n" to register in HKCU without calling DllRegisterServer
    // "regsvr32 /i:user /n /u" to unregister from HKCU
    protected virtual HRESULT Install(bool install, nint cmdLinePtr)
    {
        Trace($"Path:{DllPath}");
        if (cmdLinePtr != 0)
        {
            var cmdLine = Marshal.PtrToStringUni(cmdLinePtr);
            Trace($"CmdLine:{cmdLine}");
            if (string.Equals(cmdLine, "user", StringComparison.OrdinalIgnoreCase))
            {
                InstallInHkcu = true;
            }
        }

        return install ? RegisterServer() : UnregisterServer();
    }

    protected uint ThunkInit(nint thunkDllPathPtr)
    {
        ThunkDllPath = Marshal.PtrToStringUni(thunkDllPathPtr);
        Trace($"Path:{DllPath} ThunkDllPathPtr:{ThunkDllPath}");
        return 0;
    }

    private const string _classesRegistryKey = @"Software\Classes";
    private const string _clsidRegistryKey = _classesRegistryKey + @"\CLSID";

    public static void RegisterInProcessComObject(RegistryKey root, Type type, string assemblyPath, string? threadingModel = null)
    {
        ArgumentNullException.ThrowIfNull(root);
        ArgumentNullException.ThrowIfNull(type);
        ArgumentNullException.ThrowIfNull(assemblyPath);

        threadingModel = threadingModel?.Trim() ?? "Both";
        Trace($"Registering {type.FullName} from {assemblyPath} with threading model '{threadingModel}'...");
        using var serverKey = EnsureWritableSubKey(root, Path.Combine(_clsidRegistryKey, type.GUID.ToString("B"), "InprocServer32"));
        serverKey.SetValue(null, assemblyPath);
        serverKey.SetValue("ThreadingModel", threadingModel);

        // ProgId is optional
        var att = type.GetCustomAttribute<ProgIdAttribute>();
        if (att != null && !string.IsNullOrWhiteSpace(att.Value))
        {
            var progid = att.Value.Trim();
            using var key = EnsureWritableSubKey(root, Path.Combine(_clsidRegistryKey, type.GUID.ToString("B")));
            using var progIdKey = EnsureWritableSubKey(key, "ProgId");
            progIdKey.SetValue(null, progid);

            using var ckey = EnsureWritableSubKey(root, Path.Combine(_classesRegistryKey, progid, "CLSID"));
            ckey.SetValue(null, type.GUID.ToString("B"));
        }
        Trace($"Registered {type.FullName}.");
    }

    public static void UnregisterComObject(RegistryKey root, Type type)
    {
        ArgumentNullException.ThrowIfNull(root);
        ArgumentNullException.ThrowIfNull(type);

        Trace($"Unregistering {type.FullName}...");
        using var key = root.OpenSubKey(_clsidRegistryKey, true);
        key?.DeleteSubKeyTree(type.GUID.ToString("B"), false);

        // ProgId is optional
        var att = type.GetCustomAttribute<ProgIdAttribute>();
        if (att != null && !string.IsNullOrWhiteSpace(att.Value))
        {
            var progid = att.Value.Trim();
            using var ckey = root.OpenSubKey(_classesRegistryKey, true);
            ckey?.DeleteSubKeyTree(progid, false);
        }

        Trace($"Unregistered {type.FullName}.");
    }

    protected static RegistryKey EnsureWritableSubKey(RegistryKey root, string name)
    {
        var key = root.OpenSubKey(name, true);
        if (key != null)
            return key;

        var parentName = Path.GetDirectoryName(name);
        if (string.IsNullOrEmpty(parentName))
            return root.CreateSubKey(name);

        using var parentKey = EnsureWritableSubKey(root, parentName);
        return parentKey.CreateSubKey(Path.GetFileName(name));
    }

    public static void Trace(string? text = null, [CallerMemberName] string? methodName = null, [CallerFilePath] string? filePath = null)
    {
        var name = filePath != null ? Path.GetFileNameWithoutExtension(filePath) : null;
        EventProvider.Default.WriteMessageEvent($"[{Environment.CurrentManagedThreadId}]{name}::{methodName}:{text}");
    }

#if DEBUG
    // this only works if *not* published as AOT

    [UnconditionalSuppressMessage("SingleFile", "IL3000:Avoid accessing Assembly file path when publishing as a single file", Justification = "Only used when not published as AOT")]
    protected virtual string GetDllPath() => GetType().Assembly.Location;

#else
    // this only works if published as AOT

    private static unsafe delegate* unmanaged<uint> GetModuleFunctionPointer() => &DoNothing; // any func pointer will do

    // this won't really be exported, just used to fool AOT compiler to get a valid module function pointer
    [UnmanagedCallersOnly(EntryPoint = nameof(DoNothing))]
    public static uint DoNothing() => 0;

    protected virtual unsafe string GetDllPath()
    {
        var ptr = GetModuleFunctionPointer();
        const int GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS = 4;
        const int GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT = 2;
        if (!Functions.GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, new PWSTR((nint)ptr), out var module))
            throw new Win32Exception(Marshal.GetLastPInvokeError());

        var size = 256;
        do
        {
            using var pwstr = new AllocPwstr(size);
            Functions.GetModuleFileNameW(module, pwstr, pwstr.SizeInChars); // just to ensure the DLL is loaded
            if (Marshal.GetLastWin32Error() != (int)WIN32_ERROR.ERROR_INSUFFICIENT_BUFFER)
                return pwstr.ToString()!;

            size *= 2;
        }
        while (true);
    }

#endif
}
